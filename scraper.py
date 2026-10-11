"""
Google Maps Scraper Engine using Playwright.
Handles:
- Headed Chromium browser execution (single instance)
- Cookie consent dismissal
- Grid-based search navigation (https://www.google.com/maps/search/<keyword>/@<lat>,<lng>,15z)
- Infinite scroll of the results feed until end-of-list
- Place detail extraction: Name, Phone, Category, Address, CID, Short URL, Rating, Website
- Anti-bot random delays (1-3s)
- CAPTCHA detection and interactive user pause/resume
- Deduplication by CID and per-listing SQLite/JSON checkpointing
"""
import random
import re
import time
from typing import Dict, List, Optional, Set, Any
from urllib.parse import quote_plus

from playwright.sync_api import sync_playwright, Browser, BrowserContext, Page, TimeoutError as PlaywrightTimeout

from cid_utils import extract_cid_from_url, get_short_cid_url
from config import (
    HEADLESS,
    DEFAULT_USER_AGENT,
    VIEWPORT_WIDTH,
    VIEWPORT_HEIGHT,
    MIN_DELAY_SEC,
    MAX_DELAY_SEC,
    PAGE_LOAD_TIMEOUT_MS,
    SCROLL_PAUSE_SEC,
    MAX_CONSECUTIVE_EMPTY_SCROLLS,
    DEFAULT_ZOOM,
    SELECTORS,
)
from db import DatabaseManager
from progress_logger import console


PHONE_REGEX = re.compile(
    r"(?:\+?\d{1,3}[\s.-]?)?(?:\(?\d{2,4}\)?[\s.-]?)?\d{3,4}[\s.-]?\d{3,4}"
)


class CaptchaDetectedException(Exception):
    """Raised when Google displays a CAPTCHA or unusual traffic prompt."""
    pass


class GoogleMapsScraper:
    """Orchestrates Playwright Chromium to extract business leads from Google Maps."""

    def __init__(
        self,
        db_manager: DatabaseManager,
        filter_city: bool = True,
        delay_min: float = MIN_DELAY_SEC,
        delay_max: float = MAX_DELAY_SEC,
    ):
        self.db = db_manager
        self.filter_city = filter_city
        self.delay_min = delay_min
        self.delay_max = delay_max

        self.playwright = None
        self.browser: Optional[Browser] = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None
        self.detail_page: Optional[Page] = None

    def start_browser(self) -> None:
        """Launch a single headed Chromium browser instance with search and detail tabs."""
        if self.browser and self.page:
            return

        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(
            headless=HEADLESS,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--start-maximized",
                "--no-sandbox",
                "--disable-infobars",
                "--disable-dev-shm-usage",
                "--disable-gpu",
                "--js-flags=--max-old-space-size=256",
            ],
        )
        self.context = self.browser.new_context(
            user_agent=DEFAULT_USER_AGENT,
            viewport={"width": VIEWPORT_WIDTH, "height": VIEWPORT_HEIGHT},
            locale="en-US",
            timezone_id="America/New_York",
            permissions=["geolocation"],
        )
        self.page = self.context.new_page()
        self.page.set_default_timeout(PAGE_LOAD_TIMEOUT_MS)

        # Dedicated tab for fast place details (blocks heavy media)
        self.detail_page = None
        self._setup_detail_page()
        self._extract_count = 0

    def _setup_detail_page(self) -> None:
        """Create or recycle detail page with media blocking to flush RAM cache."""
        try:
            if self.detail_page:
                self.detail_page.close()
        except Exception:
            pass
        self.detail_page = self.context.new_page()
        self.detail_page.set_default_timeout(PAGE_LOAD_TIMEOUT_MS)
        try:
            self.detail_page.route(
                re.compile(r"\.(png|jpg|jpeg|gif|webp|svg|woff|woff2|ttf)$", re.IGNORECASE),
                lambda route: route.abort(),
            )
        except Exception:
            pass

    def close_browser(self) -> None:
        """Gracefully close browser resources."""
        try:
            if self.detail_page:
                self.detail_page.close()
            if self.page:
                self.page.close()
            if self.context:
                self.context.close()
            if self.browser:
                self.browser.close()
            if self.playwright:
                self.playwright.stop()
        except Exception:
            pass
        finally:
            self.detail_page = None
            self.page = None
            self.context = None
            self.browser = None
            self.playwright = None

    def random_delay(self, extra_min: Optional[float] = None, extra_max: Optional[float] = None) -> None:
        """Sleep for a random duration between 1 and 3 seconds (or custom bounds)."""
        low = extra_min if extra_min is not None else self.delay_min
        high = extra_max if extra_max is not None else self.delay_max
        delay = round(random.uniform(low, high), 2)
        time.sleep(delay)

    def check_for_captcha(self, target_page: Optional[Page] = None) -> bool:
        """Check if Google has presented a CAPTCHA or 'unusual traffic' block."""
        p = target_page or self.page
        if not p:
            return False

        current_url = p.url.lower()
        if "/sorry/" in current_url or "recaptcha" in current_url:
            return True

        for selector in SELECTORS["captcha_indicators"]:
            try:
                if p.locator(selector).first.is_visible(timeout=500):
                    return True
            except Exception:
                continue

        return False

    def handle_captcha_if_present(self, target_page: Optional[Page] = None) -> None:
        """If CAPTCHA is detected, pause execution and prompt user to solve it in the headed window."""
        p = target_page or self.page
        if self.check_for_captcha(p):
            # Trigger terminal bell / sound
            print("\a", end="", flush=True)
            console.print("\n" + "=" * 70, style="bold red")
            console.print(
                "[bold red]🚨 CAPTCHA / Unusual Traffic Detected in Chromium![/bold red]\n"
                "[yellow]Please switch to the Chromium window, solve the CAPTCHA manually.[/yellow]\n"
                "[bold cyan]When solved and the page loads, return here and press [ENTER] to resume...[/bold cyan]"
            )
            console.print("=" * 70 + "\n", style="bold red")

            input()

            # Wait briefly and verify if cleared
            self.random_delay(2.0, 3.0)
            if self.check_for_captcha(p):
                console.print("[yellow]⚠️ CAPTCHA still appears to be active. Retrying check...[/yellow]")
                self.handle_captcha_if_present(p)
            else:
                console.print("[bold green]✅ CAPTCHA solved! Resuming extraction workflow...[/bold green]\n")

    def dismiss_consent_dialog(self, target_page: Optional[Page] = None) -> None:
        """Dismiss Google cookies consent modal if presented."""
        p = target_page or self.page
        if not p:
            return

        for selector in SELECTORS["consent_buttons"]:
            try:
                btn = p.locator(selector).first
                if btn.is_visible(timeout=800):
                    btn.click()
                    self.random_delay(1.0, 2.0)
                    break
            except Exception:
                continue

    def search_grid_point(
        self,
        keyword: str,
        lat: float,
        lng: float,
        zoom: int = DEFAULT_ZOOM,
        city: str = "",
    ) -> List[str]:
        """
        Navigate to Google Maps search at specific coordinates:
        https://www.google.com/maps/search/<keyword>+in+<city>/@<lat>,<lng>,15z
        
        Scrolls the results panel until 'You've reached the end of the list' or feed end.
        
        Returns:
            List of discovered Google Maps place URLs.
        """
        if not self.page:
            self.start_browser()

        search_term = f"{keyword.strip()} in {city.strip()}" if city else keyword.strip()
        encoded_keyword = quote_plus(search_term)
        search_url = f"https://www.google.com/maps/search/{encoded_keyword}/@{lat},{lng},{zoom}z"

        try:
            self.page.goto(search_url, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT_MS)
        except PlaywrightTimeout:
            console.print(f"[yellow]Navigation timeout for {search_url}, attempting scroll anyway...[/yellow]")

        self.handle_captcha_if_present()
        self.dismiss_consent_dialog()
        self.random_delay(1.5, 2.5)

        # Check if the search immediately redirected to a single place page
        if "/maps/place/" in self.page.url:
            return [self.page.url]

        # Locate the scrollable results feed container
        feed_locator = None
        for selector in SELECTORS["feed_container"]:
            try:
                locator = self.page.locator(selector).first
                if locator.is_visible(timeout=1500):
                    feed_locator = locator
                    break
            except Exception:
                continue

        discovered_links: Set[str] = set()
        consecutive_empty_scrolls = 0
        scroll_attempts = 0
        max_scroll_attempts = 60

        while scroll_attempts < max_scroll_attempts:
            scroll_attempts += 1
            self.handle_captcha_if_present()

            # Extract all listing URLs currently in DOM
            try:
                links = self.page.locator(SELECTORS["listing_link"]).evaluate_all(
                    "elements => elements.map(el => el.href).filter(href => href.includes('/maps/place/'))"
                )
                initial_count = len(discovered_links)
                for link in links:
                    if link:
                        discovered_links.add(link)

                new_found = len(discovered_links) - initial_count
                if new_found == 0:
                    consecutive_empty_scrolls += 1
                else:
                    consecutive_empty_scrolls = 0
            except Exception:
                consecutive_empty_scrolls += 1

            # Check for end of list message
            end_reached = False
            for end_sel in SELECTORS["end_of_list_indicators"]:
                try:
                    if self.page.locator(end_sel).first.is_visible(timeout=300):
                        end_reached = True
                        break
                except Exception:
                    continue

            if end_reached or consecutive_empty_scrolls >= MAX_CONSECUTIVE_EMPTY_SCROLLS:
                break

            # Fast cutoff on empty outskirts: if 2 scroll attempts found 0 listings, skip point
            if scroll_attempts >= 2 and len(discovered_links) == 0:
                break

            # Perform scroll on feed container or page
            try:
                if feed_locator and feed_locator.is_visible():
                    feed_locator.evaluate(
                        "el => { el.scrollTop = el.scrollHeight; }"
                    )
                else:
                    self.page.mouse.wheel(0, 1500)
            except Exception:
                self.page.keyboard.press("PageDown")

            # Pause between scrolls
            time.sleep(SCROLL_PAUSE_SEC)

        return list(discovered_links)

    def extract_place_details(
        self,
        place_url: str,
        target_city: Optional[str] = None,
        keyword: str = "",
    ) -> Optional[Dict[str, Any]]:
        """
        Navigate to a specific place URL, extract business info, convert CID,
        and optionally verify city match. Uses dedicated fast detail_page.
        """
        if not self.page:
            self.start_browser()

        # Auto-recycle detail tab every 25 listings to release accumulated RAM back to Windows
        self._extract_count = getattr(self, "_extract_count", 0) + 1
        if self._extract_count % 25 == 0:
            self._setup_detail_page()

        p = self.detail_page if self.detail_page else self.page

        try:
            p.goto(place_url, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT_MS)
        except PlaywrightTimeout:
            pass

        self.handle_captcha_if_present(p)
        self.dismiss_consent_dialog(p)
        self.random_delay(0.5, 1.2)

        current_url = p.url

        # 1. Extract Business Name
        business_name = ""
        for name_sel in SELECTORS["business_name"]:
            try:
                elem = p.locator(name_sel).first
                if elem.is_visible(timeout=600):
                    text = elem.inner_text().strip()
                    if text:
                        business_name = text
                        break
            except Exception:
                continue

        if not business_name:
            # Fallback: Extract from document title
            try:
                doc_title = p.title()
                if " - Google Maps" in doc_title:
                    business_name = doc_title.replace(" - Google Maps", "").strip()
            except Exception:
                pass

        if not business_name:
            # Unable to extract business name, likely an invalid or blocked page
            return None

        # 2. Extract Phone Number
        phone = ""
        for phone_sel in SELECTORS["phone_buttons"]:
            try:
                btn = p.locator(phone_sel).first
                if btn.is_visible(timeout=400):
                    # Check aria-label or inner text or data-item-id
                    aria = btn.get_attribute("aria-label") or ""
                    tooltip = btn.get_attribute("data-tooltip") or ""
                    inner = btn.inner_text().strip()
                    item_id = btn.get_attribute("data-item-id") or ""

                    for text_source in [aria, inner, tooltip, item_id]:
                        match = PHONE_REGEX.search(text_source)
                        if match:
                            phone = match.group(0).strip()
                            break
                    if phone:
                        break
            except Exception:
                continue

        # Fallback phone search across visible text
        if not phone:
            try:
                body_text = p.locator("div[role='main']").inner_text()
                # Find lines that resemble phone numbers
                for line in body_text.splitlines():
                    match = PHONE_REGEX.search(line)
                    if match and len(match.group(0)) >= 8 and not any(c.isalpha() for c in match.group(0)):
                        phone = match.group(0).strip()
                        break
            except Exception:
                pass

        # 3. Extract Category / Business Type
        category = ""
        for cat_sel in SELECTORS["category_buttons"]:
            try:
                elem = p.locator(cat_sel).first
                if elem.is_visible(timeout=400):
                    text = elem.inner_text().strip()
                    if text and not text.startswith("(") and len(text) < 50:
                        category = text
                        break
            except Exception:
                continue

        # 4. Extract Address
        address = ""
        for addr_sel in SELECTORS["address_buttons"]:
            try:
                btn = p.locator(addr_sel).first
                if btn.is_visible(timeout=400):
                    aria = btn.get_attribute("aria-label") or ""
                    inner = btn.inner_text().strip()
                    if aria and "address:" in aria.lower():
                        address = aria.split(":", 1)[-1].strip()
                    elif inner:
                        address = inner
                    if address:
                        break
            except Exception:
                continue

        # 5. Extract Website
        website = ""
        for web_sel in SELECTORS["website_links"]:
            try:
                elem = p.locator(web_sel).first
                if elem.is_visible(timeout=400):
                    href = elem.get_attribute("href")
                    if href:
                        website = href
                        break
            except Exception:
                continue

        # 6. Extract Rating & Reviews Count
        rating = ""
        reviews_count = ""
        for r_sel in SELECTORS["rating_elements"]:
            try:
                elem = p.locator(r_sel).first
                if elem.is_visible(timeout=300):
                    rating = elem.inner_text().strip()
                    if rating:
                        break
            except Exception:
                continue

        for rev_sel in SELECTORS["reviews_elements"]:
            try:
                elem = p.locator(rev_sel).first
                if elem.is_visible(timeout=300):
                    text = elem.inner_text().strip()
                    if text:
                        reviews_count = text.replace("(", "").replace(")", "").replace("reviews", "").strip()
                        break
            except Exception:
                continue

        # 7. CID & Short URL Extraction
        hex_cid, decimal_cid = extract_cid_from_url(current_url)
        if not decimal_cid:
            hex_cid, decimal_cid = extract_cid_from_url(place_url)

        short_cid_url = get_short_cid_url(decimal_cid if decimal_cid else current_url)

        cid_key = decimal_cid or f"hex_{hex_cid}" if hex_cid else f"url_{hash(current_url)}"

        # 8. Optional City Filtering
        if self.filter_city and target_city:
            clean_target = target_city.split(",")[0].strip().lower()
            if address:
                if clean_target not in address.lower():
                    # Outside target city
                    return None

        lead_data = {
            "cid": cid_key,
            "business_name": business_name,
            "phone": phone,
            "category": category,
            "address": address,
            "website": website,
            "google_maps_url": current_url,
            "short_cid_url": short_cid_url,
            "rating": rating,
            "reviews_count": reviews_count,
            "keyword": keyword,
            "city": target_city or "",
        }

        return lead_data
