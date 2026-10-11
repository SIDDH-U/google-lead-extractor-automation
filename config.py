"""
Configuration and constants for Google Maps Leads Extractor.
"""
from pathlib import Path

# Base Paths
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_DB_PATH = BASE_DIR / "leads.db"
DEFAULT_JSON_PATH = BASE_DIR / "leads.json"
DEFAULT_LOG_PATH = BASE_DIR / "scraper.log"
DEFAULT_EXCEL_PATH = BASE_DIR / "leads_export.xlsx"

# OpenStreetMap Nominatim Geocoding API
NOMINATIM_BASE_URL = "https://nominatim.openstreetmap.org/search"
NOMINATIM_USER_AGENT = "GoogleMapsLeadExtractor/1.0 (local-desktop-tool; contact: contact@local.dev)"
NOMINATIM_TIMEOUT_SEC = 15

# Grid Settings
DEFAULT_GRID_STEP_KM = 1.0  # ~1 km distance between search points
DEFAULT_ZOOM = 15           # Standard Google Maps search zoom level (15z is ideal for ~1-2km neighborhood radius)

# Delays (Anti-ban / Rate limiting)
MIN_DELAY_SEC = 1.0
MAX_DELAY_SEC = 3.0
PAGE_LOAD_TIMEOUT_MS = 35000
SCROLL_PAUSE_SEC = 1.8
MAX_CONSECUTIVE_EMPTY_SCROLLS = 6

# Browser Settings
HEADLESS = False  # Headed mode required by user specifications
DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/125.0.0.0 Safari/537.36"
)
VIEWPORT_WIDTH = 1366
VIEWPORT_HEIGHT = 800

# Google Maps CSS Selectors & Matchers
SELECTORS = {
    # Consent dialogs (EU / Global)
    "consent_buttons": [
        'button[aria-label="Accept all"]',
        'button[aria-label="Reject all"]',
        'form[action*="consent"] button',
        'button:has-text("Accept all")',
        'button:has-text("Reject all")',
        'button:has-text("I agree")',
        'button:has-text("Accept")',
    ],
    # Search results feed container
    "feed_container": [
        'div[role="feed"]',
        'div[aria-label^="Results for"]',
        'div.m6QErb[aria-label]',
        'div.m6QErb.DxyBCb.kA9KIf.dS8AEf',
    ],
    # Listing items in results panel
    "listing_link": 'a[href*="/maps/place/"]',
    # End of list indicators
    "end_of_list_indicators": [
        "div.HlvSq",
        "span.HlvSq",
        "p.fontBodyMedium:has-text(\"You've reached the end of the list\")",
        "div:has-text(\"You've reached the end of the list\")",
        "span:has-text(\"You've reached the end of the list\")",
    ],
    # Place Detail Elements
    "business_name": [
        "h1.DUwDvf",
        "h1.fontHeadlineLarge",
        "h1[tabindex='-1']",
        "div.lMbq3e h1",
    ],
    "phone_buttons": [
        'button[data-tooltip*="phone" i]',
        'button[data-item-id^="phone:"]',
        'button[aria-label*="Phone:" i]',
        'button[data-tooltip*="Call" i]',
        'button[data-tooltip*="Copy phone" i]',
        'button[data-item-id*="phone"]',
    ],
    "category_buttons": [
        "button.DkEaL",
        "button[jsaction*='category']",
        "span.DkEaL",
        "button.fontBodyMedium",
    ],
    "address_buttons": [
        'button[data-item-id="address"]',
        'button[data-tooltip*="address" i]',
        'button[aria-label*="Address:" i]',
        'button[data-tooltip*="Copy address" i]',
    ],
    "website_links": [
        'a[data-item-id="authority"]',
        'a[data-tooltip*="website" i]',
        'a[aria-label*="Website:" i]',
        'a[data-tooltip*="Open website" i]',
    ],
    "rating_elements": [
        'div.F7nice span[aria-hidden="true"]',
        'span.ceNzKf[aria-hidden="true"]',
    ],
    "reviews_elements": [
        'div.F7nice span:has-text("reviews")',
        'div.F7nice span:has-text("(")',
    ],
    # CAPTCHA / Unusual Traffic detection
    "captcha_indicators": [
        'iframe[src*="recaptcha"]',
        'form[action*="sorry"]',
        'div:has-text("unusual traffic")',
        'div:has-text("Our systems have detected")',
        'div#captcha-form',
        'iframe[title*="reCAPTCHA"]',
    ],
}

# Excel Export Dropdown Values
EXCEL_DROPDOWNS = {
    "call_status": ["answered", "unanswered", "call back"],
    "interest": ["interested", "not interested"],
    "send": ["send", "not send"],
    "lead_quality": ["quality lead", "not quality lead"],
}
