"""
Google Maps Leads Extractor - Main CLI & Workflow Orchestration.
Integrates:
- OpenStreetMap Nominatim Bounding Box Geocoding
- ~1 km Coordinate Grid Generator
- Playwright Headed Chromium Browser Automation
- Infinite-scroll Lead Extraction & Decimal CID Short-URL Conversion
- Per-listing SQLite and JSON Checkpointing
- Styled Excel Export with openpyxl and Data Validation Dropdowns
"""
import argparse
import sys
import time
from pathlib import Path
from typing import List, Optional
from urllib.parse import quote_plus
from status_manager import (
    update_status,
    mark_completed,
    is_pause_requested,
    is_stop_requested,
    clear_stop_request,
)

# Ensure UTF-8 stdout on Windows
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

from rich.panel import Panel
from rich.prompt import Prompt, Confirm
from rich.table import Table
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn

from cid_utils import extract_cid_from_url
from config import (
    DEFAULT_DB_PATH,
    DEFAULT_JSON_PATH,
    DEFAULT_EXCEL_PATH,
    DEFAULT_GRID_STEP_KM,
    DEFAULT_LOG_PATH,
)
from db import DatabaseManager
from exporter import export_leads_to_excel
from geo import get_city_bounding_box, generate_grid_points, GeoError
from progress_logger import setup_logger, console
from scraper import GoogleMapsScraper


def check_pause_and_stop(city: str = "", current_url: str = "", kw: str = "") -> bool:
    """Check if the user paused or stopped the scraper from the dashboard."""
    if is_stop_requested():
        console.print("\n[bold yellow]⏹️ Stop requested from dashboard. Finishing run safely...[/bold yellow]")
        return False

    if is_pause_requested():
        console.print("\n[bold yellow]⏸️ Scraper paused from dashboard. Waiting for resume...[/bold yellow]")
        while is_pause_requested():
            if is_stop_requested():
                return False
            update_status(
                is_running=True,
                phase="Paused by User",
                current_url=current_url,
                current_keyword=kw,
                current_city=city,
                details="Scraper paused. Click Resume on the dashboard to continue.",
            )
            time.sleep(1.5)
        console.print("[bold green]▶️ Resumed! Continuing lead extraction...[/bold green]")
    return True


def display_welcome_banner():
    """Render application header banner."""
    console.print(
        Panel(
            "[bold cyan]🗺️  Google Maps Leads Extractor (Free & Local)[/bold cyan]\n"
            "[dim]No Paid APIs • Playwright Headed Chromium • OSM Nominatim Grid • Excel Dropdowns[/dim]",
            border_style="cyan",
            padding=(1, 2),
        )
    )


def parse_arguments():
    """Parse command line flags."""
    parser = argparse.ArgumentParser(
        description="Extract business leads from Google Maps using Playwright and OpenStreetMap grid sampling.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "-k", "--keywords",
        nargs="+",
        help="One or more search keywords (e.g. 'plumber' 'electrician' or 'coffee shops')",
    )
    parser.add_argument(
        "-c", "--city",
        type=str,
        help="Target city name (e.g. 'Austin, TX' or 'Seattle, WA')",
    )
    parser.add_argument(
        "--step-km",
        type=float,
        default=DEFAULT_GRID_STEP_KM,
        help="Grid spacing distance in kilometers (~1 km recommended)",
    )
    parser.add_argument(
        "--filter-city",
        action="store_true",
        default=True,
        help="Filter out businesses whose address does not contain the target city",
    )
    parser.add_argument(
        "--no-filter-city",
        dest="filter_city",
        action="store_false",
        help="Disable strict city address filtering",
    )
    parser.add_argument(
        "-o", "--output",
        type=str,
        default=str(DEFAULT_EXCEL_PATH),
        help="Output Excel file path (.xlsx)",
    )
    parser.add_argument(
        "--db",
        type=str,
        default=str(DEFAULT_DB_PATH),
        help="SQLite database path for checkpoints",
    )
    parser.add_argument(
        "--json",
        type=str,
        default=str(DEFAULT_JSON_PATH),
        help="JSON file path for real-time backup",
    )
    parser.add_argument(
        "--scrape-only",
        action="store_true",
        default=False,
        help="Skip grid scanning and scrape all pending discovered URLs already in the database",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Maximum total unique leads to scrape (optional)",
    )

    return parser.parse_args()


def interactive_prompt_inputs():
    """Prompt user interactively if CLI flags were not provided."""
    console.print("[bold yellow]No command line arguments detected. Entering interactive mode:[/bold yellow]\n")
    
    keywords_raw = Prompt.ask(
        "[bold green]Enter search keyword(s)[/bold green] (comma-separated, e.g. 'plumber, electrician')"
    )
    keywords = [k.strip() for k in keywords_raw.split(",") if k.strip()]
    while not keywords:
        console.print("[red]Please enter at least one valid keyword.[/red]")
        keywords_raw = Prompt.ask("[bold green]Enter search keyword(s)[/bold green]")
        keywords = [k.strip() for k in keywords_raw.split(",") if k.strip()]

    city = Prompt.ask(
        "[bold green]Enter target city[/bold green] (e.g. 'Austin, TX' or 'Miami, FL')"
    ).strip()
    while not city:
        console.print("[red]City name cannot be empty.[/red]")
        city = Prompt.ask("[bold green]Enter target city[/bold green]").strip()

    step_km_str = Prompt.ask(
        f"[bold green]Grid step size in km[/bold green]",
        default=str(DEFAULT_GRID_STEP_KM),
    )
    try:
        step_km = float(step_km_str)
    except ValueError:
        step_km = DEFAULT_GRID_STEP_KM

    filter_city = Confirm.ask(
        "[bold green]Filter out businesses strictly outside this city?[/bold green]",
        default=True,
    )

    output = Prompt.ask(
        "[bold green]Output Excel filename[/bold green]",
        default=str(DEFAULT_EXCEL_PATH),
    )

    return keywords, city, step_km, filter_city, output


def run_extraction_workflow(
    keywords: List[str],
    city: str,
    step_km: float,
    filter_city: bool,
    output_excel: str,
    db_path: str = str(DEFAULT_DB_PATH),
    json_path: str = str(DEFAULT_JSON_PATH),
    lead_limit: Optional[int] = None,
    scrape_only: bool = False,
):
    """Execute the full end-to-end extraction pipeline."""
    logger = setup_logger()
    logger.info("Initializing Google Maps Leads Extractor...")

    db = DatabaseManager(db_path=db_path, json_path=json_path)
    initial_leads_count = db.get_leads_count()

    # Initialize Scraper
    scraper = GoogleMapsScraper(
        db_manager=db,
        filter_city=filter_city,
    )

    clear_stop_request()

    try:
        update_status(
            is_running=True,
            phase="Initializing Browser",
            current_city=city,
            details="Starting headed Chromium browser session...",
        )
        scraper.start_browser()

        # Step 0: Drain any already-discovered pending listings from previous runs
        initial_pending = db.get_pending_urls_with_meta(city=city if city else None)
        if initial_pending:
            console.print(f"\n[bold cyan]Found {len(initial_pending)} pending discovered listings from previous run. Extracting now...[/bold cyan]")
            for idx, item in enumerate(initial_pending, 1):
                url = item["url"]
                item_city = city or item.get("city", "")
                item_kw = item.get("keyword", "")
                if lead_limit and db.get_leads_count() >= lead_limit:
                    break
                if db.is_url_processed(url):
                    continue

                if not check_pause_and_stop(city=item_city, current_url=url, kw=item_kw):
                    break

                update_status(
                    is_running=True,
                    phase=f"Extracting Pending Listings ({idx}/{len(initial_pending)})",
                    current_url=url,
                    current_keyword=item_kw,
                    current_city=item_city,
                    details=f"Extracting: {url[:50]}...",
                )
                try:
                    lead = scraper.extract_place_details(url, target_city=item_city, keyword=item_kw)
                    if lead:
                        is_new = db.save_lead(lead)
                        status_tag = "Added" if is_new else "Updated"
                        phone_tag = f" | Phone: {lead['phone']}" if lead['phone'] else ""
                        logger.info(f"[{status_tag}] {lead['business_name']} ({lead['category']}){phone_tag}")
                    else:
                        db.mark_url_status(url, "skipped")
                except Exception as e:
                    db.mark_url_status(url, "failed")

        if not scrape_only:
            # Step 1: Geocoding via OpenStreetMap Nominatim
            console.print(f"\n[bold cyan]Step 1: Geocoding city bounding box for '[white]{city}[/white]'...[/bold cyan]")
            try:
                bbox_data = get_city_bounding_box(city)
            except GeoError as e:
                console.print(f"[bold red]❌ Geocoding Error:[/bold red] {e}")
                sys.exit(1)

            table_geo = Table(title="City Geocoding Details", border_style="cyan")
            table_geo.add_column("Property", style="bold white")
            table_geo.add_column("Value", style="green")
            table_geo.add_row("City", bbox_data["city"])
            table_geo.add_row("Full Address", bbox_data["display_name"])
            table_geo.add_row("Latitude Range", f"{bbox_data['lat_min']:.5f} to {bbox_data['lat_max']:.5f}")
            table_geo.add_row("Longitude Range", f"{bbox_data['lon_min']:.5f} to {bbox_data['lon_max']:.5f}")
            table_geo.add_row("Center Point", f"{bbox_data['center_lat']:.5f}, {bbox_data['center_lon']:.5f}")
            console.print(table_geo)

            # Step 2: Generate Grid Points
            console.print(f"\n[bold cyan]Step 2: Generating search grid points (~{step_km} km spacing)...[/bold cyan]")
            grid_points = generate_grid_points(bbox_data, step_km=step_km)
            total_grid_points = len(grid_points)
            total_searches = total_grid_points * len(keywords)

            console.print(
                f"[green]✓ Generated [bold]{total_grid_points}[/bold] coordinate grid points.[/green] "
                f"Across [bold]{len(keywords)}[/bold] keyword(s), total map searches: [bold]{total_searches}[/bold]."
            )

            console.print("\n[bold cyan]Step 3: Scanning Grid Points & Streaming Live Lead Extraction...[/bold cyan]")

            # Iterate over Keywords and Grid Points
            search_idx = 0
            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                TaskProgressColumn(),
                console=console,
            ) as progress:
                search_task = progress.add_task("[cyan]Grid Search Progress", total=total_searches)

                for kw in keywords:
                    for lat, lng, gx, gy in grid_points:
                        search_idx += 1
                        point_label = f"Kw: '{kw}' | Point ({gx},{gy}) [{lat:.4f}, {lng:.4f}]"

                        # Check if already completed in previous run
                        if db.is_grid_point_completed(kw, city, gx, gy):
                            progress.update(
                                search_task,
                                advance=1,
                                description=f"[dim]Skipping completed {point_label}[/dim]",
                            )
                            continue

                        progress.update(
                            search_task,
                            description=f"[cyan]Searching {point_label}...[/cyan]",
                        )

                        search_term = f"{kw.strip()} in {city.strip()}" if city else kw.strip()
                        encoded_kw = quote_plus(search_term)
                        search_url = f"https://www.google.com/maps/search/{encoded_kw}/@{lat},{lng},15z"

                        if not check_pause_and_stop(city=city, current_url=search_url, kw=kw):
                            break

                        update_status(
                            is_running=True,
                            phase=f"Scanning Map Grid ({gx},{gy})",
                            current_url=search_url,
                            current_keyword=kw,
                            current_city=city,
                            details=f"Searching '{kw}' around [{lat:.4f}, {lng:.4f}]",
                        )

                        try:
                            urls = scraper.search_grid_point(kw, lat, lng, city=city)
                            new_urls_count = db.save_discovered_urls(urls, kw, city)
                            db.record_grid_progress(kw, city, gx, gy, lat, lng, "completed", len(urls))
                            logger.info(f"Grid point ({lat:.4f}, {lng:.4f}) found {len(urls)} URLs ({new_urls_count} new)")

                            # Stream immediate extraction of any newly discovered listings
                            point_pending = db.get_pending_urls_with_meta(keyword=kw, city=city)
                            if point_pending:
                                for p_item in point_pending:
                                    p_url = p_item["url"]
                                    if lead_limit and db.get_leads_count() >= lead_limit:
                                        break
                                    if db.is_url_processed(p_url):
                                        continue

                                    if not check_pause_and_stop(city=city, current_url=p_url, kw=kw):
                                        break

                                    update_status(
                                        is_running=True,
                                        phase="Extracting Lead Details",
                                        current_url=p_url,
                                        current_keyword=kw,
                                        current_city=city,
                                        details=f"Extracting listing details: {p_url[:50]}...",
                                    )

                                    try:
                                        lead = scraper.extract_place_details(p_url, target_city=city, keyword=kw)
                                        if lead:
                                            is_new = db.save_lead(lead)
                                            status_tag = "Added" if is_new else "Updated"
                                            phone_tag = f" | Phone: {lead['phone']}" if lead['phone'] else ""
                                            logger.info(f"[{status_tag}] {lead['business_name']} ({lead['category']}){phone_tag}")
                                        else:
                                            db.mark_url_status(p_url, "skipped")
                                    except Exception as e:
                                        db.mark_url_status(p_url, "failed")
                                        logger.error(f"Failed to scrape listing {p_url}: {e}")

                        except Exception as e:
                            logger.error(f"Error searching grid point {point_label}: {e}")
                            db.record_grid_progress(kw, city, gx, gy, lat, lng, "failed", 0)

                        progress.update(search_task, advance=1)
                        scraper.random_delay()
        else:
            console.print("\n[yellow]⚡ Skipping grid scanning (--scrape-only mode enabled)...[/yellow]")

        # Final sweep for any remaining un-scraped URLs
        remaining_pending = db.get_pending_urls_with_meta(city=city if city else None)
        if remaining_pending:
            console.print(f"\n[bold cyan]Final sweep: Scraping remaining [bold]{len(remaining_pending)}[/bold] discovered listings...[/bold cyan]")
            for item in remaining_pending:
                url = item["url"]
                item_city = city or item.get("city", "")
                item_kw = item.get("keyword", "")

                if lead_limit and db.get_leads_count() >= lead_limit:
                    break
                if db.is_url_processed(url):
                    continue

                if not check_pause_and_stop(city=item_city, current_url=url, kw=item_kw):
                    break

                update_status(
                    is_running=True,
                    phase="Final Sweep: Extracting Details",
                    current_url=url,
                    current_keyword=item_kw,
                    current_city=item_city,
                    details=f"Extracting: {url[:50]}...",
                )

                try:
                    lead = scraper.extract_place_details(url, target_city=item_city, keyword=item_kw)
                    if lead:
                        is_new = db.save_lead(lead)
                        status_tag = "Added" if is_new else "Updated"
                        phone_tag = f" | Phone: {lead['phone']}" if lead['phone'] else ""
                        logger.info(f"[{status_tag}] {lead['business_name']} ({lead['category']}){phone_tag}")
                    else:
                        db.mark_url_status(url, "skipped")
                except Exception as e:
                    db.mark_url_status(url, "failed")

    finally:
        console.print("\n[dim]Closing browser session...[/dim]")
        scraper.close_browser()

    # Step 7: Export to Excel with openpyxl
    console.print(f"\n[bold cyan]Step 5: Exporting Leads to Excel Workbook with Dropdowns...[/bold cyan]")
    excel_path = export_leads_to_excel(db, output_excel)

    stopped = is_stop_requested()
    final_leads_count = db.get_leads_count()
    new_leads_extracted = max(0, final_leads_count - initial_leads_count)

    # Record structured completion in status manager for dashboard alerts
    mark_completed(
        city=city,
        keywords=keywords,
        total_leads=final_leads_count,
        new_leads=new_leads_extracted,
        excel_path=str(excel_path.resolve()) if excel_path else "",
        stopped_by_user=stopped,
    )

    # Play completion alert chime/beep on Windows
    try:
        import winsound
        winsound.MessageBeep(winsound.MB_ICONASTERISK)
    except Exception:
        pass

    # Print Final Summary Statistics Table
    stats = db.get_stats()
    table_summary = Table(title="Extraction Results Summary", border_style="green")
    table_summary.add_column("Metric", style="bold white")
    table_summary.add_column("Count", style="bold green")
    table_summary.add_row("Total Unique Leads Extracted", str(stats["total_leads"]))
    table_summary.add_row("New Leads Added This Run", f"+{new_leads_extracted}")
    table_summary.add_row("Leads with Phone Numbers", str(stats["leads_with_phone"]))
    table_summary.add_row("Total Place URLs Discovered", str(stats["total_discovered_urls"]))
    table_summary.add_row("Completed Grid Search Points", str(stats["completed_grid_points"]))
    table_summary.add_row("SQLite Checkpoint Database", str(db_path))
    table_summary.add_row("Realtime JSON Backup", str(json_path))
    table_summary.add_row("Styled Excel Spreadsheet", str(excel_path.resolve()))
    console.print(table_summary)

    if stopped:
        console.print(
            Panel(
                f"[bold yellow]⏹️ SEARCH SAFELY STOPPED BY USER[/bold yellow]\n\n"
                f"📍 [bold white]City:[/bold white] {city}\n"
                f"🔑 [bold white]Keywords:[/bold white] {', '.join(keywords)}\n"
                f"📈 [bold green]New Leads Extracted:[/bold green] +{new_leads_extracted}\n"
                f"💾 [bold white]Total Leads in Database:[/bold white] {final_leads_count}\n"
                f"📊 [bold cyan]Excel Spreadsheet:[/bold cyan] [underline]{excel_path.resolve()}[/underline]\n",
                border_style="yellow",
                title="[bold yellow]⏹️ Scraper Stopped[/bold yellow]",
                padding=(1, 2),
            )
        )
    else:
        console.print(
            Panel(
                f"[bold green]🎉 SEARCH COMPLETED SUCCESSFULLY![/bold green]\n\n"
                f"📍 [bold white]City:[/bold white] {city}\n"
                f"🔑 [bold white]Keywords:[/bold white] {', '.join(keywords)}\n"
                f"📈 [bold green]New Leads Extracted:[/bold green] +{new_leads_extracted}\n"
                f"💾 [bold white]Total Unique Leads in DB:[/bold white] {final_leads_count}\n"
                f"📊 [bold cyan]Excel Spreadsheet:[/bold cyan] [underline]{excel_path.resolve()}[/underline]\n"
                f"[dim]Pre-configured dropdowns ready for Call Status, Interest, and Send.[/dim]",
                border_style="green",
                title="[bold green]🏁 Extraction Finished[/bold green]",
                padding=(1, 2),
            )
        )


def main():
    """Main CLI entry point."""
    display_welcome_banner()
    args = parse_arguments()

    if args.scrape_only or (args.keywords and args.city):
        keywords = args.keywords or ["all"]
        city = args.city or ""
        step_km = args.step_km
        filter_city = args.filter_city
        output_excel = args.output
    elif args.city and not args.keywords:
        keywords = ["all"]
        city = args.city
        step_km = args.step_km
        filter_city = args.filter_city
        output_excel = args.output
    else:
        keywords, city, step_km, filter_city, output_excel = interactive_prompt_inputs()

    run_extraction_workflow(
        keywords=keywords,
        city=city,
        step_km=step_km,
        filter_city=filter_city,
        output_excel=output_excel,
        db_path=args.db,
        json_path=args.json,
        lead_limit=args.limit,
        scrape_only=args.scrape_only,
    )


if __name__ == "__main__":
    main()
