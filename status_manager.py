"""
Shared Status Manager for Google Maps Leads Extractor.
Allows the background scraper and the local web dashboard to exchange real-time status.
Supports pause, resume, and stop controls.
"""
import json
import time
from pathlib import Path
from typing import Dict, Any, Optional

BASE_DIR = Path(__file__).resolve().parent
STATUS_FILE = BASE_DIR / "scraper_status.json"
PAUSE_FILE = BASE_DIR / "scraper_pause.flag"
STOP_FILE = BASE_DIR / "scraper_stop.flag"


def request_pause() -> None:
    """Signal scraper to pause."""
    try:
        PAUSE_FILE.touch()
    except Exception:
        pass


def request_resume() -> None:
    """Signal scraper to resume."""
    if PAUSE_FILE.exists():
        try:
            PAUSE_FILE.unlink()
        except Exception:
            pass


def is_pause_requested() -> bool:
    """Check if pause is active."""
    return PAUSE_FILE.exists()


def request_stop() -> None:
    """Signal scraper to stop completely."""
    try:
        STOP_FILE.touch()
    except Exception:
        pass


def is_stop_requested() -> bool:
    """Check if stop has been signaled."""
    return STOP_FILE.exists()


def clear_stop_request() -> None:
    """Clear stop signal for a fresh run."""
    if STOP_FILE.exists():
        try:
            STOP_FILE.unlink()
        except Exception:
            pass


def update_status(
    is_running: bool,
    phase: str = "Idle",
    current_url: str = "",
    current_keyword: str = "",
    current_city: str = "",
    details: str = "",
) -> None:
    """Write current scraper state to JSON file while preserving completion records."""
    existing_completion = None
    if STATUS_FILE.exists():
        try:
            with open(STATUS_FILE, "r", encoding="utf-8") as f:
                old = json.load(f)
                existing_completion = old.get("last_completion")
        except Exception:
            pass

    data = {
        "is_running": is_running,
        "is_paused": is_pause_requested(),
        "phase": phase,
        "current_url": current_url,
        "current_keyword": current_keyword,
        "current_city": current_city,
        "details": details,
        "timestamp": time.time(),
    }
    if existing_completion:
        data["last_completion"] = existing_completion

    try:
        temp_file = STATUS_FILE.with_suffix(".tmp")
        with open(temp_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        temp_file.replace(STATUS_FILE)
    except Exception:
        pass


def mark_completed(
    city: str = "",
    keywords: Optional[list] = None,
    total_leads: int = 0,
    new_leads: int = 0,
    excel_path: str = "",
    stopped_by_user: bool = False,
) -> None:
    """Explicitly record search completion with rich stats for web dashboard and CLI notifications."""
    completion_id = f"run_{int(time.time() * 1000)}"
    kw_str = ", ".join(keywords) if keywords else "all"
    status_phase = "Stopped by User" if stopped_by_user else "Search Completed"
    status_details = (
        f"Search ended. Total leads in database: {total_leads} (+{new_leads} new)"
        if not stopped_by_user
        else "Search was safely stopped by user."
    )

    data = {
        "is_running": False,
        "is_paused": False,
        "phase": status_phase,
        "current_url": "",
        "current_keyword": kw_str,
        "current_city": city,
        "details": status_details,
        "timestamp": time.time(),
        "last_completion": {
            "completion_id": completion_id,
            "completed_at": time.time(),
            "city": city,
            "keywords": keywords or [],
            "keywords_str": kw_str,
            "total_leads": total_leads,
            "new_leads": new_leads,
            "excel_path": excel_path,
            "stopped_by_user": stopped_by_user,
        },
    }

    try:
        temp_file = STATUS_FILE.with_suffix(".tmp")
        with open(temp_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        temp_file.replace(STATUS_FILE)
    except Exception:
        pass


def get_status() -> Dict[str, Any]:
    """Read current scraper state, with 25-second heartbeat check."""
    default = {
        "is_running": False,
        "is_paused": False,
        "phase": "Idle",
        "current_url": "",
        "current_keyword": "",
        "current_city": "",
        "details": "",
        "timestamp": 0,
    }
    if not STATUS_FILE.exists():
        return default

    try:
        with open(STATUS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)

        # Sync pause flag
        data["is_paused"] = is_pause_requested()

        # Heartbeat check: if last update was more than 120 seconds ago, consider it Idle
        if time.time() - data.get("timestamp", 0) > 120:
            data["is_running"] = False
            data["phase"] = "Idle / Stopped"
        return data
    except Exception:
        return default
