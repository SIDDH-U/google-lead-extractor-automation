"""
Database layer for Google Maps Leads Extractor.
Uses SQLite for thread-safe persistent storage and auto-syncs to JSON after every saved listing.
Supports checkpointing to resume runs seamlessly without duplicate work.
"""
import json
import sqlite3
from pathlib import Path
from typing import Dict, List, Optional, Any, Union
from config import DEFAULT_DB_PATH, DEFAULT_JSON_PATH


class DatabaseManager:
    """Manages SQLite database and JSON checkpointing."""

    def __init__(self, db_path: Union[str, Path] = DEFAULT_DB_PATH, json_path: Union[str, Path] = DEFAULT_JSON_PATH):
        self.db_path = Path(db_path)
        self.json_path = Path(json_path)
        self.init_db()

    def get_connection(self) -> sqlite3.Connection:
        """Returns a connection with Row factory enabled."""
        conn = sqlite3.connect(self.db_path, timeout=30.0)
        conn.row_factory = sqlite3.Row
        return conn

    def init_db(self) -> None:
        """Create tables if they don't already exist."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            # Leads Table - Deduplicated primarily by CID, fallback to google_maps_url
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS leads (
                    cid TEXT PRIMARY KEY,
                    business_name TEXT NOT NULL,
                    phone TEXT,
                    category TEXT,
                    address TEXT,
                    website TEXT,
                    google_maps_url TEXT,
                    short_cid_url TEXT,
                    rating TEXT,
                    reviews_count TEXT,
                    lead_quality TEXT,
                    keyword TEXT,
                    city TEXT,
                    scraped_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            # Search Grid Progress Table - Tracks each grid point completed for a keyword/city
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS search_progress (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    keyword TEXT NOT NULL,
                    city TEXT NOT NULL,
                    grid_x INTEGER NOT NULL,
                    grid_y INTEGER NOT NULL,
                    lat REAL NOT NULL,
                    lng REAL NOT NULL,
                    status TEXT NOT NULL, -- 'pending', 'completed', 'failed'
                    leads_found INTEGER DEFAULT 0,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(keyword, city, grid_x, grid_y)
                )
            """)

            # Discovered Place URLs Table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS discovered_urls (
                    url TEXT PRIMARY KEY,
                    cid TEXT,
                    keyword TEXT,
                    city TEXT,
                    status TEXT DEFAULT 'discovered', -- 'discovered', 'scraped', 'skipped'
                    discovered_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            conn.commit()

    def is_cid_saved(self, cid: str) -> bool:
        """Check if a business with this CID already exists in the database."""
        if not cid:
            return False
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT 1 FROM leads WHERE cid = ?", (cid,))
            return cursor.fetchone() is not None

    def is_url_processed(self, url: str) -> bool:
        """Check if a URL has already been processed or scraped."""
        if not url:
            return False
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT 1 FROM discovered_urls WHERE url = ? AND status IN ('scraped', 'skipped')", (url,))
            return cursor.fetchone() is not None

    def save_lead(self, lead: Dict[str, Any]) -> bool:
        """
        Insert or update a lead in SQLite and sync immediately to JSON.
        Returns True if newly inserted, False if updated or skipped.
        """
        cid = lead.get("cid")
        if not cid:
            # Generate fallback unique ID from URL if CID is not available
            cid = f"url_{hash(lead.get('google_maps_url', ''))}"
            lead["cid"] = cid

        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            # Check existence
            cursor.execute("SELECT 1 FROM leads WHERE cid = ?", (cid,))
            exists = cursor.fetchone() is not None

            cursor.execute("""
                INSERT INTO leads (
                    cid, business_name, phone, category, address, website,
                    google_maps_url, short_cid_url, rating, reviews_count,
                    keyword, city
                ) VALUES (
                    :cid, :business_name, :phone, :category, :address, :website,
                    :google_maps_url, :short_cid_url, :rating, :reviews_count,
                    :keyword, :city
                )
                ON CONFLICT(cid) DO UPDATE SET
                    business_name = excluded.business_name,
                    phone = COALESCE(NULLIF(excluded.phone, ''), leads.phone),
                    category = COALESCE(NULLIF(excluded.category, ''), leads.category),
                    address = COALESCE(NULLIF(excluded.address, ''), leads.address),
                    website = COALESCE(NULLIF(excluded.website, ''), leads.website),
                    google_maps_url = excluded.google_maps_url,
                    short_cid_url = excluded.short_cid_url,
                    rating = COALESCE(NULLIF(excluded.rating, ''), leads.rating),
                    reviews_count = COALESCE(NULLIF(excluded.reviews_count, ''), leads.reviews_count),
                    keyword = excluded.keyword,
                    city = excluded.city
            """, lead)

            # Mark in discovered_urls
            if lead.get("google_maps_url"):
                cursor.execute("""
                    INSERT INTO discovered_urls (url, cid, keyword, city, status)
                    VALUES (?, ?, ?, ?, 'scraped')
                    ON CONFLICT(url) DO UPDATE SET status = 'scraped', cid = ?
                """, (lead["google_maps_url"], cid, lead.get("keyword"), lead.get("city"), cid))

            conn.commit()

        # Checkpoint to JSON after every lead
        self.export_to_json()
        return not exists

    def save_discovered_urls(self, urls: List[str], keyword: str, city: str) -> int:
        """Save a batch of discovered place URLs."""
        new_count = 0
        with self.get_connection() as conn:
            cursor = conn.cursor()
            for url in urls:
                try:
                    cursor.execute("""
                        INSERT OR IGNORE INTO discovered_urls (url, keyword, city, status)
                        VALUES (?, ?, ?, 'discovered')
                    """, (url, keyword, city))
                    if cursor.rowcount > 0:
                        new_count += 1
                except sqlite3.Error:
                    pass
            conn.commit()
        return new_count

    def get_pending_urls(self, keyword: Optional[str] = None, city: Optional[str] = None) -> List[str]:
        """Fetch discovered URLs that have not been scraped yet."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            if keyword and city:
                cursor.execute(
                    "SELECT url FROM discovered_urls WHERE status = 'discovered' AND keyword = ? AND city = ?",
                    (keyword, city)
                )
            elif city:
                cursor.execute("SELECT url FROM discovered_urls WHERE status = 'discovered' AND city = ?", (city,))
            else:
                cursor.execute("SELECT url FROM discovered_urls WHERE status = 'discovered'")
            return [row["url"] for row in cursor.fetchall()]

    def get_pending_urls_with_meta(self, keyword: Optional[str] = None, city: Optional[str] = None) -> List[Dict[str, str]]:
        """Fetch discovered URLs with their associated keyword and city metadata."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            if keyword and city:
                cursor.execute(
                    "SELECT url, keyword, city FROM discovered_urls WHERE status = 'discovered' AND keyword = ? AND city = ?",
                    (keyword, city)
                )
            elif city:
                cursor.execute("SELECT url, keyword, city FROM discovered_urls WHERE status = 'discovered' AND city = ?", (city,))
            else:
                cursor.execute("SELECT url, keyword, city FROM discovered_urls WHERE status = 'discovered'")
            return [dict(row) for row in cursor.fetchall()]

    def mark_url_status(self, url: str, status: str, cid: Optional[str] = None) -> None:
        """Update status of a discovered URL ('scraped', 'skipped', 'failed')."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE discovered_urls SET status = ?, cid = COALESCE(?, cid) WHERE url = ?
            """, (status, cid, url))
            conn.commit()

    def record_grid_progress(
        self, keyword: str, city: str, grid_x: int, grid_y: int,
        lat: float, lng: float, status: str, leads_found: int = 0
    ) -> None:
        """Record the search status of a specific grid coordinate."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO search_progress (keyword, city, grid_x, grid_y, lat, lng, status, leads_found)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(keyword, city, grid_x, grid_y) DO UPDATE SET
                    status = excluded.status,
                    leads_found = excluded.leads_found,
                    updated_at = CURRENT_TIMESTAMP
            """, (keyword, city, grid_x, grid_y, lat, lng, status, leads_found))
            conn.commit()

    def is_grid_point_completed(self, keyword: str, city: str, grid_x: int, grid_y: int) -> bool:
        """Check if a specific grid point was already completed in a previous run."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT 1 FROM search_progress
                WHERE keyword = ? AND city = ? AND grid_x = ? AND grid_y = ? AND status = 'completed'
            """, (keyword, city, grid_x, grid_y))
            return cursor.fetchone() is not None

    def get_all_leads(self) -> List[Dict[str, Any]]:
        """Retrieve all leads from the database."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM leads ORDER BY scraped_at DESC")
            return [dict(row) for row in cursor.fetchall()]

    def get_leads_count(self) -> int:
        """Count total unique leads in the database."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) as cnt FROM leads")
            row = cursor.fetchone()
            return row["cnt"] if row else 0

    def delete_lead(self, cid: str) -> bool:
        """Delete a single lead by its CID and update JSON."""
        if not cid:
            return False
        clean_cid = str(cid).strip()
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM leads WHERE cid = ?", (clean_cid,))
            deleted = cursor.rowcount > 0
            conn.commit()
        if deleted:
            try:
                self.export_to_json()
            except Exception:
                pass
        return deleted

    def delete_leads_batch(self, cids: List[str]) -> int:
        """Delete multiple leads by their CIDs in a single transaction."""
        if not cids:
            return 0
        clean_cids = [str(c).strip() for c in cids if str(c).strip()]
        if not clean_cids:
            return 0
        
        deleted_count = 0
        with self.get_connection() as conn:
            cursor = conn.cursor()
            # Chunk into batches of 500 for SQLite parameter limits
            batch_size = 500
            for i in range(0, len(clean_cids), batch_size):
                chunk = clean_cids[i:i + batch_size]
                placeholders = ",".join("?" * len(chunk))
                cursor.execute(f"DELETE FROM leads WHERE cid IN ({placeholders})", chunk)
                deleted_count += cursor.rowcount
            conn.commit()

        if deleted_count > 0:
            try:
                self.export_to_json()
            except Exception:
                pass
        return deleted_count

    def delete_leads_by_section(self, keyword: Optional[str] = None, city: Optional[str] = None) -> int:
        """Delete all leads matching a specific search section (keyword and/or city)."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            if keyword and city:
                cursor.execute("DELETE FROM leads WHERE keyword = ? AND city = ?", (keyword, city))
            elif keyword:
                cursor.execute("DELETE FROM leads WHERE keyword = ?", (keyword,))
            elif city:
                cursor.execute("DELETE FROM leads WHERE city = ?", (city,))
            else:
                return 0
            deleted_count = cursor.rowcount
            conn.commit()

        if deleted_count > 0:
            try:
                self.export_to_json()
            except Exception:
                pass
        return deleted_count

    def get_sections_summary(self) -> List[Dict[str, Any]]:
        """Get summary of distinct search runs / sections grouped by keyword and city."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT 
                    COALESCE(NULLIF(keyword, ''), 'General') as keyword,
                    COALESCE(NULLIF(city, ''), 'All') as city,
                    COUNT(*) as count,
                    SUM(CASE WHEN phone IS NOT NULL AND phone != '' THEN 1 ELSE 0 END) as with_phone
                FROM leads
                GROUP BY keyword, city
                ORDER BY count DESC
            """)
            return [dict(row) for row in cursor.fetchall()]

    def get_leads_by_cids(self, cids: List[str]) -> List[Dict[str, Any]]:
        """Retrieve specific leads by their CIDs."""
        clean_cids = [str(c).strip() for c in cids if str(c).strip()]
        if not clean_cids:
            return []
        results = []
        with self.get_connection() as conn:
            cursor = conn.cursor()
            batch_size = 500
            for i in range(0, len(clean_cids), batch_size):
                chunk = clean_cids[i:i + batch_size]
                placeholders = ",".join("?" * len(chunk))
                cursor.execute(f"SELECT * FROM leads WHERE cid IN ({placeholders})", chunk)
                results.extend([dict(row) for row in cursor.fetchall()])
        return results

    def clear_all_leads(self) -> int:
        """Delete all leads from the leads table and re-sync JSON."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM leads")
            count = cursor.rowcount
            conn.commit()
        try:
            self.export_to_json()
        except Exception:
            pass
        return count

    def clear_all_data(self) -> None:
        """Reset entire database (leads, search_progress, discovered_urls)."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM leads")
            cursor.execute("DELETE FROM search_progress")
            cursor.execute("DELETE FROM discovered_urls")
            conn.commit()
        self.export_to_json()

    def export_to_json(self, json_path: Optional[Union[str, Path]] = None) -> Path:
        """Export all current leads from SQLite into a clean JSON file."""
        target_path = Path(json_path) if json_path else self.json_path
        leads = self.get_all_leads()
        with open(target_path, "w", encoding="utf-8") as f:
            json.dump(leads, f, indent=2, ensure_ascii=False)
        return target_path

    def get_stats(self) -> Dict[str, Any]:
        """Get summary metrics for current database."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) as total_leads FROM leads")
            total_leads = cursor.fetchone()["total_leads"]

            cursor.execute("SELECT COUNT(DISTINCT phone) as phones FROM leads WHERE phone IS NOT NULL AND phone != ''")
            phones = cursor.fetchone()["phones"]

            cursor.execute("SELECT COUNT(*) as discovered FROM discovered_urls")
            discovered = cursor.fetchone()["discovered"]

            cursor.execute("SELECT COUNT(*) as completed_points FROM search_progress WHERE status = 'completed'")
            completed_points = cursor.fetchone()["completed_points"]

            return {
                "total_leads": total_leads,
                "leads_with_phone": phones,
                "total_discovered_urls": discovered,
                "completed_grid_points": completed_points,
            }
