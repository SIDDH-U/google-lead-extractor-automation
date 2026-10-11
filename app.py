"""
Interactive Local Web Dashboard for Google Maps Leads Extractor.
Built with Python's standard http.server (zero extra dependencies).
Provides:
- Real-time interactive UI at http://localhost:5000
- Multi-section Lead Search Batches (preserve & switch between past searches)
- AI Category Analysis & 3-4 Dynamic Subcategories Section & Dropdown
- Multi-select Checkboxes with Bulk Delete & Bulk Export
- Instant search & filter across all extracted leads
- Live inline status dropdowns with SQLite auto-save
- Safe in-memory Excel (.xlsx) export & download
- Single lead delete with instant UI update
"""
import http.server
import json
import socketserver
import sys
import subprocess
import urllib.parse
from pathlib import Path
from typing import List, Dict, Any, Optional

# Ensure UTF-8 stdout
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

from config import DEFAULT_DB_PATH, DEFAULT_EXCEL_PATH
from db import DatabaseManager
from exporter import export_leads_to_excel, generate_excel_bytes
from status_manager import (
    get_status,
    request_pause,
    request_resume,
    request_stop,
    is_pause_requested,
    clear_stop_request,
)

PORT = 5000
DB_MANAGER = DatabaseManager(DEFAULT_DB_PATH)

HTML_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Google Maps Leads Dashboard</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <style>
    :root {
      --bg: #090d16;
      --card-bg: rgba(18, 26, 43, 0.78);
      --card-border: rgba(255, 255, 255, 0.08);
      --card-border-hover: rgba(56, 189, 248, 0.3);
      --accent: #38bdf8;
      --accent-glow: rgba(56, 189, 248, 0.25);
      --text: #f8fafc;
      --text-muted: #94a3b8;
      --success: #10b981;
      --warning: #f59e0b;
      --danger: #ef4444;
      --purple: #a855f7;
      --hover-bg: rgba(255, 255, 255, 0.04);
    }

    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
      font-family: 'Plus Jakarta Sans', -apple-system, sans-serif;
    }

    body {
      background: radial-gradient(circle at 50% 0%, #172554 0%, var(--bg) 60%);
      color: var(--text);
      min-height: 100vh;
      padding: 2rem 1.5rem 6rem;
    }

    .container {
      max-width: 1480px;
      margin: 0 auto;
    }

    header {
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      gap: 1.5rem;
      margin-bottom: 1.75rem;
      padding-bottom: 1.5rem;
      border-bottom: 1px solid var(--card-border);
    }

    .brand {
      display: flex;
      align-items: center;
      gap: 0.85rem;
    }

    .brand-icon {
      width: 48px;
      height: 48px;
      background: linear-gradient(135deg, #0284c7, #38bdf8);
      border-radius: 12px;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 1.5rem;
      box-shadow: 0 0 20px var(--accent-glow);
    }

    .brand h1 {
      font-size: 1.5rem;
      font-weight: 800;
      background: linear-gradient(to right, #ffffff, #94a3b8);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
    }

    .brand p {
      font-size: 0.85rem;
      color: var(--text-muted);
    }

    .header-actions {
      display: flex;
      align-items: center;
      flex-wrap: wrap;
      gap: 0.75rem;
    }

    .btn {
      display: inline-flex;
      align-items: center;
      gap: 0.5rem;
      padding: 0.65rem 1.25rem;
      border-radius: 10px;
      font-size: 0.88rem;
      font-weight: 600;
      cursor: pointer;
      text-decoration: none;
      transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
      border: 1px solid transparent;
      user-select: none;
    }

    .btn-sm {
      padding: 0.45rem 0.85rem;
      font-size: 0.82rem;
      border-radius: 8px;
    }

    .btn-primary {
      background: linear-gradient(135deg, #0284c7, #0ea5e9);
      color: #fff;
      box-shadow: 0 4px 14px rgba(14, 165, 233, 0.3);
    }

    .btn-primary:hover {
      transform: translateY(-1px);
      box-shadow: 0 6px 20px rgba(14, 165, 233, 0.45);
    }

    .btn-secondary {
      background: rgba(255, 255, 255, 0.05);
      border-color: var(--card-border);
      color: var(--text);
    }

    .btn-secondary:hover {
      background: rgba(255, 255, 255, 0.1);
      border-color: rgba(255, 255, 255, 0.2);
    }

    .btn-danger {
      background: rgba(239, 68, 68, 0.12);
      border-color: rgba(239, 68, 68, 0.25);
      color: #f87171;
    }

    .btn-danger:hover {
      background: rgba(239, 68, 68, 0.25);
      border-color: rgba(239, 68, 68, 0.5);
      color: #fff;
      box-shadow: 0 0 15px rgba(239, 68, 68, 0.25);
    }

    .btn-delete {
      background: transparent;
      border: 1px solid transparent;
      color: var(--text-muted);
      padding: 0.35rem 0.55rem;
      border-radius: 8px;
      cursor: pointer;
      font-size: 0.95rem;
      transition: all 0.2s;
    }

    .btn-delete:hover {
      background: rgba(239, 68, 68, 0.18);
      color: #ef4444;
      border-color: rgba(239, 68, 68, 0.35);
      transform: scale(1.1);
    }

    /* Live Scraper Status Banner */
    .status-banner {
      background: rgba(15, 23, 42, 0.85);
      border: 1px solid rgba(56, 189, 248, 0.25);
      border-radius: 16px;
      padding: 1rem 1.4rem;
      margin-bottom: 1.5rem;
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      gap: 1rem;
      backdrop-filter: blur(14px);
      box-shadow: 0 10px 25px rgba(0, 0, 0, 0.25);
    }

    .status-left {
      display: flex;
      align-items: center;
      gap: 1.1rem;
    }

    .status-pill {
      display: inline-flex;
      align-items: center;
      gap: 0.5rem;
      padding: 0.45rem 0.9rem;
      border-radius: 9999px;
      font-size: 0.8rem;
      font-weight: 800;
      letter-spacing: 0.05em;
      text-transform: uppercase;
      background: rgba(255, 255, 255, 0.06);
      border: 1px solid rgba(255, 255, 255, 0.12);
      color: var(--text-muted);
    }

    .status-pill.running {
      background: rgba(16, 185, 129, 0.18);
      border-color: rgba(16, 185, 129, 0.4);
      color: #34d399;
      box-shadow: 0 0 16px rgba(16, 185, 129, 0.2);
    }

    .status-pill.paused {
      background: rgba(245, 158, 11, 0.18);
      border-color: rgba(245, 158, 11, 0.4);
      color: #fbbf24;
      box-shadow: 0 0 16px rgba(245, 158, 11, 0.2);
    }

    .status-pill.completed {
      background: rgba(16, 185, 129, 0.22);
      border-color: rgba(52, 211, 153, 0.6);
      color: #34d399;
      box-shadow: 0 0 18px rgba(16, 185, 129, 0.35);
    }

    .status-pulse {
      width: 9px;
      height: 9px;
      border-radius: 50%;
      background: currentColor;
    }

    .status-pill.running .status-pulse {
      animation: pulse-glow 1.5s infinite;
    }

    @keyframes pulse-glow {
      0% { transform: scale(0.95); opacity: 0.8; }
      50% { transform: scale(1.4); opacity: 1; }
      100% { transform: scale(0.95); opacity: 0.8; }
    }

    .status-meta {
      display: flex;
      flex-direction: column;
      gap: 0.25rem;
    }

    .status-title {
      font-size: 1rem;
      font-weight: 700;
      color: #fff;
    }

    .status-sub {
      font-size: 0.82rem;
      color: var(--text-muted);
    }

    .status-right {
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }

    /* Cards */
    .card {
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 16px;
      backdrop-filter: blur(14px);
      box-shadow: 0 12px 30px rgba(0, 0, 0, 0.3);
      overflow: hidden;
      margin-bottom: 1.5rem;
    }

    /* ========================================================
       SECTIONS BAR (Preserved Search Batches)
       ======================================================== */
    .sections-container {
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 16px;
      backdrop-filter: blur(14px);
      padding: 1.25rem 1.5rem;
      margin-bottom: 1.5rem;
      box-shadow: 0 10px 25px rgba(0, 0, 0, 0.25);
    }

    .sections-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 1rem;
      margin-bottom: 1rem;
      padding-bottom: 0.85rem;
      border-bottom: 1px solid rgba(255, 255, 255, 0.06);
    }

    .sections-title-wrap {
      display: flex;
      align-items: center;
      gap: 0.65rem;
    }

    .sections-title-wrap h3 {
      font-size: 1.05rem;
      font-weight: 800;
      color: #fff;
    }

    .sections-title-wrap p {
      font-size: 0.8rem;
      color: var(--text-muted);
    }

    .sections-actions {
      display: flex;
      align-items: center;
      gap: 0.65rem;
    }

    .sections-nav {
      display: flex;
      align-items: center;
      gap: 0.65rem;
      overflow-x: auto;
      padding-bottom: 0.35rem;
      scrollbar-width: thin;
    }

    .sections-nav::-webkit-scrollbar {
      height: 6px;
    }
    .sections-nav::-webkit-scrollbar-thumb {
      background: rgba(255, 255, 255, 0.15);
      border-radius: 999px;
    }

    .section-tab {
      display: inline-flex;
      align-items: center;
      gap: 0.6rem;
      padding: 0.65rem 1.15rem;
      border-radius: 12px;
      font-size: 0.86rem;
      font-weight: 700;
      cursor: pointer;
      background: rgba(255, 255, 255, 0.04);
      border: 1px solid var(--card-border);
      color: var(--text-muted);
      transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
      white-space: nowrap;
      user-select: none;
    }

    .section-tab:hover {
      background: rgba(56, 189, 248, 0.1);
      border-color: rgba(56, 189, 248, 0.35);
      color: #fff;
      transform: translateY(-1px);
    }

    .section-tab.active {
      background: linear-gradient(135deg, rgba(2, 132, 199, 0.35), rgba(14, 165, 233, 0.2));
      border-color: #38bdf8;
      color: #fff;
      box-shadow: 0 0 16px rgba(56, 189, 248, 0.25);
    }

    .tab-badge {
      background: rgba(255, 255, 255, 0.1);
      padding: 0.15rem 0.55rem;
      border-radius: 9999px;
      font-size: 0.75rem;
      font-weight: 800;
      color: #fff;
    }

    .section-tab.active .tab-badge {
      background: #0284c7;
      color: #fff;
    }

    /* ========================================================
       CATEGORY ANALYSIS & 3-4 SUB-CATEGORIES SECTION
       ======================================================== */
    .categories-panel {
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 16px;
      backdrop-filter: blur(14px);
      padding: 1.25rem 1.5rem;
      margin-bottom: 1.5rem;
      box-shadow: 0 10px 25px rgba(0, 0, 0, 0.25);
    }

    .categories-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 1rem;
      margin-bottom: 1.15rem;
    }

    .categories-title-wrap {
      display: flex;
      align-items: center;
      gap: 0.65rem;
    }

    .categories-title-wrap h4 {
      font-size: 1.05rem;
      font-weight: 800;
      color: #fff;
    }

    .categories-title-wrap p {
      font-size: 0.8rem;
      color: var(--text-muted);
    }

    .categories-filter-wrap {
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }

    .categories-filter-wrap label {
      font-size: 0.82rem;
      font-weight: 600;
      color: var(--text-muted);
    }

    .category-cards-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
      gap: 1rem;
    }

    .cat-card {
      background: rgba(255, 255, 255, 0.03);
      border: 1px solid var(--card-border);
      border-radius: 14px;
      padding: 1rem 1.15rem;
      cursor: pointer;
      transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
      display: flex;
      flex-direction: column;
      gap: 0.6rem;
      user-select: none;
      position: relative;
    }

    .cat-card:hover {
      background: rgba(56, 189, 248, 0.08);
      border-color: rgba(56, 189, 248, 0.35);
      transform: translateY(-2px);
    }

    .cat-card.active {
      background: linear-gradient(135deg, rgba(14, 165, 233, 0.2), rgba(2, 132, 199, 0.1));
      border-color: #38bdf8;
      box-shadow: 0 0 16px rgba(56, 189, 248, 0.25);
    }

    .cat-card-top {
      display: flex;
      align-items: center;
      justify-content: space-between;
    }

    .cat-icon-title {
      display: flex;
      align-items: center;
      gap: 0.55rem;
    }

    .cat-icon {
      font-size: 1.25rem;
    }

    .cat-title {
      font-size: 0.9rem;
      font-weight: 700;
      color: #fff;
    }

    .cat-count {
      font-size: 0.82rem;
      font-weight: 800;
      background: rgba(255, 255, 255, 0.08);
      padding: 0.2rem 0.55rem;
      border-radius: 999px;
      color: #38bdf8;
    }

    .cat-card.active .cat-count {
      background: #0284c7;
      color: #fff;
    }

    .cat-progress-bg {
      width: 100%;
      height: 5px;
      background: rgba(255, 255, 255, 0.06);
      border-radius: 999px;
      overflow: hidden;
    }

    .cat-progress-fill {
      height: 100%;
      background: linear-gradient(90deg, #38bdf8, #818cf8);
      border-radius: 999px;
      transition: width 0.4s ease;
    }

    .cat-subtext {
      font-size: 0.75rem;
      color: var(--text-muted);
      display: flex;
      justify-content: space-between;
    }

    /* Stats Grid */
    .stats-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 1.25rem;
      margin-bottom: 1.75rem;
    }

    .stat-card {
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      backdrop-filter: blur(12px);
      padding: 1.25rem;
      border-radius: 16px;
      display: flex;
      align-items: center;
      gap: 1rem;
    }

    .stat-icon {
      width: 44px;
      height: 44px;
      border-radius: 12px;
      background: rgba(56, 189, 248, 0.1);
      color: var(--accent);
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 1.3rem;
    }

    .stat-val {
      font-size: 1.6rem;
      font-weight: 800;
      line-height: 1.2;
    }

    .stat-label {
      font-size: 0.8rem;
      color: var(--text-muted);
      font-weight: 600;
    }

    /* Table & Controls */
    .card-header {
      padding: 1.25rem 1.5rem;
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      gap: 1rem;
      border-bottom: 1px solid var(--card-border);
      background: rgba(10, 16, 28, 0.35);
    }

    .search-box {
      position: relative;
      flex: 1;
      max-width: 420px;
      min-width: 240px;
    }

    .search-box input {
      width: 100%;
      background: rgba(0, 0, 0, 0.3);
      border: 1px solid var(--card-border);
      padding: 0.65rem 1rem 0.65rem 2.4rem;
      border-radius: 10px;
      color: var(--text);
      font-size: 0.88rem;
      outline: none;
      transition: all 0.2s;
    }

    .search-box input:focus {
      border-color: var(--accent);
      box-shadow: 0 0 12px var(--accent-glow);
    }

    .search-icon {
      position: absolute;
      left: 0.85rem;
      top: 50%;
      transform: translateY(-50%);
      font-size: 0.9rem;
      color: var(--text-muted);
      pointer-events: none;
    }

    .dropdown {
      background: rgba(15, 23, 42, 0.9);
      border: 1px solid var(--card-border);
      color: var(--text);
      padding: 0.55rem 0.85rem;
      border-radius: 8px;
      font-size: 0.84rem;
      outline: none;
      cursor: pointer;
    }

    .dropdown:focus {
      border-color: var(--accent);
    }

    .table-container {
      overflow-x: auto;
      max-height: 65vh;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      text-align: left;
      font-size: 0.88rem;
    }

    thead {
      background: rgba(15, 23, 42, 0.95);
      position: sticky;
      top: 0;
      z-index: 10;
    }

    th {
      padding: 1rem 1.15rem;
      color: var(--text-muted);
      font-weight: 700;
      font-size: 0.78rem;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      border-bottom: 1px solid var(--card-border);
      white-space: nowrap;
    }

    td {
      padding: 0.9rem 1.15rem;
      border-bottom: 1px solid rgba(255, 255, 255, 0.04);
      vertical-align: middle;
      transition: background 0.15s ease;
    }

    tbody tr:hover {
      background: var(--hover-bg);
    }

    tbody tr.selected-row {
      background: rgba(56, 189, 248, 0.08) !important;
    }

    .row-deleting {
      opacity: 0;
      transform: translateX(-20px);
      transition: all 0.25s ease;
    }

    .lead-name {
      display: flex;
      flex-direction: column;
      gap: 0.2rem;
    }

    .lead-name span:first-child {
      font-weight: 700;
      color: #fff;
    }

    .lead-addr {
      font-size: 0.78rem;
      color: var(--text-muted);
      max-width: 320px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .badge {
      display: inline-block;
      padding: 0.25rem 0.65rem;
      border-radius: 9999px;
      font-size: 0.75rem;
      font-weight: 700;
      background: rgba(255, 255, 255, 0.06);
      color: var(--text-muted);
      border: 1px solid rgba(255, 255, 255, 0.08);
      white-space: nowrap;
    }

    .badge-subcat {
      background: rgba(56, 189, 248, 0.12);
      border-color: rgba(56, 189, 248, 0.3);
      color: #38bdf8;
    }

    .phone-link {
      color: #34d399;
      text-decoration: none;
      font-weight: 600;
      display: inline-flex;
      align-items: center;
      gap: 0.3rem;
      white-space: nowrap;
    }

    .phone-link:hover {
      text-decoration: underline;
    }

    .map-link {
      display: inline-flex;
      align-items: center;
      gap: 0.4rem;
      background: rgba(56, 189, 248, 0.1);
      border: 1px solid rgba(56, 189, 248, 0.25);
      color: #38bdf8;
      padding: 0.35rem 0.65rem;
      border-radius: 8px;
      font-size: 0.78rem;
      font-weight: 600;
      text-decoration: none;
      transition: all 0.2s;
      white-space: nowrap;
    }

    .map-link:hover {
      background: rgba(56, 189, 248, 0.2);
      border-color: #38bdf8;
    }

    /* Custom Checkbox */
    .checkbox-custom {
      width: 17px;
      height: 17px;
      cursor: pointer;
      accent-color: #0284c7;
      vertical-align: middle;
    }

    /* Floating Bulk Actions Toolbar */
    .bulk-bar {
      position: fixed;
      bottom: 2rem;
      left: 50%;
      transform: translateX(-50%) translateY(120px);
      z-index: 999;
      background: rgba(15, 23, 42, 0.95);
      border: 1px solid rgba(56, 189, 248, 0.4);
      padding: 0.75rem 1.4rem;
      border-radius: 9999px;
      box-shadow: 0 12px 40px rgba(0, 0, 0, 0.75), 0 0 20px rgba(56, 189, 248, 0.25);
      display: flex;
      align-items: center;
      gap: 1rem;
      backdrop-filter: blur(16px);
      transition: transform 0.3s cubic-bezier(0.34, 1.56, 0.64, 1);
      pointer-events: none;
    }

    .bulk-bar.visible {
      transform: translateX(-50%) translateY(0);
      pointer-events: auto;
    }

    .bulk-count {
      font-size: 0.88rem;
      font-weight: 800;
      color: #fff;
      display: flex;
      align-items: center;
      gap: 0.4rem;
    }

    .bulk-badge {
      background: #0284c7;
      color: #fff;
      padding: 0.2rem 0.6rem;
      border-radius: 9999px;
      font-size: 0.8rem;
    }

    /* Pagination */
    .pagination-footer {
      padding: 1.15rem 1.5rem;
      border-top: 1px solid var(--card-border);
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      gap: 1rem;
      background: rgba(10, 16, 28, 0.45);
    }

    .pagination-info {
      font-size: 0.85rem;
      color: var(--text-muted);
    }

    .pagination-pages {
      display: flex;
      align-items: center;
      gap: 0.35rem;
      flex-wrap: wrap;
    }

    .page-btn {
      background: rgba(255, 255, 255, 0.05);
      border: 1px solid var(--card-border);
      color: var(--text);
      padding: 0.4rem 0.75rem;
      border-radius: 8px;
      font-size: 0.85rem;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.2s;
      min-width: 36px;
      text-align: center;
    }

    .page-btn:hover:not(:disabled) {
      background: rgba(56, 189, 248, 0.15);
      border-color: rgba(56, 189, 248, 0.4);
      color: #38bdf8;
      transform: translateY(-1px);
    }

    .page-btn.active {
      background: linear-gradient(135deg, #0284c7, #0ea5e9);
      border-color: transparent;
      color: #fff;
      box-shadow: 0 0 12px rgba(14, 165, 233, 0.4);
    }

    .page-btn:disabled {
      opacity: 0.35;
      cursor: not-allowed;
    }

    .empty-state {
      padding: 4rem 2rem;
      text-align: center;
      color: var(--text-muted);
    }

    .toast {
      position: fixed;
      bottom: 2rem;
      right: 2rem;
      background: rgba(15, 23, 42, 0.95);
      border: 1px solid var(--card-border);
      border-radius: 12px;
      padding: 0.85rem 1.25rem;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.5);
      display: flex;
      align-items: center;
      gap: 0.75rem;
      transform: translateY(120px);
      transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      z-index: 1000;
      font-size: 0.9rem;
      font-weight: 600;
    }

    .toast.show {
      transform: translateY(0);
    }

    /* Search Completion Celebration Modal */
    .modal-overlay {
      position: fixed;
      top: 0; left: 0; right: 0; bottom: 0;
      background: rgba(4, 9, 20, 0.82);
      backdrop-filter: blur(10px);
      z-index: 9999;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 1.5rem;
      opacity: 0;
      pointer-events: none;
      transition: opacity 0.3s cubic-bezier(0.16, 1, 0.3, 1);
    }

    .modal-overlay.active {
      opacity: 1;
      pointer-events: auto;
    }

    .modal-card {
      background: linear-gradient(145deg, #111a2e, #0a0f1d);
      border: 1px solid rgba(56, 189, 248, 0.35);
      border-radius: 20px;
      max-width: 520px;
      width: 100%;
      padding: 2rem;
      box-shadow: 0 25px 60px rgba(0, 0, 0, 0.85), 0 0 40px rgba(16, 185, 129, 0.2);
      transform: scale(0.92) translateY(20px);
      transition: all 0.3s cubic-bezier(0.34, 1.56, 0.64, 1);
      position: relative;
    }

    .modal-overlay.active .modal-card {
      transform: scale(1) translateY(0);
    }

    .modal-badge-icon {
      width: 62px;
      height: 62px;
      border-radius: 16px;
      background: linear-gradient(135deg, rgba(16, 185, 129, 0.25), rgba(56, 189, 248, 0.25));
      border: 1px solid rgba(16, 185, 129, 0.5);
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 2rem;
      margin-bottom: 1.25rem;
      box-shadow: 0 0 25px rgba(16, 185, 129, 0.3);
    }

    .modal-title {
      font-size: 1.45rem;
      font-weight: 800;
      color: #fff;
      margin-bottom: 0.35rem;
    }

    .modal-subtitle {
      font-size: 0.88rem;
      color: var(--text-muted);
      margin-bottom: 1.4rem;
      line-height: 1.4;
    }

    .modal-stats-grid {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 0.85rem;
      margin-bottom: 1.35rem;
    }

    .modal-stat-box {
      background: rgba(255, 255, 255, 0.04);
      border: 1px solid var(--card-border);
      border-radius: 12px;
      padding: 0.9rem 0.6rem;
      text-align: center;
    }

    .modal-stat-val {
      font-size: 1.35rem;
      font-weight: 800;
      color: #38bdf8;
    }

    .modal-stat-val.green {
      color: #34d399;
    }

    .modal-stat-lbl {
      font-size: 0.72rem;
      font-weight: 700;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.05em;
      margin-top: 0.25rem;
    }

    .modal-notice {
      background: rgba(16, 185, 129, 0.08);
      border: 1px solid rgba(16, 185, 129, 0.25);
      border-radius: 12px;
      padding: 0.85rem 1rem;
      font-size: 0.84rem;
      color: #cbd5e1;
      margin-bottom: 1.5rem;
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }

    .modal-actions {
      display: flex;
      gap: 0.75rem;
      justify-content: flex-end;
      flex-wrap: wrap;
    }

    .modal-close-x {
      position: absolute;
      top: 1.25rem;
      right: 1.25rem;
      background: transparent;
      border: none;
      color: var(--text-muted);
      font-size: 1.3rem;
      cursor: pointer;
      line-height: 1;
      padding: 0.25rem;
    }
    .modal-close-x:hover {
      color: #fff;
    }
  </style>
</head>
<body>
  <div class="container">
    <header>
      <div class="brand">
        <div class="brand-icon">🗺️</div>
        <div>
          <h1>Google Maps Leads Extractor</h1>
          <p>Real-time Business Contact & Lead Intelligence System</p>
        </div>
      </div>
      <div class="header-actions">
        <button id="btn-sound-toggle" class="btn btn-secondary" onclick="toggleSoundNotification()" title="Toggle Completion Audio Chime">🔔 Sound: ON</button>
        <button id="btn-notify-toggle" class="btn btn-secondary" onclick="toggleDesktopNotifications()" title="Enable Windows Desktop Push Notifications">💻 Desktop Alerts</button>
        <button class="btn btn-secondary" onclick="loadLeads()">🔄 Refresh</button>
        <button class="btn btn-danger" onclick="clearAllLeads()">🗑️ Clear All Leads</button>
        <a href="/api/export" class="btn btn-primary" download="leads_export.xlsx">📊 Download All Excel</a>
      </div>
    </header>

    <!-- Live Scraper Status Banner -->
    <div class="status-banner" id="status-banner">
      <div class="status-left">
        <div class="status-pill" id="status-pill">
          <span class="status-pulse"></span>
          <span id="status-badge-text">CHECKING...</span>
        </div>
        <div class="status-meta">
          <div class="status-title" id="status-phase-text">Checking Scraper Status</div>
          <div class="status-sub" id="status-detail-text">Waiting for scraper activity...</div>
        </div>
      </div>
      <div class="status-right">
        <button id="btn-pause" class="btn btn-secondary" onclick="togglePause()" style="display: none; padding: 0.55rem 1rem;">
          <span id="btn-pause-icon">⏸️</span> <span id="btn-pause-text">Pause</span>
        </button>
        <button id="btn-stop" class="btn btn-danger" onclick="stopScraper()" style="display: none; padding: 0.55rem 1rem;">
          <span>⏹️</span> <span>Stop</span>
        </button>
        <a href="/api/export" id="status-download-btn" class="btn btn-primary" download="leads_export.xlsx" style="display: none; padding: 0.55rem 1rem;">
          <span>📊 Download Latest Excel</span>
        </a>
        <a href="#" target="_blank" id="status-maps-link" class="btn btn-secondary" style="display: none; padding: 0.55rem 1rem; color: #38bdf8;">
          <span>🗺️ Open Maps Search Tab ↗</span>
        </a>
      </div>
    </div>

    <!-- Launch New Extraction Card -->
    <div class="card" style="padding: 1.25rem 1.5rem; margin-bottom: 1.5rem;">
      <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 1rem; flex-wrap: wrap; gap: 0.75rem;">
        <div style="display: flex; align-items: center; gap: 0.6rem;">
          <span style="font-size: 1.3rem;">🎯</span>
          <div>
            <h2 style="font-size: 1.15rem; font-weight: 800; color: #fff;">Start New Lead Extraction</h2>
            <p style="font-size: 0.8rem; color: var(--text-muted);">Enter city and keywords. Past leads are automatically preserved in separate sections below!</p>
          </div>
        </div>
      </div>

      <!-- Search Inputs Grid -->
      <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)) 170px; gap: 1rem; align-items: flex-end;">
        <div>
          <label style="display: block; font-size: 0.8rem; font-weight: 600; color: var(--text-muted); margin-bottom: 0.35rem;">Target City</label>
          <input type="text" id="input-city" placeholder="e.g. Nanded or Mumbai" value="Nanded" style="width: 100%; background: rgba(0,0,0,0.35); border: 1px solid var(--card-border); padding: 0.65rem 0.9rem; border-radius: 10px; color: #fff; font-size: 0.9rem; outline: none;">
        </div>

        <div>
          <label style="display: block; font-size: 0.8rem; font-weight: 600; color: var(--text-muted); margin-bottom: 0.35rem;">Keywords (comma-separated)</label>
          <input type="text" id="input-keywords" placeholder="e.g. hospitals, clinics, cafes" style="width: 100%; background: rgba(0,0,0,0.35); border: 1px solid var(--card-border); padding: 0.65rem 0.9rem; border-radius: 10px; color: #fff; font-size: 0.9rem; outline: none;">
        </div>

        <div>
          <label style="display: block; font-size: 0.8rem; font-weight: 600; color: var(--text-muted); margin-bottom: 0.35rem;">Scan Mode</label>
          <select id="input-mode" class="dropdown" style="width: 100%; padding: 0.65rem 0.9rem;">
            <option value="4.0" selected>⚡ Fast Search (~20 min)</option>
            <option value="2.5">🔍 Deep Grid (~1 hr)</option>
          </select>
        </div>

        <div>
          <label style="display: block; font-size: 0.8rem; font-weight: 600; color: var(--text-muted); margin-bottom: 0.35rem;">Lead Limit (optional)</label>
          <input type="number" id="input-limit" placeholder="No limit" style="width: 100%; background: rgba(0,0,0,0.35); border: 1px solid var(--card-border); padding: 0.65rem 0.9rem; border-radius: 10px; color: #fff; font-size: 0.9rem; outline: none;">
        </div>

        <div>
          <button id="btn-start-scraper" class="btn btn-primary" onclick="startNewExtraction()" style="width: 100%; justify-content: center; padding: 0.65rem 1rem;">
            🚀 Launch Scraper
          </button>
        </div>
      </div>
    </div>

    <!-- ========================================================
         SECTIONS BAR (Preserved Search Batches)
         ======================================================== -->
    <div class="sections-container">
      <div class="sections-header">
        <div class="sections-title-wrap">
          <span style="font-size: 1.3rem;">📁</span>
          <div>
            <h3>Lead Search Sections</h3>
            <p>Switch between your past search batches without losing or mixing leads</p>
          </div>
        </div>
        <div class="sections-actions">
          <button class="btn btn-secondary btn-sm" id="btn-export-section" onclick="exportCurrentSection()">
            📊 Export This Section (.xlsx)
          </button>
          <button class="btn btn-danger btn-sm" id="btn-delete-section" onclick="deleteCurrentSection()" style="display: none;">
            🗑️ Delete This Section
          </button>
        </div>
      </div>
      <div class="sections-nav" id="sections-nav">
        <!-- Dynamic search tabs -->
      </div>
    </div>

    <!-- ========================================================
         CATEGORY ANALYSIS & 3-4 DYNAMIC SUB-CATEGORIES SECTION
         ======================================================== -->
    <div class="categories-panel">
      <div class="categories-header">
        <div class="categories-title-wrap">
          <span style="font-size: 1.3rem;">🏷️</span>
          <div>
            <h4 id="categories-heading">AI Category Analysis & Breakdown</h4>
            <p id="categories-subtext">Intelligently organized into 3-4 key industry categories for your leads</p>
          </div>
        </div>
        <div class="categories-filter-wrap">
          <label for="category-dropdown-select">Categories Dropdown:</label>
          <select id="category-dropdown-select" class="dropdown" onchange="onCategoryDropdownChange(this.value)">
            <!-- Populated dynamically -->
          </select>
        </div>
      </div>
      <div class="category-cards-grid" id="category-cards-grid">
        <!-- 3-4 interactive cards populated dynamically -->
      </div>
    </div>

    <!-- Stats Grid -->
    <div class="stats-grid">
      <div class="stat-card">
        <div class="stat-icon">🏢</div>
        <div>
          <div class="stat-val" id="stat-total">0</div>
          <div class="stat-label">Total Leads (Active Section)</div>
        </div>
      </div>
      <div class="stat-card">
        <div class="stat-icon" style="color: #34d399; background: rgba(52, 211, 153, 0.1);">📞</div>
        <div>
          <div class="stat-val" id="stat-phones">0</div>
          <div class="stat-label">With Phone Number</div>
        </div>
      </div>
      <div class="stat-card">
        <div class="stat-icon" style="color: #f59e0b; background: rgba(245, 158, 11, 0.1);">📍</div>
        <div>
          <div class="stat-val" id="stat-cids">0</div>
          <div class="stat-label">Decimal CID Links</div>
        </div>
      </div>
      <div class="stat-card">
        <div class="stat-icon" style="color: #a855f7; background: rgba(168, 85, 247, 0.1);">⚡</div>
        <div>
          <div class="stat-val" id="stat-contacted">0</div>
          <div class="stat-label">Answered / Actioned</div>
        </div>
      </div>
    </div>

    <!-- Leads Table Card -->
    <div class="card">
      <div class="card-header">
        <div class="search-box">
          <span class="search-icon">🔍</span>
          <input type="text" id="search-input" placeholder="Search by name, phone, or category..." oninput="filterTable()">
        </div>
        <div style="display: flex; align-items: center; gap: 1.25rem; flex-wrap: wrap;">
          <div style="display: flex; align-items: center; gap: 0.5rem; font-size: 0.85rem; color: var(--text-muted);">
            <label for="page-size-select">Per page:</label>
            <select id="page-size-select" class="dropdown" style="padding: 0.35rem 0.65rem;" onchange="changePageSize(this.value)">
              <option value="10">10</option>
              <option value="20" selected>20</option>
              <option value="50">50</option>
              <option value="100">100</option>
              <option value="all">All</option>
            </select>
          </div>
          <div style="font-size: 0.85rem; color: var(--text-muted);">
            Showing <span id="visible-count" style="color: #fff; font-weight: 700;">0</span> leads
          </div>
        </div>
      </div>

      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th style="width: 44px; text-align: center;">
                <input type="checkbox" id="select-all-header" class="checkbox-custom" onchange="toggleSelectAll(this.checked)" title="Select all on this page">
              </th>
              <th>Business Details</th>
              <th>Phone</th>
              <th>Category</th>
              <th>Subcategory Type</th>
              <th>Google Maps CID</th>
              <th>Call Status</th>
              <th>Interest</th>
              <th>Send</th>
              <th>Quality</th>
              <th style="text-align: center; width: 60px;">Action</th>
            </tr>
          </thead>
          <tbody id="leads-body">
            <tr>
              <td colspan="11" class="empty-state">Loading leads data...</td>
            </tr>
          </tbody>
        </table>
      </div>

      <div class="pagination-footer" id="pagination-footer">
        <div class="pagination-info" id="pagination-info">Showing 0 to 0 of 0 leads</div>
        <div class="pagination-pages" id="pagination-pages"></div>
      </div>
    </div>
  </div>

  <!-- Floating Bulk Actions Toolbar -->
  <div class="bulk-bar" id="bulk-bar">
    <div class="bulk-count">
      <span>Selected:</span>
      <span class="bulk-badge" id="bulk-selected-count">0</span>
    </div>
    <button class="btn btn-danger btn-sm" onclick="bulkDeleteSelected()">
      🗑️ Delete Selected
    </button>
    <button class="btn btn-primary btn-sm" onclick="bulkExportSelected()">
      📊 Export Selected (.xlsx)
    </button>
    <button class="btn btn-secondary btn-sm" onclick="bulkCopyPhones()">
      📋 Copy Phones
    </button>
    <button class="btn btn-secondary btn-sm" onclick="clearSelection()">
      ✕ Clear
    </button>
  </div>

  <!-- Toast Notification -->
  <div id="toast" class="toast">
    <span>✅</span> <span id="toast-msg">Status updated!</span>
  </div>

  <script>
    // State
    let allLeads = [];
    let activeSectionKey = 'all'; // 'all' or 'keyword__city'
    let activeCategoryFilter = 'all'; // 'all' or subcategory cluster key or exact category name
    let selectedCids = new Set();
    let currentSectionLeads = [];
    let filteredLeads = [];
    let currentPage = 1;
    let pageSize = 20;
    let lastLeadsCount = -1;
    let isRunningState = false;
    let isPausedState = false;

    // Load Leads with Cache Busting
    async function loadLeads(preservePage = false) {
      try {
        const res = await fetch('/api/leads?_t=' + Date.now());
        if (!res.ok) throw new Error('Failed to fetch leads');
        allLeads = await res.json();
        
        // Clean selection of any removed leads
        const existingCids = new Set(allLeads.map(l => String(l.cid)));
        selectedCids = new Set([...selectedCids].filter(cid => existingCids.has(cid)));
        updateBulkBar();

        renderSectionsNav();
        applySectionAndFilters(preservePage);
      } catch (err) {
        console.error('Failed to fetch leads', err);
      }
    }

    // ========================================================
    // SECTIONS ENGINE (Search Batches)
    // ========================================================
    function getSectionsMap() {
      const map = new Map();
      allLeads.forEach(l => {
        const kw = (l.keyword || 'General').trim();
        const city = (l.city || 'All').trim();
        const key = `${kw}__${city}`;
        if (!map.has(key)) {
          map.set(key, { keyword: kw, city: city, count: 0, leads: [] });
        }
        const item = map.get(key);
        item.count++;
        item.leads.push(l);
      });
      return map;
    }

    function getSectionIcon(keyword) {
      const kw = (keyword || '').toLowerCase();
      if (kw.includes('hosp') || kw.includes('medic') || kw.includes('doctor')) return '🏥';
      if (kw.includes('dent')) return '🦷';
      if (kw.includes('eye') || kw.includes('opt')) return '👁️';
      if (kw.includes('cafe') || kw.includes('coffee') || kw.includes('tea')) return '☕';
      if (kw.includes('rest') || kw.includes('food') || kw.includes('diner') || kw.includes('pizza') || kw.includes('burger')) return '🍔';
      if (kw.includes('school') || kw.includes('college') || kw.includes('coach')) return '🎓';
      if (kw.includes('hotel') || kw.includes('resort')) return '🏨';
      if (kw.includes('gym') || kw.includes('fit')) return '💪';
      if (kw.includes('car') || kw.includes('auto')) return '🚗';
      return '📍';
    }

    function renderSectionsNav() {
      const nav = document.getElementById('sections-nav');
      const sections = getSectionsMap();

      let html = `
        <button class="section-tab ${activeSectionKey === 'all' ? 'active' : ''}" onclick="switchSection('all')">
          <span class="tab-icon">🌐</span>
          <span class="tab-label">All Leads</span>
          <span class="tab-badge">${allLeads.length}</span>
        </button>
      `;

      sections.forEach((item, key) => {
        const icon = getSectionIcon(item.keyword);
        const activeClass = activeSectionKey === key ? 'active' : '';
        html += `
          <button class="section-tab ${activeClass}" onclick="switchSection('${escapeHtml(key)}')">
            <span class="tab-icon">${icon}</span>
            <span class="tab-label">${escapeHtml(item.keyword)} • ${escapeHtml(item.city)}</span>
            <span class="tab-badge">${item.count}</span>
          </button>
        `;
      });

      nav.innerHTML = html;

      // Toggle Delete Section button visibility
      const btnDelSec = document.getElementById('btn-delete-section');
      if (activeSectionKey !== 'all') {
        btnDelSec.style.display = 'inline-flex';
      } else {
        btnDelSec.style.display = 'none';
      }
    }

    function switchSection(key) {
      activeSectionKey = key;
      activeCategoryFilter = 'all'; // Reset category filter when switching section
      currentPage = 1;
      renderSectionsNav();
      applySectionAndFilters(false);
    }

    // ========================================================
    // AI CATEGORY ANALYSIS & 3-4 DYNAMIC SUB-CATEGORIES ENGINE
    // ========================================================
    function analyzeCategories(leads) {
      if (!leads || leads.length === 0) {
        return { domain: 'general', clusters: [] };
      }

      // Detect domain
      let healthScore = 0;
      let foodScore = 0;

      leads.forEach(l => {
        const text = `${l.category || ''} ${l.business_name || ''} ${l.keyword || ''}`.toLowerCase();
        if (text.includes('hosp') || text.includes('clinic') || text.includes('dental') || text.includes('dentist') || text.includes('eye') || text.includes('doctor') || text.includes('surgeon') || text.includes('health') || text.includes('ortho')) {
          healthScore++;
        }
        if (text.includes('cafe') || text.includes('coffee') || text.includes('restaurant') || text.includes('food') || text.includes('sandwich') || text.includes('bakery') || text.includes('fast food') || text.includes('bar') || text.includes('grill') || text.includes('tea')) {
          foodScore++;
        }
      });

      if (healthScore > leads.length * 0.35) {
        return buildHealthcareClusters(leads);
      } else if (foodScore > leads.length * 0.35) {
        return buildFoodClusters(leads);
      } else {
        return buildUniversalClusters(leads);
      }
    }

    function buildHealthcareClusters(leads) {
      const clusters = [
        {
          key: 'general_hospital',
          name: 'General & Multispeciality Hospitals',
          icon: '🏥',
          subtypes: new Set(),
          leads: [],
          matcher: (cat, name) => {
            const t = (cat + ' ' + name).toLowerCase();
            return (t.includes('hospital') || t.includes('multispeciality') || t.includes('general hospital') || t.includes('private hospital') || t.includes('psychiatric') || t.includes('government hospital')) && !t.includes('eye') && !t.includes('dental');
          }
        },
        {
          key: 'dental_oral',
          name: 'Dental & Oral Health Clinics',
          icon: '🦷',
          subtypes: new Set(),
          leads: [],
          matcher: (cat, name) => {
            const t = (cat + ' ' + name).toLowerCase();
            return t.includes('dental') || t.includes('dentist') || t.includes('oral') || t.includes('endodont') || t.includes('pediatric dentist');
          }
        },
        {
          key: 'eye_care',
          name: 'Eye & Ophthalmology Centers',
          icon: '👁️',
          subtypes: new Set(),
          leads: [],
          matcher: (cat, name) => {
            const t = (cat + ' ' + name).toLowerCase();
            return t.includes('eye') || t.includes('ophthalm') || t.includes('vision') || t.includes('optomet');
          }
        },
        {
          key: 'specialist_clinics',
          name: 'Specialist Clinics & Doctors',
          icon: '🩺',
          subtypes: new Set(),
          leads: [],
          matcher: (cat, name) => true // Fallback
        }
      ];

      leads.forEach(l => {
        const cat = l.category || 'Clinic';
        const name = l.business_name || '';
        let matched = false;
        for (let i = 0; i < 3; i++) {
          if (clusters[i].matcher(cat, name)) {
            clusters[i].leads.push(l);
            clusters[i].subtypes.add(cat);
            l._clusterKey = clusters[i].key;
            l._clusterName = clusters[i].name;
            l._clusterIcon = clusters[i].icon;
            matched = true;
            break;
          }
        }
        if (!matched) {
          clusters[3].leads.push(l);
          clusters[3].subtypes.add(cat);
          l._clusterKey = clusters[3].key;
          l._clusterName = clusters[3].name;
          l._clusterIcon = clusters[3].icon;
        }
      });

      return { domain: 'healthcare', clusters };
    }

    function buildFoodClusters(leads) {
      const clusters = [
        {
          key: 'cafes_coffee',
          name: 'Cafes & Coffee Shops',
          icon: '☕',
          subtypes: new Set(),
          leads: [],
          matcher: (cat, name) => {
            const t = (cat + ' ' + name).toLowerCase();
            return t.includes('cafe') || t.includes('coffee') || t.includes('tea') || t.includes('bakery');
          }
        },
        {
          key: 'fast_food',
          name: 'Fast Food & Quick Bites',
          icon: '🍔',
          subtypes: new Set(),
          leads: [],
          matcher: (cat, name) => {
            const t = (cat + ' ' + name).toLowerCase();
            return t.includes('fast food') || t.includes('sandwich') || t.includes('burger') || t.includes('pizza') || t.includes('snack');
          }
        },
        {
          key: 'dining_restaurants',
          name: 'Regional & Family Dining',
          icon: '🍛',
          subtypes: new Set(),
          leads: [],
          matcher: (cat, name) => {
            const t = (cat + ' ' + name).toLowerCase();
            return (t.includes('restaurant') || t.includes('indian') || t.includes('vegetarian') || t.includes('family') || t.includes('chinese') || t.includes('dhaba')) && !t.includes('bar');
          }
        },
        {
          key: 'bars_grills',
          name: 'Bars, Grills & Lounges',
          icon: '🍸',
          subtypes: new Set(),
          leads: [],
          matcher: (cat, name) => true
        }
      ];

      leads.forEach(l => {
        const cat = l.category || 'Restaurant';
        const name = l.business_name || '';
        let matched = false;
        for (let i = 0; i < 3; i++) {
          if (clusters[i].matcher(cat, name)) {
            clusters[i].leads.push(l);
            clusters[i].subtypes.add(cat);
            l._clusterKey = clusters[i].key;
            l._clusterName = clusters[i].name;
            l._clusterIcon = clusters[i].icon;
            matched = true;
            break;
          }
        }
        if (!matched) {
          clusters[3].leads.push(l);
          clusters[3].subtypes.add(cat);
          l._clusterKey = clusters[3].key;
          l._clusterName = clusters[3].name;
          l._clusterIcon = clusters[3].icon;
        }
      });

      return { domain: 'food', clusters };
    }

    function buildUniversalClusters(leads) {
      // Frequency map of exact categories
      const counts = {};
      leads.forEach(l => {
        const c = (l.category || 'General Business').trim();
        counts[c] = (counts[c] || 0) + 1;
      });

      const sorted = Object.keys(counts).sort((a, b) => counts[b] - counts[a]);
      const top3 = sorted.slice(0, 3);
      const top3Set = new Set(top3);

      const icons = ['🏢', '⭐', '📍', '🏷️'];
      const clusters = [
        ...top3.map((cat, idx) => ({
          key: 'cat_' + idx,
          name: cat,
          icon: icons[idx],
          subtypes: new Set([cat]),
          leads: [],
          matcher: (c) => c === cat
        })),
        {
          key: 'cat_other',
          name: 'Other Specialized Businesses',
          icon: icons[3],
          subtypes: new Set(),
          leads: [],
          matcher: () => true
        }
      ];

      leads.forEach(l => {
        const cat = (l.category || 'General Business').trim();
        const topIdx = top3.indexOf(cat);
        if (topIdx !== -1) {
          clusters[topIdx].leads.push(l);
          l._clusterKey = clusters[topIdx].key;
          l._clusterName = clusters[topIdx].name;
          l._clusterIcon = clusters[topIdx].icon;
        } else {
          clusters[clusters.length - 1].leads.push(l);
          clusters[clusters.length - 1].subtypes.add(cat);
          l._clusterKey = clusters[clusters.length - 1].key;
          l._clusterName = clusters[clusters.length - 1].name;
          l._clusterIcon = clusters[clusters.length - 1].icon;
        }
      });

      return { domain: 'universal', clusters };
    }

    function renderCategoryPanel(analysisResult) {
      const container = document.getElementById('category-cards-grid');
      const dropdown = document.getElementById('category-dropdown-select');
      const totalSecLeads = currentSectionLeads.length;

      // Update Subcategory Dropdown
      let dropHtml = `<option value="all">-- All Categories (${totalSecLeads}) --</option>`;

      // Optgroup 1: 3-4 Macro Sections
      dropHtml += `<optgroup label="3-4 Main Category Sections">`;
      analysisResult.clusters.forEach(c => {
        const isSel = activeCategoryFilter === c.key ? 'selected' : '';
        dropHtml += `<option value="${c.key}" ${isSel}>${c.icon} ${c.name} (${c.leads.length})</option>`;
      });
      dropHtml += `</optgroup>`;

      // Optgroup 2: Specific Google Maps Categories
      const allSubtypes = new Map();
      currentSectionLeads.forEach(l => {
        const c = (l.category || 'Uncategorized').trim();
        allSubtypes.set(c, (allSubtypes.get(c) || 0) + 1);
      });
      const sortedSubtypes = [...allSubtypes.entries()].sort((a, b) => b[1] - a[1]);

      dropHtml += `<optgroup label="Specific Google Maps Categories">`;
      sortedSubtypes.forEach(([catName, count]) => {
        const isSel = activeCategoryFilter === `exact:${catName}` ? 'selected' : '';
        dropHtml += `<option value="exact:${escapeHtml(catName)}" ${isSel}>${escapeHtml(catName)} (${count})</option>`;
      });
      dropHtml += `</optgroup>`;

      dropdown.innerHTML = dropHtml;

      // Render 3-4 Category Cards + 1 "All" Card
      let cardsHtml = `
        <div class="cat-card ${activeCategoryFilter === 'all' ? 'active' : ''}" onclick="selectCategoryFilter('all')">
          <div class="cat-card-top">
            <div class="cat-icon-title">
              <span class="cat-icon">🌐</span>
              <span class="cat-title">All Leads</span>
            </div>
            <span class="cat-count">${totalSecLeads}</span>
          </div>
          <div class="cat-progress-bg">
            <div class="cat-progress-fill" style="width: 100%;"></div>
          </div>
          <div class="cat-subtext">
            <span>Entire Active Section</span>
            <span>100%</span>
          </div>
        </div>
      `;

      analysisResult.clusters.forEach(c => {
        const pct = totalSecLeads > 0 ? Math.round((c.leads.length / totalSecLeads) * 100) : 0;
        const isActive = activeCategoryFilter === c.key ? 'active' : '';
        cardsHtml += `
          <div class="cat-card ${isActive}" onclick="selectCategoryFilter('${c.key}')">
            <div class="cat-card-top">
              <div class="cat-icon-title">
                <span class="cat-icon">${c.icon}</span>
                <span class="cat-title">${escapeHtml(c.name)}</span>
              </div>
              <span class="cat-count">${c.leads.length}</span>
            </div>
            <div class="cat-progress-bg">
              <div class="cat-progress-fill" style="width: ${pct}%;"></div>
            </div>
            <div class="cat-subtext">
              <span>${c.subtypes.size} sub-types</span>
              <span>${pct}%</span>
            </div>
          </div>
        `;
      });

      container.innerHTML = cardsHtml;
    }

    function selectCategoryFilter(key) {
      activeCategoryFilter = key;
      currentPage = 1;
      const dropdown = document.getElementById('category-dropdown-select');
      dropdown.value = key;
      applySectionAndFilters(false);
    }

    function onCategoryDropdownChange(val) {
      activeCategoryFilter = val;
      currentPage = 1;
      applySectionAndFilters(false);
    }

    // ========================================================
    // FILTERING & PAGINATION PIPELINE
    // ========================================================
    function applySectionAndFilters(preservePage = false) {
      // Step 1: Filter by Active Section
      if (activeSectionKey === 'all') {
        currentSectionLeads = [...allLeads];
      } else {
        const sections = getSectionsMap();
        const sec = sections.get(activeSectionKey);
        currentSectionLeads = sec ? [...sec.leads] : [];
      }

      // Step 2: Run AI Category Analysis on the current section leads
      const analysis = analyzeCategories(currentSectionLeads);
      renderCategoryPanel(analysis);

      // Step 3: Filter by Active Category / Subcategory
      let stepLeads = currentSectionLeads;
      if (activeCategoryFilter !== 'all') {
        if (activeCategoryFilter.startsWith('exact:')) {
          const targetCat = activeCategoryFilter.replace('exact:', '');
          stepLeads = stepLeads.filter(l => (l.category || '').trim() === targetCat);
        } else {
          stepLeads = stepLeads.filter(l => l._clusterKey === activeCategoryFilter);
        }
      }

      // Step 4: Filter by Text Search
      const q = document.getElementById('search-input').value.toLowerCase().trim();
      if (!q) {
        filteredLeads = stepLeads;
      } else {
        filteredLeads = stepLeads.filter(l => 
          (l.business_name && l.business_name.toLowerCase().includes(q)) ||
          (l.category && l.category.toLowerCase().includes(q)) ||
          (l.phone && l.phone.toLowerCase().includes(q)) ||
          (l.address && l.address.toLowerCase().includes(q)) ||
          (l._clusterName && l._clusterName.toLowerCase().includes(q))
        );
      }

      // Update Section & Active Stats
      updateStats();

      // Pagination
      const totalPages = Math.ceil(filteredLeads.length / pageSize) || 1;
      if (!preservePage || currentPage > totalPages) {
        currentPage = 1;
      }
      renderCurrentPage();
    }

    function filterTable() {
      applySectionAndFilters(false);
    }

    function updateStats() {
      const activeLeads = currentSectionLeads;
      document.getElementById('stat-total').innerText = activeLeads.length;
      document.getElementById('stat-phones').innerText = activeLeads.filter(l => l.phone && l.phone.trim()).length;
      document.getElementById('stat-cids').innerText = activeLeads.filter(l => l.short_cid_url && l.short_cid_url.includes('cid=')).length;
      document.getElementById('stat-contacted').innerText = activeLeads.filter(l => l.call_status === 'answered' || l.interest === 'interested').length;
    }

    function changePageSize(val) {
      if (val === 'all') {
        pageSize = 9999999;
      } else {
        pageSize = parseInt(val, 10) || 20;
      }
      currentPage = 1;
      renderCurrentPage();
    }

    function goToPage(page) {
      const totalPages = Math.ceil(filteredLeads.length / pageSize) || 1;
      if (page < 1) page = 1;
      if (page > totalPages) page = totalPages;
      currentPage = page;
      renderCurrentPage();
    }

    function renderCurrentPage() {
      const tbody = document.getElementById('leads-body');
      const totalCount = filteredLeads.length;
      document.getElementById('visible-count').innerText = totalCount;

      if (!filteredLeads || totalCount === 0) {
        tbody.innerHTML = `<tr><td colspan="11" class="empty-state">No matching leads found in this section.</td></tr>`;
        document.getElementById('pagination-info').innerText = 'Showing 0 of 0 leads';
        document.getElementById('pagination-pages').innerHTML = '';
        return;
      }

      const totalPages = Math.ceil(totalCount / pageSize) || 1;
      if (currentPage > totalPages) currentPage = totalPages;

      const startIndex = (currentPage - 1) * pageSize;
      const endIndex = Math.min(startIndex + pageSize, totalCount);
      const pageLeads = filteredLeads.slice(startIndex, endIndex);

      // Check if all on current page are selected
      const pageCids = pageLeads.map(l => String(l.cid));
      const allPageSelected = pageCids.length > 0 && pageCids.every(cid => selectedCids.has(cid));
      const selectAllHeader = document.getElementById('select-all-header');
      if (selectAllHeader) selectAllHeader.checked = allPageSelected;

      tbody.innerHTML = pageLeads.map(l => {
        const cidStr = String(l.cid);
        const isChecked = selectedCids.has(cidStr);
        const mapsLink = l.short_cid_url || l.google_maps_url || '#';
        const phoneHtml = l.phone 
          ? `<a href="tel:${escapeHtml(l.phone)}" class="phone-link">📞 ${escapeHtml(l.phone)}</a>`
          : `<span style="color: var(--text-muted); font-size: 0.8rem;">None</span>`;

        const clusterBadge = l._clusterName 
          ? `<span class="badge badge-subcat">${l._clusterIcon || '🏷️'} ${escapeHtml(l._clusterName)}</span>`
          : `<span class="badge">General</span>`;

        return `
          <tr id="row-${escapeHtml(cidStr)}" class="${isChecked ? 'selected-row' : ''}">
            <td style="text-align: center;">
              <input type="checkbox" class="checkbox-custom lead-check" data-cid="${escapeHtml(cidStr)}" 
                ${isChecked ? 'checked' : ''} onchange="toggleSelectLead('${escapeHtml(cidStr)}', this.checked)">
            </td>
            <td>
              <div class="lead-name">
                <span>${escapeHtml(l.business_name || 'Unnamed')}</span>
                <span class="lead-addr">${escapeHtml(l.address || l.city || '')}</span>
              </div>
            </td>
            <td>${phoneHtml}</td>
            <td><span class="badge">${escapeHtml(l.category || 'Business')}</span></td>
            <td>${clusterBadge}</td>
            <td>
              <a href="${mapsLink}" target="_blank" class="map-link">
                📍 Open CID Place ↗
              </a>
            </td>
            <td>
              <select class="dropdown" onchange="updateLead('${escapeHtml(cidStr)}', 'call_status', this.value)">
                <option value="" ${!l.call_status ? 'selected' : ''}>-- Status --</option>
                <option value="answered" ${l.call_status === 'answered' ? 'selected' : ''}>Answered</option>
                <option value="unanswered" ${l.call_status === 'unanswered' ? 'selected' : ''}>Unanswered</option>
                <option value="call back" ${l.call_status === 'call back' ? 'selected' : ''}>Call Back</option>
              </select>
            </td>
            <td>
              <select class="dropdown" onchange="updateLead('${escapeHtml(cidStr)}', 'interest', this.value)">
                <option value="" ${!l.interest ? 'selected' : ''}>-- Interest --</option>
                <option value="interested" ${l.interest === 'interested' ? 'selected' : ''}>Interested</option>
                <option value="not interested" ${l.interest === 'not interested' ? 'selected' : ''}>Not Interested</option>
              </select>
            </td>
            <td>
              <select class="dropdown" onchange="updateLead('${escapeHtml(cidStr)}', 'send', this.value)">
                <option value="" ${!l.send ? 'selected' : ''}>-- Send --</option>
                <option value="send" ${l.send === 'send' ? 'selected' : ''}>Send</option>
                <option value="not send" ${l.send === 'not send' ? 'selected' : ''}>Not Send</option>
              </select>
            </td>
            <td>
              <select class="dropdown" onchange="updateLead('${escapeHtml(cidStr)}', 'lead_quality', this.value)">
                <option value="" ${!l.lead_quality ? 'selected' : ''}>-- Quality --</option>
                <option value="quality lead" ${l.lead_quality === 'quality lead' ? 'selected' : ''}>Quality Lead</option>
                <option value="not quality lead" ${l.lead_quality === 'not quality lead' ? 'selected' : ''}>Not Quality Lead</option>
              </select>
            </td>
            <td style="text-align: center;">
              <button class="btn-delete" onclick="deleteLead('${escapeHtml(cidStr)}')" title="Delete this lead">🗑️</button>
            </td>
          </tr>
        `;
      }).join('');

      document.getElementById('pagination-info').innerText = 
        `Showing ${startIndex + 1} to ${endIndex} of ${totalCount} leads • Page ${currentPage} of ${totalPages}`;

      renderPaginationButtons(totalPages);
    }

    function renderPaginationButtons(totalPages) {
      const pagContainer = document.getElementById('pagination-pages');
      if (totalPages <= 1) {
        pagContainer.innerHTML = '';
        return;
      }

      let buttonsHtml = `
        <button class="page-btn" onclick="goToPage(${currentPage - 1})" ${currentPage === 1 ? 'disabled' : ''}>
          ◀ Prev
        </button>
      `;

      let startP = Math.max(1, currentPage - 2);
      let endP = Math.min(totalPages, currentPage + 2);

      if (startP > 1) {
        buttonsHtml += `<button class="page-btn" onclick="goToPage(1)">1</button>`;
        if (startP > 2) buttonsHtml += `<span style="color: var(--text-muted); padding: 0 4px;">…</span>`;
      }

      for (let p = startP; p <= endP; p++) {
        buttonsHtml += `
          <button class="page-btn ${p === currentPage ? 'active' : ''}" onclick="goToPage(${p})">
            ${p}
          </button>
        `;
      }

      if (endP < totalPages) {
        if (endP < totalPages - 1) buttonsHtml += `<span style="color: var(--text-muted); padding: 0 4px;">…</span>`;
        buttonsHtml += `<button class="page-btn" onclick="goToPage(${totalPages})">${totalPages}</button>`;
      }

      buttonsHtml += `
        <button class="page-btn" onclick="goToPage(${currentPage + 1})" ${currentPage === totalPages ? 'disabled' : ''}>
          Next ▶
        </button>
      `;

      pagContainer.innerHTML = buttonsHtml;
    }

    // ========================================================
    // MULTI-SELECT & BULK ACTIONS
    // ========================================================
    function toggleSelectLead(cid, isChecked) {
      const cidStr = String(cid);
      if (isChecked) {
        selectedCids.add(cidStr);
      } else {
        selectedCids.delete(cidStr);
      }
      
      const row = document.getElementById('row-' + cidStr);
      if (row) {
        if (isChecked) row.classList.add('selected-row');
        else row.classList.remove('selected-row');
      }

      updateBulkBar();
    }

    function toggleSelectAll(isChecked) {
      const totalCount = filteredLeads.length;
      const startIndex = (currentPage - 1) * pageSize;
      const endIndex = Math.min(startIndex + pageSize, totalCount);
      const pageLeads = filteredLeads.slice(startIndex, endIndex);

      pageLeads.forEach(l => {
        const cidStr = String(l.cid);
        if (isChecked) selectedCids.add(cidStr);
        else selectedCids.delete(cidStr);
      });

      renderCurrentPage();
      updateBulkBar();
    }

    function clearSelection() {
      selectedCids.clear();
      renderCurrentPage();
      updateBulkBar();
    }

    function updateBulkBar() {
      const bar = document.getElementById('bulk-bar');
      const badge = document.getElementById('bulk-selected-count');
      const count = selectedCids.size;
      badge.innerText = count;
      if (count > 0) {
        bar.classList.add('visible');
      } else {
        bar.classList.remove('visible');
      }
    }

    async function bulkDeleteSelected() {
      const count = selectedCids.size;
      if (count === 0) return;
      if (!confirm(`⚠️ Are you sure you want to delete ${count} selected leads? This cannot be undone.`)) return;

      const cidsArray = Array.from(selectedCids);
      try {
        const res = await fetch('/api/leads/delete-batch', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ cids: cidsArray })
        });
        const data = await res.json();
        if (res.ok && data.success) {
          // Remove from local memory immediately
          const delSet = new Set(cidsArray);
          allLeads = allLeads.filter(l => !delSet.has(String(l.cid)));
          selectedCids.clear();
          updateBulkBar();
          lastLeadsCount = allLeads.length;

          renderSectionsNav();
          applySectionAndFilters(true);
          showToast(`🗑️ Successfully deleted ${data.count || count} leads!`);
        } else {
          alert('Failed to delete leads.');
        }
      } catch (err) {
        console.error('Failed to bulk delete', err);
        alert('Server error occurred during deletion.');
      }
    }

    function bulkExportSelected() {
      if (selectedCids.size === 0) return;
      const cids = Array.from(selectedCids).join(',');
      window.location.href = `/api/export?cids=${encodeURIComponent(cids)}`;
      showToast(`📥 Exporting ${selectedCids.size} selected leads...`);
    }

    function bulkCopyPhones() {
      if (selectedCids.size === 0) return;
      const selLeads = allLeads.filter(l => selectedCids.has(String(l.cid)));
      const phones = selLeads.map(l => l.phone).filter(p => p && p.trim());
      if (phones.length === 0) {
        alert('None of the selected leads have phone numbers.');
        return;
      }
      navigator.clipboard.writeText(phones.join(', '));
      showToast(`📋 Copied ${phones.length} phone numbers to clipboard!`);
    }

    // ========================================================
    // DELETE SINGLE LEAD (INSTANT UPDATE FIX)
    // ========================================================
    async function deleteLead(cid) {
      const cidStr = String(cid);
      if (!confirm('Are you sure you want to delete this lead?')) return;

      try {
        const res = await fetch('/api/leads/delete', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ cid: cidStr })
        });
        const data = await res.json();
        if (res.ok && data.success) {
          // Animate row removal
          const row = document.getElementById('row-' + cidStr);
          if (row) row.classList.add('row-deleting');

          setTimeout(() => {
            // Update state in memory
            allLeads = allLeads.filter(l => String(l.cid) !== cidStr);
            selectedCids.delete(cidStr);
            updateBulkBar();
            lastLeadsCount = allLeads.length;

            renderSectionsNav();
            applySectionAndFilters(true);
            showToast('✅ Lead deleted successfully!');
          }, 200);
        } else {
          alert('Could not delete lead. Please try again.');
        }
      } catch (err) {
        console.error('Failed to delete lead', err);
        alert('Network or server error while deleting lead.');
      }
    }

    // Delete Entire Active Section
    async function deleteCurrentSection() {
      if (activeSectionKey === 'all') return;
      const sections = getSectionsMap();
      const sec = sections.get(activeSectionKey);
      if (!sec) return;

      if (!confirm(`⚠️ Are you sure you want to delete ALL ${sec.count} leads in the "${sec.keyword} • ${sec.city}" section?\\nOther sections will remain safe.`)) return;

      try {
        const res = await fetch('/api/leads/delete-section', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ keyword: sec.keyword, city: sec.city })
        });
        const data = await res.json();
        if (res.ok && data.success) {
          allLeads = allLeads.filter(l => !(String(l.keyword) === String(sec.keyword) && String(l.city) === String(sec.city)));
          activeSectionKey = 'all';
          lastLeadsCount = allLeads.length;
          renderSectionsNav();
          applySectionAndFilters(false);
          showToast(`🗑️ Section deleted! Removed ${data.count} leads.`);
        } else {
          alert('Failed to delete section.');
        }
      } catch (err) {
        console.error('Failed to delete section', err);
      }
    }

    // Export Current Section
    function exportCurrentSection() {
      if (activeSectionKey === 'all') {
        window.location.href = '/api/export';
        showToast('📥 Exporting all leads...');
      } else {
        const sections = getSectionsMap();
        const sec = sections.get(activeSectionKey);
        if (sec) {
          window.location.href = `/api/export?keyword=${encodeURIComponent(sec.keyword)}&city=${encodeURIComponent(sec.city)}`;
          showToast(`📥 Exporting ${sec.keyword} section...`);
        }
      }
    }

    async function clearAllLeads() {
      if (!confirm('⚠️ Are you sure you want to delete ALL extracted leads across ALL sections? This cannot be undone.')) return;
      try {
        const res = await fetch('/api/leads/clear', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ all: true })
        });
        if (res.ok) {
          allLeads = [];
          selectedCids.clear();
          updateBulkBar();
          currentPage = 1;
          lastLeadsCount = 0;
          renderSectionsNav();
          applySectionAndFilters(false);
          showToast('All leads cleared!');
        }
      } catch (err) {
        console.error('Failed to clear leads', err);
      }
    }

    async function updateLead(cid, field, value) {
      try {
        const res = await fetch('/api/leads/update', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ cid, field, value })
        });
        if (res.ok) {
          const lead = allLeads.find(l => String(l.cid) === String(cid));
          if (lead) lead[field] = value;
          showToast(`Updated ${field.replace('_', ' ')}!`);
          updateStats();
        }
      } catch (err) {
        console.error('Failed to update lead', err);
      }
    }

    // ========================================================
    // SCRAPER CONTROLS & LIVE POLLING
    // ========================================================
    async function startNewExtraction() {
      const city = document.getElementById('input-city').value.trim();
      const keywordsRaw = document.getElementById('input-keywords').value.trim();
      const stepKm = parseFloat(document.getElementById('input-mode').value) || 4.0;
      const limitRaw = document.getElementById('input-limit').value.trim();
      const limit = limitRaw ? parseInt(limitRaw, 10) : null;

      if (!city) {
        alert('Please enter a target city (e.g. Nanded).');
        return;
      }
      if (!keywordsRaw) {
        alert('Please enter at least one keyword (e.g. hospitals or cafes).');
        return;
      }

      const keywords = keywordsRaw.split(',').map(k => k.trim()).filter(Boolean);
      const btn = document.getElementById('btn-start-scraper');
      btn.disabled = true;
      btn.innerText = '⏳ Launching...';

      try {
        if ('Notification' in window && Notification.permission === 'default') {
          Notification.requestPermission();
        }
        document.title = `⏳ [Searching in ${city}...] Google Maps Leads`;
        wasRunningBefore = true;

        const res = await fetch('/api/scraper/start', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ city, keywords, step_km: stepKm, limit })
        });
        const data = await res.json();
        if (res.ok && data.success) {
          showToast(`🚀 Scraper started for ${city}! You will be alerted once completed.`);
          pollStatus();
        } else {
          alert(data.error || 'Failed to start scraper.');
        }
      } catch (err) {
        console.error('Failed to start scraper', err);
        alert('Error communicating with server.');
      } finally {
        btn.disabled = false;
        btn.innerText = '🚀 Launch Scraper';
      }
    }

    async function togglePause() {
      const btn = document.getElementById('btn-pause');
      btn.disabled = true;
      try {
        const endpoint = isPausedState ? '/api/scraper/resume' : '/api/scraper/pause';
        const res = await fetch(endpoint, { method: 'POST' });
        if (res.ok) {
          isPausedState = !isPausedState;
          showToast(isPausedState ? '⏸️ Scraper paused' : '▶️ Scraper resumed');
          await pollStatus();
        }
      } catch (e) {
        console.error('Failed to toggle pause', e);
      } finally {
        btn.disabled = false;
      }
    }

    async function stopScraper() {
      if (!confirm('Are you sure you want to stop the scraper?\\nAll leads collected so far are safely preserved in the database.')) return;
      const btn = document.getElementById('btn-stop');
      btn.disabled = true;
      try {
        const res = await fetch('/api/scraper/stop', { method: 'POST' });
        if (res.ok) {
          showToast('⏹️ Stop requested. Finishing up...');
          await pollStatus();
        }
      } catch (e) {
        console.error('Failed to stop scraper', e);
      } finally {
        btn.disabled = false;
      }
    }

    // ========================================================
    // SEARCH COMPLETION NOTIFICATION SYSTEM
    // ========================================================
    let wasRunningBefore = false;
    let handledCompletionId = localStorage.getItem('gm_leads_last_comp_id') || '';
    let soundEnabled = localStorage.getItem('gm_leads_sound') !== 'false';
    let titleFlashTimer = null;

    function initNotificationUI() {
      const soundBtn = document.getElementById('btn-sound-toggle');
      if (soundBtn) {
        soundBtn.innerText = soundEnabled ? '🔔 Sound: ON' : '🔕 Sound: OFF';
      }
      const notifBtn = document.getElementById('btn-notify-toggle');
      if (notifBtn && 'Notification' in window) {
        if (Notification.permission === 'granted') {
          notifBtn.innerText = '💻 Alerts: Active';
          notifBtn.style.color = '#34d399';
        } else if (Notification.permission === 'denied') {
          notifBtn.innerText = '💻 Alerts: Blocked';
          notifBtn.style.color = '#f87171';
        }
      }
    }

    function toggleSoundNotification() {
      soundEnabled = !soundEnabled;
      localStorage.setItem('gm_leads_sound', soundEnabled ? 'true' : 'false');
      initNotificationUI();
      if (soundEnabled) {
        playCompletionChime();
        showToast('🔔 Completion audio chime enabled');
      } else {
        showToast('🔕 Completion audio chime muted');
      }
    }

    function toggleDesktopNotifications() {
      if (!('Notification' in window)) {
        alert('Desktop notifications are not supported in this browser.');
        return;
      }
      if (Notification.permission === 'default') {
        Notification.requestPermission().then(permission => {
          initNotificationUI();
          if (permission === 'granted') {
            triggerDesktopNotification('Notifications Enabled!', 'You will be notified right here when search completes.');
            showToast('💻 Desktop notifications enabled!');
          }
        });
      } else if (Notification.permission === 'granted') {
        triggerDesktopNotification('Test Notification', 'Desktop notifications are actively working!');
        showToast('💻 Desktop alerts are already active!');
      } else {
        alert('Desktop notifications are currently blocked in your browser settings. Please click the site settings icon in the address bar to enable them.');
      }
    }

    function playCompletionChime() {
      if (!soundEnabled) return;
      try {
        const AudioCtx = window.AudioContext || window.webkitAudioContext;
        if (!AudioCtx) return;
        const ctx = new AudioCtx();
        if (ctx.state === 'suspended') {
          ctx.resume();
        }
        const now = ctx.currentTime;
        // Festive 4-note ascending chime: C5 (523Hz), E5 (659Hz), G5 (784Hz), C6 (1046Hz)
        const notes = [523.25, 659.25, 783.99, 1046.50];
        notes.forEach((freq, idx) => {
          const osc = ctx.createOscillator();
          const gain = ctx.createGain();
          osc.type = 'triangle';
          osc.frequency.setValueAtTime(freq, now + idx * 0.12);

          gain.gain.setValueAtTime(0.001, now + idx * 0.12);
          gain.gain.exponentialRampToValueAtTime(0.22, now + idx * 0.12 + 0.03);
          gain.gain.exponentialRampToValueAtTime(0.0001, now + idx * 0.12 + 0.38);

          osc.connect(gain);
          gain.connect(ctx.destination);

          osc.start(now + idx * 0.12);
          osc.stop(now + idx * 0.12 + 0.42);
        });
      } catch (err) {
        console.warn('Audio chime note error:', err);
      }
    }

    function triggerDesktopNotification(title, body) {
      if (!('Notification' in window)) return;
      if (Notification.permission === 'granted') {
        try {
          const notif = new Notification(title, {
            body: body,
            requireInteraction: true
          });
          notif.onclick = () => {
            window.focus();
            notif.close();
          };
        } catch (e) {
          console.warn('Desktop notification error:', e);
        }
      }
    }

    function flashTabTitle(alertText) {
      if (titleFlashTimer) clearInterval(titleFlashTimer);
      let count = 0;
      titleFlashTimer = setInterval(() => {
        count++;
        document.title = (count % 2 === 0) ? alertText : '✅ Google Maps Leads Extractor';
        if (count >= 18) {
          clearInterval(titleFlashTimer);
          document.title = '✅ [Search Finished] Google Maps Leads Extractor';
        }
      }, 1000);
    }

    function showCompletionModal(data) {
      const modal = document.getElementById('completion-modal');
      if (!modal) return;
      const newCount = (typeof data.new_leads === 'number') ? data.new_leads : 0;
      const totalCount = (typeof data.total_leads === 'number') ? data.total_leads : allLeads.length;
      const city = data.city || 'Finished Search';
      const kw = data.keywords_str || (data.keywords && data.keywords.join(', ')) || 'All Keywords';
      const isStopped = !!data.stopped_by_user;

      document.getElementById('modal-title-text').innerText = isStopped ? '⏹️ Search Stopped' : '🎉 Search Completed!';
      document.getElementById('modal-subtitle-text').innerText = isStopped
        ? 'Scraper was stopped safely by user. All collected leads are safely stored.'
        : `All targeted listings for ${kw} in ${city} have been completely extracted!`;
      document.getElementById('modal-stat-new').innerText = '+' + newCount;
      document.getElementById('modal-stat-total').innerText = totalCount;
      document.getElementById('modal-stat-city').innerText = city;

      modal.classList.add('active');
    }

    function closeCompletionModal() {
      const modal = document.getElementById('completion-modal');
      if (modal) modal.classList.remove('active');
    }

    function handleSearchCompleted(s, comp) {
      const completionData = comp || {
        new_leads: Math.max(0, (s.leads_count || 0) - lastLeadsCount),
        total_leads: s.leads_count || allLeads.length,
        city: s.current_city || 'Search',
        keywords_str: s.current_keyword || 'Keywords',
        stopped_by_user: s.phase && s.phase.includes('Stopped')
      };

      if (completionData.completion_id) {
        handledCompletionId = completionData.completion_id;
        localStorage.setItem('gm_leads_last_comp_id', handledCompletionId);
      }

      // 1. Festive audio chime
      playCompletionChime();

      // 2. Native Windows desktop notification
      const summaryMsg = `Extraction finished: ${completionData.total_leads} leads in database (+${completionData.new_leads || 0} new for ${completionData.city}). Excel file updated.`;
      triggerDesktopNotification('🎉 Google Maps Search Completed!', summaryMsg);

      // 3. Flashing browser tab title
      flashTabTitle(`🎉 [SEARCH FINISHED! +${completionData.new_leads || 0} LEADS]`);

      // 4. In-app celebratory toast and modal
      showToast(`🎉 Search complete! Total ${completionData.total_leads} leads saved.`);
      showCompletionModal(completionData);

      // 5. Automatically refresh leads table
      loadLeads(true);
    }

    async function pollStatus() {
      try {
        const res = await fetch('/api/status?_t=' + Date.now());
        const s = await res.json();
        const pill = document.getElementById('status-pill');
        const badgeText = document.getElementById('status-badge-text');
        const phaseText = document.getElementById('status-phase-text');
        const detailText = document.getElementById('status-detail-text');
        const mapsLink = document.getElementById('status-maps-link');
        const btnPause = document.getElementById('btn-pause');
        const btnPauseIcon = document.getElementById('btn-pause-icon');
        const btnPauseText = document.getElementById('btn-pause-text');
        const btnStop = document.getElementById('btn-stop');
        const statusDownloadBtn = document.getElementById('status-download-btn');

        const previouslyRunning = isRunningState;
        isRunningState = !!s.is_running;
        isPausedState = !!s.is_paused;

        // Check if a search just completed
        const hasNewCompletion = s.last_completion && s.last_completion.completion_id && s.last_completion.completion_id !== handledCompletionId;
        if ((previouslyRunning && !isRunningState) || (hasNewCompletion && !isRunningState)) {
          wasRunningBefore = false;
          handleSearchCompleted(s, s.last_completion);
        } else if (isRunningState) {
          wasRunningBefore = true;
        }

        if (isRunningState) {
          btnPause.style.display = 'inline-flex';
          btnStop.style.display = 'inline-flex';
          if (statusDownloadBtn) statusDownloadBtn.style.display = 'none';

          if (isPausedState) {
            pill.className = 'status-pill paused';
            badgeText.innerText = 'PAUSED';
            phaseText.innerText = 'Scraper Paused';
            detailText.innerText = 'Scraper is currently paused. Click Resume to continue extracting.';
            btnPauseIcon.innerText = '▶️';
            btnPauseText.innerText = 'Resume';
          } else {
            pill.className = 'status-pill running';
            badgeText.innerText = 'SCRAPING LIVE';
            phaseText.innerText = s.phase || 'Active';
            detailText.innerText = `${s.details || ''} (Found: ${s.discovered_count || 0} URLs | Leads: ${s.leads_count || 0} | Queue: ${s.pending_count || 0})`;
            btnPauseIcon.innerText = '⏸️';
            btnPauseText.innerText = 'Pause';
            document.title = `⏳ [Scraping...] Google Maps Leads`;
          }

          if (s.current_url) {
            mapsLink.href = s.current_url;
            mapsLink.style.display = 'inline-flex';
          } else {
            mapsLink.style.display = 'none';
          }
        } else {
          mapsLink.style.display = 'none';
          btnPause.style.display = 'none';
          btnStop.style.display = 'none';

          const isCompletedPhase = s.phase && (s.phase.includes('Completed') || s.phase.includes('Finished') || s.last_completion);
          if (isCompletedPhase) {
            pill.className = 'status-pill completed';
            badgeText.innerText = 'COMPLETED ✅';
            phaseText.innerText = s.phase || 'Search Completed';
            detailText.innerText = s.details || `Search complete! Total leads in database: ${s.leads_count || allLeads.length}`;
            if (statusDownloadBtn) statusDownloadBtn.style.display = 'inline-flex';
          } else {
            pill.className = 'status-pill';
            badgeText.innerText = 'IDLE';
            phaseText.innerText = s.phase || 'Idle / Ready';
            detailText.innerText = `Total Leads in DB: ${s.leads_count || allLeads.length} | Discovered: ${s.discovered_count || 0}`;
            if (statusDownloadBtn) statusDownloadBtn.style.display = 'none';
          }
        }

        // Auto-refresh leads table dynamically if scraper added new leads
        if (typeof s.leads_count === 'number' && s.leads_count > lastLeadsCount) {
          lastLeadsCount = s.leads_count;
          loadLeads(true);
        }
      } catch (err) {
        console.error('Failed to poll status', err);
      }
    }

    function showToast(msg) {
      const t = document.getElementById('toast');
      document.getElementById('toast-msg').innerText = msg;
      t.classList.add('show');
      setTimeout(() => t.classList.remove('show'), 3500);
    }

    function escapeHtml(str) {
      return String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    window.onload = () => {
      initNotificationUI();
      loadLeads();
      pollStatus();
      setInterval(pollStatus, 2500);
    };
  </script>

  <!-- Search Completed Celebration Modal -->
  <div id="completion-modal" class="modal-overlay" onclick="if(event.target===this) closeCompletionModal()">
    <div class="modal-card">
      <button class="modal-close-x" onclick="closeCompletionModal()" title="Close">✕</button>
      <div class="modal-badge-icon">🎉</div>
      <h2 class="modal-title" id="modal-title-text">Search Completed!</h2>
      <p class="modal-subtitle" id="modal-subtitle-text">The scraper has successfully finished extracting business leads from Google Maps.</p>

      <div class="modal-stats-grid">
        <div class="modal-stat-box">
          <div class="modal-stat-val green" id="modal-stat-new">+0</div>
          <div class="modal-stat-lbl">New Leads</div>
        </div>
        <div class="modal-stat-box">
          <div class="modal-stat-val" id="modal-stat-total">0</div>
          <div class="modal-stat-lbl">Total in DB</div>
        </div>
        <div class="modal-stat-box">
          <div class="modal-stat-val" id="modal-stat-city" style="font-size: 1rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">-</div>
          <div class="modal-stat-lbl">Location</div>
        </div>
      </div>

      <div class="modal-notice">
        <span style="font-size: 1.3rem;">📊</span>
        <div>
          <div><b>Excel & Database Saved:</b> <code>leads_export.xlsx</code> updated</div>
          <div style="font-size: 0.76rem; color: var(--text-muted); margin-top: 2px;">Pre-configured dropdowns ready for Call Status, Interest & Send</div>
        </div>
      </div>

      <div class="modal-actions">
        <button class="btn btn-secondary" onclick="closeCompletionModal()">View Extracted Leads</button>
        <a href="/api/export" class="btn btn-primary" download="leads_export.xlsx" onclick="closeCompletionModal()">📊 Download Excel Report</a>
      </div>
    </div>
  </div>
</body>
</html>
"""


class DashboardHandler(http.server.BaseHTTPRequestHandler):
    """Handles HTTP requests for the local dashboard."""

    def _set_no_cache_headers(self, content_type: str = "application/json; charset=utf-8"):
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.send_header("Access-Control-Allow-Origin", "*")

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        query = urllib.parse.parse_qs(parsed.query)

        if parsed.path in ("/", "/index.html"):
            self.send_response(200)
            self._set_no_cache_headers("text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(HTML_PAGE.encode("utf-8"))

        elif parsed.path == "/api/leads":
            leads = DB_MANAGER.get_all_leads()
            self.send_response(200)
            self._set_no_cache_headers("application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps(leads, ensure_ascii=False).encode("utf-8"))

        elif parsed.path == "/api/sections":
            sections = DB_MANAGER.get_sections_summary()
            self.send_response(200)
            self._set_no_cache_headers("application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps(sections, ensure_ascii=False).encode("utf-8"))

        elif parsed.path == "/api/status":
            status_data = get_status()
            status_data["is_paused"] = is_pause_requested()
            try:
                with DB_MANAGER.get_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("SELECT COUNT(*) FROM leads")
                    status_data["leads_count"] = cursor.fetchone()[0]
                    cursor.execute("SELECT COUNT(*) FROM discovered_urls")
                    status_data["discovered_count"] = cursor.fetchone()[0]
                    cursor.execute("SELECT COUNT(*) FROM discovered_urls WHERE status = 'discovered'")
                    status_data["pending_count"] = cursor.fetchone()[0]
            except Exception:
                status_data["leads_count"] = 0
                status_data["discovered_count"] = 0
                status_data["pending_count"] = 0

            self.send_response(200)
            self._set_no_cache_headers("application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps(status_data, ensure_ascii=False).encode("utf-8"))

        elif parsed.path == "/api/export":
            # Supports filtering by ?keyword=...&city=... OR ?cids=1,2,3
            target_leads = None
            filename = "leads_export.xlsx"

            if "cids" in query:
                cids_raw = query["cids"][0].split(",")
                target_leads = DB_MANAGER.get_leads_by_cids(cids_raw)
                filename = f"leads_selected_{len(target_leads)}.xlsx"
            elif "keyword" in query and "city" in query:
                kw = query["keyword"][0].strip()
                ct = query["city"][0].strip()
                with DB_MANAGER.get_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("SELECT * FROM leads WHERE keyword = ? AND city = ?", (kw, ct))
                    target_leads = [dict(r) for r in cursor.fetchall()]
                clean_name = f"{kw}_{ct}".replace(" ", "_").lower()
                filename = f"leads_{clean_name}.xlsx"
            else:
                target_leads = DB_MANAGER.get_all_leads()
                filename = "leads_all_export.xlsx"

            # Generate bytes directly in memory to avoid Windows file-lock collision
            excel_bytes = generate_excel_bytes(target_leads)

            # Also attempt to update DEFAULT_EXCEL_PATH on disk safely in background
            try:
                export_leads_to_excel(DB_MANAGER, DEFAULT_EXCEL_PATH)
            except Exception:
                pass

            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f"attachment; filename={filename}")
            self.send_header("Cache-Control", "no-cache")
            self.end_headers()
            self.wfile.write(excel_bytes)

        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        content_length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_length)

        if parsed.path == "/api/scraper/pause":
            request_pause()
            self.send_response(200)
            self._set_no_cache_headers("application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"success": True, "paused": True}).encode("utf-8"))
            return

        elif parsed.path == "/api/scraper/resume":
            request_resume()
            self.send_response(200)
            self._set_no_cache_headers("application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"success": True, "paused": False}).encode("utf-8"))
            return

        elif parsed.path == "/api/scraper/stop":
            request_stop()
            self.send_response(200)
            self._set_no_cache_headers("application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"success": True, "stopping": True}).encode("utf-8"))
            return

        elif parsed.path == "/api/scraper/start":
            try:
                data = json.loads(body.decode("utf-8"))
                city = data.get("city", "").strip()
                keywords = data.get("keywords", [])
                step_km = data.get("step_km", 4.0)
                limit = data.get("limit")

                if not city or not keywords:
                    self.send_response(400)
                    self._set_no_cache_headers("application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps({"error": "City and keywords are required"}).encode("utf-8"))
                    return

                # Clear stop flag and pause flag before launching
                clear_stop_request()
                request_resume()

                # Build CLI args for main.py
                script_path = str(Path(__file__).resolve().parent / "main.py")
                cmd = [
                    sys.executable,
                    script_path,
                    "--city", city,
                    "--keywords", *keywords,
                    "--step-km", str(step_km),
                ]
                if limit and int(limit) > 0:
                    cmd.extend(["--limit", str(limit)])

                subprocess.Popen(
                    cmd,
                    cwd=str(Path(__file__).resolve().parent),
                    creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if hasattr(subprocess, "CREATE_NEW_PROCESS_GROUP") else 0,
                )

                self.send_response(200)
                self._set_no_cache_headers("application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"success": True, "message": f"Scraper started for {city}!"}).encode("utf-8"))
                return
            except Exception as e:
                self.send_response(500)
                self._set_no_cache_headers("application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode("utf-8"))
                return

        elif parsed.path == "/api/leads/update":
            try:
                data = json.loads(body.decode("utf-8"))
                cid = data.get("cid")
                field = data.get("field")
                value = data.get("value")

                if cid and field in ("call_status", "interest", "send", "lead_quality"):
                    with DB_MANAGER.get_connection() as conn:
                        cursor = conn.cursor()
                        for col in ("call_status", "interest", "send", "lead_quality"):
                            try:
                                cursor.execute(f"ALTER TABLE leads ADD COLUMN {col} TEXT")
                            except Exception:
                                pass

                        cursor.execute(f"UPDATE leads SET {field} = ? WHERE cid = ?", (value, str(cid).strip()))
                        conn.commit()

                    DB_MANAGER.export_to_json()
                    try:
                        export_leads_to_excel(DB_MANAGER, DEFAULT_EXCEL_PATH)
                    except Exception:
                        pass

                    self.send_response(200)
                    self._set_no_cache_headers("application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps({"success": True}).encode("utf-8"))
                    return
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(str(e).encode("utf-8"))
                return

        elif parsed.path == "/api/leads/delete":
            try:
                data = json.loads(body.decode("utf-8"))
                cid = data.get("cid")
                if cid:
                    deleted = DB_MANAGER.delete_lead(str(cid).strip())
                    try:
                        export_leads_to_excel(DB_MANAGER, DEFAULT_EXCEL_PATH)
                    except Exception:
                        pass
                    self.send_response(200)
                    self._set_no_cache_headers("application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps({"success": deleted, "cid": cid}).encode("utf-8"))
                    return
                else:
                    self.send_response(400)
                    self.end_headers()
                    self.wfile.write(json.dumps({"error": "No cid provided"}).encode("utf-8"))
                    return
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(str(e).encode("utf-8"))
                return

        elif parsed.path == "/api/leads/delete-batch":
            try:
                data = json.loads(body.decode("utf-8"))
                cids = data.get("cids", [])
                if cids:
                    count = DB_MANAGER.delete_leads_batch(cids)
                    try:
                        export_leads_to_excel(DB_MANAGER, DEFAULT_EXCEL_PATH)
                    except Exception:
                        pass
                    self.send_response(200)
                    self._set_no_cache_headers("application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps({"success": True, "count": count}).encode("utf-8"))
                    return
                else:
                    self.send_response(400)
                    self.end_headers()
                    self.wfile.write(json.dumps({"error": "No cids provided"}).encode("utf-8"))
                    return
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(str(e).encode("utf-8"))
                return

        elif parsed.path == "/api/leads/delete-section":
            try:
                data = json.loads(body.decode("utf-8"))
                kw = data.get("keyword")
                ct = data.get("city")
                count = DB_MANAGER.delete_leads_by_section(kw, ct)
                try:
                    export_leads_to_excel(DB_MANAGER, DEFAULT_EXCEL_PATH)
                except Exception:
                    pass
                self.send_response(200)
                self._set_no_cache_headers("application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"success": True, "count": count}).encode("utf-8"))
                return
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(str(e).encode("utf-8"))
                return

        elif parsed.path == "/api/leads/clear":
            try:
                count = DB_MANAGER.clear_all_leads()
                try:
                    export_leads_to_excel(DB_MANAGER, DEFAULT_EXCEL_PATH)
                except Exception:
                    pass
                self.send_response(200)
                self._set_no_cache_headers("application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"success": True, "count": count}).encode("utf-8"))
                return
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(str(e).encode("utf-8"))
                return

        self.send_response(404)
        self.end_headers()

    def log_message(self, format, *args):
        # Suppress noisy logs
        pass


def run_dashboard_server(port=PORT):
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", port), DashboardHandler) as httpd:
        print(f"[*] Leads Dashboard running at: http://localhost:{port}")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\\nStopping dashboard server...")


if __name__ == "__main__":
    run_dashboard_server()
