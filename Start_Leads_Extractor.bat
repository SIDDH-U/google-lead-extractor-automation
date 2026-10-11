@echo off
title Google Maps Leads Extractor
echo ========================================================
echo   Starting Google Maps Leads Extractor Web Dashboard...
echo ========================================================

cd /d "%~dp0"

:: Automatically open http://localhost:5000 in your browser after 2 seconds
start "" powershell -NoProfile -Command "Start-Sleep -Seconds 2; Start-Process 'http://localhost:5000'"

:: Run local dashboard server
".venv\Scripts\python.exe" app.py
