@echo off
rem Daily use: start backend (serves page + API on port 8000)
cd /d "%~dp0backend"
python run.py
pause
