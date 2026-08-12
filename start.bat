@echo off
rem Daily use: start backend (serves page + API on port 15618)
cd /d "%~dp0backend"
python run.py
pause
