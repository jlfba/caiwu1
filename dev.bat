@echo off
rem Dev mode: open two windows - backend (15618) + frontend hot-reload (59323)
start "pdf-backend" /D "%~dp0backend" cmd /k python run.py
start "pdf-frontend" /D "%~dp0frontend" cmd /k npm run dev
