@echo off
rem Dev mode: open two windows - backend (8000) + frontend hot-reload (5173)
start "pdf-backend" /D "%~dp0backend" cmd /k python run.py
start "pdf-frontend" /D "%~dp0frontend" cmd /k npm run dev
