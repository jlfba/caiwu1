@echo off
rem Daily use: start backend (serves page + API on port 8000)
set "PDF_TOOL_PYTHON=D:\code\Python\Python312\python.exe"
if not exist "%PDF_TOOL_PYTHON%" (
  echo Python 3.12 not found: %PDF_TOOL_PYTHON%
  pause
  exit /b 1
)
"%PDF_TOOL_PYTHON%" -c "import rapidocr_onnxruntime" || (
  echo Missing OCR dependency: rapidocr-onnxruntime
  pause
  exit /b 1
)
cd /d "%~dp0backend"
rem Daily-use mode must stay stable while long OCR tasks are running.
set "RELOAD=0"
set "PORT=8000"
"%PDF_TOOL_PYTHON%" run.py
pause
