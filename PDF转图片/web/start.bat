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
rem Local daily-use mode: automatically reload after backend Python files change.
set "RELOAD=1"
set "PORT=3629"
"%PDF_TOOL_PYTHON%" run.py
pause
