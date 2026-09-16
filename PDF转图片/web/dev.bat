@echo off
rem Dev mode: open two windows - backend (8000) + frontend hot-reload (5173)
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
start "pdf-backend" /D "%~dp0backend" cmd /k ""%PDF_TOOL_PYTHON%" run.py"
start "pdf-frontend" /D "%~dp0frontend" cmd /k npm run dev
