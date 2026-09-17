@echo off
setlocal
pushd "%~dp0" || (
  echo 无法进入资金组脚本所在目录。
  pause
  exit /b 1
)

py -3.12 -c "import sys; print(sys.version)" >nul 2>nul
if not errorlevel 1 (
  py -3.12 -m pip install openpyxl pywin32
  goto :done
)

python -m pip install openpyxl pywin32

:done
if errorlevel 1 pause
popd
endlocal
