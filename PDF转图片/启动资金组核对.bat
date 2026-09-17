@echo off
setlocal
pushd "%~dp0" || (
  echo 无法进入资金组脚本所在目录。
  pause
  exit /b 1
)

py -3.12 -c "import sys; print(sys.version)" >nul 2>nul
if not errorlevel 1 (
  py -3.12 fund_check.py
  goto :done
)

python fund_check.py

:done
if errorlevel 1 pause
popd
endlocal
