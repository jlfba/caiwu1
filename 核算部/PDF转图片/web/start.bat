@echo off
rem 日常使用：一键启动（页面 + 接口都在 8000 端口，无需 npm）
rem 打开浏览器访问 http://127.0.0.1:8000
cd /d "%~dp0backend"
python run.py
pause
