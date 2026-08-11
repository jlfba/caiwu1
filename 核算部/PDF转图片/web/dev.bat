@echo off
rem 开发模式：开两个窗口，后端(8000) + 前端热更新(5173)
rem 改前端代码后浏览器自动刷新；开发完构建再交给后端托管
start "PDF工具-后端" /D "%~dp0backend" cmd /k python run.py
start "PDF工具-前端" /D "%~dp0frontend" cmd /k npm run dev
