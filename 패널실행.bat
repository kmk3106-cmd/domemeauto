@echo off
chcp 65001 >nul
cd /d "C:\Users\USER\domemeauto"
set "PY=C:\Users\USER\PycharmProjects\PythonProject\.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"
echo [control_panel] starting...  http://localhost:8001/
start "control_panel" "%PY%" -u control_panel.py
timeout /t 3 >nul
start "" "http://localhost:8001/"
exit /b 0
