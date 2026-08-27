@echo off
setlocal
cd /d "%~dp0"
start "" ".venv\Scripts\pythonw.exe" -m course_selection workbench --port 5000
endlocal
