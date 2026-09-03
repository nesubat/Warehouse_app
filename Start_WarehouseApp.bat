@echo off
cd /d "%~dp0"
call venv\Scripts\activate.bat
py app.py
echo.
echo === App exited or crashed. Read any error above. ===
pause
