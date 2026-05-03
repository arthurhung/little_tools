@echo off
setlocal
cd /d "%~dp0"
python outlook_cli.py run
echo.
pause
