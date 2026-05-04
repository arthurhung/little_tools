@echo off
setlocal
cd /d "%~dp0"
python outlook_com_cli.py --dry-run
echo.
pause
