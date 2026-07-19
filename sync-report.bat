@echo off
cd /d "%~dp0"
echo Syncing confirmed days into the Sales Report...
echo (Close the file in Excel first if it is open)
echo.
node scripts\sync-day-to-report.mjs
echo.
pause
