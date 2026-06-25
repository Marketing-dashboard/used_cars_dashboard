@echo off
echo Updating dashboard with latest Excel data...
echo.
cd /d "%~dp0"
python build.py
echo.
echo Done! Now go to GitHub Desktop and commit + push index.html to publish.
echo.
pause
