@echo off
:: ─────────────────────────────────────────────────────────────────────────────
:: start_chrome.bat
:: Kills any running Chrome, then launches it with remote debugging on port 9222.
:: ─────────────────────────────────────────────────────────────────────────────

set "CHROME_EXE=C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME_EXE%" set "CHROME_EXE=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME_EXE%" set "CHROME_EXE=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"

if not exist "%CHROME_EXE%" (
    echo ERROR: Chrome not found. Edit this file and set CHROME_EXE manually.
    pause
    exit /b 1
)

:: Kill any existing Chrome processes
echo Closing any existing Chrome windows...
taskkill /F /IM chrome.exe >nul 2>&1
timeout /t 2 /nobreak >nul

:: Start Chrome with remote debugging
echo Starting Chrome with remote debugging on port 9222...
start "" "%CHROME_EXE%" --remote-debugging-port=9222 --no-first-run --no-default-browser-check --disable-session-crashed-bubble --restore-last-session

:: Wait for Chrome to fully start and bind the port
echo Waiting for Chrome to start...
timeout /t 4 /nobreak >nul

:: Verify the port is open
netstat -an | findstr "9222" | findstr "LISTENING" >nul
if %errorlevel% == 0 (
    echo.
    echo  Chrome is ready on port 9222.
    echo.
    echo  Next steps:
    echo   1. Log in to the ordering system in the Chrome window
    echo   2. Then run:  python automation.py
) else (
    echo.
    echo  WARNING: Port 9222 is not yet listening. Chrome may still be starting.
    echo  Wait a few more seconds, then check with:
    echo    netstat -an ^| findstr 9222
    echo.
    echo  If it never appears, try running this script again.
)
echo.
pause
