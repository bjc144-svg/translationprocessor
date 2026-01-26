@echo off
REM This script adds a Windows Firewall rule to allow port 5000 connections
REM Must be run as Administrator

echo ============================================
echo Fixing Firewall for Translation Processor
echo ============================================
echo.
echo This script will add a Windows Firewall rule to allow
echo incoming connections on port 5000 (needed for network access)
echo.
echo IMPORTANT: You must run this as Administrator!
echo Right-click this file and select "Run as administrator"
echo.
pause

echo.
echo Adding firewall rule...
netsh advfirewall firewall delete rule name="Translation Processor" protocol=TCP localport=5000 >nul 2>&1
netsh advfirewall firewall add rule name="Translation Processor" dir=in action=allow protocol=TCP localport=5000

if errorlevel 1 (
    echo.
    echo ERROR: Failed to add firewall rule
    echo.
    echo Possible reasons:
    echo   - Not running as Administrator
    echo   - Firewall is disabled or managed by group policy
    echo.
    echo Please:
    echo   1. Close this window
    echo   2. Right-click 'fix-firewall.bat'
    echo   3. Select "Run as administrator"
    echo   4. Try again
    echo.
) else (
    echo.
    echo SUCCESS! Firewall rule added.
    echo.
    echo Port 5000 is now allowed through Windows Firewall.
    echo Other computers on your network should now be able to access the server.
    echo.
    echo Next steps:
    echo   1. Make sure the server is running (run start.bat)
    echo   2. From another computer, try: http://[YOUR-IP]:5000
    echo   3. Your IP address will be shown when you start the server
    echo.
)

pause
