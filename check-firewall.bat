@echo off
REM This script checks if the firewall rule exists

echo ============================================
echo Checking Firewall Configuration
echo ============================================
echo.

echo Searching for Translation Processor firewall rules...
echo.

netsh advfirewall firewall show rule name="Translation Processor"

if errorlevel 1 (
    echo.
    echo ============================================
    echo FIREWALL RULE NOT FOUND
    echo ============================================
    echo.
    echo The firewall rule for port 5000 is not configured.
    echo This is likely why other computers cannot access the server.
    echo.
    echo To fix this:
    echo   1. Right-click 'fix-firewall.bat'
    echo   2. Select "Run as administrator"
    echo   3. Follow the prompts
    echo.
) else (
    echo.
    echo ============================================
    echo FIREWALL RULE EXISTS
    echo ============================================
    echo.
    echo The firewall rule is configured correctly.
    echo.
    echo If other computers still cannot access the server, check:
    echo   - Both computers are on the same network
    echo   - Third-party antivirus/security software
    echo   - Corporate/network firewalls
    echo   - VPN connections (may route traffic differently)
    echo.
)

echo.
echo Checking if port 5000 is currently in use...
netstat -an | findstr ":5000"
echo.

if errorlevel 1 (
    echo No application is currently listening on port 5000
    echo (Server may not be running)
) else (
    echo Port 5000 is active (server may be running)
)

echo.
pause
