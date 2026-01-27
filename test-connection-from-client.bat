@echo off
REM This script should be run on CLIENT machines (not the server)
REM It tests connectivity to the Translation Processor server

echo ================================================================
echo TRANSLATION PROCESSOR - CLIENT CONNECTION TEST
echo ================================================================
echo.
echo This script tests if you can connect to the server.
echo Run this on the CLIENT machine (not the server machine).
echo.

REM Get server IP from user
set /p SERVER_IP="Enter the server's IP address (e.g., 192.168.1.100): "

echo.
echo Testing connection to %SERVER_IP%:5000...
echo.

REM ============================================
REM Test 1: Can we ping the server?
REM ============================================
echo ================================================================
echo TEST 1: Network Connectivity (Ping)
echo ================================================================
echo.

ping -n 2 %SERVER_IP% > nul
if %errorlevel% equ 0 (
    echo [PASS] Can ping server at %SERVER_IP%
    echo        Basic network connectivity works
) else (
    echo [FAIL] Cannot ping server at %SERVER_IP%
    echo        Possible causes:
    echo        - Wrong IP address
    echo        - Not on same network
    echo        - ICMP/ping is blocked
    echo        - Server is offline
    echo.
    echo        Try continuing anyway - sometimes ping is blocked but HTTP works
)

REM ============================================
REM Test 2: Is port 5000 reachable?
REM ============================================
echo.
echo ================================================================
echo TEST 2: Port 5000 Connectivity
echo ================================================================
echo.
echo Attempting to connect to %SERVER_IP%:5000...
echo (This may take a few seconds)
echo.

REM Use PowerShell to test TCP connection
powershell -Command "$result = Test-NetConnection -ComputerName %SERVER_IP% -Port 5000 -InformationLevel Quiet -WarningAction SilentlyContinue; if ($result) { exit 0 } else { exit 1 }" > nul 2>&1

if %errorlevel% equ 0 (
    echo [PASS] Port 5000 is reachable!
    echo        The server should be accessible
    echo.
    echo        Try opening this URL in your browser:
    echo        http://%SERVER_IP%:5000
    goto :success
) else (
    echo [FAIL] Cannot connect to port 5000
    echo.
    echo        PROBLEM: Port 5000 is blocked or server is not running
    goto :troubleshoot
)

:success
echo.
echo ================================================================
echo SUCCESS!
echo ================================================================
echo.
echo The server appears to be accessible from this machine.
echo.
echo Open your web browser and go to:
echo.
echo     http://%SERVER_IP%:5000
echo.
echo You should see a green success page.
echo.
goto :end

:troubleshoot
echo.
echo ================================================================
echo TROUBLESHOOTING
echo ================================================================
echo.
echo The server is NOT accessible from this machine. Possible causes:
echo.
echo 1. SERVER NOT RUNNING
echo    - Make sure start.bat is running on the server
echo    - Check for error messages
echo.
echo 2. WRONG IP ADDRESS
echo    - Verify the IP address with the server operator
echo    - Server should run diagnose-connection.bat to get correct IP
echo    - Do NOT use 127.0.0.1 or localhost
echo.
echo 3. DIFFERENT NETWORKS
echo    - Both computers must be on the same network
echo    - Check if this computer's IP starts with same numbers
echo    - Run 'ipconfig' on both machines to compare
echo.
echo 4. SERVER FIREWALL (Most Common!)
echo    - Server needs to run fix-firewall.bat as administrator
echo    - Server needs Windows Firewall rule for port 5000
echo.
echo 5. CLIENT FIREWALL
echo    - This computer may be blocking outbound connections
echo    - Check Windows Firewall on THIS computer
echo    - Check antivirus/security software on THIS computer
echo.
echo 6. ROUTER/NETWORK FIREWALL
echo    - Corporate networks may block port 5000
echo    - Contact IT department
echo    - May need to use different approach
echo.
echo 7. VPN CONNECTIONS
echo    - If either computer uses VPN, disconnect and test
echo    - VPNs route traffic differently
echo.
echo ================================================================
echo CLIENT MACHINE DIAGNOSTICS
echo ================================================================
echo.
echo Your IP address:
ipconfig | findstr /C:"IPv4"
echo.
echo Your firewall status:
netsh advfirewall show currentprofile state
echo.

:end
echo.
echo For more help, read FIREWALL-FIX.md on the server machine.
echo.
pause
