@echo off
REM Advanced diagnostic script for network connection issues
REM This script checks everything needed for remote access

echo ================================================================
echo TRANSLATION PROCESSOR - NETWORK DIAGNOSTIC TOOL
echo ================================================================
echo.
echo This tool will help identify why other users cannot connect.
echo.
pause

REM ============================================
REM Check 1: Is the server running?
REM ============================================
echo.
echo ================================================================
echo CHECK 1: Is the server running?
echo ================================================================
echo.

netstat -an | findstr "0.0.0.0:5000" > nul
if %errorlevel% equ 0 (
    echo [PASS] Server is running and listening on ALL interfaces
    echo        This means it SHOULD be accessible from other machines
    goto :check2
)

netstat -an | findstr "127.0.0.1:5000" > nul
if %errorlevel% equ 0 (
    echo [FAIL] Server is running but ONLY on localhost ^(127.0.0.1^)
    echo        Other machines CANNOT access it this way!
    echo.
    echo PROBLEM FOUND: Server is bound to localhost only
    echo SOLUTION: Check app.py - it should have host='0.0.0.0'
    echo           Current code might not be running correctly
    goto :end
)

netstat -an | findstr ":5000" > nul
if %errorlevel% equ 0 (
    echo [WARN] Port 5000 is in use, but binding is unclear
    echo        Here are all connections on port 5000:
    echo.
    netstat -an | findstr ":5000"
    goto :check2
) else (
    echo [FAIL] Server is NOT running! No process listening on port 5000
    echo.
    echo PROBLEM FOUND: The server is not running
    echo SOLUTION: Start the server by running start.bat
    echo           Make sure you see "Running on http://0.0.0.0:5000"
    goto :end
)

:check2
REM ============================================
REM Check 2: What's the server IP address?
REM ============================================
echo.
echo ================================================================
echo CHECK 2: Network IP Addresses
echo ================================================================
echo.
echo Your computer's IP addresses:
echo.

ipconfig | findstr /C:"IPv4"

echo.
echo IMPORTANT: Give other users one of these IP addresses ^(NOT 127.0.0.1^)
echo Example: If your IP is 192.168.1.100, they should use:
echo          http://192.168.1.100:5000
echo.

REM ============================================
REM Check 3: Windows Firewall
REM ============================================
echo.
echo ================================================================
echo CHECK 3: Windows Firewall Rule
echo ================================================================
echo.

netsh advfirewall firewall show rule name="Translation Processor" > nul 2>&1
if %errorlevel% equ 0 (
    echo [PASS] Firewall rule exists for Translation Processor
) else (
    echo [FAIL] Firewall rule is missing!
    echo        Run fix-firewall.bat as administrator
)

REM ============================================
REM Check 4: Detailed network status
REM ============================================
echo.
echo ================================================================
echo CHECK 4: Detailed Port 5000 Status
echo ================================================================
echo.
echo All connections involving port 5000:
echo.
netstat -an | findstr ":5000"
echo.

REM ============================================
REM Check 5: Third-party security software check
REM ============================================
echo.
echo ================================================================
echo CHECK 5: Third-Party Security Software
echo ================================================================
echo.
echo Checking for common antivirus/security software...
echo.

tasklist | findstr /I "norton mcafee kaspersky avast avg bitdefender" > nul
if %errorlevel% equ 0 (
    echo [WARN] Third-party security software detected!
    echo        These programs may have their own firewalls:
    tasklist | findstr /I "norton mcafee kaspersky avast avg bitdefender"
    echo.
    echo        Check their settings and allow Python or port 5000
) else (
    echo [PASS] No common third-party security software detected
)

REM ============================================
REM Check 6: Network connectivity test
REM ============================================
echo.
echo ================================================================
echo CHECK 6: Network Configuration
echo ================================================================
echo.
echo Your network adapter details:
echo.
ipconfig /all | findstr /C:"Description" /C:"IPv4" /C:"Subnet" /C:"Default Gateway"

echo.
echo ================================================================
echo SUMMARY AND RECOMMENDATIONS
echo ================================================================
echo.

REM Provide summary based on findings
netstat -an | findstr "0.0.0.0:5000" > nul
if %errorlevel% equ 0 (
    echo Server Status: RUNNING ^(listening on all interfaces^)
    echo Firewall: Check the output above
    echo.
    echo If users still cannot connect:
    echo.
    echo 1. VERIFY IP ADDRESS
    echo    - Make sure they're using your actual IP ^(from ipconfig above^)
    echo    - NOT using 127.0.0.1 or localhost
    echo    - Format: http://[YOUR-IP]:5000
    echo.
    echo 2. SAME NETWORK
    echo    - Both computers must be on same network
    echo    - Check: Do IP addresses start with same numbers?
    echo      Example: Both start with 192.168.1.x = good
    echo      Example: One is 192.168.1.x, other is 10.0.0.x = bad
    echo.
    echo 3. THIRD-PARTY FIREWALL
    echo    - Check antivirus/security software settings ^(see above^)
    echo    - Temporarily disable to test ^(remember to re-enable^)
    echo.
    echo 4. ROUTER/NETWORK FIREWALL
    echo    - Corporate networks may block custom ports
    echo    - Contact IT department if on corporate network
    echo.
    echo 5. VPN CONNECTIONS
    echo    - If either computer is on VPN, disconnect and test
    echo    - VPNs often route traffic differently
    echo.
    echo 6. CLIENT-SIDE FIREWALL
    echo    - The OTHER computers may be blocking outbound connections
    echo    - Check their firewall settings too
    echo.
) else (
    echo Server Status: NOT RUNNING or bound to localhost only
    echo.
    echo NEXT STEPS:
    echo 1. Make sure start.bat is running
    echo 2. Check for error messages in the server window
    echo 3. Look for "Running on http://0.0.0.0:5000" in output
    echo 4. If you see "Running on http://127.0.0.1:5000" instead,
    echo    the server is misconfigured
)

:end
echo.
echo ================================================================
echo.
echo If you're still having issues after checking all of the above,
echo read FIREWALL-FIX.md for more detailed troubleshooting.
echo.
echo You may also want to run this diagnostic on the CLIENT machines
echo to check for outbound firewall blocks.
echo.
pause
