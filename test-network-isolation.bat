@echo off
REM Quick network isolation test - checks if clients are on same subnet

echo ================================================================
echo NETWORK ISOLATION TEST
echo ================================================================
echo.
echo Your server IP: 192.168.1.244
echo Your subnet: 192.168.1.0/24
echo.
echo This script helps determine if client machines are on the same
echo network or if there's network isolation preventing access.
echo.
pause

echo.
echo ================================================================
echo YOUR NETWORK INFORMATION
echo ================================================================
echo.
echo Server IP Address: 192.168.1.244
echo Subnet Mask: 255.255.255.0
echo Network Range: 192.168.1.1 - 192.168.1.254
echo.
echo For other users to connect, their IP address should start with:
echo   192.168.1.x
echo.
echo If their IP starts with something else (like 10.x.x.x or
echo 172.x.x.x or different 192.168.x.x), they are on a DIFFERENT
echo network and cannot connect without routing.
echo.

echo.
echo ================================================================
echo TESTING NETWORK REACH
echo ================================================================
echo.
echo Let's test if we can reach other machines on the 192.168.1.x network
echo.
echo Common devices to try pinging (router, etc.):
ping -n 1 192.168.1.1 > nul 2>&1
if %errorlevel% equ 0 (
    echo [FOUND] 192.168.1.1 - likely your router/gateway
) else (
    echo [NOT FOUND] 192.168.1.1
)

ping -n 1 192.168.1.254 > nul 2>&1
if %errorlevel% equ 0 (
    echo [FOUND] 192.168.1.254
)

echo.
echo Checking for active devices on your network...
echo (This may take a moment)
echo.

REM Try a few common addresses
for /L %%i in (1,1,10) do (
    ping -n 1 -w 100 192.168.1.%%i > nul 2>&1
    if not errorlevel 1 (
        echo [ACTIVE] 192.168.1.%%i
    )
)

echo.
echo ================================================================
echo CLIENT MACHINE REQUIREMENTS
echo ================================================================
echo.
echo For users to connect to your server, they MUST:
echo.
echo 1. Be on the same network (192.168.1.x range)
echo 2. Be able to ping 192.168.1.244
echo 3. Not have client-side firewall blocking port 5000
echo 4. Not be on VPN (VPN may route traffic differently)
echo 5. Not be on different VLAN (corporate networks)
echo.
echo ================================================================
echo WHAT TO ASK THE OTHER USERS
echo ================================================================
echo.
echo Ask them to:
echo.
echo 1. Open Command Prompt
echo 2. Run: ipconfig
echo 3. Find their "IPv4 Address"
echo 4. Report back what it is
echo.
echo Then:
echo.
echo IF their IP is 192.168.1.something:
echo   → They ARE on same network
echo   → Have them run: ping 192.168.1.244
echo   → Have them run: test-connection-from-client.bat
echo   → Problem is likely client-side firewall
echo.
echo IF their IP is NOT 192.168.1.something:
echo   → They are on a DIFFERENT network
echo   → They CANNOT connect (network isolation)
echo   → Solutions:
echo      a) Connect to same network (same WiFi, etc.)
echo      b) Use desktop app approach instead
echo      c) Contact IT if corporate network
echo.
echo ================================================================
echo.
pause
