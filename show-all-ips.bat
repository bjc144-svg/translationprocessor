@echo off
REM This script shows all available IP addresses and helps identify the correct one

echo ================================================================
echo TRANSLATION PROCESSOR - IP ADDRESS HELPER
echo ================================================================
echo.
echo This tool shows ALL IP addresses on this computer.
echo Use this to find the correct IP for other users to connect to.
echo.
pause

echo ================================================================
echo ALL NETWORK ADAPTERS AND IP ADDRESSES
echo ================================================================
echo.

ipconfig /all

echo.
echo ================================================================
echo SUMMARY - IPv4 ADDRESSES ONLY
echo ================================================================
echo.
echo These are your IPv4 addresses:
echo.

ipconfig | findstr /C:"IPv4"

echo.
echo ================================================================
echo WHICH IP ADDRESS SHOULD OTHER USERS USE?
echo ================================================================
echo.
echo Look for the IP address that matches your local network:
echo.
echo - Typically starts with: 192.168.x.x or 10.x.x.x
echo - NOT 127.0.0.1 (that's localhost only)
echo - NOT 169.254.x.x (that's a self-assigned error address)
echo - If you have multiple adapters (WiFi + Ethernet), use the
echo   one that's actually connected to your network
echo.
echo Example: If you see "192.168.1.105", other users should use:
echo          http://192.168.1.105:5000
echo.

echo ================================================================
echo VERIFY SERVER IS LISTENING
echo ================================================================
echo.
echo Checking if server is running on ALL interfaces (0.0.0.0):
echo.

netstat -an | findstr ":5000"

echo.
echo What you want to see:
echo   TCP    0.0.0.0:5000           0.0.0.0:0              LISTENING
echo.
echo What's BAD (if you see this instead):
echo   TCP    127.0.0.1:5000         0.0.0.0:0              LISTENING
echo   ^(This means server is only listening on localhost!^)
echo.

REM Check what we actually see
netstat -an | findstr "0.0.0.0:5000" > nul
if %errorlevel% equ 0 (
    echo [GOOD] Server is listening on ALL interfaces (0.0.0.0)
    echo        Other machines SHOULD be able to connect
) else (
    netstat -an | findstr "127.0.0.1:5000" > nul
    if %errorlevel% equ 0 (
        echo [BAD] Server is only listening on localhost (127.0.0.1)
        echo       Other machines CANNOT connect this way!
        echo.
        echo       This is a configuration problem - the server is not
        echo       accepting connections from the network.
    ) else (
        echo [UNKNOWN] Server may not be running, or port 5000 isn't in use
        echo           Make sure start.bat is running
    )
)

echo.
echo ================================================================
echo.
pause
