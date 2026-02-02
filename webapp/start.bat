@echo off
REM Translation Processor Web Application - Startup Script

echo ============================================
echo Translation Processor - Web Application
echo Park Evaluation Services
echo ============================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

echo Python found!
python --version
echo.

REM Check if virtual environment exists
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment
        pause
        exit /b 1
    )
    echo Virtual environment created!
    echo.
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

REM Install/update requirements
echo.
echo Installing dependencies...
echo This may take a few minutes on first run...
echo.

python -m pip install --upgrade pip setuptools wheel --quiet
pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo WARNING: Some dependencies may not have installed correctly
    echo The application may not work properly
    echo.
    pause
)

echo.
echo Dependencies installed!
echo.

REM Get local IP address
echo ============================================
echo Network Information
echo ============================================
echo.
echo Your computer's IP addresses:
ipconfig | findstr /C:"IPv4"
echo.
echo ============================================
echo.

REM Start application
echo Starting Translation Processor web server...
echo.
echo The application will be accessible at:
echo   - On this computer: http://localhost:5000
echo   - From other computers: http://[YOUR-IP]:5000
echo.
echo Press CTRL+C to stop the server
echo ============================================
echo.

python app.py

REM Deactivate virtual environment
deactivate

echo.
echo Server stopped.
pause
