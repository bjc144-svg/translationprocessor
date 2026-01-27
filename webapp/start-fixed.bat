@echo off
REM Enhanced startup script with Python 3.14 compatibility

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

REM Check Python version
for /f "tokens=2" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo Detected Python version: %PYVER%
echo.

REM Warn if Python 3.14+
echo %PYVER% | findstr /B "3.14" >nul
if not errorlevel 1 (
    echo WARNING: Python 3.14 detected - some packages may have compatibility issues
    echo Recommended: Python 3.11 or 3.12 for best compatibility
    echo.
    echo Press any key to continue anyway, or CTRL+C to cancel...
    pause >nul
)

REM Remove old venv if exists and has issues
if exist "venv\Scripts\python.exe" (
    echo Checking existing virtual environment...
    venv\Scripts\pip --version >nul 2>&1
    if errorlevel 1 (
        echo Old virtual environment is broken, removing...
        rmdir /s /q venv
    )
)

REM Create virtual environment if needed
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

REM Upgrade pip first
echo.
echo Upgrading pip...
python -m pip install --upgrade pip --quiet

REM Install requirements one by one with better error handling
echo.
echo Installing dependencies...
echo This may take a few minutes on first run...
echo.

echo [1/6] Installing Flask...
pip install "Flask>=3.0.0" --quiet
if errorlevel 1 (
    echo ERROR: Failed to install Flask
    goto :error
)

echo [2/6] Installing Werkzeug...
pip install "Werkzeug>=3.0.0" --quiet

echo [3/6] Installing python-docx...
pip install "python-docx>=1.0.0" --quiet
if errorlevel 1 (
    echo ERROR: Failed to install python-docx
    goto :error
)

echo [4/6] Installing Pillow...
pip install "Pillow>=9.0.0" --quiet

echo [5/6] Installing pywin32...
pip install pywin32 --quiet
if errorlevel 1 (
    echo WARNING: pywin32 failed to install (may need older Python version)
    echo PDF conversion may not work - will save as .docx instead
)

echo [6/6] Installing docx2pdf...
pip install docx2pdf --quiet
if errorlevel 1 (
    echo WARNING: docx2pdf failed to install
    echo PDF conversion may not work - will save as .docx instead
)

echo.
echo Dependencies installed!
echo.

REM Verify Flask is actually installed
python -c "import flask" 2>nul
if errorlevel 1 (
    echo ERROR: Flask installation verification failed
    echo.
    echo Trying alternative installation method...
    pip install Flask Werkzeug python-docx Pillow
    if errorlevel 1 (
        goto :error
    )
)

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

if errorlevel 1 (
    echo.
    echo ============================================
    echo ERROR: Application failed to start
    echo ============================================
    echo.
    echo Please check the error message above.
    echo.
    goto :error
)

REM Deactivate virtual environment
deactivate

echo.
echo Server stopped.
pause
exit /b 0

:error
echo.
echo ============================================
echo INSTALLATION TROUBLESHOOTING
echo ============================================
echo.
echo Common solutions:
echo.
echo 1. Use Python 3.11 or 3.12 (more compatible than 3.14):
echo    - Download from: https://www.python.org/downloads/
echo    - Install and use that version instead
echo.
echo 2. Delete venv folder and try again:
echo    rmdir /s /q venv
echo    start.bat
echo.
echo 3. Try manual installation:
echo    python -m venv venv
echo    venv\Scripts\activate
echo    pip install Flask python-docx Pillow
echo    python app.py
echo.
echo 4. Run as Administrator (right-click start.bat)
echo.
pause
exit /b 1
