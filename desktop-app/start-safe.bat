@echo off
REM Translation Processor - Enhanced Startup Script with Better Error Handling

echo ============================================
echo Translation Processor
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

REM Install core requirements first
echo Installing core packages...
pip install --upgrade pip setuptools wheel --quiet

echo Installing python-docx...
pip install python-docx==1.1.0

echo Installing Pillow (for image support)...
pip install Pillow>=9.0.0

echo Installing pywin32 (for Word COM)...
pip install pywin32==306

echo Installing docx2pdf...
pip install docx2pdf==0.1.8

echo.
echo Core installation complete!
echo.

REM Check if installation was successful
python -c "import docx; import PIL; import docx2pdf; import win32com.client" 2>nul
if errorlevel 1 (
    echo.
    echo WARNING: Some packages may not have installed correctly.
    echo The application will attempt to start anyway.
    echo If you encounter errors, you may need to:
    echo   1. Install Microsoft Visual C++ Build Tools
    echo   2. Or use a different Python version
    echo.
    pause
)

REM Start application
echo.
echo ============================================
echo Starting Translation Processor...
echo ============================================
echo.

cd src
python main.py

if errorlevel 1 (
    echo.
    echo ============================================
    echo ERROR: Application failed to start
    echo ============================================
    echo.
    echo Please check the error message above.
    echo.
    echo Common solutions:
    echo   1. Make sure Microsoft Word is installed
    echo   2. Try running as Administrator
    echo   3. Check that all dependencies installed correctly
    echo.
)

REM Deactivate virtual environment
cd ..
deactivate

echo.
echo Application closed.
pause
