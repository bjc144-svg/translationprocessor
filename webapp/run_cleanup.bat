@echo off
REM ============================================================================
REM Translation Processor - Cleanup Script Wrapper
REM ============================================================================
REM
REM This script activates the virtual environment and runs the cleanup script.
REM Schedule this script in Windows Task Scheduler to run daily.
REM
REM Recommended schedule: Daily at 2:00 AM
REM ============================================================================

echo ============================================
echo Translation Processor - File Cleanup
echo ============================================
echo.
echo Starting cleanup at %DATE% %TIME%
echo.

REM Change to script directory
cd /d "%~dp0"

REM Check if virtual environment exists
if not exist "venv\Scripts\activate.bat" (
    echo ERROR: Virtual environment not found!
    echo Please run start.bat first to create the virtual environment.
    echo.
    pause
    exit /b 1
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

REM Run cleanup script
echo Running cleanup script...
echo.
python cleanup_output.py

REM Check exit code
if errorlevel 1 (
    echo.
    echo ERROR: Cleanup script failed!
    echo Check logs\cleanup.log for details.
    echo.
) else (
    echo.
    echo Cleanup completed successfully.
    echo.
)

REM Deactivate virtual environment
deactivate

echo Finished at %TIME%
echo.

REM Exit (comment out 'pause' if running via Task Scheduler)
REM pause
