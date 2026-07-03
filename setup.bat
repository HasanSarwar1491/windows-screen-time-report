@echo off
setlocal enabledelayedexpansion

echo =======================================================
echo   Windows Accountability Tracker - First Time Setup
echo =======================================================
echo.

:: 1. Check for Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python from https://www.python.org/ and try again.
    pause
    exit /b 1
)
echo [OK] Python detected.

:: 2. Install Dependencies
echo.
echo [1/4] Installing required Python libraries...
python -m pip install --upgrade pip
python -m pip install pywin32 rich psutil
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

:: 3. Create Startup Script (Silent Background Runner)
echo.
echo [2/4] Configuring background tracker to start with Windows...
set "STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "VBS_SCRIPT=%STARTUP_FOLDER%\start_activity_logger.vbs"
set "LOGGER_PATH=%~dp0activity_logger.py"

:: Get pythonw.exe path (runs without console window)
for /f "delims=" %%i in ('where python') do set "PYTHON_EXE=%%i"
set "PYTHONW_EXE=%PYTHON_EXE:python.exe=pythonw.exe%"

(
echo Set WshShell = CreateObject^("WScript.Shell"^)
echo WshShell.Run "%PYTHONW_EXE% %LOGGER_PATH%", 0, False
) > "%VBS_SCRIPT%"

echo [OK] Startup script created at: %VBS_SCRIPT%

:: 4. Start the logger immediately
echo.
echo [3/4] Launching background activity logger...
start "" /B "%PYTHONW_EXE%" "%LOGGER_PATH%"
echo [OK] Logger is now running silently in the background.

:: 5. Final Run
echo.
echo [4/4] Generating your first report...
echo.
timeout /t 2 >nul
python "%~dp0screen_time.py"

echo.
echo =======================================================
echo   SETUP COMPLETE! 
echo =======================================================
echo.
echo  * The tracker will now start automatically when you log in.
echo  * To see your report anytime, run 'python screen_time.py'
echo.
pause
