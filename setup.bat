@echo off
setlocal enabledelayedexpansion

echo =======================================================
echo   Windows Accountability Tracker - Setup & Restart
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

:: 2. Kill existing logger instances
echo.
echo [1/5] Stopping any running tracker instances...
taskkill /F /IM pythonw.exe /FI "WINDOWTITLE eq activity_logger.py" >nul 2>&1
:: Also try killing by matching python processes if they are invisible
powershell -Command "Get-Process pythonw -ErrorAction SilentlyContinue | Stop-Process -Force" >nul 2>&1
echo [OK] Old instances stopped.

:: 3. Install Dependencies
echo.
echo [2/5] Installing required Python libraries...
python -m pip install --upgrade pip
python -m pip install pywin32 rich psutil
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

:: 4. Create Startup Script (Silent Background Runner)
echo.
echo [3/5] Configuring background tracker to start with Windows...
set "STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "VBS_SCRIPT=%STARTUP_FOLDER%\start_activity_logger.vbs"
set "LOGGER_PATH=%~dp0activity_logger.py"

:: Get pythonw.exe path (runs without console window)
for /f "delims=" %%i in ('where pythonw') do (
    set "PYTHONW_EXE=%%i"
    goto :found_pythonw
)
:: Fallback if where pythonw fails
for /f "delims=" %%i in ('where python') do set "PYTHON_EXE=%%i"
set "PYTHONW_EXE=%PYTHON_EXE:python.exe=pythonw.exe%"

:found_pythonw
echo [OK] Using PythonW at: %PYTHONW_EXE%

(
echo Set WshShell = CreateObject^("WScript.Shell"^)
echo WshShell.Run "%PYTHONW_EXE% %LOGGER_PATH%", 0, False
) > "%VBS_SCRIPT%"

echo [OK] Startup script created at: %VBS_SCRIPT%

:: 5. Start the logger immediately
echo.
echo [4/5] Launching background activity logger...
:: Remove lock file if it exists to ensure clean start
if exist "%USERPROFILE%\.screen_time\activity_logger.lock" del /F "%USERPROFILE%\.screen_time\activity_logger.lock" >nul 2>&1
start "" /B "%PYTHONW_EXE%" "%LOGGER_PATH%"
echo [OK] Logger is now running silently in the background.

:: 6. Final Run
echo.
echo [5/5] Generating your first report...
echo.
timeout /t 2 >nul
python "%~dp0screen_time.py"

echo.
echo =======================================================
echo   SETUP & RESTART COMPLETE! 
echo =======================================================
echo.
echo  * The tracker will now start automatically when you log in.
echo  * To see your report anytime, run 'python screen_time.py'
echo.
pause