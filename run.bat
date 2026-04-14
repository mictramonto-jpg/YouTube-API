@echo off
setlocal EnableDelayedExpansion

echo ========================================
echo   YouTube Comment Extractor
echo ========================================
echo.

REM ==== Detect Python command ====
set PYTHON_CMD=
where py >nul 2>&1
if %ERRORLEVEL%==0 (
    set PYTHON_CMD=py
    goto python_found
)
where python >nul 2>&1
if %ERRORLEVEL%==0 (
    set PYTHON_CMD=python
    goto python_found
)
where python3 >nul 2>&1
if %ERRORLEVEL%==0 (
    set PYTHON_CMD=python3
    goto python_found
)

echo [ERROR] Python not found.
echo.
echo Please install Python:
echo   1. Download from https://www.python.org/downloads/
echo   2. During install, CHECK "Add python.exe to PATH"
echo   3. Restart PowerShell / Command Prompt
echo   4. Run this batch file again
echo.
pause
exit /b 1

:python_found
echo Python command: %PYTHON_CMD%
%PYTHON_CMD% --version
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to run Python.
    echo Microsoft Store alias may be interfering.
    echo Go to: Settings ^> Apps ^> App execution aliases
    echo Turn OFF: App Installer python.exe and python3.exe
    echo.
    pause
    exit /b 1
)
echo.

REM ==== Install dependencies ====
echo Installing dependencies...
%PYTHON_CMD% -m pip install -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Failed to install dependencies.
    echo Check your internet connection or proxy settings.
    echo.
    pause
    exit /b 1
)
echo.

REM ==== Launch app ====
echo Starting app...
echo.
%PYTHON_CMD% main.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] App execution failed.
    echo Please check the error messages above.
    echo.
)

pause
