@echo off
setlocal enabledelayedexpansion

title MRMV Report Content Migration Tool - Launcher

echo =======================================================================
echo          MRMV Report Content Migration Tool
echo =======================================================================
echo.

:: 1. Detect Python executable
where python >nul 2>nul
if %errorlevel% equ 0 (
    set PYTHON_CMD=python
    goto :PYTHON_FOUND
)

where py >nul 2>nul
if %errorlevel% equ 0 (
    set PYTHON_CMD=py
    goto :PYTHON_FOUND
)

:: Python not detected
echo [ERROR] Python was not found on this computer.
echo.
echo Please ensure Python 3.8 or higher is installed and added to your PATH.
echo Contact your IT Department or Helpdesk to install standard Python.
echo.
pause
exit /b 1

:PYTHON_FOUND
echo [OK] Python detected:
%PYTHON_CMD% --version
echo.

:: 2. Check pywin32 library
echo [INFO] Checking required dependency (pywin32)...
%PYTHON_CMD% -c "import win32com.client, pythoncom" >nul 2>nul
if %errorlevel% neq 0 (
    echo [INFO] Installing 'pywin32' dependency...
    %PYTHON_CMD% -m pip install pywin32 --quiet
    if !errorlevel! neq 0 (
        echo.
        echo [WARNING] Automatic pip install failed or was blocked by bank proxy.
        echo Trying user installation: pip install pywin32 --user
        %PYTHON_CMD% -m pip install pywin32 --user --quiet
    )
    
    :: Re-verify pywin32
    %PYTHON_CMD% -c "import win32com.client, pythoncom" >nul 2>nul
    if !errorlevel! neq 0 (
        echo.
        echo [ERROR] 'pywin32' is not installed or could not be loaded.
        echo Please ask your IT representative to install pywin32 for Python.
        echo.
        pause
        exit /b 1
    )
)
echo [OK] pywin32 is ready.

:: Optional Drag-and-Drop enhancement
%PYTHON_CMD% -c "import tkinterdnd2" >nul 2>nul
if %errorlevel% neq 0 (
    echo [INFO] Checking optional drag-and-drop support...
    %PYTHON_CMD% -m pip install tkinterdnd2 --quiet 2>nul
)
echo.

:: 3. Launch the desktop GUI
echo [INFO] Launching Desktop GUI Application...
echo.

%PYTHON_CMD% "%~dp0gui_app.py"

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] The application exited with an error code.
    pause
)

exit /b 0

