@echo off
setlocal enabledelayedexpansion

title MRMV Report Content Migration Tool - Launcher

:: 1. Fast Launch: Try windowless Python directly (instant launch, closes terminal immediately)
where pythonw >nul 2>nul
if %errorlevel% equ 0 (
    start "" pythonw "%~dp0gui_app.py"
    exit /b 0
)

where pyw >nul 2>nul
if %errorlevel% equ 0 (
    start "" pyw "%~dp0gui_app.py"
    exit /b 0
)

:: 2. Standard Python Launch
where python >nul 2>nul
if %errorlevel% equ 0 (
    start "" python "%~dp0gui_app.py"
    exit /b 0
)

where py >nul 2>nul
if %errorlevel% equ 0 (
    start "" py "%~dp0gui_app.py"
    exit /b 0
)

:: 3. Error Handling if Python is missing
echo =======================================================================
echo          MRMV Report Content Migration Tool
echo =======================================================================
echo.
echo [ERROR] Python was not found on this computer.
echo.
echo Please ensure Python 3.8 or higher is installed and added to your PATH.
echo Contact your IT Department or Helpdesk to install standard Python.
echo.
pause
exit /b 1
