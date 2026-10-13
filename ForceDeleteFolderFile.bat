@echo off
setlocal EnableDelayedExpansion
title Advanced Force Delete & Take Ownership Utility - htmqng.blogspot.com
:: ====================================================================
:: STEP 1: CHECK & ELEVATE ADMINISTRATOR PRIVILEGES
:: ====================================================================
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo [INFO] Requesting administrative privileges...
    goto UACPrompt
) else ( 
    goto gotAdmin 
)

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
    pushd "%CD%"
    CD /D "%~dp0"

:: ====================================================================
:: STEP 2: MAIN SELECTION MENU
:: ====================================================================
:MENU
cls
echo ===================================================================
echo             ADVANCED FORCE DELETE & OWNERSHIP TOOL
echo ===================================================================
echo  [1] Select Target via GUI (Folder / File Picker)
echo  [2] Enter Target Path Manually
echo  [3] Exit Utility
echo ===================================================================
set /p "choice=Select an option [1-3]: "

if "%choice%"=="1" goto GUI_SELECT
if "%choice%"=="2" goto MANUAL_SELECT
if "%choice%"=="3" exit /b
goto MENU

:: ====================================================================
:: STEP 3: GUI PATH SELECTION (POWERPOLL WINDOWS FORMS)
:: ====================================================================
:GUI_SELECT
cls
echo Select Target Type:
echo  [1] Folder
echo  [2] File
echo  [3] Back to Menu
set /p "gtype=Choice [1-3]: "

set "TARGET="
if "%gtype%"=="1" (
    for /f "delims=" %%I in ('powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; $f = New-Object System.Windows.Forms.FolderBrowserDialog; $f.Description = 'Select target folder to delete'; if ($f.ShowDialog() -eq 'OK') { Write-Output $f.SelectedPath }"') do set "TARGET=%%I"
) else if "%gtype%"=="2" (
    for /f "delims=" %%I in ('powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; $f = New-Object System.Windows.Forms.OpenFileDialog; $f.Title = 'Select target file to delete'; if ($f.ShowDialog() -eq 'OK') { Write-Output $f.FileName }"') do set "TARGET=%%I"
) else (
    goto MENU
)

if "%TARGET%"=="" (
    echo [WARNING] No path selected. Returning to menu...
    timeout /t 2 >nul
    goto MENU
)
goto CONFIRM_DELETE

:: ====================================================================
:: STEP 4: MANUAL PATH SELECTION
:: ====================================================================
:MANUAL_SELECT
cls
echo ===================================================================
echo MANUAL PATH ENTRY
echo ===================================================================
set /p "TARGET=Enter full path of file or folder: "
:: Remove enclosing quotes if user dragged and dropped into cmd
set "TARGET=%TARGET:"=%"

if "%TARGET%"=="" goto MANUAL_SELECT
goto CONFIRM_DELETE

:: ====================================================================
:: STEP 5: DELETION & TAKE OWNERSHIP LOGIC
:: ====================================================================
:CONFIRM_DELETE
cls
echo ===================================================================
echo TARGET IDENTIFIED:
echo %TARGET%
echo ===================================================================
set /p "confirm=Are you SURE you want to force delete this? (Y/N): "
if /i "%confirm%" neq "Y" goto MENU

echo.
echo [1/3] Attempting standard deletion...
if exist "%TARGET%\" (
    rd /s /q "%TARGET%"
) else if exist "%TARGET%" (
    del /f /q /s "%TARGET%"
)

:: Check if deletion succeeded
if not exist "%TARGET%" (
    echo.
    echo [SUCCESS] Target deleted successfully!
    goto END_PROMPT
)

echo [WARNING] Access denied or files locked. Taking ownership...

if exist "%TARGET%\" (
    takeown /f "%TARGET%" /r /d y >nul 2>&1
    icacls "%TARGET%" /grant administrators:F /t /c /q >nul 2>&1
    icacls "%TARGET%" /grant "%USERNAME%":F /t /c /q >nul 2>&1
    rd /s /q "%TARGET%"
) else if exist "%TARGET%" (
    takeown /f "%TARGET%" /d y >nul 2>&1
    icacls "%TARGET%" /grant administrators:F /c /q >nul 2>&1
    icacls "%TARGET%" /grant "%USERNAME%":F /c /q >nul 2>&1
    del /f /q /s "%TARGET%"
)

if not exist "%TARGET%" (
    echo.
    echo [SUCCESS] Target force-deleted after taking ownership!
) else (
    echo.
    echo [ERROR] Failed to delete target. It may be locked by an active process, 
    echo system service, or requires a reboot to purge.
)

:END_PROMPT
echo.
pause
goto MENU