@echo off
setlocal enabledelayedexpansion
title USB Virus Remover Tool - htmqng.blogspot.com

:input_drive
cls
echo ===================================================================
echo                   USB VIRUS REMOVER
echo 	Malicious shortcut (.LNK) ^& autorun.inf
echo ===================================================================
echo.
set "user_input="
set /p "user_input=Enter target USB Drive Letter (e.g., E, F, G): "

:: Remove trailing spaces, colons, and backslashes
set "drive=%user_input: =%"
set "drive=%drive::=%"
set "drive=%drive:\=%"

:: Validate non-empty input
if "%drive%"=="" (
    echo.
    echo [ERROR] Input cannot be empty. Please enter a valid drive letter.
    timeout /t 3 >nul
    goto input_drive
)

set "target_drive=%drive%:"
set "sys_drive=%SystemDrive::=%"

:: Safety Verification: Prevent running against System Drive (C:)
if /i "%drive%"=="%sys_drive%" (
    echo.
    echo ===================================================================
    echo [CRITICAL ERROR] SYSTEM DRIVE DETECTED (%target_drive%)
    echo Operation aborted to prevent damage to Windows OS system files!
    echo You cannot target the active Windows installation drive.
    echo ===================================================================
    echo.
    pause
    goto input_drive
)

:: Safety Verification: Confirm drive exists and is accessible
if not exist "%target_drive%\" (
    echo.
    echo [ERROR] Drive "%target_drive%" was not found or is inaccessible.
    echo Please verify the drive letter and try again.
    echo.
    pause
    goto input_drive
)

echo.
echo Target confirmed: %target_drive%
echo Starting cleanup scan and immunization...
echo -------------------------------------------------------------------
echo.

:: 1. Strip system, hidden, and read-only attributes
echo [1/4] Restoring hidden directories and files...
attrib -h -r -s /s /d "%target_drive%\*.*"

:: 2. Delete malicious shortcut files
echo [2/4] Removing malicious shortcut (.LNK) files...
del /f /q /s "%target_drive%\*.lnk" >nul 2>&1

:: 3. Purge existing autorun infection files or folders
echo [3/4] Deleting autorun infection files...
if exist "%target_drive%\autorun.inf" (
    attrib -r -a -s -h "%target_drive%\autorun.inf" >nul 2>&1
    del /f /q /a "%target_drive%\autorun.inf" >nul 2>&1
    rmdir /s /q "%target_drive%\autorun.inf" >nul 2>&1
)

:: 4. Apply Vaccine: Create an unmodifiable autorun.inf directory
echo [4/4] Applying Autorun vaccine folder...
if not exist "%target_drive%\autorun.inf" (
    mkdir "%target_drive%\autorun.inf" >nul 2>&1
    attrib +r +s +h "%target_drive%\autorun.inf" >nul 2>&1
)

echo.
echo ===================================================================
echo SUCCESS: USB Drive %target_drive% cleaned and immunized successfully!
echo ===================================================================
echo Note: Check %target_drive%\ for any unnamed/hidden folders where
echo malware originally stashed your legitimate files, and move them back.
echo.
pause
goto input_drive