@echo off
setlocal
chcp 65001 >nul
title Fix Macro XLSM
color 0A

echo =====================================================
echo.
echo    CONG CU XOA MACRO VA CHUYEN DOI XLSM SANG XLSX
echo.
echo =====================================================
echo.

set "PS1=%~dp0Fix_Macro_XLSM.ps1"

if not exist "%PS1%" (
    color 0C
    echo [LỖI] Không tìm thấy file Fix_Macro_XLSM.ps1.
    echo Hãy đảm bảo file BAT và PS1 nằm cùng một thư mục.
    echo.
    pause
    exit /b 1
)

powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%PS1%"
set "EC=%ERRORLEVEL%"

if not "%EC%"=="0" (
    color 0E
    echo.
    echo [CẢNH BÁO] Tiến trình kết thúc với mã lỗi %EC%.
    pause
)

endlocal & exit /b %EC%