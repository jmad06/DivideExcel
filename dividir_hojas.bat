@echo off
chcp 65001 >nul
setlocal

set "SCRIPT=%~dp0dividir_hojas.py"

if "%~1"=="" (
    echo Arrastra uno o varios archivos .xlsx sobre este .bat
    pause
    exit /b 1
)

py "%SCRIPT%" %*
if errorlevel 9009 python "%SCRIPT%" %*

endlocal