@echo off
cd /d "%~dp0"
chcp 65001 >nul 2>&1
title Organizador Automatico de CXP - Excel

echo ======================================================
echo    ORGANIZADOR AUTOMATICO DE ARCHIVOS CXP (EXCEL)
echo ======================================================
echo.

REM 1. Verificar entorno virtual .venv
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "import sys; exit(0)" >nul 2>&1
    if %errorlevel% neq 0 (
        echo [AVISO] El entorno virtual de esta carpeta necesita configurarse para este equipo.
        call instalar_requisitos.bat
    )
) else (
    echo [AVISO] Configurando el entorno por primera vez...
    call instalar_requisitos.bat
)

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] No se pudo iniciar el entorno de Python.
    pause
    exit /b 1
)

REM 2. Ejecutar script Python
if "%~1"=="" (
    echo Procesando archivos Excel en esta carpeta...
    ".venv\Scripts\python.exe" organizar_cxp.py
) else (
    echo Procesando archivo: "%~1"
    ".venv\Scripts\python.exe" organizar_cxp.py "%~1"
)

:fin
echo.
echo ======================================================
echo    Proceso finalizado.
echo ======================================================
echo.
pause