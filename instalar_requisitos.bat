@echo off
cd /d "%~dp0"
chcp 65001 >nul 2>&1
title Instalador de Requisitos - CXP

echo ======================================================
echo    INSTALADOR DE REQUISITOS - AUTOMATIZACION CXP
echo ======================================================
echo.

echo [1/3] Detectando instalacion de Python...

set "PYTHON_CMD="

REM 1. Probar lanzador oficial 'py -3'
py -3 -c "import sys; exit(0)" >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON_CMD=py -3"
    goto :python_encontrado
)

REM 2. Probar 'python' en PATH
python -c "import sys; exit(0)" >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON_CMD=python"
    goto :python_encontrado
)

REM 3. Probar lanzador generico 'py'
py -c "import sys; exit(0)" >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON_CMD=py"
    goto :python_encontrado
)

REM 4. Probar rutas estandar en AppData o Program Files
for %%V in (Python314 Python313 Python312 Python311 Python310) do (
    if exist "%LOCALAPPDATA%\Programs\Python\%%V\python.exe" (
        set "PYTHON_CMD="%LOCALAPPDATA%\Programs\Python\%%V\python.exe""
        goto :python_encontrado
    )
    if exist "%ProgramFiles%\Python%%V\python.exe" (
        set "PYTHON_CMD="%ProgramFiles%\Python%%V\python.exe""
        goto :python_encontrado
    )
    if exist "%SystemDrive%\Python%%V\python.exe" (
        set "PYTHON_CMD="%SystemDrive%\Python%%V\python.exe""
        goto :python_encontrado
    )
)

echo.
echo [ERROR] No se encontro Python instalado en este equipo.
echo Por favor descarga e instala Python desde: https://www.python.org/downloads/
echo (Recuerda marcar la opcion "Add python.exe to PATH").
echo.
pause
exit /b 1

:python_encontrado
echo [OK] Python detectado: %PYTHON_CMD%
%PYTHON_CMD% --version
echo.

echo [2/3] Configurando entorno virtual (.venv)...
if exist ".venv" (
    ".venv\Scripts\python.exe" -c "import sys; exit(0)" >nul 2>&1
    if %errorlevel% neq 0 (
        echo [AVISO] El entorno virtual existente no es compatible. Recreando...
        rmdir /s /q ".venv" >nul 2>&1
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo Creando entorno virtual local...
    %PYTHON_CMD% -m venv .venv
    if not exist ".venv\Scripts\python.exe" (
        echo.
        echo [ERROR] No se pudo crear el entorno virtual.
        pause
        exit /b 1
    )
    echo [OK] Entorno virtual creado.
) else (
    echo [OK] Entorno virtual verificado.
)
echo.

echo [3/3] Instalando librerias necesarias (openpyxl, xlrd)...
".venv\Scripts\python.exe" -m pip install --upgrade pip >nul 2>&1
".venv\Scripts\python.exe" -m pip install openpyxl xlrd
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Ocurrio un problema instalando openpyxl o xlrd con pip.
    pause
    exit /b 1
)
echo [OK] Librerias instaladas correctamente.
echo.

echo ======================================================
echo    ¡Configuracion completada con exito!
echo ======================================================
echo Ya puedes procesar tus archivos usando:
echo   ejecutar_automatizacion.bat
echo.
pause