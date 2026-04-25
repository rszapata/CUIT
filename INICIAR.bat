@echo off
title Buscador de Jurisdicciones CUIT - ARCA

echo.
echo  ======================================================
echo   Buscador de Jurisdicciones CUIT - ARCA
echo  ======================================================
echo.

:: Verificar Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  Python no encontrado. Instalando...
    winget install --id Python.Python.3.12 --silent --accept-source-agreements --accept-package-agreements
    set "PATH=%PATH%;%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts"
    python --version >nul 2>&1
    if %errorlevel% neq 0 (
        echo  ERROR: Instala Python desde https://www.python.org/downloads/
        echo  Marca "Add Python to PATH" durante la instalacion.
        pause
        exit /b 1
    )
)

python --version

:: Instalar dependencias si faltan
echo.
echo  Verificando dependencias...

python -c "import selenium" >nul 2>&1
if %errorlevel% neq 0 (
    echo  Instalando selenium...
    python -m pip install selenium --quiet
)

python -c "import pandas" >nul 2>&1
if %errorlevel% neq 0 (
    echo  Instalando pandas...
    python -m pip install pandas --quiet
)

python -c "import psutil" >nul 2>&1
if %errorlevel% neq 0 (
    echo  Instalando psutil...
    python -m pip install psutil --quiet
)

python -c "import openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo  Instalando openpyxl...
    python -m pip install openpyxl --quiet
)

echo  Dependencias OK.

:: Verificar que el script existe
if not exist "%~dp0buscar_jurisdicciones.py" (
    echo  ERROR: No se encontro buscar_jurisdicciones.py
    echo  Asegurate de que INICIAR.bat y buscar_jurisdicciones.py esten en la misma carpeta.
    pause
    exit /b 1
)

:: Ejecutar
echo.
echo  Iniciando...
echo.

cd /d "%~dp0"
python buscar_jurisdicciones.py

if %errorlevel% neq 0 (
    echo.
    echo  El script termino con un error.
    pause
)
