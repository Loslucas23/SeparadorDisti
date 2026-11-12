@echo off
chcp 65001 >nul
title Backup por Distribuidor
cd /d "%~dp0"

echo ============================================================
echo   BACKUP POR DISTRIBUIDOR
echo ============================================================
echo.
echo  Toma TODOS los archivos .xlsb / .xlsx de esta carpeta
echo  y crea un Excel por distribuidor en "Por Distribuidor\"
echo  (una hoja por año fiscal: FY24, FY25, ... FY28).
echo.
echo  - FY nuevo, FY anteriores o un mes suelto: copiar el archivo
echo    aca (hojas llamadas FY28, FY24, etc.).
echo  - Por cada mes y distribuidor se usa el archivo mas nuevo.
echo  - Los archivos que no cambiaron no se vuelven a leer.
echo  - Si un Excel de distribuidor esta abierto, se saltea y se avisa.
echo  - Al final: control.txt (origen vs generado) y log.txt (historial).
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] No se encontro Python. Instalarlo desde python.org
    echo         marcando "Add Python to PATH".
    pause
    exit /b 1
)

python -c "import pandas, pyarrow, pyxlsb, xlsxwriter, openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Instalando librerias necesarias por unica vez...
    python -m pip install --quiet pandas pyarrow pyxlsb xlsxwriter openpyxl
)

echo  1 = Generar TODOS los distribuidores
echo  2 = Generar UN distribuidor
echo.
set "OPC=1"
set /p "OPC=Opcion [1]: "

set "FILTRO="
if "%OPC%"=="2" set /p "FILTRO=Nombre (o parte del nombre) del distribuidor: "

echo.
python "%~dp0Automatizacion\separar_por_distribuidor.py" "%FILTRO%"
set "RC=%errorlevel%"
echo.
if "%RC%"=="0" (
    echo [OK] Todo generado y controlado. Ver "Por Distribuidor\control.txt"
    start "" "%~dp0Por Distribuidor"
) else if "%RC%"=="2" (
    echo [ATENCION] Termino con avisos: archivos abiertos, posibles archivos
    echo            incompletos o diferencias. Revisar "Por Distribuidor\control.txt".
    start "" "%~dp0Por Distribuidor"
) else (
    echo [ERROR] Algo fallo. Revisar el mensaje de arriba o "Por Distribuidor\log.txt"
)
echo.
pause
