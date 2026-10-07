@echo off
rem Doble clic: descarga los reportes del portal WMS, corre el ETL, abre el
rem dashboard y deja el resultado en descarga_log.txt.
rem Para una tarea programada usar "ejecutar_dashboard_cyber.bat auto"
rem (no abre el navegador ni espera una tecla al final).
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
rem %LOCALAPPDATA% apunta al Python del usuario que lo ejecuta (sirve en cualquier equipo).
set PYTHON="%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
set SALIDA="%TEMP%\dashboard_cyber_salida.txt"

echo Actualizando dashboard Cyber (descarga + ETL, tarda unos minutos)...
echo ===== %date% %time% ===== >> descarga_log.txt
%PYTHON% dashboard_cyber.py > %SALIDA% 2>&1
set CODIGO=%errorlevel%
type %SALIDA%
type %SALIDA% >> descarga_log.txt
echo codigo de salida: %CODIGO% >> descarga_log.txt

if /i "%~1"=="auto" exit /b %CODIGO%

rem Si la descarga falló, el ETL igual arma el dashboard con los Excel anteriores.
findstr /c:"se usan los" %SALIDA% >nul && (
    echo.
    echo ATENCION: no se pudieron descargar los reportes del portal. El dashboard
    echo se armo con los ultimos Excel descargados. Revisa la conexion/VPN y reintenta.
)

if %CODIGO%==0 (
    echo.
    echo Listo. Abriendo el dashboard...
    start "" "cyber_dashboard.html"
) else (
    echo.
    echo ERROR: el dashboard no se pudo actualizar (codigo %CODIGO%). Revisa el detalle arriba.
)
echo.
pause
