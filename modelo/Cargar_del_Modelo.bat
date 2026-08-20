@echo off
REM ============================================================
REM  CARGA REAL - escribe en la base los datos del modelo
REM  (hace copia de seguridad de la base antes de tocarla)
REM ============================================================
cd /d "%~dp0"
echo.
echo   Se van a ESCRIBIR en la base los datos marcados SI en mapa_modelo.csv
echo   Se hara una copia de seguridad automatica antes.
echo.
set /p CONF="   Continuar? (S/N): "
if /i not "%CONF%"=="S" goto :fin
echo.
python extraer_modelo.py --escribir
:fin
echo.
pause
