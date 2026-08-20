@echo off
REM ============================================================
REM  VISTA PREVIA - que datos traeria del modelo (NO escribe nada)
REM ============================================================
cd /d "%~dp0"
python extraer_modelo.py
echo.
echo ============================================================
echo   Esto fue solo una VISTA PREVIA: no se toco la base.
echo   Para cargarlos de verdad usa: Cargar_del_Modelo.bat
echo ============================================================
pause
