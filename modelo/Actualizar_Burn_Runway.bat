@echo off
REM ============================================================
REM  Actualiza el informe Burn & Runway desde el Modelo
REM  Usalo despues de reemplazar Modelo_COCO.xlsx por la version nueva
REM ============================================================
cd /d "%~dp0"
echo.
echo   Leyendo la hoja "Burn ^& RunW" de Modelo_COCO.xlsx...
echo.
python extraer_burn_runway.py
echo.
echo   Listo. Abre el dashboard y ve a la pestana "Burn ^& Runway".
echo.
pause
