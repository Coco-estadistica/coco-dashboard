@echo off
REM ============================================================
REM  CORREGIR desalineaciones entre hojas de la base
REM  Alinea Subsidiarias_PyG con BD_Indicadores (la que lee el tablero).
REM  Hace copia de seguridad antes de tocar nada.
REM ============================================================
cd /d "%~dp0migracion"
cls
echo.
echo   ================================================================
echo     CORREGIR LA BASE - COCO
echo   ================================================================
echo.
echo   Primero te muestro QUE cambiaria (todavia no toca nada):
echo.
python sincronizar_subsidiarias.py
echo.
echo   ================================================================
set /p CONF="   Aplicar estos cambios? (S/N): "
if /i not "%CONF%"=="S" goto :fin
echo.
python sincronizar_subsidiarias.py --escribir
echo.
echo   Verificando que todo quedo alineado...
echo.
python conciliar_capas.py
:fin
echo.
pause
