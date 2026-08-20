@echo off
REM ============================================================
REM  REVISAR LA BASE antes de publicar un cierre
REM  Comprueba que las capas de la base digan lo mismo.
REM  Doble clic. No modifica nada.
REM ============================================================
cd /d "%~dp0migracion"
cls
echo.
echo   ================================================================
echo     REVISION DE LA BASE - COCO
echo   ================================================================
echo.
echo   Revisando que todas las hojas digan lo mismo...
echo.
python conciliar_capas.py
echo.
echo   ================================================================
echo.
echo   Si dice "todas las capas estan alineadas", puedes publicar.
echo.
echo   Si aparecen diferencias en Subsidiarias_PyG, puedes corregirlas
echo   automaticamente ejecutando:  Corregir_Base.bat
echo.
echo   ================================================================
echo.
pause
