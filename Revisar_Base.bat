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
echo   ----------------------------------------------------------------
echo   Reglas del negocio (espejo SQLite)...
echo.
python exportar_sqlite.py
echo.
echo   ================================================================
echo.
echo   Si dice "todas las capas estan alineadas" y las validaciones
echo   V1..V6 salen en OK, puedes publicar.
echo.
echo   OJO con Subsidiarias_PyG y BD_PYG_OFICIAL: sus diferencias son
echo   ARCHIVO HISTORICO y estan asi a proposito desde el 13-ago-2026
echo   (ver el encabezado de conciliar_capas.py). NO correr
echo   Corregir_Base.bat para "arreglarlas": eso borra el antes de la
echo   homologacion. El control que importa es el A.
echo.
echo   ================================================================
echo.
pause
