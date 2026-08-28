@echo off
REM ============================================================
REM  REVISAR LA BASE antes de publicar un cierre
REM  Comprueba que las capas de la base digan lo mismo.
REM  Doble clic. No modifica nada.
REM ============================================================
cd /d "%~dp0"
cls
echo.
echo   ================================================================
echo     REVISION DE LA BASE - COCO
echo   ================================================================
echo.
echo   1/3  Cifras del cierre y reglas que no dependen del mes...
echo.
python -m pytest tests -q --no-header -p no:cacheprovider
echo.
echo   ----------------------------------------------------------------
echo   2/3  Revisando que todas las hojas digan lo mismo...
echo.
cd migracion
python conciliar_capas.py
echo.
echo   ----------------------------------------------------------------
echo   3/3  Reglas del negocio (espejo SQLite)...
echo.
python exportar_sqlite.py
echo.
echo   ================================================================
echo.
echo   Para publicar hacen falta las tres:
echo     - los tests en "passed" (si alguno falla, mira cual y por que)
echo     - "todas las capas estan alineadas"
echo     - las validaciones V1..V6 en OK
echo.
echo   Si un test de tests	est_cifras.py falla despues de cargar un mes
echo   nuevo, NO lo ajustes a ojo: mira que cifra se movio y por que.
echo   Esa revision ES el cierre. Los de test_invariantes.py no deberian
echo   fallar nunca; si uno falla, algo se rompio.
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
