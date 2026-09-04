@echo off
REM ============================================================
REM  COCO - Dashboard local (doble clic para abrir)
REM  Levanta un servidor local y abre el tablero en el navegador.
REM  Para cerrarlo: cierra esta ventana negra.
REM ============================================================
cd /d "%~dp0"

echo.
echo   COCO Tecnologias - Cockpit Financiero
echo   -------------------------------------------------
echo   Archivo: index.html  (unico oficial)
echo   Base:    migracion\BD_MAESTRA_COCO.xlsx
echo.
echo   Abriendo el tablero en tu navegador...
echo   NO cierres esta ventana mientras uses el tablero.
echo   Para terminar: cierra esta ventana.
echo.

REM Abre el navegador tras un breve arranque del servidor
start "" cmd /c "timeout /t 2 >nul & start http://localhost:8740/index.html"

REM Arranca el servidor (queda corriendo en esta ventana)
python -m http.server 8740
