@echo off
setlocal
REM ============================================================
REM  COCO - SUBIR TUS CAMBIOS A GITHUB
REM  Doble clic DESPUES de cargar cifras y revisar la base.
REM  Muestra que va a subir y pide confirmacion.
REM ============================================================
cd /d "%~dp0"
cls
echo.
echo   ================================================================
echo     SUBIR CAMBIOS A GITHUB - COCO
echo   ================================================================
echo.

where git >nul 2>&1
if errorlevel 1 (
  echo   NO ENCUENTRO GIT en este computador.
  echo   Instalalo desde https://git-scm.com/download/win y vuelve a intentar.
  echo.
  pause
  exit /b 1
)

for /f "delims=" %%b in ('git rev-parse --abbrev-ref HEAD 2^>nul') do set RAMA=%%b
if "%RAMA%"=="" (
  echo   Esta carpeta no es un repositorio de GitHub. Nada que subir.
  echo.
  pause
  exit /b 1
)
echo   Rama: %RAMA%
echo.

REM --- 0. Revisar la base ANTES de subir -------------------------------
REM     Vuelca la base a SQLite, corre los controles y regenera coco.sql
REM     (el volcado de texto que si se puede comparar en GitHub).
REM
REM     Si un control encuentra una cifra rota DENTRO de la base --el
REM     consolidado que no es la suma de los paises, una utilidad que no
REM     cuadra, la misma llave dos veces-- no se sube nada. Los avisos que
REM     dependen de terceros (cartera) no detienen: se anotan y siguen.
REM     Se usa el codigo de salida del script: 1 = no subir, 0 = adelante.
where python >nul 2>&1
if errorlevel 1 goto SINPYTHON

echo   Revisando las cifras de la base...
echo.
python migracion\exportar_sqlite.py --volcado
if errorlevel 1 goto BASEROTA
echo.
goto TRASREVISION

:SINPYTHON
echo   AVISO: no encuentro Python en este computador, asi que NO pude
echo   revisar las cifras antes de subir. Se puede subir igual, pero sin
echo   esa comprobacion. Instala Python desde https://python.org para
echo   recuperarla.
echo.
set "SIGO="
set /p "SIGO=  Subir de todos modos, sin revisar? (S/N): "
if /i not "%SIGO%"=="S" goto CANCELADO
goto TRASREVISION

:BASEROTA
echo.
echo   ------------------------------------------------------------
echo   NO SE SUBIO NADA.
echo.
echo   Los controles encontraron cifras rotas dentro de la base.
echo   Arriba esta el detalle de cada una: dicen el mes, el pais y el
echo   indicador. Corrigelas en el Excel y vuelve a ejecutar este boton.
echo.
echo   Tu trabajo sigue en tu computador, intacto. No se perdio nada.
echo   ------------------------------------------------------------
echo.
pause
exit /b 1

:CANCELADO
echo.
echo   Cancelado. No se subio nada.
echo.
pause
exit /b 0

:TRASREVISION

REM --- 1. Ver si hay algo que subir ------------------------------------
set HAYCAMBIOS=
for /f "delims=" %%s in ('git status --porcelain') do set HAYCAMBIOS=1
if not defined HAYCAMBIOS (
  echo   No hay ningun cambio pendiente. Todo esta ya en GitHub.
  echo.
  pause
  exit /b 0
)

echo   Estos son los archivos que cambiaron:
echo.
git status --short
echo.

REM --- 2. Descripcion y confirmacion -----------------------------------
set "DESC="
set /p "DESC=  Describe en pocas palabras que cambiaste (ej: cierre agosto 2026): "
if not defined DESC set "DESC=Actualizacion de la base"

echo.
set "OK="
set /p "OK=  Confirmas que quieres subir esto a GitHub? (S/N): "
if /i not "%OK%"=="S" (
  echo.
  echo   Cancelado. No se subio nada.
  echo.
  pause
  exit /b 0
)

REM --- 3. Guardar tu trabajo PRIMERO -----------------------------------
REM     Registrar el cambio antes de traer nada de GitHub. Asi tu trabajo
REM     queda a salvo pase lo que pase despues, y si algo sale mal siempre
REM     se puede volver a este punto.
echo.
echo   Guardando tu trabajo...
git add -A
git commit -q -m "%DESC%"
if errorlevel 1 (
  echo.
  echo   No se pudo registrar el cambio. Nada se subio.
  echo.
  pause
  exit /b 1
)
echo   Guardado.

REM --- 4. Traer lo que haya nuevo en GitHub ----------------------------
echo.
echo   Comprobando si hay algo nuevo en GitHub...
git fetch origin %RAMA% >nul 2>&1
set PENDIENTES=0
for /f %%n in ('git rev-list --count HEAD..origin/%RAMA% 2^>nul') do set PENDIENTES=%%n

if not "%PENDIENTES%"=="0" (
  echo   Hay %PENDIENTES% cambio^(s^) nuevo^(s^) en GitHub. Juntandolos con el tuyo...
  git pull --no-rebase --no-edit origin %RAMA%
  if errorlevel 1 (
    echo.
    echo   ------------------------------------------------------------
    echo   NO SE PUDIERON JUNTAR AUTOMATICAMENTE.
    echo   Tu y GitHub cambiaron lo mismo, y un Excel no se puede fusionar.
    echo.
    echo   Tu trabajo NO se perdio: quedo guardado en el paso anterior.
    echo   Dejo la carpeta como estaba y no subo nada. Pide ayuda.
    echo   ------------------------------------------------------------
    git merge --abort
    echo.
    pause
    exit /b 1
  )
  echo   Juntados sin problema.
)

REM --- 5. Subir --------------------------------------------------------
echo.
echo   Enviando a GitHub...
git push -u origin %RAMA%
if errorlevel 1 (
  echo.
  echo   Se guardo tu cambio en el computador pero NO se pudo enviar a GitHub.
  echo   Revisa tu conexion y vuelve a ejecutar este boton: reintentara el envio.
  echo.
  pause
  exit /b 1
)

echo.
echo   ================================================================
echo     LISTO. Tus cambios ya estan en GitHub.
echo   ================================================================
echo.
pause
