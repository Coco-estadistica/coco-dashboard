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

REM --- 2. Comprobar que nadie subio algo mientras tanto -----------------
REM     Un .xlsx es binario: Git no puede fusionarlo. Si subimos encima de
REM     una version mas nueva, una de las dos se pierde entera.
echo   Comprobando que no haya versiones mas nuevas en GitHub...
git fetch origin %RAMA% >nul 2>&1
set PENDIENTES=0
for /f %%n in ('git rev-list --count HEAD..origin/%RAMA% 2^>nul') do set PENDIENTES=%%n
if not "%PENDIENTES%"=="0" (
  echo.
  echo   ALTO. En GitHub hay %PENDIENTES% version^(es^) mas nueva^(s^) que la tuya.
  echo   Si subes ahora se pueden perder cifras: un Excel no se puede fusionar.
  echo.
  echo   Que hacer: guarda una copia de tu Excel FUERA de esta carpeta,
  echo   ejecuta Actualizar_Desde_GitHub.bat, y vuelve a cargar tus cifras
  echo   sobre la version actualizada.
  echo.
  pause
  exit /b 1
)

REM --- 3. Descripcion y confirmacion -----------------------------------
echo.
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

REM --- 4. Subir --------------------------------------------------------
echo.
git add -A
git commit -m "%DESC%"
if errorlevel 1 (
  echo.
  echo   No se pudo registrar el cambio. Nada se subio.
  echo.
  pause
  exit /b 1
)

git push -u origin %RAMA%
if errorlevel 1 (
  echo.
  echo   Se registro el cambio en tu computador pero NO se pudo enviar a GitHub.
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
