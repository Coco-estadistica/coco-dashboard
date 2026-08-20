@echo off
setlocal
REM ============================================================
REM  COCO - TRAER LA ULTIMA VERSION DESDE GITHUB
REM  Doble clic ANTES de empezar a cargar cifras.
REM  No sube nada. Solo baja.
REM ============================================================
cd /d "%~dp0"
cls
echo.
echo   ================================================================
echo     TRAER LA ULTIMA VERSION - COCO
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
  echo   Esta carpeta no es un repositorio de GitHub. Nada que actualizar.
  echo.
  pause
  exit /b 1
)
echo   Rama: %RAMA%
echo.

REM --- 1. Si hay cambios sin subir, NO se baja nada -------------------
REM     Un .xlsx es binario: Git no puede fusionarlo. Si bajaramos encima
REM     de cambios sin subir, uno de los dos archivos se perderia entero.
set HAYCAMBIOS=
for /f "delims=" %%s in ('git status --porcelain') do set HAYCAMBIOS=1
if defined HAYCAMBIOS (
  echo   ALTO. Tienes cambios en esta carpeta que aun no has subido:
  echo.
  git status --short
  echo.
  echo   Sube primero con  Subir_A_GitHub.bat  y vuelve a ejecutar este boton.
  echo   No se bajo nada: tus cambios estan intactos.
  echo.
  pause
  exit /b 1
)

REM --- 2. Bajar --------------------------------------------------------
echo   Consultando GitHub...
git fetch origin %RAMA%
if errorlevel 1 (
  echo.
  echo   No pude conectarme a GitHub. Revisa tu conexion y reintenta.
  echo.
  pause
  exit /b 1
)

set PENDIENTES=0
for /f %%n in ('git rev-list --count HEAD..origin/%RAMA% 2^>nul') do set PENDIENTES=%%n
if "%PENDIENTES%"=="0" (
  echo.
  echo   Ya tienes la ultima version. No habia nada nuevo.
  echo.
  pause
  exit /b 0
)

echo.
echo   Hay %PENDIENTES% actualizacion(es) nueva(s):
echo.
git log --oneline HEAD..origin/%RAMA%
echo.
git pull --ff-only origin %RAMA%
if errorlevel 1 (
  echo.
  echo   NO SE PUDO ACTUALIZAR AUTOMATICAMENTE.
  echo   Tu copia y la de GitHub siguieron caminos distintos.
  echo   NO toques nada y pide ayuda: unir archivos de Excel a mano
  echo   hace perder cifras.
  echo.
  pause
  exit /b 1
)

echo.
echo   ================================================================
echo     LISTO. Ya tienes la ultima version.
echo.
echo     Siguiente paso: Revisar_Base.bat  o  Abrir_Dashboard_COCO.bat
echo   ================================================================
echo.
pause
