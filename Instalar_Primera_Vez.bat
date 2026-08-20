@echo off
setlocal
REM ============================================================
REM  COCO - INSTALACION (se hace UNA SOLA VEZ)
REM  Crea la carpeta de trabajo conectada a GitHub.
REM  No borra ni toca tu carpeta actual de OneDrive.
REM ============================================================
cls
echo.
echo   ================================================================
echo     INSTALACION DE LA CARPETA DE TRABAJO - COCO
echo     (esto se hace una sola vez)
echo   ================================================================
echo.

where git >nul 2>&1
if errorlevel 1 (
  echo   NO ENCUENTRO GIT en este computador.
  echo   Instalalo desde https://git-scm.com/download/win
  echo   Acepta todas las opciones por defecto, reinicia el computador
  echo   y vuelve a hacer doble clic aqui.
  echo.
  pause
  exit /b 1
)

cd /d "%~dp0"
echo   Se va a crear la carpeta:
echo      %CD%\coco-dashboard
echo.

REM --- Aviso si estamos dentro de OneDrive -----------------------------
echo %CD% | find /i "OneDrive" >nul
if not errorlevel 1 (
  echo   ------------------------------------------------------------
  echo   AVISO IMPORTANTE
  echo.
  echo   Estas dentro de OneDrive. No es buena idea: OneDrive y GitHub
  echo   se pelean por los mismos archivos y la carpeta se puede danar.
  echo.
  echo   Recomendado: cancela, mueve este archivo a  C:\COCO
  echo   y vuelve a hacer doble clic alli.
  echo   ------------------------------------------------------------
  echo.
)

if exist "coco-dashboard" (
  echo   Ya existe una carpeta llamada coco-dashboard aqui.
  echo   No hago nada para no pisarla. Si quieres empezar de cero,
  echo   renombrala o muevela y vuelve a intentar.
  echo.
  pause
  exit /b 1
)

set "OK="
set /p "OK=  Continuo con la instalacion? (S/N): "
if /i not "%OK%"=="S" (
  echo.
  echo   Cancelado. No se hizo nada.
  echo.
  pause
  exit /b 0
)

echo.
echo   Descargando el proyecto desde GitHub. Puede tardar un minuto...
echo.
git clone --branch claude/coco-dashboard-cloud-setup-sbqa86 https://github.com/Moscorrofio/coco-dashboard.git coco-dashboard
if errorlevel 1 (
  echo.
  echo   NO SE PUDO DESCARGAR.
  echo   Puede ser la conexion, o que GitHub te pida usuario y clave.
  echo   Copia lo que dice arriba y pidelo revisar.
  echo.
  pause
  exit /b 1
)

echo.
echo   ================================================================
echo     LISTO. Tu carpeta de trabajo quedo en:
echo        %CD%\coco-dashboard
echo   ================================================================
echo.
echo   FALTAN DOS COSAS, hazlas ahora:
echo.
echo   1. Copia la carpeta  modelo  desde tu carpeta vieja de OneDrive
echo      a la carpeta nueva. Esa carpeta no esta en GitHub y hace
echo      falta para la pestana Burn ^& Runway.
echo.
echo   2. Compara tu Excel viejo con el nuevo:
echo         viejo:  tu carpeta de OneDrive \migracion\BD_MAESTRA_COCO.xlsx
echo         nuevo:  %CD%\coco-dashboard\migracion\BD_MAESTRA_COCO.xlsx
echo      Si el tuyo tiene cifras mas recientes, copialo encima del nuevo
echo      y despues haz doble clic en Subir_A_GitHub.bat
echo.
echo   NO borres tu carpeta de OneDrive todavia. Dejala como respaldo
echo   hasta que compruebes que todo funciona en la carpeta nueva.
echo.
echo   Desde manana trabajas SIEMPRE en la carpeta nueva.
echo.
pause
