@echo off
TITLE Lanzador del Instalador de Finanzas

:: -----------------------------------------------------------------------------
:: Este script ejecuta el instalador de PowerShell con los permisos correctos.
:: -----------------------------------------------------------------------------

echo =================================================================
echo ==    Lanzador para el Instalador del Sistema de Finanzas      ==
echo =================================================================
echo.
echo Este programa ejecutara el instalador principal (instalador_finanzas.ps1).
echo.
echo IMPORTANTE:
echo Si no has ejecutado este archivo como Administrador, por favor,
echo cierra esta ventana, haz clic derecho sobre "lanzador.bat"
echo y selecciona "Ejecutar como administrador".
echo.

REM --- Verifica si tiene permisos de administrador ---
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo.
    echo ******************************************************
    echo ** ERROR: Se requieren permisos de Administrador. **
    echo ******************************************************
    echo Por favor, vuelve a ejecutar como Administrador.
    echo.
    pause
    exit
)

echo Permisos de Administrador detectados. Continuando...
echo Presiona cualquier tecla para iniciar la instalacion.
pause
echo.

:: --- Ejecuta el script de PowerShell ---
:: -ExecutionPolicy Bypass: Permite que el script se ejecute en esta sesion.
:: -File "%~dp0...": Le dice a PowerShell que ejecute tu script.
::                   %~dp0 se asegura de que encuentre el script en la misma carpeta.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalador_finanzas.ps1"

echo.
echo =================================================================
echo El script de instalacion ha finalizado.
echo Puedes cerrar esta ventana.
echo =================================================================
pause