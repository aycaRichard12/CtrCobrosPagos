@echo off
setlocal enabledelayedexpansion

:: ==============================
:: CONFIGURACIÓN
:: ==============================
set repoURL=https://github.com/aycaRichard12/CtrCobrosPagos.git
set rutaDestino=C:\Sistema-Finanzas
set nombreAccesoDirecto=Sistema de Finanzas.lnk
set userName=aycaRichard12
set userEmail=aycarichard7@gmail.com

:: ==============================
:: VERIFICAR ADMINISTRADOR
:: ==============================
net session >nul 2>&1
if %errorlevel% neq 0 (
  echo ERROR: Este script requiere permisos de Administrador.
  echo Por favor, ejecuta el 'lanzador.bat' como administrador.
  pause
  exit /b
)

cls
echo =========================================
echo   INSTALADOR DEL SISTEMA DE FINANZAS
echo =========================================

:: ==============================
:: VERIFICAR SI GIT ESTÁ INSTALADO
:: ==============================
where git >nul 2>&1
if %errorlevel% neq 0 (
  echo Git no encontrado. Instalando...
  set gitInstallerURL=https://github.com/git-for-windows/git/releases/download/v2.46.0.windows.1/Git-2.46.0-64-bit.exe
  set installerPath=%TEMP%\Git-Installer.exe
  powershell -Command "Invoke-WebRequest -Uri %gitInstallerURL% -OutFile %installerPath%"
  start /wait "" "%installerPath%" /SILENT /NORESTART /CLOSEAPPLICATIONS /SUPPRESSMSGBOXES
)

:: ==============================
:: CONFIGURAR GIT
:: ==============================
git config --global user.name "%userName%"
git config --global user.email "%userEmail%"

:: ==============================
:: CLONAR O ACTUALIZAR REPOSITORIO
:: ==============================
if exist "%rutaDestino%" (
  echo La carpeta ya existe. Actualizando...
  cd /d "%rutaDestino%"
  git pull origin main
) else (
  echo Clonando repositorio...
  git clone %repoURL% "%rutaDestino%"
)

:: ==============================
:: CREAR ACCESO DIRECTO (usando PowerShell)
:: ==============================
for /f "delims=" %%i in ('powershell -Command "Get-ChildItem -Path \"%rutaDestino%\" -Filter *.xlsm -Recurse | Select-Object -First 1 | ForEach-Object { $_.FullName }"') do (
  set rutaExcel=%%i
)

if "%rutaExcel%"=="" (
  echo No se encontró ningún archivo .xlsm en "%rutaDestino%"
  pause
  exit /b
)

:: Verificar que existe el ejecutable de Excel
if not exist "%rutaExcel%" (
    echo Error: No se encuentra Excel en la ruta especificada.
    echo Buscando instalaciones alternativas de Excel...
    
    :: Búsqueda alternativa de Excel
    set "rutaExcel="
    for %%d in (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
        if exist "%%d:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" (
            set "rutaExcel=%%d:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
        ) else if exist "%%d:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE" (
            set "rutaExcel=%%d:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
        )
    )
    
    if "!rutaExcel!"=="" (
        echo No se pudo encontrar ninguna instalación de Excel.
        pause
        exit /b 1
    )
)

:: Crear el acceso directo
set "psCommand=$WshShell = New-Object -ComObject WScript.Shell; $lnk = $WshShell.CreateShortcut('%rutaEscritorio%\%nombreAccesoDirecto%'); $lnk.TargetPath = '%rutaExcel%'; $lnk.WorkingDirectory = '%rutaDestino%'; $lnk.Save()"

powershell -Command "& {%psCommand%}"

if exist "%rutaEscritorio%\%nombreAccesoDirecto%" (
    echo Acceso directo creado exitosamente en:
    echo %rutaEscritorio%\%nombreAccesoDirecto%
) else (
    echo Error: No se pudo crear el acceso directo.
)

pause

:: ==============================
:: FINALIZADO
:: ==============================
echo.
echo =========================================
echo   INSTALACION COMPLETADA EXITOSAMENTE
echo =========================================
echo El acceso directo "%nombreAccesoDirecto%" ha sido creado en tu escritorio.
pause
