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
set rutaEscritorio=%USERPROFILE%\Desktop

:: ==============================
:: VERIFICAR ADMINISTRADOR
:: ==============================
net session >nul 2>&1
if %errorlevel% neq 0 (
  echo ERROR: Este script requiere permisos de Administrador.
  echo Por favor, ejecuta el script como administrador.
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
  
  echo Descargando Git...
  powershell -Command "Invoke-WebRequest -Uri '%gitInstallerURL%' -OutFile '%installerPath%'"
  
  echo Instalando Git...
  start /wait "" "%installerPath%" /SILENT /NORESTART /CLOSEAPPLICATIONS /SUPPRESSMSGBOXES
  
  :: Agregar Git al PATH
  setx PATH "%PATH%;C:\Program Files\Git\cmd" /M
)

:: ==============================
:: CONFIGURAR GIT
:: ==============================
echo Configurando Git...
git config --global user.name "%userName%"
git config --global user.email "%userEmail%"

:: Agregar el directorio como seguro (esto soluciona el error de ownership)
git config --global --add safe.directory "%rutaDestino%"

:: ==============================
:: CLONAR O ACTUALIZAR REPOSITORIO
:: ==============================
if exist "%rutaDestino%\.git" (
  echo La carpeta del repositorio ya existe. Actualizando...
  cd /d "%rutaDestino%"
  
  :: Verificar si hay cambios locales antes de hacer pull
  git fetch origin
  git status | findstr "modified" >nul
  if %errorlevel% equ 0 (
    echo Se detectaron cambios locales. Haciendo backup...
    set backupDir=%rutaDestino%-backup-!DATE!/!TIME!
    set backupDir=!backupDir:/=-!
    set backupDir=!backupDir::=-!
    xcopy "%rutaDestino%" "!backupDir!" /E /I /H /Y
    echo Backup creado en: !backupDir!
  )
  
  git pull origin main
) else (
  echo Clonando repositorio...
  if exist "%rutaDestino%" (
    echo Eliminando directorio existente...
    rmdir /s /q "%rutaDestino%"
  )
  git clone %repoURL% "%rutaDestino%"
)

:: Verificar si el clonado fue exitoso
if %errorlevel% neq 0 (
  echo ERROR: No se pudo clonar/actualizar el repositorio.
  echo Verifica tu conexión a Internet y los permisos del directorio.
  pause
  exit /b
)

:: ==============================
:: BUSCAR ARCHIVO EXCEL PRINCIPAL
:: ==============================
echo Buscando archivo Excel principal...
set rutaExcel=
for /f "delims=" %%i in ('dir /s /b "%rutaDestino%\*.xlsm" 2^>nul') do (
  set "rutaExcel=%%i"
  goto :encontrado
)

:encontrado
if "%rutaExcel%"=="" (
  echo No se encontró ningún archivo .xlsm en "%rutaDestino%"
  echo Buscando archivos .xlsx...
  for /f "delims=" %%i in ('dir /s /b "%rutaDestino%\*.xlsx" 2^>nul') do (
    set "rutaExcel=%%i"
    goto :encontrado2
  )
)

:encontrado2
if "%rutaExcel%"=="" (
  echo ERROR: No se encontró ningún archivo Excel en el repositorio.
  pause
  exit /b
)

echo Archivo Excel encontrado: %rutaExcel%

:: ==============================
:: VERIFICAR QUE EXCEL ESTÉ INSTALADO
:: ==============================
echo Verificando instalación de Microsoft Excel...
set excelPath=
for %%d in (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
  if exist "%%d:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" (
    set "excelPath=%%d:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
  ) else if exist "%%d:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE" (
    set "excelPath=%%d:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
  ) else if exist "%%d:\Program Files\Microsoft Office\Office16\EXCEL.EXE" (
    set "excelPath=%%d:\Program Files\Microsoft Office\Office16\EXCEL.EXE"
  ) else if exist "%%d:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE" (
    set "excelPath=%%d:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE"
  )
)

if "!excelPath!"=="" (
  echo ERROR: No se pudo encontrar Microsoft Excel instalado en el sistema.
  echo Es necesario tener Microsoft Excel 2016 o superior instalado.
  pause
  exit /b 1
)

echo Excel encontrado en: !excelPath!

:: ==============================
:: CREAR ACCESO DIRECTO
:: ==============================
echo Creando acceso directo en el escritorio...
set "psCommand=$WshShell = New-Object -ComObject WScript.Shell; $lnk = $WshShell.CreateShortcut('%rutaEscritorio%\%nombreAccesoDirecto%'); $lnk.TargetPath = '!excelPath!'; $lnk.Arguments = '"%rutaExcel%"'; $lnk.WorkingDirectory = '%rutaDestino%'; $lnk.IconLocation = '!excelPath!'; $lnk.Save()"

powershell -Command "& {%psCommand%}"

if exist "%rutaEscritorio%\%nombreAccesoDirecto%" (
    echo Acceso directo creado exitosamente en:
    echo %rutaEscritorio%\%nombreAccesoDirecto%
) else (
    echo Error: No se pudo crear el acceso directo.
    echo Intentando método alternativo...
    
    :: Método alternativo usando mklink (simbólico)
    mklink "%rutaEscritorio%\%nombreAccesoDirecto%" "%rutaExcel%" >nul 2>&1
    if %errorlevel% equ 0 (
        echo Acceso directo creado usando método alternativo.
    ) else (
        echo No se pudo crear el acceso directo con ningún método.
    )
)

:: ==============================
:: CONFIGURACIÓN ADICIONAL
:: ==============================
echo Configurando permisos del directorio...
icacls "%rutaDestino%" /grant "%USERNAME%":F /T >nul 2>&1

:: ==============================
:: FINALIZADO
:: ==============================
echo.
echo =========================================
echo   INSTALACION COMPLETADA EXITOSAMENTE
echo =========================================
echo.
echo Resumen:
echo - Repositorio clonado/actualizado en: %rutaDestino%
echo - Acceso directo creado en: %rutaEscritorio%\%nombreAccesoDirecto%
echo - Archivo Excel principal: %rutaExcel%
echo.
echo Ejecuta el acceso directo para iniciar el sistema.
echo.
pause