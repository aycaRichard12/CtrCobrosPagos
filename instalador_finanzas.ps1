$repoURL = "https://github.com/aycaRichard12/CtrCobrosPagos.git"
$userName = "aycaRichard12"
$userEmail = "aycarichard7@gmail.com"
$rutaDestino = "C:\Sistema-Finanzas" 
$nombreAccesoDirecto = "Sistema de Finanzas.lnk"

function Write-ColoredMessage {
    param(
        [string]$Message,
        [System.ConsoleColor]$Color
    )
    Write-Host $Message -ForegroundColor $Color
}

if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-ColoredMessage -Message "ERROR: Este script requiere permisos de Administrador." -Color Red
    Write-ColoredMessage -Message "Por favor, ejecuta el 'lanzador.bat' como administrador." -Color Red
    pause
    exit
}

Clear-Host
Write-ColoredMessage -Message "=========================================" -Color Green
Write-ColoredMessage -Message "  INSTALADOR DEL SISTEMA DE FINANZAS" -Color Green
Write-ColoredMessage -Message "=========================================" -Color Green

try {    
    Write-ColoredMessage -Message "`n[1/5] Verificando si Git está instalado..." -Color Cyan
    $gitPath = Get-Command git -ErrorAction SilentlyContinue
    if (-not $gitPath) {
        Write-Host "Git no encontrado. Se procederá con la instalación..."
        $gitInstallerURL = "https://github.com/git-for-windows/git/releases/download/v2.46.0.windows.1/Git-2.46.0-64-bit.exe"
        $installerPath = Join-Path $env:TEMP "Git-Installer.exe"
        Write-Host "Descargando Git... (Esto puede tomar unos minutos)"
        Invoke-WebRequest -Uri $gitInstallerURL -OutFile $installerPath
        Write-Host "Instalando Git silenciosamente... Por favor, espera."
        $installArgs = "/SILENT /NORESTART /CLOSEAPPLICATIONS /SUPPRESSMSGBOXES"
        Start-Process -FilePath $installerPath -ArgumentList $installArgs -Wait
        Write-ColoredMessage -Message "Git ha sido instalado. Actualizando la variable de entorno PATH..." -Color Green
        $gitInstallPath = "$env:ProgramFiles\Git\cmd"
        $env:Path += ";$gitInstallPath"
        $gitPath = Get-Command git -ErrorAction SilentlyContinue
        if (-not $gitPath) {
            throw "No se pudo encontrar Git después de la instalación. Por favor, reinicia la terminal y vuelve a ejecutar."
        }
    } else {
        Write-ColoredMessage -Message "Git ya está instalado. Continuando..." -Color Green
    }
    Write-ColoredMessage -Message "`n[2/5] Configurando Git con tu nombre y email..." -Color Cyan
    git config --global user.name "$userName"
    git config --global user.email "$userEmail"
    Write-ColoredMessage -Message "Configuración de Git completada." -Color Green
    Write-ColoredMessage -Message "`n[3/5] Clonando o actualizando el repositorio del sistema..." -Color Cyan
    if (Test-Path $rutaDestino) {
        Write-Host "La carpeta de destino ya existe. Actualizando con 'git pull'..."
        Set-Location $rutaDestino
        git pull origin main
    } else {
        Write-Host "Clonando el repositorio por primera vez en '$rutaDestino'..."
        git clone $repoURL $rutaDestino
    }
    Write-ColoredMessage -Message "Repositorio clonado/actualizado correctamente." -Color Green
    Write-ColoredMessage -Message "`n[4/5] Creando acceso directo en el escritorio..." -Color Cyan
    $rutaExcelFile = Get-ChildItem -Path $rutaDestino -Filter "*.xlsm" -Recurse | Select-Object -First 1
    if (-not $rutaExcelFile) {
        throw "No se encontró ningún archivo .xlsm en la carpeta '$rutaDestino'."
    }
    $rutaExcel = $rutaExcelFile.FullName
    $rutaEscritorio = [Environment]::GetFolderPath("Desktop")
    $shell = New-Object -ComObject WScript.Shell
    $accesoDirectoPath = Join-Path $rutaEscritorio $nombreAccesoDirecto
    $accesoDirecto = $shell.CreateShortcut($accesoDirectoPath)
    $accesoDirecto.TargetPath = $rutaExcel
    $accesoDirecto.WorkingDirectory = $rutaDestino
    $accesoDirecto.Save()
    Write-ColoredMessage -Message "Acceso directo creado en el escritorio." -Color Green
    Write-ColoredMessage -Message "`n[5/5] Finalizando..." -Color Cyan
    Write-ColoredMessage -Message "`n=========================================" -Color Green
    Write-ColoredMessage -Message "  INSTALACIÓN COMPLETADA EXITOSAMENTE" -Color Green
    Write-ColoredMessage -Message "=========================================" -Color Green
    Write-Host "`nEl acceso directo '$nombreAccesoDirecto' ha sido creado en tu escritorio."
    Write-Host "`nSe recomienda reiniciar el equipo para asegurar que todos los cambios se apliquen correctamente."
    pause
}
catch {
    Write-ColoredMessage -Message "`n-----------------------------------------" -Color Red
    Write-ColoredMessage -Message "          OCURRIÓ UN ERROR" -Color Red
    Write-ColoredMessage -Message "-----------------------------------------" -Color Red
    Write-ColoredMessage -Message "MENSAJE: $($_.Exception.Message)" -Color Yellow
    Write-ColoredMessage -Message "LÍNEA: $($_.InvocationInfo.ScriptLineNumber)" -Color Yellow
    Write-ColoredMessage -Message "`nLa instalación no pudo completarse." -Color Red
    pause
}