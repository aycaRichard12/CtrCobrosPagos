# INSTALADOR AUTOMÁTICO - Sistema de Finanzas
# CONFIGURACIÓN: Edita estas variables con tu información
$repoURL = "https://github.com/tuusuario/turepositorio.git"
$userName = "Nombre del Usuario"
$userEmail = "usuario@ejemplo.com"
$rutaDestino = "C:\Sistema-Finanzas" # Carpeta donde se clonará el proyecto
$nombreAccesoDirecto = "Sistema de Finanzas.lnk"

# --- NO EDITAR A PARTIR DE AQUÍ (a menos que sepas lo que haces) ---

# Función para escribir mensajes en color
function Write-ColorOutput($ForegroundColor) {
    $fc = $Host.UI.RawUI.ForegroundColor
    $Host.UI.RawUI.ForegroundColor = $ForegroundColor
    if ($args) {
        Write-Output $args
    }
    $Host.UI.RawUI.ForegroundColor = $fc
}

Write-ColorOutput Green "========================================="
Write-ColorOutput Green "  INSTALADOR DEL SISTEMA DE FINANZAS"
Write-ColorOutput Green "========================================="

# 1. VERIFICAR/INSTALAR GIT
Write-Output "`n[1/4] Verificando si Git está instalado..."
$gitPath = Get-Command git -ErrorAction SilentlyContinue

if (-not $gitPath) {
    Write-Output "Git no encontrado. Procediendo con la instalación..."
    
    # URL de descarga del instalador oficial de Git para Windows (64-bit)
    $gitInstallerURL = "https://github.com/git-for-windows/git/releases/download/v2.43.0.windows.1/Git-2.43.0-64-bit.exe"
    $installerPath = "$env:TEMP\Git-Installer.exe"
    
    # Descargar el instalador
    Write-Output "Descargando Git... (Esto puede tomar unos minutos)"
    Invoke-WebRequest -Uri $gitInstallerURL -OutFile $installerPath
    
    # Instalar Git en modo silencioso
    Write-Output "Instalando Git silenciosamente..."
    $installArgs = "/SILENT /NORESTART /CLOSEAPPLICATIONS /SUPPRESSMSGBOXES"
    Start-Process -FilePath $installerPath -ArgumentList $installArgs -Wait
    
    # Agregar Git al PATH del sistema (necesario para que esté disponible inmediatamente)
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    Write-Output "Git instalado correctamente."
} else {
    Write-Output "Git ya está instalado. Continuando..."
}

# 2. CONFIGURAR GIT
Write-Output "`n[2/4] Configurando Git con tu nombre y email..."
git config --global user.name "$userName"
git config --global user.email "$userEmail"

# 3. CLONAR EL REPOSITORIO
Write-Output "`n[3/4] Clonando el repositorio del sistema..."
if (Test-Path $rutaDestino) {
    # Si la carpeta ya existe, mejor hacer un pull para actualizar
    Set-Location $rutaDestino
    git pull origin main
} else {
    # Clonar el repositorio por primera vez
    git clone $repoURL $rutaDestino
}

# 4. CREAR ACCESO DIRECTO en el Escritorio
Write-Output "`n[4/4] Creando acceso directo en el escritorio..."
$rutaExcel = (Get-ChildItem -Path $rutaDestino -Filter "*.xlsm" | Select-Object -First 1).FullName
$rutaEscritorio = [Environment]::GetFolderPath("Desktop")
$shell = New-Object -ComObject WScript.Shell
$accesoDirecto = $shell.CreateShortcut("$rutaEscritorio\$nombreAccesoDirecto")
$accesoDirecto.TargetPath = "excel.exe"
$accesoDirecto.Arguments = "`"$rutaExcel`""
$accesoDirecto.WorkingDirectory = $rutaDestino
$accesoDirecto.IconLocation = "excel.exe, 0"
$accesoDirecto.Save()

Write-ColorOutput Green "`n========================================="
Write-ColorOutput Green "   INSTALACIÓN COMPLETADA EXITOSAMENTE"
Write-ColorOutput Green "========================================="
Write-Output "`nEl acceso directo '$nombreAccesoDirecto' ha sido creado en tu escritorio."
Write-Output "`nPor favor, CIERRA todas las ventanas y REINICIA tu computadora para que todos los cambios surtan efecto."
Write-Output "Después del reinicio, usa el acceso directo para abrir el sistema."
pause