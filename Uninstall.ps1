<#
.SYNOPSIS
    Desinstallateur M365 Monster
    Supprime les fichiers d'installation et les raccourcis.

.DESCRIPTION
    Executer en tant qu'administrateur :
    powershell -ExecutionPolicy Bypass -File .\Uninstall.ps1

    Le script se copie dans un dossier temporaire puis se relance
    depuis la-bas, afin de pouvoir supprimer son propre repertoire.

.AUTEUR
    [Equipe IT - M365 Monster]
#>

[CmdletBinding()]
param(
    [string]$InstallPath = "C:\Program Files\M365Monster",
    [switch]$RunFromTemp,
    [switch]$KeepClients
)

# === Verification admin ===
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host ""
    Write-Host "  Ce script doit etre execute en administrateur." -ForegroundColor Red
    Write-Host ""
    pause
    exit 1
}

# ================================================================
# PHASE 1 : Interaction utilisateur + copie vers temp
# ================================================================
if (-not $RunFromTemp) {

    # Si lance depuis le dossier d'installation, deduire le chemin
    $scriptDir = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($scriptDir)) {
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    }
    if ((Test-Path (Join-Path $scriptDir "Main.ps1"))) {
        $InstallPath = $scriptDir
    }

    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "  ║      M365 Monster — Desinstallation              ║" -ForegroundColor Yellow
    Write-Host "  ╚══════════════════════════════════════════════════╝" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Emplacement : $InstallPath" -ForegroundColor Gray
    Write-Host ""

    if (-not (Test-Path $InstallPath)) {
        Write-Host "  M365 Monster n'est pas installe a cet emplacement." -ForegroundColor Yellow
        pause
        exit 0
    }

    # Demander confirmation
    $confirm = Read-Host "  Confirmer la desinstallation ? (O/N)"
    if ($confirm -notmatch "^[oOyY]") {
        Write-Host "  Desinstallation annulee." -ForegroundColor Yellow
        pause
        exit 0
    }

    # Demander si on garde les configs client
    $keepConfigs = Read-Host "  Conserver les fichiers de configuration client ? (O/N)"

    # Sauvegarder les configs si demande
    if ($keepConfigs -match "^[oOyY]") {
        $backupDir = Join-Path -Path ([Environment]::GetFolderPath("Desktop")) -ChildPath "M365Monster_Configs_Backup"
        $clientsDir = Join-Path -Path $InstallPath -ChildPath "Clients"
        if (Test-Path $clientsDir) {
            Copy-Item -Path $clientsDir -Destination $backupDir -Recurse -Force
            Write-Host "  Configurations sauvegardees dans : $backupDir" -ForegroundColor Green
        }
    }

    # Copier ce script dans le dossier temp et relancer depuis la-bas
    $tempDir = Join-Path -Path $env:TEMP -ChildPath "M365Monster_Uninstall_$(Get-Date -Format 'yyyyMMddHHmmss')"
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
    $tempScript = Join-Path -Path $tempDir -ChildPath "Uninstall.ps1"
    Copy-Item -Path $MyInvocation.MyCommand.Definition -Destination $tempScript -Force

    Write-Host ""
    Write-Host "  Relance depuis le dossier temporaire..." -ForegroundColor Gray

    # Changer le repertoire courant pour liberer le verrou sur le dossier d'installation
    Set-Location -Path $env:TEMP

    # Detecter le PowerShell en cours (pwsh.exe ou powershell.exe)
    $psExe = (Get-Process -Id $PID).Path
    $psArgs = @("-ExecutionPolicy", "Bypass", "-File", $tempScript, "-InstallPath", $InstallPath, "-RunFromTemp")
    if ($keepConfigs -match "^[oOyY]") {
        $psArgs += "-KeepClients"
    }
    & $psExe @psArgs
    exit $LASTEXITCODE
}

# ================================================================
# PHASE 2 : Suppression effective (execute depuis le dossier temp)
# ================================================================

Write-Host ""
Write-Host "  Suppression des raccourcis..." -ForegroundColor White

# --- Raccourcis Bureau ---
$desktopShortcut = Join-Path -Path ([Environment]::GetFolderPath("CommonDesktopDirectory")) -ChildPath "M365 Monster.lnk"
if (Test-Path $desktopShortcut) {
    Remove-Item $desktopShortcut -Force
    Write-Host "    Raccourci Bureau supprime." -ForegroundColor Green
}

# --- Raccourcis Menu Demarrer ---
$startMenuFolder = Join-Path -Path ([Environment]::GetFolderPath("CommonStartMenu")) -ChildPath "Programs\M365 Monster"
if (Test-Path $startMenuFolder) {
    Remove-Item $startMenuFolder -Recurse -Force
    Write-Host "    Menu Demarrer supprime." -ForegroundColor Green
}

# --- Supprimer le dossier d'installation ---
Write-Host "  Suppression de $InstallPath..." -ForegroundColor White
if (Test-Path $InstallPath) {
    try {
        # Se deplacer hors du dossier pour eviter le verrou "in use"
        Set-Location -Path $env:TEMP

        if ($KeepClients) {
            # Supprimer tout sauf le dossier Clients/
            Get-ChildItem -Path $InstallPath -Force | Where-Object {
                $_.Name -ne "Clients"
            } | Remove-Item -Recurse -Force -ErrorAction Stop

            Write-Host "  Fichiers d'installation supprimes (Clients/ conserve)." -ForegroundColor Green
        }
        else {
            Remove-Item -Path $InstallPath -Recurse -Force -ErrorAction Stop
            Write-Host "  Dossier d'installation supprime." -ForegroundColor Green
        }
    }
    catch {
        Write-Host "  [ERREUR] $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Supprimez manuellement : $InstallPath" -ForegroundColor Yellow
    }
}

# --- Supprimer les donnees utilisateur dans AppData ---
$appDataDir = Join-Path -Path $env:APPDATA -ChildPath "M365Monster"
if (Test-Path $appDataDir) {
    try {
        Remove-Item -Path $appDataDir -Recurse -Force -ErrorAction Stop
        Write-Host "  Donnees utilisateur (AppData) supprimees." -ForegroundColor Green
    }
    catch {
        Write-Host "  [AVERTISSEMENT] Impossible de supprimer $appDataDir" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "  ║      M365 Monster desinstalle avec succes !      ║" -ForegroundColor Green
Write-Host "  ╚══════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
pause
