<#
.SYNOPSIS
    Installateur M365 Monster
    Copie les fichiers, installe les prerequis, crée les raccourcis
    et configure la mise a jour automatique depuis GitHub.

.DESCRIPTION
    Exécuter en tant qu'administrateur :
    powershell -ExecutionPolicy Bypass -File .\Install.ps1

    Options :
    -InstallPath "C:\Chemin\Custom"   Chemin d'installation personnalise
    -SkipModules                       Ne pas installer les modules PowerShell
    -SkipShortcuts                     Ne pas creer les raccourcis
    -SkipUpdateConfig                  Ne pas configurer l'auto-update

.AUTEUR
    [Equipe IT - M365 Monster]
#>

[CmdletBinding()]
param(
    [string]$InstallPath = "C:\Program Files\M365Monster",
    [switch]$SkipModules,
    [switch]$SkipShortcuts,
    [switch]$SkipUpdateConfig
)

# === Verification des droits administrateur ===
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host ""
    Write-Host "  =============================================" -ForegroundColor Red
    Write-Host "  Ce script doit etre execute en administrateur" -ForegroundColor Red
    Write-Host "  =============================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Clic-droit sur PowerShell > Executer en tant qu'administrateur" -ForegroundColor Yellow
    Write-Host ""
    pause
    exit 1
}

# === Banniere ===
Clear-Host
Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════╗" -ForegroundColor Magenta
Write-Host "  ║                                                  ║" -ForegroundColor Magenta
Write-Host "  ║           M365 Monster — Installateur            ║" -ForegroundColor Magenta
Write-Host "  ║           v1.0.0                                 ║" -ForegroundColor Magenta
Write-Host "  ║                                                  ║" -ForegroundColor Magenta
Write-Host "  ╚══════════════════════════════════════════════════╝" -ForegroundColor Magenta
Write-Host ""

# === Determination du dossier source ===
$sourcePath = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($sourcePath)) {
    $sourcePath = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

if (-not (Test-Path (Join-Path $sourcePath "Main.ps1"))) {
    Write-Host "  [ERREUR] Main.ps1 introuvable dans le dossier source." -ForegroundColor Red
    Write-Host "  Executez Install.ps1 depuis le dossier du projet." -ForegroundColor Yellow
    pause
    exit 1
}

Write-Host "  Source          : $sourcePath" -ForegroundColor Gray
Write-Host "  Destination     : $InstallPath" -ForegroundColor Gray
Write-Host ""

# ================================================================
# ETAPE 1 : Modules PowerShell
# ================================================================
if (-not $SkipModules) {
    Write-Host "  [1/5] Verification des modules PowerShell..." -ForegroundColor White

    $requiredModules = @("Microsoft.Graph")

    foreach ($moduleName in $requiredModules) {
        $installed = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue
        if ($installed) {
            $ver = ($installed | Sort-Object Version -Descending | Select-Object -First 1).Version
            Write-Host "        $moduleName v$ver — deja installe." -ForegroundColor Green
        }
        else {
            Write-Host "        Installation de $moduleName (peut prendre quelques minutes)..." -ForegroundColor Yellow
            try {
                Install-Module -Name $moduleName -Scope AllUsers -Force -AllowClobber -ErrorAction Stop
                Write-Host "        $moduleName installe avec succes." -ForegroundColor Green
            }
            catch {
                Write-Host "        [AVERTISSEMENT] Echec : $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "        Installez manuellement : Install-Module $moduleName -Scope CurrentUser -Force" -ForegroundColor Yellow
            }
        }
    }
}
else {
    Write-Host "  [1/5] Modules — ignore (-SkipModules)." -ForegroundColor DarkGray
}
Write-Host ""

# ================================================================
# ETAPE 2 : Copie des fichiers
# ================================================================
Write-Host "  [2/5] Copie des fichiers..." -ForegroundColor White

try {
    if (-not (Test-Path $InstallPath)) {
        New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null
    }

    # Elements a copier (pas de Clients/*.json sauf le template)
    $itemsToCopy = @(
        "Main.ps1",
        "version.json",
        "update_config.json",
        "Core",
        "Modules",
        "Scripts",
        "Assets",
        "Lang"
    )

    foreach ($item in $itemsToCopy) {
        $src = Join-Path -Path $sourcePath -ChildPath $item
        $dst = Join-Path -Path $installPath -ChildPath $item

        if (Test-Path $src) {
            if ((Get-Item $src).PSIsContainer) {
                if (Test-Path $dst) { Remove-Item $dst -Recurse -Force }
                Copy-Item -Path $src -Destination $dst -Recurse -Force
                $count = (Get-ChildItem $dst -Recurse -File).Count
                Write-Host "        $item/ ($count fichiers)" -ForegroundColor Green
            }
            else {
                Copy-Item -Path $src -Destination $dst -Force
                Write-Host "        $item" -ForegroundColor Green
            }
        }
    }

    # Copier Uninstall.ps1 s'il existe
    $uninstallSrc = Join-Path -Path $sourcePath -ChildPath "Uninstall.ps1"
    if (Test-Path $uninstallSrc) {
        Copy-Item -Path $uninstallSrc -Destination $InstallPath -Force
        Write-Host "        Uninstall.ps1" -ForegroundColor Green
    }

    # Creer le dossier Clients/ avec le template uniquement
    $clientsDir = Join-Path -Path $InstallPath -ChildPath "Clients"
    if (-not (Test-Path $clientsDir)) {
        New-Item -Path $clientsDir -ItemType Directory -Force | Out-Null
    }
    $templateSrc = Join-Path -Path $sourcePath -ChildPath "Clients\_Template.json"
    if (Test-Path $templateSrc) {
        Copy-Item -Path $templateSrc -Destination $clientsDir -Force
        Write-Host "        Clients/_Template.json" -ForegroundColor Green
    }

    # Donner les droits d'ecriture sur Clients/ aux utilisateurs
    # Utilise icacls avec le SID S-1-5-32-545 (BUILTIN\Users / BUILTIN\Utilisateurs)
    # pour fonctionner quelle que soit la langue de Windows
    try {
        $icaclsResult = icacls $clientsDir /grant "*S-1-5-32-545:(OI)(CI)M" /T 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "        Permissions Clients/ configurees (ecriture utilisateurs)." -ForegroundColor Green
        }
        else {
            Write-Host "        [AVERTISSEMENT] icacls a retourne une erreur : $icaclsResult" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "        [AVERTISSEMENT] Impossible de configurer les permissions sur Clients/" -ForegroundColor Yellow
        Write-Host "        Les utilisateurs devront editer les fichiers en administrateur." -ForegroundColor Yellow
    }

    # Debloquer tous les fichiers
    Get-ChildItem -Path $InstallPath -Recurse -File | Unblock-File -ErrorAction SilentlyContinue

    Write-Host "        Copie terminee." -ForegroundColor Green
}
catch {
    Write-Host "  [ERREUR] Echec de la copie : $($_.Exception.Message)" -ForegroundColor Red
    pause
    exit 1
}
Write-Host ""

# ================================================================
# ETAPE 3 : Raccourcis (Bureau + Menu Demarrer)
# ================================================================
if (-not $SkipShortcuts) {
    Write-Host "  [3/5] Creation des raccourcis..." -ForegroundColor White

    $mainScript = Join-Path -Path $InstallPath -ChildPath "Main.ps1"
    $iconFile = Join-Path -Path $InstallPath -ChildPath "Assets\M365Monster.ico"

    # Detecter PowerShell 7 (pwsh.exe) en priorite, sinon fallback sur powershell.exe (5.1)
    $pwshPath = Get-Command "pwsh.exe" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1
    if ($pwshPath) {
        $targetPath = $pwshPath
        Write-Host "        PowerShell 7 detecte : $pwshPath" -ForegroundColor Green
    }
    else {
        $targetPath = "powershell.exe"
        Write-Host "        PowerShell 7 non trouve, utilisation de PowerShell 5.1" -ForegroundColor Yellow
        Write-Host "        (Recommande : installer PowerShell 7 pour une meilleure compatibilite)" -ForegroundColor Yellow
    }
    $arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$mainScript`""

    $shell = New-Object -ComObject WScript.Shell

    # --- Raccourci Bureau ---
    try {
        $desktopPath = [Environment]::GetFolderPath("CommonDesktopDirectory")
        $shortcutPath = Join-Path -Path $desktopPath -ChildPath "M365 Monster.lnk"
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $targetPath
        $shortcut.Arguments = $arguments
        $shortcut.WorkingDirectory = $InstallPath
        if (Test-Path $iconFile) {
            $shortcut.IconLocation = "$iconFile,0"
        }
        $shortcut.Description = "M365 Monster — Gestion du cycle de vie employe"
        $shortcut.Save()
        Write-Host "        Bureau          : $shortcutPath" -ForegroundColor Green
    }
    catch {
        Write-Host "        [AVERTISSEMENT] Raccourci Bureau non cree : $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # --- Raccourci Menu Demarrer ---
    try {
        $startMenuPath = Join-Path -Path ([Environment]::GetFolderPath("CommonStartMenu")) -ChildPath "Programs"
        $m365Folder = Join-Path -Path $startMenuPath -ChildPath "M365 Monster"
        if (-not (Test-Path $m365Folder)) {
            New-Item -Path $m365Folder -ItemType Directory -Force | Out-Null
        }

        # Raccourci principal
        $shortcutPath = Join-Path -Path $m365Folder -ChildPath "M365 Monster.lnk"
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $targetPath
        $shortcut.Arguments = $arguments
        $shortcut.WorkingDirectory = $InstallPath
        if (Test-Path $iconFile) {
            $shortcut.IconLocation = "$iconFile,0"
        }
        $shortcut.Description = "M365 Monster — Gestion du cycle de vie employe"
        $shortcut.Save()
        Write-Host "        Menu Demarrer   : $shortcutPath" -ForegroundColor Green

        # Raccourci desinstallation
        $uninstallScript = Join-Path -Path $InstallPath -ChildPath "Uninstall.ps1"
        if (Test-Path $uninstallScript) {
            $uninstShortcut = Join-Path -Path $m365Folder -ChildPath "Desinstaller M365 Monster.lnk"
            $sc = $shell.CreateShortcut($uninstShortcut)
            $sc.TargetPath = $targetPath
            $sc.Arguments = "-ExecutionPolicy Bypass -File `"$uninstallScript`""
            $sc.WorkingDirectory = $InstallPath
            $sc.Description = "Desinstaller M365 Monster"
            $sc.Save()
            Write-Host "        Desinstallateur : $uninstShortcut" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "        [AVERTISSEMENT] Menu Demarrer non cree : $($_.Exception.Message)" -ForegroundColor Yellow
    }

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
}
else {
    Write-Host "  [3/5] Raccourcis — ignore (-SkipShortcuts)." -ForegroundColor DarkGray
}
Write-Host ""

# ================================================================
# ETAPE 4 : Configuration de la mise a jour automatique
# ================================================================
if (-not $SkipUpdateConfig) {
    Write-Host "  [4/5] Configuration de la mise a jour automatique" -ForegroundColor White
    Write-Host ""
    Write-Host "        M365 Monster peut verifier les mises a jour depuis GitHub" -ForegroundColor Gray
    Write-Host "        a chaque lancement. Pour cela, indiquez votre repo." -ForegroundColor Gray
    Write-Host ""

    $configureUpdate = Read-Host "        Configurer l'auto-update maintenant ? (O/N)"

    if ($configureUpdate -match "^[oOyY]") {
        Write-Host ""
        Write-Host "        Format du repo : proprietaire/nom-du-repo" -ForegroundColor Gray
        Write-Host "        Exemple : monorg/M365Monster" -ForegroundColor Gray
        $githubRepo = Read-Host "        Repo GitHub"

        $branch = Read-Host "        Branche (defaut: main)"
        if ([string]::IsNullOrWhiteSpace($branch)) { $branch = "main" }

        Write-Host ""
        Write-Host "        Pour un repo prive, un Personal Access Token est necessaire." -ForegroundColor Gray
        Write-Host "        Pour un repo public, laissez vide." -ForegroundColor Gray
        $githubToken = Read-Host "        GitHub Token (ou vide)"

        Write-Host ""
        Write-Host "        URL de telechargement du .zip de mise a jour." -ForegroundColor Gray
        Write-Host "        Defaut : GitHub Releases (latest/M365Monster.zip)" -ForegroundColor Gray
        $downloadUrl = Read-Host "        URL du .zip (ou vide pour le defaut)"

        $checkInterval = Read-Host "        Intervalle de verification en heures (defaut: 24)"
        if ([string]::IsNullOrWhiteSpace($checkInterval)) { $checkInterval = "24" }

        # Construire et ecrire le fichier
        $updateConfig = [ordered]@{
            github_repo          = $githubRepo
            branch               = $branch
            github_token         = $githubToken
            download_url         = $downloadUrl
            check_interval_hours = [int]$checkInterval
        }

        $updateConfigPath = Join-Path -Path $InstallPath -ChildPath "update_config.json"
        $updateConfig | ConvertTo-Json -Depth 3 | Out-File -FilePath $updateConfigPath -Encoding UTF8 -Force

        Write-Host ""
        if (-not [string]::IsNullOrWhiteSpace($githubRepo)) {
            Write-Host "        Auto-update configure pour : $githubRepo ($branch)" -ForegroundColor Green
        }
        else {
            Write-Host "        Auto-update desactive (repo vide)." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "        Auto-update non configure. Vous pourrez editer" -ForegroundColor Yellow
        Write-Host "        update_config.json dans le dossier d'installation." -ForegroundColor Yellow
    }
}
else {
    Write-Host "  [4/5] Auto-update — ignore (-SkipUpdateConfig)." -ForegroundColor DarkGray
}
Write-Host ""

# ================================================================
# ETAPE 5 : Resume
# ================================================================
Write-Host "  [5/5] Installation terminee !" -ForegroundColor Green
Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║                                                             ║" -ForegroundColor Cyan
Write-Host "  ║   M365 Monster installe avec succes !                       ║" -ForegroundColor Cyan
Write-Host "  ║                                                             ║" -ForegroundColor Cyan
Write-Host "  ╠══════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan
Write-Host "  ║                                                             ║" -ForegroundColor Cyan
Write-Host "  ║   Emplacement : $($InstallPath.PadRight(44))║" -ForegroundColor Cyan
Write-Host "  ║                                                             ║" -ForegroundColor Cyan
Write-Host "  ║   Prochaines etapes :                                       ║" -ForegroundColor Cyan
Write-Host "  ║                                                             ║" -ForegroundColor Cyan
Write-Host "  ║   1. Ouvrez le dossier Clients\                             ║" -ForegroundColor Cyan
Write-Host "  ║   2. Copiez _Template.json → VotreClient.json               ║" -ForegroundColor Cyan
Write-Host "  ║   3. Renseignez tenant_id, client_id, smtp_domain           ║" -ForegroundColor Cyan
Write-Host "  ║   4. Lancez M365 Monster depuis le Bureau                   ║" -ForegroundColor Cyan
Write-Host "  ║                                                             ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# Proposer d'ouvrir le dossier Clients
$openFolder = Read-Host "  Ouvrir le dossier Clients pour configurer votre premier client ? (O/N)"
if ($openFolder -match "^[oOyY]") {
    $clientsDir = Join-Path -Path $InstallPath -ChildPath "Clients"
    Start-Process "explorer.exe" -ArgumentList $clientsDir
}

Write-Host ""
pause
