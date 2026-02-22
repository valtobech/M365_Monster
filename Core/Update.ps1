<#
.FICHIER
    Core/Update.ps1

.ROLE
    Vérification automatique des mises à jour au lancement.
    Compare la version locale (version.json) avec la version distante sur GitHub.
    Si une nouvelle version est disponible, propose le téléchargement et l'installation.

    Supporte :
    - Repo GitHub public (téléchargement direct)
    - Repo GitHub privé (avec Personal Access Token)

    L'URL du repo est configurable dans update_config.json à la racine du projet.
    Si le fichier n'existe pas ou si l'URL est vide, la vérification est silencieusement ignorée.

.DEPENDANCES
    - version.json (racine du projet)
    - update_config.json (racine du projet, optionnel)

.AUTEUR
    [Equipe IT - M365 Monster]
#>

function Get-LocalVersion {
    <#
    .SYNOPSIS
        Lit la version locale depuis version.json.
    .OUTPUTS
        [PSCustomObject] ou $null
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $versionFile = Join-Path -Path $RootPath -ChildPath "version.json"
    if (-not (Test-Path -Path $versionFile)) {
        return $null
    }

    try {
        $content = Get-Content -Path $versionFile -Raw -Encoding UTF8 | ConvertFrom-Json
        return $content
    }
    catch {
        return $null
    }
}

function Get-UpdateConfig {
    <#
    .SYNOPSIS
        Lit la configuration de mise à jour depuis update_config.json.
    .OUTPUTS
        [PSCustomObject] ou $null
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $configFile = Join-Path -Path $RootPath -ChildPath "update_config.json"
    if (-not (Test-Path -Path $configFile)) {
        return $null
    }

    try {
        $content = Get-Content -Path $configFile -Raw -Encoding UTF8 | ConvertFrom-Json
        return $content
    }
    catch {
        return $null
    }
}

function Test-UpdateAvailable {
    <#
    .SYNOPSIS
        Vérifie si une mise à jour est disponible sur GitHub.
        Compare la version locale avec le version.json distant.

    .PARAMETER RootPath
        Chemin racine du projet.

    .OUTPUTS
        [PSCustomObject] — {UpdateAvailable: bool, RemoteVersion: string, ReleaseNotes: string, DownloadUrl: string}
                           ou $null si la vérification échoue ou est désactivée.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    # Charger la config de mise à jour
    $updateConfig = Get-UpdateConfig -RootPath $RootPath
    if ($null -eq $updateConfig -or [string]::IsNullOrWhiteSpace($updateConfig.github_repo)) {
        # Pas de config ou repo non défini — silencieux
        return $null
    }

    # Vérifier le délai entre vérifications (éviter de spammer GitHub)
    # Fichier de throttling dans %APPDATA% (pas dans Program Files)
    $appDataPath   = Join-Path -Path $env:APPDATA -ChildPath "M365Monster"
    $lastCheckFile = Join-Path -Path $appDataPath -ChildPath ".last_update_check"

    if (Test-Path -Path $lastCheckFile) {
        $lastCheck = Get-Content -Path $lastCheckFile -Raw -ErrorAction SilentlyContinue
        try {
            $lastCheckDate = [datetime]::Parse($lastCheck)
            $hoursElapsed  = ((Get-Date) - $lastCheckDate).TotalHours
            $checkInterval = [int]$updateConfig.check_interval_hours
            if ($checkInterval -gt 0 -and $hoursElapsed -lt $checkInterval) {
                return $null
            }
        }
        catch {
            # Fichier corrompu — on continue
        }
    }

    # Version locale
    $localVersion = Get-LocalVersion -RootPath $RootPath
    if ($null -eq $localVersion) { return $null }

    # Construire l'URL du version.json distant (raw GitHub content)
    $repo = $updateConfig.github_repo  # format: "owner/repo"
    $branch = if ($updateConfig.branch) { $updateConfig.branch } else { "main" }
    $remoteVersionUrl = "https://raw.githubusercontent.com/$repo/$branch/version.json"

    # Construire les headers (token pour repo privé)
    $headers = @{ "User-Agent" = "M365Monster-Updater" }
    if (-not [string]::IsNullOrWhiteSpace($updateConfig.github_token)) {
        $headers["Authorization"] = "token $($updateConfig.github_token)"
    }

    try {
        $response = Invoke-RestMethod -Uri $remoteVersionUrl -Headers $headers -TimeoutSec 10 -ErrorAction Stop

        # Enregistrer la date de vérification
        if (-not (Test-Path -Path $appDataPath)) { New-Item -Path $appDataPath -ItemType Directory -Force | Out-Null }
        (Get-Date).ToString("o") | Out-File -FilePath $lastCheckFile -Encoding UTF8 -Force

        # Comparer les versions
        $localVer = [version]$localVersion.version
        $remoteVer = [version]$response.version

        if ($remoteVer -gt $localVer) {
            # Construire l'URL de téléchargement du .zip
            $downloadUrl = "https://github.com/$repo/releases/latest/download/M365Monster.zip"
            if (-not [string]::IsNullOrWhiteSpace($updateConfig.download_url)) {
                $downloadUrl = $updateConfig.download_url
            }

            return [PSCustomObject]@{
                UpdateAvailable = $true
                LocalVersion    = $localVersion.version
                RemoteVersion   = $response.version
                ReleaseNotes    = $response.release_notes
                ReleaseDate     = $response.release_date
                DownloadUrl     = $downloadUrl
                Headers         = $headers
            }
        }
        else {
            return [PSCustomObject]@{
                UpdateAvailable = $false
                LocalVersion    = $localVersion.version
                RemoteVersion   = $response.version
            }
        }
    }
    catch {
        # Échec silencieux — pas de réseau, repo inaccessible, etc.
        return $null
    }
}

function Test-AdminRights {
    <#
    .SYNOPSIS
        Vérifie si le processus courant a les droits administrateur.
    .OUTPUTS
        [bool]
    #>
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-AutoUpdate {
    <#
    .SYNOPSIS
        Vérifie et propose la mise à jour au lancement de l'application.
        Affiche une popup si une mise à jour est disponible.

        Stratégie d'élévation :
        - Téléchargement et extraction dans %TEMP% — pas de droits requis.
        - Génération d'un script de copie temporaire dans %TEMP%.
        - Si l'installation est dans Program Files et que le processus
          n'est pas admin : relance le script de copie avec Start-Process -Verb RunAs
          (prompt UAC unique, visible et attendu par l'utilisateur).
        - Si déjà admin (ou hors Program Files) : copie directe.

    .PARAMETER RootPath
        Chemin racine du projet.

    .OUTPUTS
        [bool] — $true si une mise à jour a été appliquée (nécessite un redémarrage),
                 $false sinon.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $updateInfo = Test-UpdateAvailable -RootPath $RootPath
    if ($null -eq $updateInfo -or -not $updateInfo.UpdateAvailable) {
        return $false
    }

    # Popup de proposition
    $message  = "Une nouvelle version de M365 Monster est disponible !`n`n"
    $message += "Version actuelle  : $($updateInfo.LocalVersion)`n"
    $message += "Nouvelle version  : $($updateInfo.RemoteVersion)`n"
    $message += "Date de sortie    : $($updateInfo.ReleaseDate)`n"
    if (-not [string]::IsNullOrWhiteSpace($updateInfo.ReleaseNotes)) {
        $message += "`nNotes : $($updateInfo.ReleaseNotes)`n"
    }
    $message += "`nVoulez-vous mettre à jour maintenant ?"

    $result = [System.Windows.Forms.MessageBox]::Show(
        $message,
        "Mise à jour disponible — M365 Monster",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )

    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
        return $false
    }

    # Chemins temporaires (accessibles sans droits admin)
    $tempZip      = Join-Path -Path $env:TEMP -ChildPath "M365Monster_update_$($updateInfo.RemoteVersion).zip"
    $tempExtract  = Join-Path -Path $env:TEMP -ChildPath "M365Monster_update_extract"
    $tempScript   = Join-Path -Path $env:TEMP -ChildPath "M365Monster_apply_update.ps1"
    $stampFile    = Join-Path -Path $env:TEMP -ChildPath "M365Monster_update_done.txt"

    try {
        # Nettoyage préalable
        @($tempZip, $tempExtract, $tempScript, $stampFile) | ForEach-Object {
            if (Test-Path $_) { Remove-Item $_ -Recurse -Force -ErrorAction SilentlyContinue }
        }

        # --- Étape 1 : Téléchargement (sans droits admin) ---
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "M365Monster-Updater")
        if ($updateInfo.Headers.ContainsKey("Authorization")) {
            $webClient.Headers.Add("Authorization", $updateInfo.Headers["Authorization"])
        }
        $webClient.DownloadFile($updateInfo.DownloadUrl, $tempZip)

        # --- Étape 2 : Extraction (sans droits admin) ---
        Expand-Archive -Path $tempZip -DestinationPath $tempExtract -Force
        Remove-Item $tempZip -Force -ErrorAction SilentlyContinue

        # Trouver le dossier racine dans le zip
        $extractedContent = Get-ChildItem -Path $tempExtract
        $sourceDir = $tempExtract
        if ($extractedContent.Count -eq 1 -and $extractedContent[0].PSIsContainer) {
            $sourceDir = $extractedContent[0].FullName
        }

        # Chemins de sauvegarde et restauration
        $backupDir        = Join-Path -Path $env:TEMP -ChildPath "M365Monster_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        $clientsDir       = Join-Path -Path $RootPath -ChildPath "Clients"
        $updateConfigPath = Join-Path -Path $RootPath -ChildPath "update_config.json"

        # --- Étape 3 : Générer le script de copie dans %TEMP% ---
        # Ce script sera exécuté avec ou sans élévation selon le contexte.
        $applyScript = @"
# Script de copie généré automatiquement par M365 Monster Update
# Exécuté avec droits admin si installation dans Program Files

`$sourceDir        = '$($sourceDir -replace "'","''")'
`$rootPath         = '$($RootPath -replace "'","''")'
`$backupDir        = '$($backupDir -replace "'","''")'
`$clientsDir       = '$($clientsDir -replace "'","''")'
`$updateConfigPath = '$($updateConfigPath -replace "'","''")'
`$stampFile        = '$($stampFile -replace "'","''")'

try {
    # Sauvegarde des fichiers client
    New-Item -Path `$backupDir -ItemType Directory -Force | Out-Null
    if (Test-Path `$clientsDir) {
        Copy-Item -Path `$clientsDir -Destination `$backupDir -Recurse -Force
    }
    if (Test-Path `$updateConfigPath) {
        Copy-Item -Path `$updateConfigPath -Destination `$backupDir -Force
    }

    # Copie des nouveaux fichiers (Clients/ et Logs/ exclus)
    Get-ChildItem -Path `$sourceDir -Exclude 'Clients','Logs' | ForEach-Object {
        `$dest = Join-Path -Path `$rootPath -ChildPath `$_.Name
        if (`$_.PSIsContainer) {
            Copy-Item -Path `$_.FullName -Destination `$dest -Recurse -Force
        }
        else {
            Copy-Item -Path `$_.FullName -Destination `$dest -Force
        }
    }

    # Restauration des fichiers client
    `$backupClients = Join-Path -Path `$backupDir -ChildPath 'Clients'
    if (Test-Path `$backupClients) {
        Copy-Item -Path "`$backupClients\*" -Destination `$clientsDir -Force -ErrorAction SilentlyContinue
    }
    `$backupConfig = Join-Path -Path `$backupDir -ChildPath 'update_config.json'
    if (Test-Path `$backupConfig) {
        Copy-Item -Path `$backupConfig -Destination `$rootPath -Force
    }

    # Débloquer les fichiers téléchargés
    Get-ChildItem -Path `$rootPath -Recurse -File | Unblock-File -ErrorAction SilentlyContinue

    # Écrire le fichier de confirmation
    'OK' | Out-File -FilePath `$stampFile -Encoding UTF8 -Force
}
catch {
    `$_.Exception.Message | Out-File -FilePath `$stampFile -Encoding UTF8 -Force
}
"@
        $applyScript | Out-File -FilePath $tempScript -Encoding UTF8 -Force

        # --- Étape 4 : Exécution du script de copie ---
        # Détecter si une élévation est nécessaire :
        # Program Files nécessite des droits admin ; un chemin custom (AppData, Bureau…) non.
        $needsElevation = ($RootPath -like "*Program Files*") -and (-not (Test-AdminRights))

        if ($needsElevation) {
            # Informer l'utilisateur qu'une demande UAC va apparaître
            [System.Windows.Forms.MessageBox]::Show(
                "Les fichiers sont dans Program Files.`n`nWindows va demander une confirmation administrateur (UAC) pour copier les fichiers mis à jour.`n`nCliquez Oui dans la fenêtre qui va apparaître.",
                "Autorisation requise",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null

            # Relance avec élévation — déclenche le prompt UAC standard Windows
            $pwsh = if (Get-Command pwsh -ErrorAction SilentlyContinue) { "pwsh" } else { "powershell" }
            $proc = Start-Process $pwsh `
                -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$tempScript`"" `
                -Verb RunAs `
                -Wait `
                -PassThru

            if ($proc.ExitCode -ne 0) {
                throw "Le script d'élévation s'est terminé avec le code $($proc.ExitCode). L'utilisateur a peut-être refusé l'UAC."
            }
        }
        else {
            # Déjà admin ou chemin non protégé — exécution directe
            $pwsh = if (Get-Command pwsh -ErrorAction SilentlyContinue) { "pwsh" } else { "powershell" }
            Start-Process $pwsh `
                -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$tempScript`"" `
                -Wait `
                -WindowStyle Hidden
        }

        # --- Étape 5 : Vérifier que la copie a réussi ---
        if (-not (Test-Path $stampFile)) {
            throw "Fichier de confirmation absent. La copie n'a peut-être pas abouti."
        }
        $stampContent = Get-Content $stampFile -Raw -ErrorAction SilentlyContinue
        if ($stampContent.Trim() -ne "OK") {
            throw "Erreur pendant la copie : $stampContent"
        }

        # Nettoyage
        @($tempExtract, $tempScript, $stampFile) | ForEach-Object {
            Remove-Item $_ -Recurse -Force -ErrorAction SilentlyContinue
        }

        [System.Windows.Forms.MessageBox]::Show(
            "Mise à jour vers la version $($updateInfo.RemoteVersion) réussie !`n`nL'application va redémarrer.",
            "Mise à jour terminée",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Erreur lors de la mise à jour :`n$($_.Exception.Message)`n`nL'application va démarrer avec la version actuelle.",
            "Erreur de mise à jour",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null

        # Nettoyage en cas d'erreur
        @($tempZip, $tempExtract, $tempScript, $stampFile) | ForEach-Object {
            Remove-Item $_ -Recurse -Force -ErrorAction SilentlyContinue
        }

        return $false
    }
}