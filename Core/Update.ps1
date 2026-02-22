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

function Invoke-AutoUpdate {
    <#
    .SYNOPSIS
        Vérifie et propose la mise à jour au lancement de l'application.
        Affiche une popup si une mise à jour est disponible.
        Télécharge, extrait et remplace les fichiers si l'utilisateur accepte.

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

    # Téléchargement
    $tempZip = Join-Path -Path $env:TEMP -ChildPath "M365Monster_update_$($updateInfo.RemoteVersion).zip"
    $tempExtract = Join-Path -Path $env:TEMP -ChildPath "M365Monster_update_extract"

    try {
        # Nettoyage préalable
        if (Test-Path $tempZip) { Remove-Item $tempZip -Force }
        if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force }

        # Téléchargement du .zip
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "M365Monster-Updater")
        if ($updateInfo.Headers.ContainsKey("Authorization")) {
            $webClient.Headers.Add("Authorization", $updateInfo.Headers["Authorization"])
        }
        $webClient.DownloadFile($updateInfo.DownloadUrl, $tempZip)

        # Extraction
        Expand-Archive -Path $tempZip -DestinationPath $tempExtract -Force

        # Trouver le dossier racine dans le zip (peut être M365Monster/ ou direct)
        $extractedContent = Get-ChildItem -Path $tempExtract
        $sourceDir = $tempExtract
        if ($extractedContent.Count -eq 1 -and $extractedContent[0].PSIsContainer) {
            $sourceDir = $extractedContent[0].FullName
        }

        # Sauvegarder les fichiers client (Clients/*.json et Logs/)
        $backupDir = Join-Path -Path $env:TEMP -ChildPath "M365Monster_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        New-Item -Path $backupDir -ItemType Directory -Force | Out-Null

        $clientsDir = Join-Path -Path $RootPath -ChildPath "Clients"
        if (Test-Path $clientsDir) {
            Copy-Item -Path $clientsDir -Destination $backupDir -Recurse -Force
        }

        $updateConfigPath = Join-Path -Path $RootPath -ChildPath "update_config.json"
        if (Test-Path $updateConfigPath) {
            Copy-Item -Path $updateConfigPath -Destination $backupDir -Force
        }

        # Copier les nouveaux fichiers (sauf Clients/ et Logs/)
        $itemsToCopy = Get-ChildItem -Path $sourceDir -Exclude "Clients", "Logs"
        foreach ($item in $itemsToCopy) {
            $destPath = Join-Path -Path $RootPath -ChildPath $item.Name
            if ($item.PSIsContainer) {
                Copy-Item -Path $item.FullName -Destination $destPath -Recurse -Force
            }
            else {
                Copy-Item -Path $item.FullName -Destination $destPath -Force
            }
        }

        # Restaurer les fichiers client depuis la sauvegarde
        $backupClients = Join-Path -Path $backupDir -ChildPath "Clients"
        if (Test-Path $backupClients) {
            Copy-Item -Path "$backupClients\*" -Destination $clientsDir -Force -ErrorAction SilentlyContinue
        }

        $backupUpdateConfig = Join-Path -Path $backupDir -ChildPath "update_config.json"
        if (Test-Path $backupUpdateConfig) {
            Copy-Item -Path $backupUpdateConfig -Destination $RootPath -Force
        }

        # Débloquer les nouveaux fichiers
        Get-ChildItem -Path $RootPath -Recurse -File | Unblock-File -ErrorAction SilentlyContinue

        # Nettoyage temp
        Remove-Item $tempZip -Force -ErrorAction SilentlyContinue
        Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue

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
        Remove-Item $tempZip -Force -ErrorAction SilentlyContinue
        Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue

        return $false
    }
}