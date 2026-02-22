<#
.FICHIER
    Main.ps1

.ROLE
    Point d'entrée unique de l'application M365 Monster.
    Orchestre : auto-update, chargement des modules, sélection du client,
    connexion Graph, puis lancement de la fenêtre principale.

.DEPENDANCES
    - Core/Update.ps1
    - Core/Config.ps1
    - Core/Connect.ps1
    - Core/GraphAPI.ps1
    - Core/Functions.ps1
    - Modules/GUI_Main.ps1

.AUTEUR
    [Equipe IT - M365 Monster]

.USAGE
    Exécuter directement : .\Main.ps1
#>

# === Détermination du répertoire racine du script ===
$script:RootPath = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($script:RootPath)) {
    $script:RootPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

# === Débloquer tous les fichiers du projet (supprime le Zone Identifier NTFS) ===
Get-ChildItem -Path $script:RootPath -Recurse -File | Unblock-File -ErrorAction SilentlyContinue

# === Chargement des assemblies WinForms ===
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# === Auto-update (avant tout chargement de module) ===
. "$script:RootPath\Core\Update.ps1"

$updateApplied = Invoke-AutoUpdate -RootPath $script:RootPath
if ($updateApplied) {
    # Redémarrer le script avec la nouvelle version
    # On utilise le meme executable PowerShell que celui en cours (pwsh.exe ou powershell.exe)
    $psExe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $psExe -ArgumentList "-ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Definition)`""
    exit
}

# === Système de langue (i18n) ===
. "$script:RootPath\Core\Lang.ps1"

# Le dossier des donnees utilisateur (settings, logs) est dans AppData
# pour eviter les problemes de permissions dans Program Files
$script:UserDataPath = Join-Path -Path $env:APPDATA -ChildPath "M365Monster"
if (-not (Test-Path $script:UserDataPath)) {
    New-Item -Path $script:UserDataPath -ItemType Directory -Force | Out-Null
}

$langOk = Initialize-Language -RootPath $script:RootPath -UserDataPath $script:UserDataPath
if (-not $langOk) {
    $langFolder = Join-Path -Path $script:RootPath -ChildPath "Lang"
    $langExists = Test-Path $langFolder
    $langFiles = if ($langExists) { (Get-ChildItem -Path $langFolder -Filter "*.json" -ErrorAction SilentlyContinue).Count } else { 0 }
    $debugMsg  = "Impossible de charger la langue.`nLanguage loading failed.`n`n"
    $debugMsg += "RootPath: $($script:RootPath)`n"
    $debugMsg += "Lang folder exists: $langExists`n"
    $debugMsg += "Lang files found: $langFiles"
    [System.Windows.Forms.MessageBox]::Show(
        $debugMsg,
        "M365 Monster",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit
}

# === Dot-sourcing des modules Core ===
. "$script:RootPath\Core\Config.ps1"
. "$script:RootPath\Core\Functions.ps1"
. "$script:RootPath\Core\GraphAPI.ps1"
. "$script:RootPath\Core\Connect.ps1"

# === Initialisation du fichier de log ===
# Les logs sont stockés dans AppData pour eviter les problemes de permissions
# dans Program Files (qui est en lecture seule pour les utilisateurs non-admin)
$logFolder = Join-Path -Path $env:APPDATA -ChildPath "M365Monster\Logs"
Initialize-LogFile -LogFolder $logFolder
Write-Log -Level "INFO" -Action "DEMARRAGE" -Message "M365 Monster démarré."

# === Chargement anticipé de GUI_Settings (nécessaire pour le sélecteur de client) ===
. "$script:RootPath\Modules\GUI_Settings.ps1"

# === Fenêtre de sélection du client ===
function Show-ClientSelector {
    <#
    .SYNOPSIS
        Affiche une fenêtre de sélection du client avec liste déroulante.
        Inclut un bouton pour créer un nouveau client via GUI_Settings.
    .OUTPUTS
        [string] — Chemin complet du fichier JSON du client sélectionné, ou $null si annulé.
    #>

    $clientsFolder = Join-Path -Path $script:RootPath -ChildPath "Clients"

    # Fonction interne pour charger la liste des clients dans la combobox
    function Update-ClientComboBox {
        param($ComboBox)
        $ComboBox.Items.Clear()
        $script:SelectorClients = Get-ClientList -ClientsFolder $clientsFolder
        foreach ($client in $script:SelectorClients) {
            $ComboBox.Items.Add("$($client.Name) ($($client.FileName))") | Out-Null
        }
        if ($ComboBox.Items.Count -gt 0) { $ComboBox.SelectedIndex = 0 }
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = Get-Text "client_selector.title"
    $form.Size = New-Object System.Drawing.Size(450, 300)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::WhiteSmoke

    $lblTitre = New-Object System.Windows.Forms.Label
    $lblTitre.Text = Get-Text "client_selector.heading"
    $lblTitre.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitre.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $lblTitre.Location = New-Object System.Drawing.Point(20, 20)
    $lblTitre.Size = New-Object System.Drawing.Size(400, 35)
    $form.Controls.Add($lblTitre)

    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = Get-Text "client_selector.description"
    $lblDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblDesc.ForeColor = [System.Drawing.Color]::Gray
    $lblDesc.Location = New-Object System.Drawing.Point(20, 55)
    $lblDesc.Size = New-Object System.Drawing.Size(400, 20)
    $form.Controls.Add($lblDesc)

    $cboClient = New-Object System.Windows.Forms.ComboBox
    $cboClient.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $cboClient.Location = New-Object System.Drawing.Point(20, 90)
    $cboClient.Size = New-Object System.Drawing.Size(390, 30)
    $cboClient.DropDownStyle = "DropDownList"
    $form.Controls.Add($cboClient)

    # Charger la liste initiale
    Update-ClientComboBox -ComboBox $cboClient

    # --- Bouton Nouveau client (ouvre le formulaire de paramétrage) ---
    $btnNouveau = New-Object System.Windows.Forms.Button
    $btnNouveau.Text = Get-Text "client_selector.btn_new_client"
    $btnNouveau.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnNouveau.Location = New-Object System.Drawing.Point(20, 135)
    $btnNouveau.Size = New-Object System.Drawing.Size(390, 32)
    $btnNouveau.FlatStyle = "Flat"
    $btnNouveau.BackColor = [System.Drawing.Color]::FromArgb(111, 66, 193)
    $btnNouveau.ForeColor = [System.Drawing.Color]::White
    $btnNouveau.Cursor = [System.Windows.Forms.Cursors]::Hand
    $form.Controls.Add($btnNouveau)

    $btnNouveau.Add_Click({
        # Ouvrir le formulaire de paramétrage client
        Show-SettingsForm
        # Rafraîchir la liste après fermeture du formulaire
        Update-ClientComboBox -ComboBox $cboClient
    })

    # --- Boutons Connecter / Quitter ---
    $btnValider = New-Object System.Windows.Forms.Button
    $btnValider.Text = Get-Text "client_selector.btn_connect"
    $btnValider.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnValider.Location = New-Object System.Drawing.Point(20, 190)
    $btnValider.Size = New-Object System.Drawing.Size(185, 40)
    $btnValider.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $btnValider.ForeColor = [System.Drawing.Color]::White
    $btnValider.FlatStyle = "Flat"
    $btnValider.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnValider)

    $btnQuitter = New-Object System.Windows.Forms.Button
    $btnQuitter.Text = Get-Text "client_selector.btn_quit"
    $btnQuitter.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnQuitter.Location = New-Object System.Drawing.Point(225, 190)
    $btnQuitter.Size = New-Object System.Drawing.Size(185, 40)
    $btnQuitter.FlatStyle = "Flat"
    $btnQuitter.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnQuitter)

    $form.AcceptButton = $btnValider
    $form.CancelButton = $btnQuitter

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $cboClient.SelectedIndex -ge 0) {
        $selectedPath = $script:SelectorClients[$cboClient.SelectedIndex].FullPath
        $form.Dispose()
        return $selectedPath
    }

    $form.Dispose()
    return $null
}

# === Flux principal ===

# Étape 1 : Sélection du client
$clientConfigPath = Show-ClientSelector

if ($null -eq $clientConfigPath) {
    Write-Log -Level "INFO" -Action "DEMARRAGE" -Message "Aucun client sélectionné. Fermeture."
    exit
}

# Étape 2 : Chargement de la configuration
try {
    $script:Config = Load-ClientConfig -ConfigPath $clientConfigPath
    Write-Log -Level "SUCCESS" -Action "CONFIG" -Message "Configuration chargée : $($Config.client_name)"
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Erreur lors du chargement de la configuration :`n$($_.Exception.Message)",
        "Erreur de configuration",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    Write-Log -Level "ERROR" -Action "CONFIG" -Message "Erreur chargement config : $($_.Exception.Message)"
    exit
}

# Étape 3 : Connexion à Microsoft Graph
Write-Log -Level "INFO" -Action "CONNEXION" -Message "Connexion à Microsoft Graph pour '$($Config.client_name)'..."

$connectResult = Connect-GraphAPI
if (-not $connectResult.Success) {
    [System.Windows.Forms.MessageBox]::Show(
        "Impossible de se connecter à Microsoft Graph :`n$($connectResult.Error)`n`nVérifiez la configuration du client et réessayez.",
        "Erreur de connexion",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    Write-Log -Level "ERROR" -Action "CONNEXION" -Message "Connexion échouée. Fermeture."
    exit
}

# Étape 4 : Dot-sourcing des modules GUI
. "$script:RootPath\Modules\GUI_Main.ps1"
. "$script:RootPath\Modules\GUI_Onboarding.ps1"
. "$script:RootPath\Modules\GUI_Offboarding.ps1"
. "$script:RootPath\Modules\GUI_Modification.ps1"
# GUI_Settings.ps1 déjà chargé avant le sélecteur de client

# Étape 5 : Affichage de la fenêtre principale
Write-Log -Level "INFO" -Action "GUI" -Message "Ouverture de la fenêtre principale."
Show-MainWindow

# Étape 6 : Nettoyage à la fermeture
Write-Log -Level "INFO" -Action "FERMETURE" -Message "M365 Monster fermé."
Disconnect-GraphAPI
