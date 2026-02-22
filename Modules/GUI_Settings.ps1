<#
.FICHIER
    Modules/GUI_Settings.ps1

.ROLE
    Interface de parametrage client.
    Permet de lister, creer, editer et tester les fichiers de configuration JSON.
    Adapte a la structure JSON actuelle :
      - license_groups (noms de groupes pour assignation de licence)
      - membership_groups (groupes d'appartenance)
      - Plus de naming_convention, license_map, departments, contract_types
        (ces valeurs sont desormais dynamiques depuis Azure AD)

.DEPENDANCES
    - Core/Config.ps1 (Get-ClientList)
    - Core/Functions.ps1 (Write-Log, Show-ConfirmDialog, Show-ResultDialog)
    - Variable globale $Config, $RootPath

.AUTEUR
    [Equipe IT - GestionRH-AzureAD]
#>

function Show-SettingsForm {
    <#
    .SYNOPSIS
        Affiche l'interface de gestion des configurations client.
    .OUTPUTS
        [void] - Formulaire modal.
    #>

    $clientsFolder = Join-Path -Path $script:RootPath -ChildPath "Clients"
    $templatePath = Join-Path -Path $clientsFolder -ChildPath "_Template.json"

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Parametrage client - M365 Monster"
    $form.Size = New-Object System.Drawing.Size(750, 700)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::WhiteSmoke
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # =================================================================
    # SECTION : Liste des clients existants
    # =================================================================
    $lblSection = New-Object System.Windows.Forms.Label
    $lblSection.Text = "Fichiers de configuration client"
    $lblSection.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblSection.Location = New-Object System.Drawing.Point(15, 15)
    $lblSection.Size = New-Object System.Drawing.Size(400, 25)
    $form.Controls.Add($lblSection)

    $lstClients = New-Object System.Windows.Forms.ListBox
    $lstClients.Location = New-Object System.Drawing.Point(15, 45)
    $lstClients.Size = New-Object System.Drawing.Size(280, 110)
    $form.Controls.Add($lstClients)

    function Update-ClientList {
        $lstClients.Items.Clear()
        $clients = Get-ClientList -ClientsFolder $clientsFolder
        foreach ($c in $clients) {
            $lstClients.Items.Add("$($c.Name) ($($c.FileName))") | Out-Null
        }
        $script:ClientListData = $clients
    }
    $script:ClientListData = @()
    Update-ClientList

    # Boutons de gestion
    $btnNouveau = New-Object System.Windows.Forms.Button
    $btnNouveau.Text = "Nouveau client"
    $btnNouveau.Location = New-Object System.Drawing.Point(310, 45)
    $btnNouveau.Size = New-Object System.Drawing.Size(130, 30)
    $btnNouveau.FlatStyle = "Flat"
    $btnNouveau.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $btnNouveau.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($btnNouveau)

    $btnEditer = New-Object System.Windows.Forms.Button
    $btnEditer.Text = "Editer"
    $btnEditer.Location = New-Object System.Drawing.Point(310, 80)
    $btnEditer.Size = New-Object System.Drawing.Size(130, 30)
    $btnEditer.FlatStyle = "Flat"
    $form.Controls.Add($btnEditer)

    $btnSupprimer = New-Object System.Windows.Forms.Button
    $btnSupprimer.Text = "Supprimer"
    $btnSupprimer.Location = New-Object System.Drawing.Point(310, 115)
    $btnSupprimer.Size = New-Object System.Drawing.Size(130, 30)
    $btnSupprimer.FlatStyle = "Flat"
    $btnSupprimer.ForeColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $form.Controls.Add($btnSupprimer)

    # Separateur
    $sep = New-Object System.Windows.Forms.Label
    $sep.BorderStyle = "Fixed3D"
    $sep.Location = New-Object System.Drawing.Point(15, 165)
    $sep.Size = New-Object System.Drawing.Size(700, 2)
    $form.Controls.Add($sep)

    # =================================================================
    # SECTION : Formulaire d'edition (scrollable)
    # =================================================================
    $lblEdit = New-Object System.Windows.Forms.Label
    $lblEdit.Text = "Edition de la configuration"
    $lblEdit.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblEdit.Location = New-Object System.Drawing.Point(15, 175)
    $lblEdit.Size = New-Object System.Drawing.Size(400, 25)
    $form.Controls.Add($lblEdit)

    $panelEdit = New-Object System.Windows.Forms.Panel
    $panelEdit.Location = New-Object System.Drawing.Point(15, 205)
    $panelEdit.Size = New-Object System.Drawing.Size(705, 400)
    $panelEdit.AutoScroll = $true
    $form.Controls.Add($panelEdit)

    $yEdit = 0
    $lblW = 165
    $fldX = 175
    $fldW = 500

    # --- Fonction utilitaire : ajouter un champ texte avec label ---
    function Add-EditField {
        param($Parent, $LabelText, [ref]$YPos, $DefaultValue = "")

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $LabelText
        $lbl.Location = New-Object System.Drawing.Point(0, ($YPos.Value + 3))
        $lbl.Size = New-Object System.Drawing.Size($lblW, 20)
        $Parent.Controls.Add($lbl)

        $ctrl = New-Object System.Windows.Forms.TextBox
        $ctrl.Text = $DefaultValue
        $ctrl.Location = New-Object System.Drawing.Point($fldX, $YPos.Value)
        $ctrl.Size = New-Object System.Drawing.Size($fldW, 25)
        $Parent.Controls.Add($ctrl)

        $YPos.Value += 30
        return $ctrl
    }

    # --- Fonction utilitaire : ajouter un label de section ---
    function Add-SectionLabel {
        param($Parent, $Text, [ref]$YPos)
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $Text
        $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $lbl.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
        $lbl.Location = New-Object System.Drawing.Point(0, $YPos.Value)
        $lbl.Size = New-Object System.Drawing.Size(680, 20)
        $Parent.Controls.Add($lbl)
        $YPos.Value += 24
    }

    # --- Fonction utilitaire : ajouter un label d'aide ---
    function Add-HelpLabel {
        param($Parent, $Text, [ref]$YPos)
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $Text
        $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
        $lbl.ForeColor = [System.Drawing.Color]::Gray
        $lbl.Location = New-Object System.Drawing.Point($fldX, $YPos.Value)
        $lbl.Size = New-Object System.Drawing.Size($fldW, 16)
        $Parent.Controls.Add($lbl)
        $YPos.Value += 18
    }

    # =================================================================
    # Champs du formulaire
    # =================================================================

    # --- CONNEXION ---
    Add-SectionLabel -Parent $panelEdit -Text "CONNEXION" -YPos ([ref]$yEdit)

    $txtClientName = Add-EditField -Parent $panelEdit -LabelText "Nom du client *" -YPos ([ref]$yEdit)
    $txtTenantId   = Add-EditField -Parent $panelEdit -LabelText "Tenant ID *" -YPos ([ref]$yEdit)
    $txtClientId   = Add-EditField -Parent $panelEdit -LabelText "Client ID (App) *" -YPos ([ref]$yEdit)

    # Auth method (ComboBox)
    $lblAuth = New-Object System.Windows.Forms.Label
    $lblAuth.Text = "Methode auth *"
    $lblAuth.Location = New-Object System.Drawing.Point(0, ($yEdit + 3))
    $lblAuth.Size = New-Object System.Drawing.Size($lblW, 20)
    $panelEdit.Controls.Add($lblAuth)

    $cboAuth = New-Object System.Windows.Forms.ComboBox
    $cboAuth.Location = New-Object System.Drawing.Point($fldX, $yEdit)
    $cboAuth.Size = New-Object System.Drawing.Size(220, 25)
    $cboAuth.DropDownStyle = "DropDownList"
    $cboAuth.Items.Add("interactive_browser") | Out-Null
    $cboAuth.Items.Add("device_code") | Out-Null
    $cboAuth.Items.Add("client_secret") | Out-Null
    $cboAuth.SelectedIndex = 0
    $panelEdit.Controls.Add($cboAuth)
    $yEdit += 30

    $txtSmtpDomain = Add-EditField -Parent $panelEdit -LabelText "Domaine SMTP *" -YPos ([ref]$yEdit) -DefaultValue "@domaine.com"
    Add-HelpLabel -Parent $panelEdit -Text "Doit commencer par @ (ex: @entreprise.com)" -YPos ([ref]$yEdit)

    $yEdit += 8

    # --- GROUPES DE LICENCE ---
    Add-SectionLabel -Parent $panelEdit -Text "GROUPES DE LICENCE" -YPos ([ref]$yEdit)

    $txtLicenseGroups = Add-EditField -Parent $panelEdit -LabelText "Groupes de licence" -YPos ([ref]$yEdit)
    Add-HelpLabel -Parent $panelEdit -Text "Noms des groupes Entra ID pour les licences, separes par des virgules" -YPos ([ref]$yEdit)

    $yEdit += 8

    # --- GROUPES D'APPARTENANCE ---
    Add-SectionLabel -Parent $panelEdit -Text "GROUPES D'APPARTENANCE" -YPos ([ref]$yEdit)

    $txtMembershipGroups = Add-EditField -Parent $panelEdit -LabelText "Groupes membre" -YPos ([ref]$yEdit)
    Add-HelpLabel -Parent $panelEdit -Text "Noms des groupes pour l'ajout des nouveaux employes, separes par des virgules" -YPos ([ref]$yEdit)

    $yEdit += 8

    # --- OFFBOARDING ---
    Add-SectionLabel -Parent $panelEdit -Text "OFFBOARDING" -YPos ([ref]$yEdit)

    $txtDisabledGroup   = Add-EditField -Parent $panelEdit -LabelText "Groupe desactives" -YPos ([ref]$yEdit) -DefaultValue "Comptes-Desactives"
    Add-HelpLabel -Parent $panelEdit -Text "Groupe ou placer les comptes desactives lors de l'offboarding" -YPos ([ref]$yEdit)

    $yEdit += 8

    # --- NOTIFICATIONS ---
    Add-SectionLabel -Parent $panelEdit -Text "NOTIFICATIONS" -YPos ([ref]$yEdit)

    $txtNotifRecipients = Add-EditField -Parent $panelEdit -LabelText "Destinataires" -YPos ([ref]$yEdit)
    Add-HelpLabel -Parent $panelEdit -Text "Adresses email separees par des virgules (vide = notifications desactivees)" -YPos ([ref]$yEdit)

    $yEdit += 8

    # --- MOT DE PASSE ---
    Add-SectionLabel -Parent $panelEdit -Text "POLITIQUE DE MOT DE PASSE" -YPos ([ref]$yEdit)

    $txtPasswordLength = Add-EditField -Parent $panelEdit -LabelText "Longueur MDP" -YPos ([ref]$yEdit) -DefaultValue "14"

    # Force change at login (CheckBox)
    $chkForceChange = New-Object System.Windows.Forms.CheckBox
    $chkForceChange.Text = "Forcer le changement au premier login"
    $chkForceChange.Location = New-Object System.Drawing.Point($fldX, $yEdit)
    $chkForceChange.Size = New-Object System.Drawing.Size(350, 22)
    $chkForceChange.Checked = $true
    $panelEdit.Controls.Add($chkForceChange)
    $yEdit += 26

    # Include special chars (CheckBox)
    $chkSpecialChars = New-Object System.Windows.Forms.CheckBox
    $chkSpecialChars.Text = "Inclure des caracteres speciaux"
    $chkSpecialChars.Location = New-Object System.Drawing.Point($fldX, $yEdit)
    $chkSpecialChars.Size = New-Object System.Drawing.Size(350, 22)
    $chkSpecialChars.Checked = $true
    $panelEdit.Controls.Add($chkSpecialChars)
    $yEdit += 30

    # =================================================================
    # Fonctions de remplissage / construction du JSON
    # =================================================================

    $script:EditingFilePath = $null

    function Set-FormFromConfig {
        param($ConfigObj)
        $txtClientName.Text      = $ConfigObj.client_name
        $txtTenantId.Text        = $ConfigObj.tenant_id
        $txtClientId.Text        = $ConfigObj.client_id
        $cboAuth.SelectedItem    = $ConfigObj.auth_method
        $txtSmtpDomain.Text      = $ConfigObj.smtp_domain
        $txtLicenseGroups.Text   = ($ConfigObj.license_groups -join ", ")
        $txtMembershipGroups.Text = ($ConfigObj.membership_groups -join ", ")
        $txtDisabledGroup.Text   = $ConfigObj.offboarding.disabled_ou_group
        $txtNotifRecipients.Text = ($ConfigObj.notifications.recipients -join ", ")
        $txtPasswordLength.Text  = $ConfigObj.password_policy.length.ToString()
        $chkForceChange.Checked  = $ConfigObj.password_policy.force_change_at_login
        $chkSpecialChars.Checked = $ConfigObj.password_policy.include_special_chars
    }

    function Build-ConfigFromForm {
        # Transformer les champs texte separes par virgules en tableaux
        $licGroups = @($txtLicenseGroups.Text.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
        $memGroups = @($txtMembershipGroups.Text.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
        $notifList = @($txtNotifRecipients.Text.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })

        $obj = [ordered]@{
            client_name       = $txtClientName.Text.Trim()
            tenant_id         = $txtTenantId.Text.Trim()
            client_id         = $txtClientId.Text.Trim()
            auth_method       = $cboAuth.SelectedItem.ToString()
            smtp_domain       = $txtSmtpDomain.Text.Trim()
            license_groups    = $licGroups
            membership_groups = $memGroups
            offboarding       = [ordered]@{
                disabled_ou_group  = $txtDisabledGroup.Text.Trim()
                mailbox_forward_to = ""
                revoke_licenses    = $true
                remove_all_groups  = $true
                retention_days     = 30
            }
            notifications     = [ordered]@{
                enabled    = ($notifList.Count -gt 0)
                recipients = $notifList
            }
            password_policy   = [ordered]@{
                length                = [int]$txtPasswordLength.Text.Trim()
                force_change_at_login = $chkForceChange.Checked
                include_special_chars = $chkSpecialChars.Checked
            }
        }
        return $obj
    }

    # =================================================================
    # Actions des boutons
    # =================================================================

    # Nouveau client
    $btnNouveau.Add_Click({
        if (Test-Path -Path $templatePath) {
            $template = Get-Content -Path $templatePath -Raw -Encoding UTF8 | ConvertFrom-Json
            Set-FormFromConfig -ConfigObj $template
        }
        else {
            $txtClientName.Text = ""
            $txtTenantId.Text = ""
            $txtClientId.Text = ""
            $txtSmtpDomain.Text = "@domaine.com"
            $txtLicenseGroups.Text = ""
            $txtMembershipGroups.Text = ""
        }
        $script:EditingFilePath = $null
        $lblEdit.Text = "Nouveau client"
    })

    # Editer un client existant
    $btnEditer.Add_Click({
        if ($lstClients.SelectedIndex -ge 0 -and $lstClients.SelectedIndex -lt $script:ClientListData.Count) {
            $selectedClient = $script:ClientListData[$lstClients.SelectedIndex]
            try {
                $configObj = Get-Content -Path $selectedClient.FullPath -Raw -Encoding UTF8 | ConvertFrom-Json
                Set-FormFromConfig -ConfigObj $configObj
                $script:EditingFilePath = $selectedClient.FullPath
                $lblEdit.Text = "Edition : $($selectedClient.Name)"
            }
            catch {
                Show-ResultDialog -Titre "Erreur" -Message "Impossible de lire le fichier : $($_.Exception.Message)" -IsSuccess $false
            }
        }
    })

    # Supprimer un client
    $btnSupprimer.Add_Click({
        if ($lstClients.SelectedIndex -ge 0 -and $lstClients.SelectedIndex -lt $script:ClientListData.Count) {
            $selectedClient = $script:ClientListData[$lstClients.SelectedIndex]
            $confirm = Show-ConfirmDialog -Titre "Suppression" -Message "Supprimer la configuration '$($selectedClient.Name)' ?`n`nFichier : $($selectedClient.FileName)" -Icon ([System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($confirm) {
                try {
                    Remove-Item -Path $selectedClient.FullPath -Force
                    Write-Log -Level "INFO" -Action "SETTINGS" -Message "Configuration supprimee : $($selectedClient.FileName)"
                    Update-ClientList
                    Show-ResultDialog -Titre "Succes" -Message "Configuration supprimee." -IsSuccess $true
                }
                catch {
                    Show-ResultDialog -Titre "Erreur" -Message $_.Exception.Message -IsSuccess $false
                }
            }
        }
    })

    # =================================================================
    # Boutons du bas
    # =================================================================

    # Tester la connexion
    $btnTester = New-Object System.Windows.Forms.Button
    $btnTester.Text = "Tester la connexion"
    $btnTester.Location = New-Object System.Drawing.Point(15, 615)
    $btnTester.Size = New-Object System.Drawing.Size(150, 35)
    $btnTester.FlatStyle = "Flat"
    $btnTester.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $btnTester.ForeColor = [System.Drawing.Color]::White
    $btnTester.Add_Click({
        $tenantId = $txtTenantId.Text.Trim()
        $clientId = $txtClientId.Text.Trim()

        $guidRegex = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
        if ($tenantId -notmatch $guidRegex -or $clientId -notmatch $guidRegex) {
            Show-ResultDialog -Titre "Validation" -Message "Le Tenant ID et/ou Client ID ne sont pas des GUID valides." -IsSuccess $false
            return
        }

        Show-ResultDialog -Titre "Test" -Message "Le format des identifiants est valide.`n`nTenant ID : $tenantId`nClient ID : $clientId`n`nPour tester la connexion complete, sauvegardez et relancez l'application avec ce client." -IsSuccess $true
    })
    $form.Controls.Add($btnTester)

    # Sauvegarder
    $btnSauvegarder = New-Object System.Windows.Forms.Button
    $btnSauvegarder.Text = "Sauvegarder"
    $btnSauvegarder.Location = New-Object System.Drawing.Point(400, 615)
    $btnSauvegarder.Size = New-Object System.Drawing.Size(140, 35)
    $btnSauvegarder.FlatStyle = "Flat"
    $btnSauvegarder.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $btnSauvegarder.ForeColor = [System.Drawing.Color]::White
    $btnSauvegarder.Add_Click({
        # Validation des champs obligatoires
        if ([string]::IsNullOrWhiteSpace($txtClientName.Text) -or
            [string]::IsNullOrWhiteSpace($txtTenantId.Text) -or
            [string]::IsNullOrWhiteSpace($txtClientId.Text)) {
            Show-ResultDialog -Titre "Validation" -Message "Les champs Nom, Tenant ID et Client ID sont obligatoires." -IsSuccess $false
            return
        }

        if (-not $txtSmtpDomain.Text.Trim().StartsWith("@")) {
            Show-ResultDialog -Titre "Validation" -Message "Le domaine SMTP doit commencer par @ (ex: @entreprise.com)." -IsSuccess $false
            return
        }

        $configData = Build-ConfigFromForm
        $jsonContent = $configData | ConvertTo-Json -Depth 5

        # Determiner le chemin de sauvegarde
        if ($script:EditingFilePath) {
            $savePath = $script:EditingFilePath
        }
        else {
            $fileName = ($txtClientName.Text.Trim() -replace '[^a-zA-Z0-9]', '') + ".json"
            $savePath = Join-Path -Path $clientsFolder -ChildPath $fileName
        }

        try {
            $jsonContent | Out-File -FilePath $savePath -Encoding UTF8 -Force
            Write-Log -Level "SUCCESS" -Action "SETTINGS" -Message "Configuration sauvegardee : $savePath"
            Show-ResultDialog -Titre "Succes" -Message "Configuration sauvegardee dans :`n$savePath" -IsSuccess $true
            Update-ClientList
        }
        catch {
            Show-ResultDialog -Titre "Erreur" -Message "Echec de la sauvegarde : $($_.Exception.Message)" -IsSuccess $false
        }
    })
    $form.Controls.Add($btnSauvegarder)

    # Fermer
    $btnFermer = New-Object System.Windows.Forms.Button
    $btnFermer.Text = "Fermer"
    $btnFermer.Location = New-Object System.Drawing.Point(560, 615)
    $btnFermer.Size = New-Object System.Drawing.Size(140, 35)
    $btnFermer.FlatStyle = "Flat"
    $btnFermer.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnFermer)
    $form.CancelButton = $btnFermer

    $form.ShowDialog() | Out-Null
    $form.Dispose()
}
