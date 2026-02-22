<#
.FICHIER
    Modules/GUI_Onboarding.ps1

.ROLE
    Formulaire d'arrivée (onboarding) d'un nouvel employé.
    Champs dynamiques alimentés depuis Azure AD (department, jobTitle, etc.).
    Génération automatique du username avec 3 formats au choix et vérification d'unicité.
    Licences et groupes d'appartenance référencés depuis le JSON client.

.DEPENDANCES
    - Core/Functions.ps1 (New-UsernameVariants, New-SecurePassword, Write-Log,
      Show-ConfirmDialog, Show-ResultDialog, Show-PasswordDialog, Get-MailNickname,
      Remove-Diacritics)
    - Core/GraphAPI.ps1 (New-AzUser, Add-AzUserToGroup, Set-AzUserManager,
      Search-AzUsers, Get-AzDistinctValues, Test-AzUserExists)
    - Variable globale $Config

.AUTEUR
    [Equipe IT - GestionRH-AzureAD]
#>

function Show-OnboardingForm {
    <#
    .SYNOPSIS
        Affiche le formulaire d'onboarding pour créer un nouvel employé.
        Tous les champs dynamiques sont chargés depuis Azure AD au lancement.
    .OUTPUTS
        [void] - Formulaire modal.
    #>

    # =================================================================
    # Chargement des données dynamiques depuis Azure AD
    # =================================================================
    $script:DynDepartments   = @()
    $script:DynJobTitles     = @()
    $script:DynEmployeeTypes = @()
    $script:DynUsageLocations = @()

    $resDept = Get-AzDistinctValues -Property "department"
    if ($resDept.Success) { $script:DynDepartments = $resDept.Data }

    $resJob = Get-AzDistinctValues -Property "jobTitle"
    if ($resJob.Success) { $script:DynJobTitles = $resJob.Data }

    $resEmpType = Get-AzDistinctValues -Property "employeeType"
    if ($resEmpType.Success) { $script:DynEmployeeTypes = $resEmpType.Data }

    $resLoc = Get-AzDistinctValues -Property "usageLocation"
    if ($resLoc.Success) { $script:DynUsageLocations = $resLoc.Data }

    # =================================================================
    # Construction du formulaire
    # =================================================================
    $form = New-Object System.Windows.Forms.Form
    $form.Text = Get-Text "onboarding.title" $Config.client_name
    $form.Size = New-Object System.Drawing.Size(680, 780)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::WhiteSmoke
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # Panel scrollable pour le contenu
    $panelMain = New-Object System.Windows.Forms.Panel
    $panelMain.Location = New-Object System.Drawing.Point(0, 0)
    $panelMain.Size = New-Object System.Drawing.Size(665, 690)
    $panelMain.AutoScroll = $true
    $form.Controls.Add($panelMain)

    $yPos = 10
    $lblX = 15
    $fldX = 200
    $fldW = 430
    $lblW = 180
    $lineH = 32

    # -----------------------------------------------------------------
    # Fonction utilitaire : créer un ComboBox éditable (dropdown + saisie libre)
    # -----------------------------------------------------------------
    function New-EditableCombo {
        param([string[]]$Items, [int]$Y)
        $cbo = New-Object System.Windows.Forms.ComboBox
        $cbo.Location = New-Object System.Drawing.Point($fldX, $Y)
        $cbo.Size = New-Object System.Drawing.Size($fldW, 25)
        $cbo.DropDownStyle = "DropDown"  # Editable — permet la saisie libre
        $cbo.AutoCompleteMode = "SuggestAppend"
        $cbo.AutoCompleteSource = "ListItems"
        foreach ($item in $Items) {
            $cbo.Items.Add($item) | Out-Null
        }
        return $cbo
    }

    # -----------------------------------------------------------------
    # Fonction utilitaire : ajouter un label
    # -----------------------------------------------------------------
    function Add-Label {
        param([string]$Text, [int]$Y, [bool]$Required = $false)
        $lbl = New-Object System.Windows.Forms.Label
        $displayText = if ($Required) { "$Text *" } else { $Text }
        $lbl.Text = $displayText
        $lbl.Location = New-Object System.Drawing.Point($lblX, ($Y + 4))
        $lbl.Size = New-Object System.Drawing.Size($lblW, 20)
        if ($Required) {
            $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        }
        $panelMain.Controls.Add($lbl)
    }

    # =================================================================
    # SECTION : Identité
    # =================================================================
    $lblSectionId = New-Object System.Windows.Forms.Label
    $lblSectionId.Text = Get-Text "onboarding.section_identity"
    $lblSectionId.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblSectionId.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $lblSectionId.Location = New-Object System.Drawing.Point($lblX, $yPos)
    $lblSectionId.Size = New-Object System.Drawing.Size(620, 22)
    $panelMain.Controls.Add($lblSectionId)
    $yPos += 28

    # --- Nom ---
    Add-Label -Text (Get-Text "onboarding.field_lastname") -Y $yPos -Required $true
    $txtNom = New-Object System.Windows.Forms.TextBox
    $txtNom.Location = New-Object System.Drawing.Point($fldX, $yPos)
    $txtNom.Size = New-Object System.Drawing.Size($fldW, 25)
    $panelMain.Controls.Add($txtNom)
    $yPos += $lineH

    # --- Prénom ---
    Add-Label -Text (Get-Text "onboarding.field_firstname") -Y $yPos -Required $true
    $txtPrenom = New-Object System.Windows.Forms.TextBox
    $txtPrenom.Location = New-Object System.Drawing.Point($fldX, $yPos)
    $txtPrenom.Size = New-Object System.Drawing.Size($fldW, 25)
    $panelMain.Controls.Add($txtPrenom)
    $yPos += $lineH

    # --- Courriel / Logon Name (ComboBox alimenté dynamiquement) ---
    Add-Label -Text (Get-Text "onboarding.field_email") -Y $yPos -Required $true
    $cboUsername = New-Object System.Windows.Forms.ComboBox
    $cboUsername.Location = New-Object System.Drawing.Point($fldX, $yPos)
    $cboUsername.Size = New-Object System.Drawing.Size(340, 25)
    $cboUsername.DropDownStyle = "DropDownList"
    $panelMain.Controls.Add($cboUsername)

    $btnGenererUPN = New-Object System.Windows.Forms.Button
    $btnGenererUPN.Text = Get-Text "onboarding.btn_generate"
    $btnGenererUPN.Location = New-Object System.Drawing.Point(($fldX + 348), $yPos)
    $btnGenererUPN.Size = New-Object System.Drawing.Size(80, 25)
    $btnGenererUPN.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $btnGenererUPN.ForeColor = [System.Drawing.Color]::White
    $btnGenererUPN.FlatStyle = "Flat"
    $panelMain.Controls.Add($btnGenererUPN)
    $yPos += 5

    $lblUpnInfo = New-Object System.Windows.Forms.Label
    $lblUpnInfo.Text = ""
    $lblUpnInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblUpnInfo.ForeColor = [System.Drawing.Color]::Gray
    $lblUpnInfo.Location = New-Object System.Drawing.Point($fldX, ($yPos + 24))
    $lblUpnInfo.Size = New-Object System.Drawing.Size(430, 16)
    $panelMain.Controls.Add($lblUpnInfo)
    $yPos += 42

    # Variable pour stocker les variantes générées
    $script:UsernameVariants = @()

    # Action du bouton Générer
    $btnGenererUPN.Add_Click({
        $nom = $txtNom.Text.Trim()
        $prenom = $txtPrenom.Text.Trim()

        if ([string]::IsNullOrWhiteSpace($nom) -or [string]::IsNullOrWhiteSpace($prenom)) {
            [System.Windows.Forms.MessageBox]::Show(
                    (Get-Text "onboarding.upn_required"),
                    (Get-Text "onboarding.upn_required_title"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return
        }

        $cboUsername.Items.Clear()
        $lblUpnInfo.Text = Get-Text "onboarding.upn_checking"
        $form.Refresh()

        $script:UsernameVariants = New-UsernameVariants -Prenom $prenom -Nom $nom

        foreach ($variant in $script:UsernameVariants) {
            $cboUsername.Items.Add("$($variant.UPN)  [$($variant.Label)]") | Out-Null
        }

        if ($cboUsername.Items.Count -gt 0) {
            $cboUsername.SelectedIndex = 0
            $lblUpnInfo.Text = Get-Text "onboarding.upn_available" $cboUsername.Items.Count
            $lblUpnInfo.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
        }
        else {
            $lblUpnInfo.Text = Get-Text "onboarding.upn_none"
            $lblUpnInfo.ForeColor = [System.Drawing.Color]::Red
        }
    })

    # =================================================================
    # SECTION : Poste
    # =================================================================
    $lblSectionPoste = New-Object System.Windows.Forms.Label
    $lblSectionPoste.Text = Get-Text "onboarding.section_position"
    $lblSectionPoste.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblSectionPoste.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $lblSectionPoste.Location = New-Object System.Drawing.Point($lblX, $yPos)
    $lblSectionPoste.Size = New-Object System.Drawing.Size(620, 22)
    $panelMain.Controls.Add($lblSectionPoste)
    $yPos += 28

    # --- Titre du poste (Job Title) — dynamique + saisie libre ---
    Add-Label -Text (Get-Text "onboarding.field_job_title") -Y $yPos -Required $true
    $cboJobTitle = New-EditableCombo -Items $script:DynJobTitles -Y $yPos
    $panelMain.Controls.Add($cboJobTitle)
    $yPos += $lineH

    # --- Département — dynamique + saisie libre ---
    Add-Label -Text (Get-Text "onboarding.field_department") -Y $yPos -Required $true
    $cboDepartment = New-EditableCombo -Items $script:DynDepartments -Y $yPos
    $panelMain.Controls.Add($cboDepartment)
    $yPos += $lineH

    # --- Employee Type — dynamique + saisie libre ---
    Add-Label -Text (Get-Text "onboarding.field_employee_type") -Y $yPos -Required $true
    $cboEmployeeType = New-EditableCombo -Items $script:DynEmployeeTypes -Y $yPos
    $panelMain.Controls.Add($cboEmployeeType)
    $yPos += $lineH

    # --- Matricule (Employee ID) — optionnel ---
    Add-Label -Text (Get-Text "onboarding.field_employee_id") -Y $yPos -Required $false
    $txtEmployeeId = New-Object System.Windows.Forms.TextBox
    $txtEmployeeId.Location = New-Object System.Drawing.Point($fldX, $yPos)
    $txtEmployeeId.Size = New-Object System.Drawing.Size($fldW, 25)
    $panelMain.Controls.Add($txtEmployeeId)
    $yPos += $lineH

    # --- Gestionnaire (Manager) — recherche dynamique ---
    Add-Label -Text (Get-Text "onboarding.field_manager") -Y $yPos -Required $false
    $txtManagerSearch = New-Object System.Windows.Forms.TextBox
    $txtManagerSearch.Location = New-Object System.Drawing.Point($fldX, $yPos)
    $txtManagerSearch.Size = New-Object System.Drawing.Size(340, 25)
    $panelMain.Controls.Add($txtManagerSearch)

    $btnManagerSearch = New-Object System.Windows.Forms.Button
    $btnManagerSearch.Text = Get-Text "onboarding.btn_search"
    $btnManagerSearch.Location = New-Object System.Drawing.Point(($fldX + 348), $yPos)
    $btnManagerSearch.Size = New-Object System.Drawing.Size(80, 25)
    $btnManagerSearch.FlatStyle = "Flat"
    $panelMain.Controls.Add($btnManagerSearch)
    $yPos += 28

    $lstManager = New-Object System.Windows.Forms.ListBox
    $lstManager.Location = New-Object System.Drawing.Point($fldX, $yPos)
    $lstManager.Size = New-Object System.Drawing.Size(430, 60)
    $lstManager.Visible = $false
    $panelMain.Controls.Add($lstManager)

    $lblManagerSelected = New-Object System.Windows.Forms.Label
    $lblManagerSelected.Text = ""
    $lblManagerSelected.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblManagerSelected.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $lblManagerSelected.Location = New-Object System.Drawing.Point($fldX, $yPos)
    $lblManagerSelected.Size = New-Object System.Drawing.Size(430, 16)
    $panelMain.Controls.Add($lblManagerSelected)

    $script:SelectedManagerId = $null
    $script:SelectedManagerName = $null
    $script:ManagerSearchResults = @()

    $btnManagerSearch.Add_Click({
        $terme = $txtManagerSearch.Text.Trim()
        if ($terme.Length -lt 2) {
            [System.Windows.Forms.MessageBox]::Show(    (Get-Text "onboarding.manager_min_chars"), (Get-Text "onboarding.manager_search_title"), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }
        $lstManager.Items.Clear()
        $result = Search-AzUsers -SearchTerm $terme -MaxResults 10
        if ($result.Success -and $result.Data) {
            $script:ManagerSearchResults = @($result.Data)
            foreach ($user in $script:ManagerSearchResults) {
                $lstManager.Items.Add("$($user.DisplayName) — $($user.UserPrincipalName)") | Out-Null
            }
            $lstManager.Visible = $true
            $lblManagerSelected.Visible = $false
        }
        else {
            $lstManager.Visible = $false
            [System.Windows.Forms.MessageBox]::Show(    (Get-Text "onboarding.manager_no_result"), (Get-Text "onboarding.manager_search_title"), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
    })

    $txtManagerSearch.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $btnManagerSearch.PerformClick()
            $_.SuppressKeyPress = $true
        }
    })

    $lstManager.Add_SelectedIndexChanged({
        if ($lstManager.SelectedIndex -ge 0 -and $lstManager.SelectedIndex -lt $script:ManagerSearchResults.Count) {
            $sel = $script:ManagerSearchResults[$lstManager.SelectedIndex]
            $script:SelectedManagerId = $sel.Id
            $script:SelectedManagerName = $sel.DisplayName
            $txtManagerSearch.Text = $sel.DisplayName
            $lstManager.Visible = $false
            $lblManagerSelected.Text = "✓ $($sel.DisplayName) ($($sel.UserPrincipalName))"
            $lblManagerSelected.Visible = $true
        }
    })

    $yPos += 22

    # --- Usage Location — dynamique + saisie libre ---
    Add-Label -Text (Get-Text "onboarding.field_usage_location") -Y $yPos -Required $true
    $cboUsageLocation = New-EditableCombo -Items $script:DynUsageLocations -Y $yPos
    # Pré-sélectionner CA si disponible
    $idxCA = $cboUsageLocation.Items.IndexOf("CA")
    if ($idxCA -ge 0) { $cboUsageLocation.SelectedIndex = $idxCA }
    elseif ($cboUsageLocation.Items.Count -gt 0) { $cboUsageLocation.SelectedIndex = 0 }
    $panelMain.Controls.Add($cboUsageLocation)
    $yPos += $lineH

    # =================================================================
    # SECTION : Licence
    # =================================================================
    $lblSectionLic = New-Object System.Windows.Forms.Label
    $lblSectionLic.Text = Get-Text "onboarding.section_license"
    $lblSectionLic.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblSectionLic.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $lblSectionLic.Location = New-Object System.Drawing.Point($lblX, $yPos)
    $lblSectionLic.Size = New-Object System.Drawing.Size(620, 22)
    $panelMain.Controls.Add($lblSectionLic)
    $yPos += 28

    # --- License Pack — dropdown depuis le JSON ---
    Add-Label -Text (Get-Text "onboarding.field_license_group") -Y $yPos -Required $false
    $cboLicenseGroup = New-Object System.Windows.Forms.ComboBox
    $cboLicenseGroup.Location = New-Object System.Drawing.Point($fldX, $yPos)
    $cboLicenseGroup.Size = New-Object System.Drawing.Size($fldW, 25)
    $cboLicenseGroup.DropDownStyle = "DropDownList"
    $cboLicenseGroup.Items.Add((Get-Text "onboarding.license_none")) | Out-Null
    foreach ($grp in $Config.license_groups) {
        $cboLicenseGroup.Items.Add($grp) | Out-Null
    }
    $cboLicenseGroup.SelectedIndex = 0
    $panelMain.Controls.Add($cboLicenseGroup)
    $yPos += $lineH

    # =================================================================
    # SECTION : Groupes d'appartenance
    # =================================================================
    $lblSectionGrp = New-Object System.Windows.Forms.Label
    $lblSectionGrp.Text = Get-Text "onboarding.section_groups"
    $lblSectionGrp.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblSectionGrp.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $lblSectionGrp.Location = New-Object System.Drawing.Point($lblX, $yPos)
    $lblSectionGrp.Size = New-Object System.Drawing.Size(620, 22)
    $panelMain.Controls.Add($lblSectionGrp)
    $yPos += 26

    # CheckedListBox pour sélection multiple des groupes
    $clbGroupes = New-Object System.Windows.Forms.CheckedListBox
    $clbGroupes.Location = New-Object System.Drawing.Point($fldX, $yPos)
    $clbGroupes.Size = New-Object System.Drawing.Size($fldW, 80)
    $clbGroupes.CheckOnClick = $true
    foreach ($grp in $Config.membership_groups) {
        $clbGroupes.Items.Add($grp, $false) | Out-Null
    }
    $panelMain.Controls.Add($clbGroupes)
    $yPos += 88

    # =================================================================
    # SECTION : Mot de passe (info)
    # =================================================================
    $lblSectionPwd = New-Object System.Windows.Forms.Label
    $lblSectionPwd.Text = Get-Text "onboarding.section_password"
    $lblSectionPwd.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblSectionPwd.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $lblSectionPwd.Location = New-Object System.Drawing.Point($lblX, $yPos)
    $lblSectionPwd.Size = New-Object System.Drawing.Size(620, 22)
    $panelMain.Controls.Add($lblSectionPwd)
    $yPos += 26

    $lblPwdInfo = New-Object System.Windows.Forms.Label
    $lblPwdInfo.Text = Get-Text "onboarding.password_info" $Config.password_policy.length
    $lblPwdInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblPwdInfo.ForeColor = [System.Drawing.Color]::Gray
    $lblPwdInfo.Location = New-Object System.Drawing.Point($fldX, $yPos)
    $lblPwdInfo.Size = New-Object System.Drawing.Size(430, 30)
    $panelMain.Controls.Add($lblPwdInfo)
    $yPos += 35

    # =================================================================
    # Label erreur et chargement
    # =================================================================
    $lblErreur = New-Object System.Windows.Forms.Label
    $lblErreur.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $lblErreur.ForeColor = [System.Drawing.Color]::Red
    $lblErreur.Location = New-Object System.Drawing.Point($lblX, $yPos)
    $lblErreur.Size = New-Object System.Drawing.Size(620, 20)
    $lblErreur.Visible = $false
    $panelMain.Controls.Add($lblErreur)

    $lblChargement = New-Object System.Windows.Forms.Label
    $lblChargement.Text = Get-Text "onboarding.creating"
    $lblChargement.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $lblChargement.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $lblChargement.Location = New-Object System.Drawing.Point($lblX, $yPos)
    $lblChargement.Size = New-Object System.Drawing.Size(300, 20)
    $lblChargement.Visible = $false
    $panelMain.Controls.Add($lblChargement)

    # =================================================================
    # Boutons Créer / Annuler
    # =================================================================
    $btnCreer = New-Object System.Windows.Forms.Button
    $btnCreer.Text = Get-Text "onboarding.btn_create"
    $btnCreer.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnCreer.Location = New-Object System.Drawing.Point(310, 700)
    $btnCreer.Size = New-Object System.Drawing.Size(170, 38)
    $btnCreer.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $btnCreer.ForeColor = [System.Drawing.Color]::White
    $btnCreer.FlatStyle = "Flat"
    $form.Controls.Add($btnCreer)

    $btnAnnuler = New-Object System.Windows.Forms.Button
    $btnAnnuler.Text = Get-Text "onboarding.btn_cancel"
    $btnAnnuler.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnAnnuler.Location = New-Object System.Drawing.Point(490, 700)
    $btnAnnuler.Size = New-Object System.Drawing.Size(140, 38)
    $btnAnnuler.FlatStyle = "Flat"
    $btnAnnuler.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnAnnuler)
    $form.CancelButton = $btnAnnuler

    # =================================================================
    # ACTION : Création du compte
    # =================================================================
    $btnCreer.Add_Click({
        $lblErreur.Visible = $false

        # --- Validation des champs obligatoires ---
        $nom       = $txtNom.Text.Trim()
        $prenom    = $txtPrenom.Text.Trim()
        $jobTitle  = $cboJobTitle.Text.Trim()
        $dept      = $cboDepartment.Text.Trim()
        $empType   = $cboEmployeeType.Text.Trim()
        $usageLoc  = $cboUsageLocation.Text.Trim()

        if ([string]::IsNullOrWhiteSpace($nom)) {
            $lblErreur.Text = Get-Text "onboarding.validation_lastname"
            $lblErreur.Visible = $true; return
        }
        if ([string]::IsNullOrWhiteSpace($prenom)) {
            $lblErreur.Text = Get-Text "onboarding.validation_firstname"
            $lblErreur.Visible = $true; return
        }
        if ($cboUsername.SelectedIndex -lt 0 -or $script:UsernameVariants.Count -eq 0) {
            $lblErreur.Text = Get-Text "onboarding.validation_email"
            $lblErreur.Visible = $true; return
        }
        if ([string]::IsNullOrWhiteSpace($jobTitle)) {
            $lblErreur.Text = Get-Text "onboarding.validation_job_title"
            $lblErreur.Visible = $true; return
        }
        if ([string]::IsNullOrWhiteSpace($dept)) {
            $lblErreur.Text = Get-Text "onboarding.validation_department"
            $lblErreur.Visible = $true; return
        }
        if ([string]::IsNullOrWhiteSpace($empType)) {
            $lblErreur.Text = Get-Text "onboarding.validation_employee_type"
            $lblErreur.Visible = $true; return
        }
        if ([string]::IsNullOrWhiteSpace($usageLoc)) {
            $lblErreur.Text = Get-Text "onboarding.validation_usage_location"
            $lblErreur.Visible = $true; return
        }

        # Récupérer la variante sélectionnée
        $selectedVariant = $script:UsernameVariants[$cboUsername.SelectedIndex]
        $upn = $selectedVariant.UPN
        $mailNickname = $selectedVariant.MailNickname

        # --- Récapitulatif et confirmation ---
        $confirmMsg  = "$(Get-Text 'onboarding.confirm_header')`n`n"
        $confirmMsg += "$(Get-Text 'onboarding.confirm_fullname')     : $prenom $nom`n"
        $confirmMsg += "$(Get-Text 'onboarding.confirm_upn')             : $upn`n"
        $confirmMsg += "$(Get-Text 'onboarding.confirm_title_field')           : $jobTitle`n"
        $confirmMsg += "$(Get-Text 'onboarding.confirm_department')     : $dept`n"
        $confirmMsg += "$(Get-Text 'onboarding.confirm_employee_type')   : $empType`n"
        $confirmMsg += "$(Get-Text 'onboarding.confirm_usage_location')  : $usageLoc`n"
        if (-not [string]::IsNullOrWhiteSpace($txtEmployeeId.Text)) {
            $confirmMsg += "$(Get-Text 'onboarding.confirm_employee_id')       : $($txtEmployeeId.Text.Trim())`n"
        }
        if ($script:SelectedManagerName) {
            $confirmMsg += "$(Get-Text 'onboarding.confirm_manager')    : $($script:SelectedManagerName)`n"
        }
        # Licence
        $licGroup = $null
        if ($cboLicenseGroup.SelectedIndex -gt 0) {
            $licGroup = $cboLicenseGroup.SelectedItem.ToString()
            $confirmMsg += "$(Get-Text 'onboarding.confirm_license'): $licGroup`n"
        }
        # Groupes
        $groupesSelectionnes = @()
        for ($i = 0; $i -lt $clbGroupes.Items.Count; $i++) {
            if ($clbGroupes.GetItemChecked($i)) {
                $groupesSelectionnes += $clbGroupes.Items[$i].ToString()
            }
        }
        if ($groupesSelectionnes.Count -gt 0) {
            $confirmMsg += "$(Get-Text 'onboarding.confirm_groups')         : $($groupesSelectionnes -join ', ')`n"
        }

        $confirm = Show-ConfirmDialog -Titre (Get-Text "onboarding.confirm_title") -Message $confirmMsg
        if (-not $confirm) { return }

        # --- Exécution ---
        $btnCreer.Enabled = $false
        $lblChargement.Visible = $true
        $form.Refresh()

        try {
            # 1. Générer le mot de passe
            $password = New-SecurePassword

            # 2. Construire les paramètres
            $userParams = @{
                DisplayName       = "$prenom $nom"
                GivenName         = $prenom
                Surname           = $nom
                UserPrincipalName = $upn
                MailNickname      = $mailNickname
                Password          = $password
                Department        = $dept
                JobTitle          = $jobTitle
                EmployeeType      = $empType
                UsageLocation     = $usageLoc
                ForceChangePasswordNextSignIn = $Config.password_policy.force_change_at_login
            }

            if (-not [string]::IsNullOrWhiteSpace($txtEmployeeId.Text)) {
                $userParams.EmployeeId = $txtEmployeeId.Text.Trim()
            }

            # 3. Création du compte
            $createResult = New-AzUser -UserParams $userParams
            if (-not $createResult.Success) {
                throw (Get-Text "onboarding.error_creation_failed" $createResult.Error)
            }

            $newUserId = $createResult.Data.Id
            $erreurs = @()

            # 4. Attribution du manager
            if ($script:SelectedManagerId) {
                $mgrResult = Set-AzUserManager -UserId $newUserId -ManagerId $script:SelectedManagerId
                if (-not $mgrResult.Success) { $erreurs += "Manager : $($mgrResult.Error)" }
            }

            # 5. Ajout au groupe de licence
            if ($licGroup) {
                $licResult = Add-AzUserToGroup -UserId $newUserId -GroupName $licGroup
                if (-not $licResult.Success) { $erreurs += "Licence ($licGroup) : $($licResult.Error)" }
            }

            # 6. Ajout aux groupes d'appartenance
            foreach ($grp in $groupesSelectionnes) {
                $grpResult = Add-AzUserToGroup -UserId $newUserId -GroupName $grp
                if (-not $grpResult.Success) { $erreurs += "Groupe '$grp' : $($grpResult.Error)" }
            }

            # 7. Notification
            if ($Config.notifications.enabled) {
                $sujet = "Onboarding - $prenom $nom ($upn)"
                $corps = "<h2>Nouveau compte employé créé</h2>"
                $corps += "<table style='border-collapse:collapse'>"
                $corps += "<tr><td><b>Nom :</b></td><td>$prenom $nom</td></tr>"
                $corps += "<tr><td><b>UPN :</b></td><td>$upn</td></tr>"
                $corps += "<tr><td><b>Poste :</b></td><td>$jobTitle</td></tr>"
                $corps += "<tr><td><b>Département :</b></td><td>$dept</td></tr>"
                $corps += "<tr><td><b>Type :</b></td><td>$empType</td></tr>"
                if ($script:SelectedManagerName) {
                    $corps += "<tr><td><b>Gestionnaire :</b></td><td>$($script:SelectedManagerName)</td></tr>"
                }
                $corps += "</table>"
                $corps += "<p><em>Mot de passe temporaire transmis séparément.</em></p>"
                Send-Notification -Sujet $sujet -Corps $corps
            }

            # 8. Résultat
            $lblChargement.Visible = $false

            if ($erreurs.Count -gt 0) {
                $warnMsg = (Get-Text "onboarding.result_partial_msg" ($erreurs -join "`n"))
                Show-ResultDialog -Titre (Get-Text "onboarding.result_partial_title") -Message $warnMsg -IsSuccess $false
            }

            Write-Log -Level "SUCCESS" -Action "ONBOARDING" -UPN $upn -Message "Onboarding terminé. Erreurs: $($erreurs.Count)"

            # 9. Affichage du mot de passe (une seule fois)
            Show-PasswordDialog -UPN $upn -InitialToken $password

            $form.Close()
        }
        catch {
            $lblChargement.Visible = $false
            $btnCreer.Enabled = $true
            $errMsg = $_.Exception.Message
            Write-Log -Level "ERROR" -Action "ONBOARDING" -UPN $upn -Message "Erreur onboarding : $errMsg"
            Show-ResultDialog -Titre (Get-Text "onboarding.result_error_title") -Message (Get-Text "onboarding.result_error_msg" $errMsg) -IsSuccess $false
        }
    })

    # =================================================================
    # Affichage
    # =================================================================
    $form.ShowDialog() | Out-Null
    $form.Dispose()
}

# Point d'attention :
# - Les champs dynamiques (department, jobTitle, etc.) sont chargés une fois à l'ouverture du formulaire
# - Les ComboBox sont en mode "DropDown" (éditable) pour permettre la saisie libre de nouvelles valeurs
# - Le bouton "Générer" construit les 3 variantes de UPN et vérifie leur unicité dans Azure AD
# - Si un UPN existe déjà, le script ajoute une incrémentation (1, 2, 3...)
# - La licence est gérée par ajout à un groupe (et non par assignation directe de SKU)
# - Les groupes d'appartenance utilisent un CheckedListBox pour la sélection multiple
# - Le mot de passe est généré UNIQUEMENT à la soumission et affiché une seule fois
