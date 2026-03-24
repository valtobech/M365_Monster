<#
.FICHIER
    Modules/GUI_Onboarding.ps1

.ROLE
    Formulaire d'arrivée (onboarding) d'un nouvel employé.
    Champs dynamiques alimentés depuis Azure AD (department, jobTitle, etc.).
    Génération automatique du username avec 3 formats au choix et vérification d'unicité.
    Licences via recherche dynamique par préfixe (configurable dans Settings).
    Groupes d'appartenance via recherche dynamique dans Entra ID.

.DEPENDANCES
    - Core/Functions.ps1 (New-UsernameVariants, New-SecurePassword, Write-Log,
      Show-ConfirmDialog, Show-ResultDialog, Show-PasswordDialog, Get-MailNickname,
      Remove-Diacritics)
    - Core/GraphAPI.ps1 (New-AzUser, Add-AzUserToGroup, Set-AzUserManager,
      Search-AzUsers, Search-AzGroups, Get-AzDistinctValues, Test-AzUserExists)
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
    $dynData = @{}
    foreach ($prop in @("department", "jobTitle", "employeeType", "usageLocation")) {
        $res = Get-AzDistinctValues -Property $prop
        $dynData[$prop] = if ($res.Success) { $res.Data } else { @() }
    }

    # =================================================================
    # Construction du formulaire — redimensionnable
    # =================================================================
    $form = New-Object System.Windows.Forms.Form
    $form.Text = Get-Text "onboarding.title" $Config.client_name
    $form.Size = New-Object System.Drawing.Size(1060, 920)
    $form.MinimumSize = New-Object System.Drawing.Size(900, 700)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "Sizable"
    $form.MaximizeBox = $true
    $form.MinimizeBox = $true
    $form.BackColor = [System.Drawing.Color]::WhiteSmoke
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # Panel scrollable pour le contenu — ancré sur les 4 côtés
    $panelMain = New-Object System.Windows.Forms.Panel
    $panelMain.Location = New-Object System.Drawing.Point(0, 0)
    $panelMain.Size = New-Object System.Drawing.Size(1040, 830)
    $panelMain.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right
    $panelMain.AutoScroll = $true
    $form.Controls.Add($panelMain)

    $yPos = 10
    $lblX = 15
    $fldX = 200
    $fldW = 430
    $lblW = 180
    $lineH = 32

    # -----------------------------------------------------------------
    # Fonctions utilitaires : création de contrôles GUI réutilisables
    # -----------------------------------------------------------------
    function New-EditableCombo {
        param([string[]]$Items, [int]$Y)
        $cbo = New-Object System.Windows.Forms.ComboBox
        $cbo.Location = New-Object System.Drawing.Point($fldX, $Y)
        $cbo.Size = New-Object System.Drawing.Size($fldW, 25)
        $cbo.DropDownStyle = "DropDown"
        $cbo.AutoCompleteMode = "SuggestAppend"
        $cbo.AutoCompleteSource = "ListItems"
        foreach ($item in $Items) { $cbo.Items.Add($item) | Out-Null }
        return $cbo
    }

    function Add-Label {
        param([string]$Text, [int]$Y, [bool]$Required = $false)
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = if ($Required) { "$Text *" } else { $Text }
        $lbl.Location = New-Object System.Drawing.Point($lblX, ($Y + 4))
        $lbl.Size = New-Object System.Drawing.Size($lblW, 20)
        if ($Required) {
            $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        }
        $panelMain.Controls.Add($lbl)
    }

    function Add-SectionHeader {
        param([string]$TextKey, [ref]$YRef)
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = Get-Text $TextKey
        $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $lbl.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
        $lbl.Location = New-Object System.Drawing.Point($lblX, $YRef.Value)
        $lbl.Size = New-Object System.Drawing.Size(620, 22)
        $panelMain.Controls.Add($lbl)
        $YRef.Value += 28
    }

    # =================================================================
    # SECTION : Identité
    # =================================================================
    Add-SectionHeader -TextKey "onboarding.section_identity" -YRef ([ref]$yPos)

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

    # --- Courriel / Logon Name ---
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
    Add-SectionHeader -TextKey "onboarding.section_position" -YRef ([ref]$yPos)

    # --- Titre du poste ---
    Add-Label -Text (Get-Text "onboarding.field_job_title") -Y $yPos -Required $true
    $cboJobTitle = New-EditableCombo -Items $dynData["jobTitle"] -Y $yPos
    $panelMain.Controls.Add($cboJobTitle)
    $yPos += $lineH

    # --- Département ---
    Add-Label -Text (Get-Text "onboarding.field_department") -Y $yPos -Required $true
    $cboDepartment = New-EditableCombo -Items $dynData["department"] -Y $yPos
    $panelMain.Controls.Add($cboDepartment)
    $yPos += $lineH

    # --- Employee Type ---
    Add-Label -Text (Get-Text "onboarding.field_employee_type") -Y $yPos -Required $true
    $cboEmployeeType = New-EditableCombo -Items $dynData["employeeType"] -Y $yPos
    $panelMain.Controls.Add($cboEmployeeType)
    $yPos += $lineH

    # --- Matricule (Employee ID) — optionnel ---
    Add-Label -Text (Get-Text "onboarding.field_employee_id") -Y $yPos -Required $false
    $txtEmployeeId = New-Object System.Windows.Forms.TextBox
    $txtEmployeeId.Location = New-Object System.Drawing.Point($fldX, $yPos)
    $txtEmployeeId.Size = New-Object System.Drawing.Size($fldW, 25)
    $panelMain.Controls.Add($txtEmployeeId)
    $yPos += $lineH

    # --- Gestionnaire (Manager) — recherche dynamique — OBLIGATOIRE ---
    Add-Label -Text (Get-Text "onboarding.field_manager") -Y $yPos -Required $true
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
            [System.Windows.Forms.MessageBox]::Show(
                (Get-Text "onboarding.manager_min_chars"),
                (Get-Text "onboarding.manager_search_title"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
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
            [System.Windows.Forms.MessageBox]::Show(
                (Get-Text "onboarding.manager_no_result"),
                (Get-Text "onboarding.manager_search_title"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
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

    # --- Usage Location ---
    Add-Label -Text (Get-Text "onboarding.field_usage_location") -Y $yPos -Required $true
    $cboUsageLocation = New-EditableCombo -Items $dynData["usageLocation"] -Y $yPos
    $idxCA = $cboUsageLocation.Items.IndexOf("CA")
    if ($idxCA -ge 0) { $cboUsageLocation.SelectedIndex = $idxCA }
    elseif ($cboUsageLocation.Items.Count -gt 0) { $cboUsageLocation.SelectedIndex = 0 }
    $panelMain.Controls.Add($cboUsageLocation)
    $yPos += $lineH

    # =================================================================
    # SECTION : Licence — startsWith préfixe + multi-sélection
    # =================================================================
    Add-SectionHeader -TextKey "onboarding.section_license" -YRef ([ref]$yPos)

    # Fonction locale : charger les groupes filtrés strictement par startsWith
    $script:LicenseGroupNames = @()

    function Update-LicenseList {
        $clbLicenses.Items.Clear()
        $script:LicenseGroupNames = @()
        $prefix = if ($Config.PSObject.Properties["license_group_prefix"]) { $Config.license_group_prefix } else { "" }
        if ([string]::IsNullOrWhiteSpace($prefix)) { return }

        $res = Search-AzGroups -SearchTerm $prefix -MaxResults 50
        if ($res.Success -and $res.Data) {
            # Filtrage strict startsWith — Graph $search fait un "contains"
            $filtered = $res.Data | Where-Object { $_.DisplayName -like "$prefix*" } | Sort-Object DisplayName
            foreach ($grp in $filtered) {
                $clbLicenses.Items.Add($grp.DisplayName, $false) | Out-Null
                $script:LicenseGroupNames += $grp.DisplayName
            }
        }
    }

    Add-Label -Text (Get-Text "onboarding.field_license_group") -Y $yPos -Required $false
    $clbLicenses = New-Object System.Windows.Forms.CheckedListBox
    $clbLicenses.Location = New-Object System.Drawing.Point($fldX, $yPos)
    $clbLicenses.Size = New-Object System.Drawing.Size(340, 80)
    $clbLicenses.CheckOnClick = $true
    $panelMain.Controls.Add($clbLicenses)

    # Bouton rafraîchir
    $btnRefreshLic = New-Object System.Windows.Forms.Button
    $btnRefreshLic.Text = "⟳"
    $btnRefreshLic.Location = New-Object System.Drawing.Point(($fldX + 348), $yPos)
    $btnRefreshLic.Size = New-Object System.Drawing.Size(35, 25)
    $btnRefreshLic.FlatStyle = "Flat"
    $btnRefreshLic.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnRefreshLic.Add_Click({ Update-LicenseList })
    $panelMain.Controls.Add($btnRefreshLic)

    # Chargement initial des licences
    Update-LicenseList

    # Info préfixe
    $licPrefix = if ($Config.PSObject.Properties["license_group_prefix"]) { $Config.license_group_prefix } else { "" }
    $lblLicInfo = New-Object System.Windows.Forms.Label
    $lblLicInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblLicInfo.ForeColor = [System.Drawing.Color]::Gray
    $lblLicInfo.Location = New-Object System.Drawing.Point($fldX, ($yPos + 82))
    $lblLicInfo.Size = New-Object System.Drawing.Size(430, 16)
    if (-not [string]::IsNullOrWhiteSpace($licPrefix)) {
        $lblLicInfo.Text = Get-Text "onboarding.license_prefix_info" $licPrefix
    }
    else {
        $lblLicInfo.Text = Get-Text "onboarding.license_no_prefix"
    }
    $panelMain.Controls.Add($lblLicInfo)
    $yPos += 104

    # =================================================================
    # SECTION : Profils d'accès (si configurés)
    # =================================================================
    $script:OnbProfileKeys = @()
    $clbProfiles = $null

    if ($Config.PSObject.Properties["access_profiles"]) {
        $onbProfiles = Get-AccessProfiles -ExcludeBaseline
        $onbBaseline = Get-BaselineProfile

        if ($onbProfiles.Count -gt 0 -or $onbBaseline) {
            Add-SectionHeader -TextKey "onboarding.section_access_profiles" -YRef ([ref]$yPos)

            if ($onbBaseline) {
                Add-Label -Text (Get-Text "onboarding.profile_baseline_label") -Y $yPos -Required $false
                $lblBaselineVal = New-Object System.Windows.Forms.Label
                $lblBaselineVal.Text = "$($onbBaseline.DisplayName) ($($onbBaseline.Groups.Count) groupe(s))"
                $lblBaselineVal.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
                $lblBaselineVal.ForeColor = [System.Drawing.Color]::FromArgb(40, 120, 40)
                $lblBaselineVal.Location = New-Object System.Drawing.Point($fldX, ($yPos + 2))
                $lblBaselineVal.Size = New-Object System.Drawing.Size($fldW, 20)
                $panelMain.Controls.Add($lblBaselineVal)
                $yPos += $lineH
            }

            if ($onbProfiles.Count -gt 0) {
                Add-Label -Text (Get-Text "onboarding.profile_select_label") -Y $yPos -Required $false
                $clbProfiles = New-Object System.Windows.Forms.CheckedListBox
                $clbProfiles.Location = New-Object System.Drawing.Point($fldX, $yPos)
                $clbProfiles.Size = New-Object System.Drawing.Size($fldW, ([Math]::Min(($onbProfiles.Count * 20 + 4), 120)))
                $clbProfiles.CheckOnClick = $true
                foreach ($ap in $onbProfiles) {
                    $clbProfiles.Items.Add("$($ap.DisplayName) — $($ap.Description)", $false) | Out-Null
                }
                $panelMain.Controls.Add($clbProfiles)
                $yPos += $clbProfiles.Size.Height + 8
            }

            # Bouton Prévisualiser
            $btnPreviewGrp = New-Object System.Windows.Forms.Button
            $btnPreviewGrp.Text = Get-Text "onboarding.profile_preview_btn"
            $btnPreviewGrp.Location = New-Object System.Drawing.Point($fldX, $yPos)
            $btnPreviewGrp.Size = New-Object System.Drawing.Size(200, 28)
            $btnPreviewGrp.FlatStyle = "Flat"
            $btnPreviewGrp.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
            $btnPreviewGrp.ForeColor = [System.Drawing.Color]::White
            $btnPreviewGrp.Add_Click({
                $selectedKeys = @()
                if ($clbProfiles) {
                    for ($i = 0; $i -lt $clbProfiles.Items.Count; $i++) {
                        if ($clbProfiles.GetItemChecked($i)) {
                            $selectedKeys += $onbProfiles[$i].Key
                        }
                    }
                }
                $diff = Compare-AccessProfileGroups -OldProfileKeys @() -NewProfileKeys $selectedKeys -IncludeBaseline
                $previewLines = @()
                foreach ($g in $diff.ToAdd) {
                    $source = ""
                    if ($onbBaseline) {
                        $baseGrpIds = @($onbBaseline.Groups | ForEach-Object { $_.id })
                        if ($g.id -in $baseGrpIds) { $source = Get-Text "onboarding.profile_preview_source" $onbBaseline.DisplayName }
                    }
                    if (-not $source) {
                        foreach ($ap in $onbProfiles) {
                            if ($ap.Key -in $selectedKeys) {
                                $apGrpIds = @($ap.Groups | ForEach-Object { $_.id })
                                if ($g.id -in $apGrpIds) { $source = Get-Text "onboarding.profile_preview_source" $ap.DisplayName; break }
                            }
                        }
                    }
                    $previewLines += "+ $($g.display_name)  $source"
                }
                $total = Get-Text "onboarding.profile_preview_total" $diff.ToAdd.Count
                $msg = "$total`n`n$($previewLines -join "`n")"
                [System.Windows.Forms.MessageBox]::Show($msg, (Get-Text "onboarding.profile_preview_title"),
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            })
            $panelMain.Controls.Add($btnPreviewGrp)
            $yPos += 36
        }
    }

    # =================================================================
    # SECTION : Groupes d'appartenance — layout côte à côte
    # Résultats de recherche à GAUCHE, groupes sélectionnés à DROITE
    # =================================================================
    Add-SectionHeader -TextKey "onboarding.section_groups" -YRef ([ref]$yPos)

    # --- Colonne gauche : Recherche ---
    $grpLeftX = $fldX
    $grpColW = 380
    $grpRightX = $fldX + $grpColW + 20  # Espacement de 20px entre les deux colonnes

    Add-Label -Text (Get-Text "onboarding.group_search_label") -Y $yPos -Required $false
    $txtGroupSearch = New-Object System.Windows.Forms.TextBox
    $txtGroupSearch.Location = New-Object System.Drawing.Point($grpLeftX, $yPos)
    $txtGroupSearch.Size = New-Object System.Drawing.Size(290, 25)
    $panelMain.Controls.Add($txtGroupSearch)

    $btnGroupSearch = New-Object System.Windows.Forms.Button
    $btnGroupSearch.Text = Get-Text "onboarding.btn_search"
    $btnGroupSearch.Location = New-Object System.Drawing.Point(($grpLeftX + 298), $yPos)
    $btnGroupSearch.Size = New-Object System.Drawing.Size(80, 25)
    $btnGroupSearch.FlatStyle = "Flat"
    $panelMain.Controls.Add($btnGroupSearch)

    # Label "Groupes ajoutés" à droite (même ligne que le champ de recherche)
    $lblGrpSelectedTitle = New-Object System.Windows.Forms.Label
    $lblGrpSelectedTitle.Text = Get-Text "onboarding.group_selected_label"
    $lblGrpSelectedTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblGrpSelectedTitle.Location = New-Object System.Drawing.Point($grpRightX, ($yPos + 4))
    $lblGrpSelectedTitle.Size = New-Object System.Drawing.Size($grpColW, 20)
    $panelMain.Controls.Add($lblGrpSelectedTitle)
    $yPos += 28

    # Info aide recherche
    $lblGroupHelp = New-Object System.Windows.Forms.Label
    $lblGroupHelp.Text = Get-Text "onboarding.group_search_help"
    $lblGroupHelp.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblGroupHelp.ForeColor = [System.Drawing.Color]::Gray
    $lblGroupHelp.Location = New-Object System.Drawing.Point($grpLeftX, $yPos)
    $lblGroupHelp.Size = New-Object System.Drawing.Size($grpColW, 16)
    $panelMain.Controls.Add($lblGroupHelp)
    $yPos += 20

    # Sauvegarder le Y de départ pour les deux colonnes
    $grpRowY = $yPos

    # --- Colonne gauche : ListBox résultats de recherche (toujours visible, vide par défaut) ---
    $lstGroupResults = New-Object System.Windows.Forms.ListBox
    $lstGroupResults.Location = New-Object System.Drawing.Point($grpLeftX, $grpRowY)
    $lstGroupResults.Size = New-Object System.Drawing.Size($grpColW, 140)
    $panelMain.Controls.Add($lstGroupResults)
    $script:GroupSearchResults = @()

    # --- Colonne droite : CheckedListBox des groupes sélectionnés ---
    $clbGroupes = New-Object System.Windows.Forms.CheckedListBox
    $clbGroupes.Location = New-Object System.Drawing.Point($grpRightX, $grpRowY)
    $clbGroupes.Size = New-Object System.Drawing.Size($grpColW, 140)
    $clbGroupes.CheckOnClick = $true
    $panelMain.Controls.Add($clbGroupes)

    # Label compteur sous la colonne droite
    $lblGroupCount = New-Object System.Windows.Forms.Label
    $lblGroupCount.Text = Get-Text "onboarding.group_selected_count" 0
    $lblGroupCount.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblGroupCount.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $lblGroupCount.Location = New-Object System.Drawing.Point($grpRightX, ($grpRowY + 144))
    $lblGroupCount.Size = New-Object System.Drawing.Size($grpColW, 16)
    $panelMain.Controls.Add($lblGroupCount)

    $yPos = $grpRowY + 166

    # Action recherche de groupes
    $btnGroupSearch.Add_Click({
        $terme = $txtGroupSearch.Text.Trim()
        if ($terme.Length -lt 2) {
            [System.Windows.Forms.MessageBox]::Show(
                (Get-Text "onboarding.group_min_chars"),
                (Get-Text "onboarding.group_search_title"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }
        $lstGroupResults.Items.Clear()
        $result = Search-AzGroups -SearchTerm $terme -MaxResults 20
        if ($result.Success -and $result.Data) {
            $script:GroupSearchResults = @($result.Data | Sort-Object DisplayName)
            foreach ($grp in $script:GroupSearchResults) {
                $lstGroupResults.Items.Add($grp.DisplayName) | Out-Null
            }
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                (Get-Text "onboarding.group_no_result"),
                (Get-Text "onboarding.group_search_title"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        }
    })

    $txtGroupSearch.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $btnGroupSearch.PerformClick()
            $_.SuppressKeyPress = $true
        }
    })

    # Double-clic sur un résultat → ajoute au CheckedListBox à droite (coché par défaut)
    $lstGroupResults.Add_DoubleClick({
        if ($lstGroupResults.SelectedIndex -ge 0) {
            $grpName = $script:GroupSearchResults[$lstGroupResults.SelectedIndex].DisplayName
            $existing = @()
            for ($i = 0; $i -lt $clbGroupes.Items.Count; $i++) {
                $existing += $clbGroupes.Items[$i].ToString()
            }
            if ($grpName -notin $existing) {
                $clbGroupes.Items.Add($grpName, $true) | Out-Null
                $checkedCount = 0
                for ($i = 0; $i -lt $clbGroupes.Items.Count; $i++) {
                    if ($clbGroupes.GetItemChecked($i)) { $checkedCount++ }
                }
                $lblGroupCount.Text = Get-Text "onboarding.group_selected_count" $checkedCount
            }
        }
    })

    # Mise à jour du compteur quand on coche/décoche
    $clbGroupes.Add_ItemCheck({
        $futureCount = 0
        for ($i = 0; $i -lt $clbGroupes.Items.Count; $i++) {
            if ($i -eq $_.Index) {
                if ($_.NewValue -eq [System.Windows.Forms.CheckState]::Checked) { $futureCount++ }
            }
            elseif ($clbGroupes.GetItemChecked($i)) { $futureCount++ }
        }
        $lblGroupCount.Text = Get-Text "onboarding.group_selected_count" $futureCount
    })

    # =================================================================
    # SECTION : Mot de passe (info)
    # =================================================================
    Add-SectionHeader -TextKey "onboarding.section_password" -YRef ([ref]$yPos)

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
    # Boutons Créer / Annuler — ancrés en bas à droite
    # =================================================================
    $btnCreer = New-Object System.Windows.Forms.Button
    $btnCreer.Text = Get-Text "onboarding.btn_create"
    $btnCreer.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnCreer.Size = New-Object System.Drawing.Size(170, 38)
    $btnCreer.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $btnCreer.ForeColor = [System.Drawing.Color]::White
    $btnCreer.FlatStyle = "Flat"
    $btnCreer.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCreer.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 340), ($form.ClientSize.Height - 48))
    $form.Controls.Add($btnCreer)

    $btnAnnuler = New-Object System.Windows.Forms.Button
    $btnAnnuler.Text = Get-Text "onboarding.btn_cancel"
    $btnAnnuler.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnAnnuler.Size = New-Object System.Drawing.Size(140, 38)
    $btnAnnuler.FlatStyle = "Flat"
    $btnAnnuler.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnAnnuler.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnAnnuler.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 160), ($form.ClientSize.Height - 48))
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

        $validations = @(
            @{ Value = $nom;      Key = "onboarding.validation_lastname" },
            @{ Value = $prenom;   Key = "onboarding.validation_firstname" },
            @{ Value = $jobTitle; Key = "onboarding.validation_job_title" },
            @{ Value = $dept;     Key = "onboarding.validation_department" },
            @{ Value = $empType;  Key = "onboarding.validation_employee_type" },
            @{ Value = $usageLoc; Key = "onboarding.validation_usage_location" }
        )
        foreach ($v in $validations) {
            if ([string]::IsNullOrWhiteSpace($v.Value)) {
                $lblErreur.Text = Get-Text $v.Key
                $lblErreur.Visible = $true; return
            }
        }

        # Validation UPN
        if ($cboUsername.SelectedIndex -lt 0 -or $script:UsernameVariants.Count -eq 0) {
            $lblErreur.Text = Get-Text "onboarding.validation_email"
            $lblErreur.Visible = $true; return
        }

        # Validation gestionnaire (obligatoire)
        if (-not $script:SelectedManagerId) {
            $lblErreur.Text = Get-Text "onboarding.validation_manager"
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
        $confirmMsg += "$(Get-Text 'onboarding.confirm_manager')    : $($script:SelectedManagerName)`n"

        # Licences sélectionnées (multi-sélection)
        $licGroupes = @()
        for ($i = 0; $i -lt $clbLicenses.Items.Count; $i++) {
            if ($clbLicenses.GetItemChecked($i)) {
                $licGroupes += $clbLicenses.Items[$i].ToString()
            }
        }
        if ($licGroupes.Count -gt 0) {
            $confirmMsg += "$(Get-Text 'onboarding.confirm_license'): $($licGroupes -join ', ')`n"
        }

        # Groupes d'appartenance sélectionnés
        $groupesSelectionnes = @()
        for ($i = 0; $i -lt $clbGroupes.Items.Count; $i++) {
            if ($clbGroupes.GetItemChecked($i)) {
                $groupesSelectionnes += $clbGroupes.Items[$i].ToString()
            }
        }
        if ($groupesSelectionnes.Count -gt 0) {
            $confirmMsg += "$(Get-Text 'onboarding.confirm_groups')         : $($groupesSelectionnes -join ', ')`n"
        }

        # Profils d'accès sélectionnés
        $script:OnbProfileKeys = @()
        if ($Config.PSObject.Properties["access_profiles"] -and $clbProfiles) {
            for ($i = 0; $i -lt $clbProfiles.Items.Count; $i++) {
                if ($clbProfiles.GetItemChecked($i)) {
                    $script:OnbProfileKeys += $onbProfiles[$i].Key
                }
            }
            if ($script:OnbProfileKeys.Count -gt 0) {
                $profileNames = ($script:OnbProfileKeys | ForEach-Object { $Config.access_profiles.$_.display_name }) -join ', '
                $confirmMsg += "$(Get-Text 'onboarding.confirm_profiles')  : $profileNames`n"
            }
            $baseline = Get-BaselineProfile
            if ($baseline) {
                $confirmMsg += "$(Get-Text 'onboarding.profile_baseline_label') $($baseline.DisplayName)`n"
            }
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

            # 4. Attribution du manager (toujours — champ obligatoire)
            $mgrResult = Set-AzUserManager -UserId $newUserId -ManagerId $script:SelectedManagerId
            if (-not $mgrResult.Success) { $erreurs += "Manager : $($mgrResult.Error)" }

            # 5. Ajout aux groupes de licence (multi-sélection)
            foreach ($licGrp in $licGroupes) {
                $licResult = Add-AzUserToGroup -UserId $newUserId -GroupName $licGrp
                if (-not $licResult.Success) { $erreurs += "Licence ($licGrp) : $($licResult.Error)" }
            }

            # 6. Ajout aux groupes d'appartenance
            foreach ($grp in $groupesSelectionnes) {
                $grpResult = Add-AzUserToGroup -UserId $newUserId -GroupName $grp
                if (-not $grpResult.Success) { $erreurs += "Groupe '$grp' : $($grpResult.Error)" }
            }

            # 7. Application des profils d'accès
            if ($script:OnbProfileKeys.Count -gt 0 -or ($Config.PSObject.Properties["access_profiles"] -and (Get-BaselineProfile))) {
                $profileResult = Invoke-AccessProfileChange -UserId $newUserId -UPN $upn `
                    -OldProfileKeys @() -NewProfileKeys $script:OnbProfileKeys
                if (-not $profileResult.Success) {
                    foreach ($err in $profileResult.Errors) { $erreurs += "Profil : $err" }
                }
            }

            # 8. Notification
            if ($Config.notifications.enabled) {
                $sujet = "Onboarding - $prenom $nom ($upn)"
                $corps = "<h2>Nouveau compte employé créé</h2>"
                $corps += "<table style='border-collapse:collapse'>"
                $corps += "<tr><td><b>Nom :</b></td><td>$prenom $nom</td></tr>"
                $corps += "<tr><td><b>UPN :</b></td><td>$upn</td></tr>"
                $corps += "<tr><td><b>Poste :</b></td><td>$jobTitle</td></tr>"
                $corps += "<tr><td><b>Département :</b></td><td>$dept</td></tr>"
                $corps += "<tr><td><b>Type :</b></td><td>$empType</td></tr>"
                $corps += "<tr><td><b>Gestionnaire :</b></td><td>$($script:SelectedManagerName)</td></tr>"
                if ($licGroupes.Count -gt 0) {
                    $corps += "<tr><td><b>Licences :</b></td><td>$($licGroupes -join ', ')</td></tr>"
                }
                $corps += "</table>"
                $corps += "<p><em>Mot de passe temporaire transmis séparément.</em></p>"
                Send-Notification -Sujet $sujet -Corps $corps
            }

            # 9. Résultat
            $lblChargement.Visible = $false

            if ($erreurs.Count -gt 0) {
                $warnMsg = (Get-Text "onboarding.result_partial_msg" ($erreurs -join "`n"))
                Show-ResultDialog -Titre (Get-Text "onboarding.result_partial_title") -Message $warnMsg -IsSuccess $false
            }

            Write-Log -Level "SUCCESS" -Action "ONBOARDING" -UPN $upn -Message "Onboarding terminé. Erreurs: $($erreurs.Count)"

            # 10. Affichage du mot de passe (une seule fois)
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