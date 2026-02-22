<#
.FICHIER
    Modules/GUI_Modification.ps1

.ROLE
    Formulaire de modification des attributs d'un employe existant.
    Sous-menus : departement, manager, contrat, telephone/titre, reset MDP, activer/desactiver.

.DEPENDANCES
    - Core/Functions.ps1 (Write-Log, Show-ConfirmDialog, Show-ResultDialog, New-SecurePassword, Show-PasswordDialog)
    - Core/GraphAPI.ps1 (Search-AzUsers, Get-AzUser, Set-AzUser, Set-AzUserManager,
      Set-AzUserLicense, Reset-AzUserPassword, Disable-AzUser, Enable-AzUser, Get-AzUserManager)
    - Variable globale $Config

.AUTEUR
    [Equipe IT - GestionRH-AzureAD]
#>

function Show-ModificationForm {
    <#
    .SYNOPSIS
        Affiche le formulaire de modification avec sous-menu de choix.
    .OUTPUTS
        [void] - Formulaire modal.
    #>

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Modification - $($Config.client_name)"
    $form.Size = New-Object System.Drawing.Size(700, 620)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::WhiteSmoke

    # === Section Recherche ===
    $lblSection = New-Object System.Windows.Forms.Label
    $lblSection.Text = "Rechercher l'employe a modifier"
    $lblSection.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblSection.Location = New-Object System.Drawing.Point(15, 15)
    $lblSection.Size = New-Object System.Drawing.Size(650, 25)
    $form.Controls.Add($lblSection)

    $txtRecherche = New-Object System.Windows.Forms.TextBox
    $txtRecherche.Location = New-Object System.Drawing.Point(15, 50)
    $txtRecherche.Size = New-Object System.Drawing.Size(480, 25)
    $txtRecherche.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($txtRecherche)

    $btnRecherche = New-Object System.Windows.Forms.Button
    $btnRecherche.Text = "Chercher"
    $btnRecherche.Location = New-Object System.Drawing.Point(505, 50)
    $btnRecherche.Size = New-Object System.Drawing.Size(90, 25)
    $form.Controls.Add($btnRecherche)

    $lstResultats = New-Object System.Windows.Forms.ListBox
    $lstResultats.Location = New-Object System.Drawing.Point(15, 80)
    $lstResultats.Size = New-Object System.Drawing.Size(580, 70)
    $lstResultats.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lstResultats.Visible = $false
    $form.Controls.Add($lstResultats)

    # Info utilisateur selectionne + valeurs actuelles
    $lblUserInfo = New-Object System.Windows.Forms.Label
    $lblUserInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblUserInfo.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $lblUserInfo.Location = New-Object System.Drawing.Point(15, 155)
    $lblUserInfo.Size = New-Object System.Drawing.Size(650, 45)
    $lblUserInfo.Visible = $false
    $form.Controls.Add($lblUserInfo)

    $script:SelectedUserId = $null
    $script:SelectedUserUPN = $null
    $script:SelectedUserData = $null
    $script:SearchResults = @()

    $btnRecherche.Add_Click({
        $terme = $txtRecherche.Text.Trim()
        if ($terme.Length -lt 2) { return }
        $lstResultats.Items.Clear()
        $result = Search-AzUsers -SearchTerm $terme -MaxResults 15
        if ($result.Success -and $result.Data) {
            $script:SearchResults = @($result.Data)
            foreach ($user in $script:SearchResults) {
                $statut = if ($user.AccountEnabled) { "Actif" } else { "Desactive" }
                $lstResultats.Items.Add("$($user.DisplayName) - $($user.UserPrincipalName) [$statut]") | Out-Null
            }
            $lstResultats.Visible = $true
        }
    })

    $txtRecherche.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $btnRecherche.PerformClick()
            $_.SuppressKeyPress = $true
        }
    })

    # Panel d'actions (masque tant qu'aucun utilisateur n'est selectionne)
    $panelActions = New-Object System.Windows.Forms.Panel
    $panelActions.Location = New-Object System.Drawing.Point(15, 210)
    $panelActions.Size = New-Object System.Drawing.Size(655, 370)
    $panelActions.Visible = $false
    $form.Controls.Add($panelActions)

    $lstResultats.Add_SelectedIndexChanged({
        if ($lstResultats.SelectedIndex -ge 0 -and $lstResultats.SelectedIndex -lt $script:SearchResults.Count) {
            $selected = $script:SearchResults[$lstResultats.SelectedIndex]
            $script:SelectedUserId = $selected.Id
            $script:SelectedUserUPN = $selected.UserPrincipalName

            # Charger les details complets
            $detailResult = Get-AzUser -UserId $selected.Id
            if ($detailResult.Success) {
                $script:SelectedUserData = $detailResult.Data
                $lblUserInfo.Text = "Selectionne : $($selected.DisplayName) | Dept: $($detailResult.Data.Department) | Poste: $($detailResult.Data.JobTitle) | Statut: $(if ($detailResult.Data.AccountEnabled) { 'Actif' } else { 'Desactive' })"
            }
            else {
                $lblUserInfo.Text = "Selectionne : $($selected.DisplayName) ($($selected.UserPrincipalName))"
            }

            $lblUserInfo.Visible = $true
            $lstResultats.Visible = $false
            $panelActions.Visible = $true
        }
    })

    # === Sous-menu des modifications ===
    $lblChoix = New-Object System.Windows.Forms.Label
    $lblChoix.Text = "Choisir la modification :"
    $lblChoix.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblChoix.Location = New-Object System.Drawing.Point(0, 0)
    $lblChoix.Size = New-Object System.Drawing.Size(300, 25)
    $panelActions.Controls.Add($lblChoix)

    # Boutons de sous-menu
    $actions = @(
        @{ Text = "Departement";          Tag = "dept" },
        @{ Text = "Manager";              Tag = "manager" },
        @{ Text = "Type de contrat";      Tag = "contrat" },
        @{ Text = "Telephone / Titre";    Tag = "infos" },
        @{ Text = "Reset mot de passe";   Tag = "password" },
        @{ Text = "Activer / Desactiver"; Tag = "toggle" }
    )

    $btnY = 30
    foreach ($action in $actions) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $action.Text
        $btn.Tag = $action.Tag
        $btn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $btn.Location = New-Object System.Drawing.Point(0, $btnY)
        $btn.Size = New-Object System.Drawing.Size(180, 30)
        $btn.FlatStyle = "Flat"
        $btn.BackColor = [System.Drawing.Color]::White
        $btn.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $btn.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)

        $btn.Add_Click({
            $tag = $this.Tag
            $userId = $script:SelectedUserId
            $upn = $script:SelectedUserUPN
            $userData = $script:SelectedUserData

            switch ($tag) {
                "dept" {
                    Show-ModifyDepartment -UserId $userId -UPN $upn -CurrentValue $userData.Department
                }
                "manager" {
                    Show-ModifyManager -UserId $userId -UPN $upn
                }
                "contrat" {
                    Show-ModifyContract -UserId $userId -UPN $upn -CurrentValue $userData.EmployeeType
                }
                "infos" {
                    Show-ModifyInfos -UserId $userId -UPN $upn -CurrentPhone $userData.MobilePhone -CurrentTitle $userData.JobTitle
                }
                "password" {
                    Show-ResetPassword -UserId $userId -UPN $upn
                }
                "toggle" {
                    Show-ToggleAccount -UserId $userId -UPN $upn -IsEnabled $userData.AccountEnabled
                }
            }
        })

        $panelActions.Controls.Add($btn)
        $btnY += 35
    }

    # Bouton Annuler
    $btnAnnuler = New-Object System.Windows.Forms.Button
    $btnAnnuler.Text = "Fermer"
    $btnAnnuler.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnAnnuler.Location = New-Object System.Drawing.Point(530, 530)
    $btnAnnuler.Size = New-Object System.Drawing.Size(140, 40)
    $btnAnnuler.FlatStyle = "Flat"
    $btnAnnuler.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnAnnuler)
    $form.CancelButton = $btnAnnuler

    $form.ShowDialog() | Out-Null
    $form.Dispose()
}

# === Sous-formulaires de modification ===

function Show-ModifyDepartment {
    param([string]$UserId, [string]$UPN, [string]$CurrentValue)

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Changement de departement - $UPN"
    $subForm.Size = New-Object System.Drawing.Size(400, 200)
    $subForm.StartPosition = "CenterScreen"
    $subForm.FormBorderStyle = "FixedDialog"
    $subForm.MaximizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Departement actuel : $CurrentValue"
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lbl.Location = New-Object System.Drawing.Point(15, 15)
    $lbl.Size = New-Object System.Drawing.Size(350, 20)
    $subForm.Controls.Add($lbl)

    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = "Nouveau departement :"
    $lblNew.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblNew.Location = New-Object System.Drawing.Point(15, 50)
    $lblNew.Size = New-Object System.Drawing.Size(150, 20)
    $subForm.Controls.Add($lblNew)

    $cbo = New-Object System.Windows.Forms.ComboBox
    $cbo.Location = New-Object System.Drawing.Point(170, 47)
    $cbo.Size = New-Object System.Drawing.Size(200, 25)
    $cbo.DropDownStyle = "DropDownList"
    foreach ($dept in $Config.departments) { $cbo.Items.Add($dept) | Out-Null }
    $subForm.Controls.Add($cbo)

    $btnAppliquer = New-Object System.Windows.Forms.Button
    $btnAppliquer.Text = "Appliquer"
    $btnAppliquer.Location = New-Object System.Drawing.Point(100, 100)
    $btnAppliquer.Size = New-Object System.Drawing.Size(100, 35)
    $btnAppliquer.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $btnAppliquer.ForeColor = [System.Drawing.Color]::White
    $btnAppliquer.FlatStyle = "Flat"
    $btnAppliquer.Add_Click({
        if ($cbo.SelectedItem) {
            $confirm = Show-ConfirmDialog -Titre "Confirmation" -Message "Changer le departement de $UPN de '$CurrentValue' a '$($cbo.SelectedItem)' ?"
            if ($confirm) {
                $result = Set-AzUser -UserId $UserId -Properties @{ Department = $cbo.SelectedItem.ToString() }
                if ($result.Success) {
                    Show-ResultDialog -Titre "Succes" -Message "Departement mis a jour." -IsSuccess $true
                    Write-Log -Level "SUCCESS" -Action "MODIFY_DEPT" -UPN $UPN -Message "Departement change : $CurrentValue -> $($cbo.SelectedItem)"
                    $subForm.Close()
                } else {
                    Show-ResultDialog -Titre "Erreur" -Message $result.Error -IsSuccess $false
                }
            }
        }
    })
    $subForm.Controls.Add($btnAppliquer)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(210, 100)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 35)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $subForm.Controls.Add($btnCancel)
    $subForm.CancelButton = $btnCancel

    $subForm.ShowDialog() | Out-Null
    $subForm.Dispose()
}

function Show-ModifyManager {
    param([string]$UserId, [string]$UPN)

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Changement de manager - $UPN"
    $subForm.Size = New-Object System.Drawing.Size(450, 260)
    $subForm.StartPosition = "CenterScreen"
    $subForm.FormBorderStyle = "FixedDialog"
    $subForm.MaximizeBox = $false

    # Afficher le manager actuel
    $mgrResult = Get-AzUserManager -UserId $UserId
    $currentMgr = if ($mgrResult.Success -and $mgrResult.Data) { $mgrResult.Data.AdditionalProperties.displayName } else { "(aucun)" }

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Manager actuel : $currentMgr"
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lbl.Location = New-Object System.Drawing.Point(15, 15)
    $lbl.Size = New-Object System.Drawing.Size(400, 20)
    $subForm.Controls.Add($lbl)

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text = "Nouveau manager :"
    $lblSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSearch.Location = New-Object System.Drawing.Point(15, 50)
    $lblSearch.Size = New-Object System.Drawing.Size(120, 20)
    $subForm.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(140, 47)
    $txtSearch.Size = New-Object System.Drawing.Size(210, 25)
    $subForm.Controls.Add($txtSearch)

    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Text = "Chercher"
    $btnSearch.Location = New-Object System.Drawing.Point(355, 47)
    $btnSearch.Size = New-Object System.Drawing.Size(70, 25)
    $subForm.Controls.Add($btnSearch)

    $lstMgr = New-Object System.Windows.Forms.ListBox
    $lstMgr.Location = New-Object System.Drawing.Point(140, 78)
    $lstMgr.Size = New-Object System.Drawing.Size(285, 70)
    $lstMgr.Visible = $false
    $subForm.Controls.Add($lstMgr)

    $script:MgrSearchResults = @()
    $script:NewManagerId = $null

    $btnSearch.Add_Click({
        $terme = $txtSearch.Text.Trim()
        if ($terme.Length -lt 2) { return }
        $lstMgr.Items.Clear()
        $res = Search-AzUsers -SearchTerm $terme -MaxResults 10
        if ($res.Success -and $res.Data) {
            $script:MgrSearchResults = @($res.Data)
            foreach ($u in $script:MgrSearchResults) {
                $lstMgr.Items.Add("$($u.DisplayName) ($($u.UserPrincipalName))") | Out-Null
            }
            $lstMgr.Visible = $true
        }
    })

    $lstMgr.Add_SelectedIndexChanged({
        if ($lstMgr.SelectedIndex -ge 0) {
            $script:NewManagerId = $script:MgrSearchResults[$lstMgr.SelectedIndex].Id
            $txtSearch.Text = $script:MgrSearchResults[$lstMgr.SelectedIndex].DisplayName
            $lstMgr.Visible = $false
        }
    })

    $btnAppliquer = New-Object System.Windows.Forms.Button
    $btnAppliquer.Text = "Appliquer"
    $btnAppliquer.Location = New-Object System.Drawing.Point(120, 170)
    $btnAppliquer.Size = New-Object System.Drawing.Size(100, 35)
    $btnAppliquer.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $btnAppliquer.ForeColor = [System.Drawing.Color]::White
    $btnAppliquer.FlatStyle = "Flat"
    $btnAppliquer.Add_Click({
        if ($script:NewManagerId) {
            $confirm = Show-ConfirmDialog -Titre "Confirmation" -Message "Definir '$($txtSearch.Text)' comme manager de $UPN ?"
            if ($confirm) {
                $result = Set-AzUserManager -UserId $UserId -ManagerId $script:NewManagerId
                if ($result.Success) {
                    Show-ResultDialog -Titre "Succes" -Message "Manager mis a jour." -IsSuccess $true
                    Write-Log -Level "SUCCESS" -Action "MODIFY_MANAGER" -UPN $UPN -Message "Manager change -> $($txtSearch.Text)"
                    $subForm.Close()
                } else {
                    Show-ResultDialog -Titre "Erreur" -Message $result.Error -IsSuccess $false
                }
            }
        }
    })
    $subForm.Controls.Add($btnAppliquer)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(230, 170)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 35)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $subForm.Controls.Add($btnCancel)
    $subForm.CancelButton = $btnCancel

    $subForm.ShowDialog() | Out-Null
    $subForm.Dispose()
}

function Show-ModifyContract {
    param([string]$UserId, [string]$UPN, [string]$CurrentValue)

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Changement de contrat - $UPN"
    $subForm.Size = New-Object System.Drawing.Size(400, 220)
    $subForm.StartPosition = "CenterScreen"
    $subForm.FormBorderStyle = "FixedDialog"
    $subForm.MaximizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Type actuel : $CurrentValue"
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lbl.Location = New-Object System.Drawing.Point(15, 15)
    $lbl.Size = New-Object System.Drawing.Size(350, 20)
    $subForm.Controls.Add($lbl)

    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = "Nouveau type :"
    $lblNew.Location = New-Object System.Drawing.Point(15, 50)
    $lblNew.Size = New-Object System.Drawing.Size(120, 20)
    $subForm.Controls.Add($lblNew)

    $cbo = New-Object System.Windows.Forms.ComboBox
    $cbo.Location = New-Object System.Drawing.Point(140, 47)
    $cbo.Size = New-Object System.Drawing.Size(230, 25)
    $cbo.DropDownStyle = "DropDownList"
    foreach ($ct in $Config.contract_types) { $cbo.Items.Add($ct) | Out-Null }
    $subForm.Controls.Add($cbo)

    $chkLicense = New-Object System.Windows.Forms.CheckBox
    $chkLicense.Text = "Mettre a jour la licence associee"
    $chkLicense.Location = New-Object System.Drawing.Point(15, 85)
    $chkLicense.Size = New-Object System.Drawing.Size(350, 22)
    $chkLicense.Checked = $true
    $subForm.Controls.Add($chkLicense)

    $btnAppliquer = New-Object System.Windows.Forms.Button
    $btnAppliquer.Text = "Appliquer"
    $btnAppliquer.Location = New-Object System.Drawing.Point(100, 125)
    $btnAppliquer.Size = New-Object System.Drawing.Size(100, 35)
    $btnAppliquer.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $btnAppliquer.ForeColor = [System.Drawing.Color]::White
    $btnAppliquer.FlatStyle = "Flat"
    $btnAppliquer.Add_Click({
        if ($cbo.SelectedItem) {
            $newType = $cbo.SelectedItem.ToString()
            $confirm = Show-ConfirmDialog -Titre "Confirmation" -Message "Changer le type de contrat de $UPN a '$newType' ?"
            if ($confirm) {
                $result = Set-AzUser -UserId $UserId -Properties @{ EmployeeType = $newType }
                if ($result.Success) {
                    # Mise a jour licence si coche
                    if ($chkLicense.Checked) {
                        $contractKey = $newType.ToLower()
                        $licenseMap = @{}
                        $Config.license_map.PSObject.Properties | ForEach-Object { $licenseMap[$_.Name] = $_.Value }
                        if ($licenseMap.ContainsKey($contractKey)) {
                            Set-AzUserLicense -UserId $UserId -SkuId $licenseMap[$contractKey]
                        }
                    }
                    Show-ResultDialog -Titre "Succes" -Message "Type de contrat mis a jour." -IsSuccess $true
                    Write-Log -Level "SUCCESS" -Action "MODIFY_CONTRACT" -UPN $UPN -Message "Contrat change : $CurrentValue -> $newType"
                    $subForm.Close()
                } else {
                    Show-ResultDialog -Titre "Erreur" -Message $result.Error -IsSuccess $false
                }
            }
        }
    })
    $subForm.Controls.Add($btnAppliquer)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(210, 125)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 35)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $subForm.Controls.Add($btnCancel)
    $subForm.CancelButton = $btnCancel

    $subForm.ShowDialog() | Out-Null
    $subForm.Dispose()
}

function Show-ModifyInfos {
    param([string]$UserId, [string]$UPN, [string]$CurrentPhone, [string]$CurrentTitle)

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Mise a jour infos - $UPN"
    $subForm.Size = New-Object System.Drawing.Size(420, 220)
    $subForm.StartPosition = "CenterScreen"
    $subForm.FormBorderStyle = "FixedDialog"
    $subForm.MaximizeBox = $false

    $lblPhone = New-Object System.Windows.Forms.Label
    $lblPhone.Text = "Telephone :"
    $lblPhone.Location = New-Object System.Drawing.Point(15, 20)
    $lblPhone.Size = New-Object System.Drawing.Size(100, 20)
    $subForm.Controls.Add($lblPhone)

    $txtPhone = New-Object System.Windows.Forms.TextBox
    $txtPhone.Text = $CurrentPhone
    $txtPhone.Location = New-Object System.Drawing.Point(120, 17)
    $txtPhone.Size = New-Object System.Drawing.Size(260, 25)
    $subForm.Controls.Add($txtPhone)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "Titre de poste :"
    $lblTitle.Location = New-Object System.Drawing.Point(15, 55)
    $lblTitle.Size = New-Object System.Drawing.Size(100, 20)
    $subForm.Controls.Add($lblTitle)

    $txtTitle = New-Object System.Windows.Forms.TextBox
    $txtTitle.Text = $CurrentTitle
    $txtTitle.Location = New-Object System.Drawing.Point(120, 52)
    $txtTitle.Size = New-Object System.Drawing.Size(260, 25)
    $subForm.Controls.Add($txtTitle)

    $btnAppliquer = New-Object System.Windows.Forms.Button
    $btnAppliquer.Text = "Appliquer"
    $btnAppliquer.Location = New-Object System.Drawing.Point(100, 110)
    $btnAppliquer.Size = New-Object System.Drawing.Size(100, 35)
    $btnAppliquer.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $btnAppliquer.ForeColor = [System.Drawing.Color]::White
    $btnAppliquer.FlatStyle = "Flat"
    $btnAppliquer.Add_Click({
        $props = @{}
        if ($txtPhone.Text.Trim() -ne $CurrentPhone) { $props.MobilePhone = $txtPhone.Text.Trim() }
        if ($txtTitle.Text.Trim() -ne $CurrentTitle) { $props.JobTitle = $txtTitle.Text.Trim() }
        if ($props.Count -eq 0) {
            Show-ResultDialog -Titre "Info" -Message "Aucune modification detectee." -IsSuccess $true
            return
        }
        $confirm = Show-ConfirmDialog -Titre "Confirmation" -Message "Appliquer les modifications a $UPN ?"
        if ($confirm) {
            $result = Set-AzUser -UserId $UserId -Properties $props
            if ($result.Success) {
                Show-ResultDialog -Titre "Succes" -Message "Informations mises a jour." -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "MODIFY_INFOS" -UPN $UPN -Message "Infos mises a jour : $($props.Keys -join ', ')"
                $subForm.Close()
            } else {
                Show-ResultDialog -Titre "Erreur" -Message $result.Error -IsSuccess $false
            }
        }
    })
    $subForm.Controls.Add($btnAppliquer)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(210, 110)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 35)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $subForm.Controls.Add($btnCancel)
    $subForm.CancelButton = $btnCancel

    $subForm.ShowDialog() | Out-Null
    $subForm.Dispose()
}

function Show-ResetPassword {
    param([string]$UserId, [string]$UPN)

    $confirm = Show-ConfirmDialog -Titre "Reset mot de passe" -Message "Reinitialiser le mot de passe de $UPN ?`n`nUn nouveau mot de passe temporaire sera genere."
    if (-not $confirm) { return }

    $newPassword = New-SecurePassword
    $result = Reset-AzUserPassword -UserId $UserId -NewPassword $newPassword -ForceChange $Config.password_policy.force_change_at_login

    if ($result.Success) {
        Write-Log -Level "SUCCESS" -Action "RESET_PASSWORD" -UPN $UPN -Message "Mot de passe reinitialise."
        Show-PasswordDialog -UPN $UPN -Password $newPassword
    }
    else {
        Show-ResultDialog -Titre "Erreur" -Message "Echec du reset : $($result.Error)" -IsSuccess $false
    }
}

function Show-ToggleAccount {
    param([string]$UserId, [string]$UPN, [bool]$IsEnabled)

    $action = if ($IsEnabled) { "DESACTIVER" } else { "REACTIVER" }
    $confirm = Show-ConfirmDialog -Titre "$action le compte" -Message "Voulez-vous $($action.ToLower()) le compte de $UPN ?"

    if (-not $confirm) { return }

    if ($IsEnabled) {
        $result = Disable-AzUser -UserId $UserId
    }
    else {
        $result = Enable-AzUser -UserId $UserId
    }

    if ($result.Success) {
        Show-ResultDialog -Titre "Succes" -Message "Compte $($action.ToLower()) avec succes." -IsSuccess $true
        Write-Log -Level "SUCCESS" -Action "TOGGLE_ACCOUNT" -UPN $UPN -Message "Compte $action"
    }
    else {
        Show-ResultDialog -Titre "Erreur" -Message $result.Error -IsSuccess $false
    }
}

# Point d'attention :
# - Chaque sous-formulaire est independant et modal
# - La confirmation est requise avant chaque modification
# - Le changement de contrat peut impacter la licence (case a cocher)
# - Le reset de mot de passe utilise Show-PasswordDialog (affichage unique)
