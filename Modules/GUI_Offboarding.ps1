<#
.FICHIER
    Modules/GUI_Offboarding.ps1

.ROLE
    Formulaire de depart (offboarding) d'un employe.
    Gere la desactivation du compte, la revocation des licences et sessions,
    le retrait des groupes (avec skip des groupes dynamiques),
    la conversion en boite partagee, le masquage du GAL, et la notification.

.DEPENDANCES
    - Core/Functions.ps1 (Write-Log, Show-ConfirmDialog, Show-ResultDialog, Send-Notification)
    - Core/GraphAPI.ps1 (Search-AzUsers, Disable-AzUser, Revoke-AzUserSessions,
      Remove-AzUserGroups, Remove-AzUserLicenses, Add-AzUserToGroup,
      Get-AzMailboxSize, Convert-AzMailboxToShared, Hide-AzMailboxFromGAL,
      Search-AzGroups)
    - Module ExchangeOnlineManagement (Get-EXOMailboxStatistics, Set-Mailbox)
    - Variable globale $Config

.AUTEUR
    [Equipe IT - GestionRH-AzureAD]
#>

function Show-OffboardingForm {
    <#
    .SYNOPSIS
        Affiche le formulaire d'offboarding pour gerer le depart d'un employe.
    .OUTPUTS
        [void] - Formulaire modal.
    #>

    # === Variables de scope ===
    $script:SelectedUserId = $null
    $script:SelectedUserUPN = $null
    $script:SelectedUserName = $null
    $script:SearchResults = @()
    $script:MailboxSizeGB = 0
    $script:ExchangeAvailable = $false
    $script:LicenseGroupNames = @()

    # Vérifier si Exchange Online est disponible
    try {
        $exoSession = Get-Command "Get-EXOMailboxStatistics" -ErrorAction SilentlyContinue
        $script:ExchangeAvailable = ($null -ne $exoSession)
    }
    catch { $script:ExchangeAvailable = $false }

    # === Formulaire principal ===
    $form = New-Object System.Windows.Forms.Form
    $form.Text = Get-Text "offboarding.title" $Config.client_name
    $form.Size = New-Object System.Drawing.Size(660, 760)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::WhiteSmoke

    $yPos = 15

    # =================================================================
    # SECTION : Recherche de l'employe
    # =================================================================
    $lblSection = New-Object System.Windows.Forms.Label
    $lblSection.Text = Get-Text "offboarding.section_search"
    $lblSection.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblSection.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $lblSection.Location = New-Object System.Drawing.Point(15, $yPos)
    $lblSection.Size = New-Object System.Drawing.Size(600, 25)
    $form.Controls.Add($lblSection)
    $yPos += 35

    # Barre de recherche
    $lblRecherche = New-Object System.Windows.Forms.Label
    $lblRecherche.Text = Get-Text "offboarding.search_label"
    $lblRecherche.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblRecherche.Location = New-Object System.Drawing.Point(15, ($yPos + 3))
    $lblRecherche.Size = New-Object System.Drawing.Size(90, 20)
    $form.Controls.Add($lblRecherche)

    $txtRecherche = New-Object System.Windows.Forms.TextBox
    $txtRecherche.Location = New-Object System.Drawing.Point(110, $yPos)
    $txtRecherche.Size = New-Object System.Drawing.Size(420, 25)
    $txtRecherche.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($txtRecherche)

    $btnRecherche = New-Object System.Windows.Forms.Button
    $btnRecherche.Text = Get-Text "offboarding.btn_search"
    $btnRecherche.Location = New-Object System.Drawing.Point(540, $yPos)
    $btnRecherche.Size = New-Object System.Drawing.Size(90, 25)
    $form.Controls.Add($btnRecherche)
    $yPos += 33

    # Liste de resultats
    $lstResultats = New-Object System.Windows.Forms.ListBox
    $lstResultats.Location = New-Object System.Drawing.Point(110, $yPos)
    $lstResultats.Size = New-Object System.Drawing.Size(420, 80)
    $lstResultats.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lstResultats.Visible = $false
    $form.Controls.Add($lstResultats)

    # Info utilisateur selectionne
    $lblUserInfo = New-Object System.Windows.Forms.Label
    $lblUserInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblUserInfo.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $lblUserInfo.Location = New-Object System.Drawing.Point(15, ($yPos + 85))
    $lblUserInfo.Size = New-Object System.Drawing.Size(600, 20)
    $lblUserInfo.Visible = $false
    $form.Controls.Add($lblUserInfo)

    # Info taille BAL (affichée après sélection)
    $lblMailboxInfo = New-Object System.Windows.Forms.Label
    $lblMailboxInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblMailboxInfo.Location = New-Object System.Drawing.Point(15, ($yPos + 107))
    $lblMailboxInfo.Size = New-Object System.Drawing.Size(600, 20)
    $lblMailboxInfo.Visible = $false
    $form.Controls.Add($lblMailboxInfo)

    # =================================================================
    # SECTION : Détails du départ
    # =================================================================
    $ySection2 = 265

    $lblSection2 = New-Object System.Windows.Forms.Label
    $lblSection2.Text = Get-Text "offboarding.section_details"
    $lblSection2.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblSection2.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $lblSection2.Location = New-Object System.Drawing.Point(15, $ySection2)
    $lblSection2.Size = New-Object System.Drawing.Size(600, 25)
    $form.Controls.Add($lblSection2)
    $ySection2 += 35

    # Date de depart
    $lblDate = New-Object System.Windows.Forms.Label
    $lblDate.Text = Get-Text "offboarding.field_date"
    $lblDate.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblDate.Location = New-Object System.Drawing.Point(15, ($ySection2 + 3))
    $lblDate.Size = New-Object System.Drawing.Size(120, 20)
    $form.Controls.Add($lblDate)

    $dtpDepart = New-Object System.Windows.Forms.DateTimePicker
    $dtpDepart.Location = New-Object System.Drawing.Point(150, $ySection2)
    $dtpDepart.Size = New-Object System.Drawing.Size(200, 25)
    $dtpDepart.Format = "Short"
    $form.Controls.Add($dtpDepart)
    $ySection2 += 33

    # Raison du depart
    $lblRaison = New-Object System.Windows.Forms.Label
    $lblRaison.Text = Get-Text "offboarding.field_reason"
    $lblRaison.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblRaison.Location = New-Object System.Drawing.Point(15, ($ySection2 + 3))
    $lblRaison.Size = New-Object System.Drawing.Size(120, 20)
    $form.Controls.Add($lblRaison)

    $cboRaison = New-Object System.Windows.Forms.ComboBox
    $cboRaison.Location = New-Object System.Drawing.Point(150, $ySection2)
    $cboRaison.Size = New-Object System.Drawing.Size(200, 25)
    $cboRaison.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cboRaison.DropDownStyle = "DropDownList"
    @(
        (Get-Text "offboarding.reason_resignation"),
        (Get-Text "offboarding.reason_termination"),
        (Get-Text "offboarding.reason_contract_end"),
        (Get-Text "offboarding.reason_other")
    ) | ForEach-Object { $cboRaison.Items.Add($_) | Out-Null }
    $cboRaison.SelectedIndex = 0
    $form.Controls.Add($cboRaison)
    $ySection2 += 33

    # Déléguer accès à la boîte (Read & Manage / FullAccess)
    $lblDelegate = New-Object System.Windows.Forms.Label
    $lblDelegate.Text = Get-Text "offboarding.field_delegate_access"
    $lblDelegate.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblDelegate.Location = New-Object System.Drawing.Point(15, ($ySection2 + 3))
    $lblDelegate.Size = New-Object System.Drawing.Size(130, 20)
    $form.Controls.Add($lblDelegate)

    $txtDelegate = New-Object System.Windows.Forms.TextBox
    $txtDelegate.Location = New-Object System.Drawing.Point(150, $ySection2)
    $txtDelegate.Size = New-Object System.Drawing.Size(300, 25)
    $txtDelegate.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($txtDelegate)

    $lblDelegateHint = New-Object System.Windows.Forms.Label
    $lblDelegateHint.Text = Get-Text "offboarding.delegate_hint"
    $lblDelegateHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblDelegateHint.ForeColor = [System.Drawing.Color]::Gray
    $lblDelegateHint.Location = New-Object System.Drawing.Point(150, ($ySection2 + 27))
    $lblDelegateHint.Size = New-Object System.Drawing.Size(400, 16)
    $form.Controls.Add($lblDelegateHint)
    $ySection2 += 48

    # =================================================================
    # SECTION : Actions d'offboarding (checkboxes)
    # =================================================================
    $chkDesactiver = New-Object System.Windows.Forms.CheckBox
    $chkDesactiver.Text = Get-Text "offboarding.chk_disable"
    $chkDesactiver.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkDesactiver.Location = New-Object System.Drawing.Point(30, $ySection2)
    $chkDesactiver.Size = New-Object System.Drawing.Size(270, 22)
    $chkDesactiver.Checked = $true
    $form.Controls.Add($chkDesactiver)

    $chkRevoquerSessions = New-Object System.Windows.Forms.CheckBox
    $chkRevoquerSessions.Text = Get-Text "offboarding.chk_revoke_sessions"
    $chkRevoquerSessions.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkRevoquerSessions.Location = New-Object System.Drawing.Point(320, $ySection2)
    $chkRevoquerSessions.Size = New-Object System.Drawing.Size(300, 22)
    $chkRevoquerSessions.Checked = $true
    $form.Controls.Add($chkRevoquerSessions)
    $ySection2 += 28

    $chkRetirerGroupes = New-Object System.Windows.Forms.CheckBox
    $chkRetirerGroupes.Text = Get-Text "offboarding.chk_remove_groups"
    $chkRetirerGroupes.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkRetirerGroupes.Location = New-Object System.Drawing.Point(30, $ySection2)
    $chkRetirerGroupes.Size = New-Object System.Drawing.Size(270, 22)
    $chkRetirerGroupes.Checked = $Config.offboarding.remove_all_groups
    $form.Controls.Add($chkRetirerGroupes)

    $chkRevoquerLicences = New-Object System.Windows.Forms.CheckBox
    $chkRevoquerLicences.Text = Get-Text "offboarding.chk_revoke_licenses"
    $chkRevoquerLicences.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkRevoquerLicences.Location = New-Object System.Drawing.Point(320, $ySection2)
    $chkRevoquerLicences.Size = New-Object System.Drawing.Size(300, 22)
    $chkRevoquerLicences.Checked = $Config.offboarding.revoke_licenses
    $form.Controls.Add($chkRevoquerLicences)
    $ySection2 += 28

    # Masquer du carnet d'adresses (GAL) — coché par défaut
    $chkHideGAL = New-Object System.Windows.Forms.CheckBox
    $chkHideGAL.Text = Get-Text "offboarding.chk_hide_gal"
    $chkHideGAL.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkHideGAL.Location = New-Object System.Drawing.Point(30, $ySection2)
    $chkHideGAL.Size = New-Object System.Drawing.Size(270, 22)
    $chkHideGAL.Checked = $true
    $form.Controls.Add($chkHideGAL)

    # Convertir en boîte partagée — activé seulement après check Exchange, décoché par défaut
    $chkConvertShared = New-Object System.Windows.Forms.CheckBox
    $chkConvertShared.Text = Get-Text "offboarding.chk_convert_shared"
    $chkConvertShared.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkConvertShared.Location = New-Object System.Drawing.Point(320, $ySection2)
    $chkConvertShared.Size = New-Object System.Drawing.Size(300, 22)
    $chkConvertShared.Checked = $false
    $chkConvertShared.Enabled = $false
    $form.Controls.Add($chkConvertShared)
    $ySection2 += 35

    # =================================================================
    # SECTION : Sélecteur de licence Exchange (visible si BAL > 50 Go)
    # =================================================================
    $lblLicWarning = New-Object System.Windows.Forms.Label
    $lblLicWarning.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblLicWarning.ForeColor = [System.Drawing.Color]::FromArgb(220, 130, 0)
    $lblLicWarning.Location = New-Object System.Drawing.Point(30, $ySection2)
    $lblLicWarning.Size = New-Object System.Drawing.Size(600, 20)
    $lblLicWarning.Visible = $false
    $form.Controls.Add($lblLicWarning)

    $clbLicenses = New-Object System.Windows.Forms.CheckedListBox
    $clbLicenses.Location = New-Object System.Drawing.Point(30, ($ySection2 + 24))
    $clbLicenses.Size = New-Object System.Drawing.Size(500, 60)
    $clbLicenses.CheckOnClick = $true
    $clbLicenses.Visible = $false
    $form.Controls.Add($clbLicenses)

    $btnRefreshLic = New-Object System.Windows.Forms.Button
    $btnRefreshLic.Text = "⟳"
    $btnRefreshLic.Location = New-Object System.Drawing.Point(538, ($ySection2 + 24))
    $btnRefreshLic.Size = New-Object System.Drawing.Size(35, 25)
    $btnRefreshLic.FlatStyle = "Flat"
    $btnRefreshLic.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnRefreshLic.Visible = $false
    $form.Controls.Add($btnRefreshLic)

    $lblLicInfo = New-Object System.Windows.Forms.Label
    $lblLicInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblLicInfo.ForeColor = [System.Drawing.Color]::Gray
    $lblLicInfo.Location = New-Object System.Drawing.Point(30, ($ySection2 + 86))
    $lblLicInfo.Size = New-Object System.Drawing.Size(540, 16)
    $lblLicInfo.Visible = $false
    $form.Controls.Add($lblLicInfo)

    # Fonction : charger les groupes de licence par préfixe (même pattern que onboarding)
    function Update-OffboardLicenseList {
        $clbLicenses.Items.Clear()
        $script:LicenseGroupNames = @()
        $prefix = if ($Config.PSObject.Properties["license_group_prefix"]) { $Config.license_group_prefix } else { "" }
        if ([string]::IsNullOrWhiteSpace($prefix)) { return }

        $res = Search-AzGroups -SearchTerm $prefix -MaxResults 50
        if ($res.Success -and $res.Data) {
            $filtered = $res.Data | Where-Object { $_.DisplayName -like "$prefix*" } | Sort-Object DisplayName
            foreach ($grp in $filtered) {
                $clbLicenses.Items.Add($grp.DisplayName, $false) | Out-Null
                $script:LicenseGroupNames += $grp.DisplayName
            }
        }
    }

    $btnRefreshLic.Add_Click({ Update-OffboardLicenseList })

    # Fonction : afficher/masquer le panneau licence selon conditions
    function Update-LicensePanelVisibility {
        $showPanel = ($chkConvertShared.Checked -and $script:MailboxSizeGB -gt 50)
        $lblLicWarning.Visible = $showPanel
        $clbLicenses.Visible = $showPanel
        $btnRefreshLic.Visible = $showPanel
        $lblLicInfo.Visible = $showPanel
    }

    $chkConvertShared.Add_CheckedChanged({ Update-LicensePanelVisibility })

    # =================================================================
    # Labels de chargement et erreur
    # =================================================================
    $yBottom = $ySection2 + 110

    $lblChargement = New-Object System.Windows.Forms.Label
    $lblChargement.Text = Get-Text "offboarding.processing"
    $lblChargement.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $lblChargement.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $lblChargement.Location = New-Object System.Drawing.Point(15, $yBottom)
    $lblChargement.Size = New-Object System.Drawing.Size(400, 20)
    $lblChargement.Visible = $false
    $form.Controls.Add($lblChargement)

    $lblErreur = New-Object System.Windows.Forms.Label
    $lblErreur.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblErreur.ForeColor = [System.Drawing.Color]::Red
    $lblErreur.Location = New-Object System.Drawing.Point(15, ($yBottom + 22))
    $lblErreur.Size = New-Object System.Drawing.Size(600, 20)
    $lblErreur.Visible = $false
    $form.Controls.Add($lblErreur)

    # =================================================================
    # Boutons d'action
    # =================================================================
    $btnExecuter = New-Object System.Windows.Forms.Button
    $btnExecuter.Text = Get-Text "offboarding.btn_execute"
    $btnExecuter.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnExecuter.Location = New-Object System.Drawing.Point(280, 670)
    $btnExecuter.Size = New-Object System.Drawing.Size(210, 40)
    $btnExecuter.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $btnExecuter.ForeColor = [System.Drawing.Color]::White
    $btnExecuter.FlatStyle = "Flat"
    $form.Controls.Add($btnExecuter)

    $btnAnnuler = New-Object System.Windows.Forms.Button
    $btnAnnuler.Text = Get-Text "offboarding.btn_cancel"
    $btnAnnuler.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnAnnuler.Location = New-Object System.Drawing.Point(500, 670)
    $btnAnnuler.Size = New-Object System.Drawing.Size(130, 40)
    $btnAnnuler.FlatStyle = "Flat"
    $btnAnnuler.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnAnnuler)

    $form.CancelButton = $btnAnnuler

    # =================================================================
    # ÉVÉNEMENT : Recherche d'utilisateur
    # =================================================================
    $btnRecherche.Add_Click({
        $terme = $txtRecherche.Text.Trim()
        if ($terme.Length -lt 2) {
            [System.Windows.Forms.MessageBox]::Show(
                (Get-Text "offboarding.search_min_chars"),
                (Get-Text "offboarding.btn_search"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }
        $lstResultats.Items.Clear()
        $result = Search-AzUsers -SearchTerm $terme -MaxResults 15
        if ($result.Success -and $result.Data) {
            $script:SearchResults = @($result.Data)
            foreach ($user in $script:SearchResults) {
                $statut = if ($user.AccountEnabled) { Get-Text "offboarding.status_active" } else { Get-Text "offboarding.status_disabled" }
                $lstResultats.Items.Add("$($user.DisplayName) - $($user.UserPrincipalName) [$statut]") | Out-Null
            }
            $lstResultats.Visible = $true
        }
        else {
            $lstResultats.Visible = $false
            [System.Windows.Forms.MessageBox]::Show(
                (Get-Text "offboarding.search_no_result"),
                (Get-Text "offboarding.btn_search"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        }
    })

    $txtRecherche.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $btnRecherche.PerformClick()
            $_.SuppressKeyPress = $true
        }
    })

    # =================================================================
    # ÉVÉNEMENT : Sélection d'un utilisateur + vérification BAL
    # =================================================================
    $lstResultats.Add_SelectedIndexChanged({
        if ($lstResultats.SelectedIndex -ge 0 -and $lstResultats.SelectedIndex -lt $script:SearchResults.Count) {
            $selected = $script:SearchResults[$lstResultats.SelectedIndex]
            $script:SelectedUserId = $selected.Id
            $script:SelectedUserUPN = $selected.UserPrincipalName
            $script:SelectedUserName = $selected.DisplayName
            $lblUserInfo.Text = (Get-Text "offboarding.selected_user") -f $selected.DisplayName, $selected.UserPrincipalName
            $lblUserInfo.Visible = $true
            $lstResultats.Visible = $false

            # Réinitialiser l'état BAL
            $script:MailboxSizeGB = 0
            $lblMailboxInfo.Visible = $false
            $chkConvertShared.Enabled = $false
            Update-LicensePanelVisibility

            # Vérification de la taille de la BAL via Exchange Online
            if ($script:ExchangeAvailable) {
                $lblMailboxInfo.Text = Get-Text "offboarding.mailbox_checking"
                $lblMailboxInfo.ForeColor = [System.Drawing.Color]::Gray
                $lblMailboxInfo.Visible = $true
                $form.Refresh()

                $sizeResult = Get-AzMailboxSize -Identity $script:SelectedUserUPN

                if ($sizeResult.Success) {
                    $script:MailboxSizeGB = $sizeResult.SizeGB

                    if ($sizeResult.SizeGB -gt 50) {
                        # BAL > 50 Go — Warning orange
                        $lblMailboxInfo.Text = (Get-Text "offboarding.mailbox_over_limit") -f $sizeResult.SizeGB
                        $lblMailboxInfo.ForeColor = [System.Drawing.Color]::FromArgb(220, 130, 0)

                        # Charger les groupes de licence et afficher le panneau
                        Update-OffboardLicenseList
                        $licPrefix = if ($Config.PSObject.Properties["license_group_prefix"]) { $Config.license_group_prefix } else { "" }
                        if (-not [string]::IsNullOrWhiteSpace($licPrefix)) {
                            $lblLicInfo.Text = (Get-Text "offboarding.license_prefix_info") -f $licPrefix
                        }
                        $lblLicWarning.Text = Get-Text "offboarding.mailbox_license_required"
                    }
                    else {
                        # BAL <= 50 Go — OK vert
                        $lblMailboxInfo.Text = (Get-Text "offboarding.mailbox_ok") -f $sizeResult.SizeGB
                        $lblMailboxInfo.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
                    }

                    $chkConvertShared.Enabled = $true
                    $chkConvertShared.Checked = $false
                }
                else {
                    # Erreur Exchange — BAL peut-être inexistante
                    $lblMailboxInfo.Text = (Get-Text "offboarding.mailbox_error") -f $sizeResult.Error
                    $lblMailboxInfo.ForeColor = [System.Drawing.Color]::FromArgb(220, 130, 0)
                    $chkConvertShared.Enabled = $false
                    $chkConvertShared.Checked = $false
                }

                $lblMailboxInfo.Visible = $true
            }
            else {
                # Exchange Online non connecté
                $lblMailboxInfo.Text = Get-Text "offboarding.exchange_unavailable"
                $lblMailboxInfo.ForeColor = [System.Drawing.Color]::Gray
                $lblMailboxInfo.Visible = $true
                $chkConvertShared.Enabled = $false
                $chkConvertShared.Checked = $false
            }

            Update-LicensePanelVisibility
        }
    })

    # =================================================================
    # ÉVÉNEMENT : Exécution de l'offboarding
    # =================================================================
    $btnExecuter.Add_Click({
        $lblErreur.Visible = $false

        # Validation : utilisateur selectionne
        if ($null -eq $script:SelectedUserId) {
            $lblErreur.Text = Get-Text "offboarding.validation_no_user"
            $lblErreur.Visible = $true
            return
        }

        # Vérification qu'au moins une action est cochée
        $hasAction = $chkDesactiver.Checked -or $chkRevoquerLicences.Checked -or `
                     $chkRetirerGroupes.Checked -or $chkRevoquerSessions.Checked -or `
                     $chkHideGAL.Checked -or $chkConvertShared.Checked
        if (-not $hasAction) {
            $lblErreur.Text = Get-Text "offboarding.validation_no_action"
            $lblErreur.Visible = $true
            return
        }

        # Vérification : BAL > 50 Go + conversion cochée + aucune licence sélectionnée
        if ($chkConvertShared.Checked -and $script:MailboxSizeGB -gt 50 -and $clbLicenses.CheckedItems.Count -eq 0) {
            $confirmNoLic = Show-ConfirmDialog `
                -Titre (Get-Text "offboarding.confirm_no_license_title") `
                -Message (Get-Text "offboarding.confirm_no_license_msg") `
                -Icon ([System.Windows.Forms.MessageBoxIcon]::Warning)
            if (-not $confirmNoLic) { return }
        }

        # DOUBLE CONFIRMATION (securite offboarding)
        $confirmMsg = (Get-Text "offboarding.confirm_warning") + "`n`n"
        $confirmMsg += (Get-Text "offboarding.confirm_user") + " : $($script:SelectedUserName)`n"
        $confirmMsg += "UPN : $($script:SelectedUserUPN)`n"
        $confirmMsg += (Get-Text "offboarding.confirm_reason") + " : $($cboRaison.SelectedItem)`n`n"
        $confirmMsg += (Get-Text "offboarding.confirm_actions") + " :`n"
        if ($chkDesactiver.Checked)       { $confirmMsg += "  - " + (Get-Text "offboarding.chk_disable") + "`n" }
        if ($chkRevoquerSessions.Checked) { $confirmMsg += "  - " + (Get-Text "offboarding.chk_revoke_sessions") + "`n" }
        if ($chkRetirerGroupes.Checked)   { $confirmMsg += "  - " + (Get-Text "offboarding.chk_remove_groups") + "`n" }
        if ($chkHideGAL.Checked)          { $confirmMsg += "  - " + (Get-Text "offboarding.chk_hide_gal") + "`n" }
        if ($chkConvertShared.Checked)    { $confirmMsg += "  - " + (Get-Text "offboarding.chk_convert_shared") + "`n" }
        if ($chkRevoquerLicences.Checked) { $confirmMsg += "  - " + (Get-Text "offboarding.chk_revoke_licenses") + "`n" }
        $confirmMsg += "`n" + (Get-Text "offboarding.confirm_irreversible")

        $confirm1 = Show-ConfirmDialog -Titre (Get-Text "offboarding.confirm_title") -Message $confirmMsg -Icon ([System.Windows.Forms.MessageBoxIcon]::Warning)
        if (-not $confirm1) { return }

        # Deuxieme confirmation
        $confirm2 = Show-ConfirmDialog `
            -Titre (Get-Text "offboarding.confirm_title_2") `
            -Message ((Get-Text "offboarding.confirm_final") -f $script:SelectedUserName) `
            -Icon ([System.Windows.Forms.MessageBoxIcon]::Exclamation)
        if (-not $confirm2) { return }

        # =================================================================
        # Exécution séquentielle des actions
        # Ordre : Disable+JobTitle → Sessions → Groups(skip dynamic) →
        #         Hide GAL → Convert Shared → Add License → Remove Licenses
        #         → Add to disabled group → Notification
        # =================================================================
        $btnExecuter.Enabled = $false
        $lblChargement.Visible = $true
        $form.Refresh()

        $userId = $script:SelectedUserId
        $upn = $script:SelectedUserUPN
        $erreurs = @()
        $actionsReussies = @()

        try {
            # 1. Désactivation du compte + suffixe "- DISABLED" sur jobTitle
            if ($chkDesactiver.Checked) {
                $lblChargement.Text = Get-Text "offboarding.step_disable"
                $form.Refresh()
                $result = Disable-AzUser -UserId $userId
                if ($result.Success) {
                    $actionsReussies += Get-Text "offboarding.result_disabled"
                    if ($result.OriginalJobTitle) {
                        $actionsReussies += (Get-Text "offboarding.result_jobtitle") -f $result.OriginalJobTitle
                    }
                }
                else { $erreurs += (Get-Text "offboarding.error_disable") + " : $($result.Error)" }
            }

            # 2. Révocation des sessions
            if ($chkRevoquerSessions.Checked) {
                $lblChargement.Text = Get-Text "offboarding.step_sessions"
                $form.Refresh()
                $result = Revoke-AzUserSessions -UserId $userId
                if ($result.Success) { $actionsReussies += Get-Text "offboarding.result_sessions" }
                else { $erreurs += (Get-Text "offboarding.error_sessions") + " : $($result.Error)" }
            }

            # 3. Retrait des groupes (skip dynamiques)
            if ($chkRetirerGroupes.Checked) {
                $lblChargement.Text = Get-Text "offboarding.step_groups"
                $form.Refresh()
                $result = Remove-AzUserGroups -UserId $userId
                if ($result.Success) {
                    $msg = (Get-Text "offboarding.result_groups") -f $result.RemovedCount
                    if ($result.SkippedDynamic -gt 0) {
                        $msg += " " + ((Get-Text "offboarding.result_groups_dynamic") -f $result.SkippedDynamic)
                    }
                    $actionsReussies += $msg
                }
                else { $erreurs += (Get-Text "offboarding.error_groups") + " : $($result.Error)" }
            }

            # 4. Masquer du carnet d'adresses (GAL)
            if ($chkHideGAL.Checked) {
                $lblChargement.Text = Get-Text "offboarding.step_hide_gal"
                $form.Refresh()
                $result = Hide-AzMailboxFromGAL -Identity $upn -UserId $userId
                if ($result.Success) { $actionsReussies += (Get-Text "offboarding.result_hide_gal") + " ($($result.Method))" }
                else { $erreurs += (Get-Text "offboarding.error_hide_gal") + " : $($result.Error)" }
            }

            # 5. Conversion en boîte partagée (AVANT révocation licences)
            if ($chkConvertShared.Checked) {
                $lblChargement.Text = Get-Text "offboarding.step_convert"
                $form.Refresh()
                $result = Convert-AzMailboxToShared -Identity $upn
                if ($result.Success) { $actionsReussies += Get-Text "offboarding.result_convert" }
                else { $erreurs += (Get-Text "offboarding.error_convert") + " : $($result.Error)" }
            }

            # 5b. Ajout de licence Exchange si BAL > 50 Go + licence sélectionnée
            if ($chkConvertShared.Checked -and $clbLicenses.CheckedItems.Count -gt 0) {
                $lblChargement.Text = Get-Text "offboarding.step_add_license"
                $form.Refresh()
                foreach ($licGrp in $clbLicenses.CheckedItems) {
                    $licResult = Add-AzUserToGroup -UserId $userId -GroupName $licGrp
                    if ($licResult.Success) { $actionsReussies += (Get-Text "offboarding.result_add_license") -f $licGrp }
                    else { $erreurs += (Get-Text "offboarding.error_add_license") -f $licGrp, $licResult.Error }
                }
            }

            # 5c. Délégation FullAccess (Read & Manage) sur la boîte
            $delegateUpn = $txtDelegate.Text.Trim()
            if (-not [string]::IsNullOrWhiteSpace($delegateUpn)) {
                $lblChargement.Text = Get-Text "offboarding.step_delegate"
                $form.Refresh()
                $result = Grant-AzMailboxFullAccess -MailboxIdentity $upn -DelegateUPN $delegateUpn
                if ($result.Success) { $actionsReussies += (Get-Text "offboarding.result_delegate") -f $delegateUpn }
                else { $erreurs += (Get-Text "offboarding.error_delegate") -f $delegateUpn, $result.Error }
            }

            # 6. Révocation des licences (les héritées de groupes sont ignorées)
            if ($chkRevoquerLicences.Checked) {
                $lblChargement.Text = Get-Text "offboarding.step_licenses"
                $form.Refresh()
                $result = Remove-AzUserLicenses -UserId $userId
                if ($result.Success) {
                    $msg = (Get-Text "offboarding.result_licenses") -f $result.RemovedCount
                    if ($result.SkippedInherited -gt 0) {
                        $msg += " " + ((Get-Text "offboarding.result_licenses_inherited") -f $result.SkippedInherited)
                    }
                    $actionsReussies += $msg
                }
                else { $erreurs += (Get-Text "offboarding.error_licenses") + " : $($result.Error)" }
            }

            # 7. Ajout au groupe des comptes désactivés
            if ($chkDesactiver.Checked -and -not [string]::IsNullOrWhiteSpace($Config.offboarding.disabled_ou_group)) {
                $lblChargement.Text = Get-Text "offboarding.step_disabled_group"
                $form.Refresh()
                $grpResult = Add-AzUserToGroup -UserId $userId -GroupName $Config.offboarding.disabled_ou_group
                if ($grpResult.Success) { $actionsReussies += (Get-Text "offboarding.result_disabled_group") -f $Config.offboarding.disabled_ou_group }
                else { $erreurs += (Get-Text "offboarding.error_disabled_group") + " : $($grpResult.Error)" }
            }

            # 8. Notification
            if ($Config.notifications.enabled) {
                $sujet = "Offboarding - $($script:SelectedUserName) ($upn)"
                $corps = "<h2>Offboarding effectué</h2>"
                $corps += "<p><strong>Employé :</strong> $($script:SelectedUserName)</p>"
                $corps += "<p><strong>UPN :</strong> $upn</p>"
                $corps += "<p><strong>Raison :</strong> $($cboRaison.SelectedItem)</p>"
                $corps += "<p><strong>Date de départ :</strong> $($dtpDepart.Value.ToString('yyyy-MM-dd'))</p>"
                $corps += "<p><strong>Actions effectuées :</strong></p><ul>"
                foreach ($action in $actionsReussies) { $corps += "<li>$action</li>" }
                $corps += "</ul>"
                if ($erreurs.Count -gt 0) {
                    $corps += "<p style='color:red'><strong>Erreurs :</strong></p><ul>"
                    foreach ($err in $erreurs) { $corps += "<li>$err</li>" }
                    $corps += "</ul>"
                }
                Send-Notification -Sujet $sujet -Corps $corps
            }

            # 9. Résultat
            $lblChargement.Visible = $false

            $recapMsg = ((Get-Text "offboarding.result_summary") -f $script:SelectedUserName) + "`n`n"
            $recapMsg += (Get-Text "offboarding.result_success_label") + " :`n" + ($actionsReussies -join "`n") + "`n"
            if ($erreurs.Count -gt 0) {
                $recapMsg += "`n" + (Get-Text "offboarding.result_error_label") + " :`n" + ($erreurs -join "`n")
            }

            $isSuccess = ($erreurs.Count -eq 0)
            Show-ResultDialog -Titre (Get-Text "offboarding.result_title") -Message $recapMsg -IsSuccess $isSuccess

            Write-Log -Level $(if ($isSuccess) { "SUCCESS" } else { "WARNING" }) -Action "OFFBOARDING" -UPN $upn -Message "Offboarding terminé. Réussites: $($actionsReussies.Count), Erreurs: $($erreurs.Count)"

            $form.Close()
        }
        catch {
            $lblChargement.Visible = $false
            $btnExecuter.Enabled = $true
            $errMsg = $_.Exception.Message
            Write-Log -Level "ERROR" -Action "OFFBOARDING" -UPN $upn -Message "Erreur offboarding : $errMsg"
            Show-ResultDialog -Titre (Get-Text "offboarding.result_error_title") -Message ((Get-Text "offboarding.result_error_msg") -f $errMsg) -IsSuccess $false
        }
    })

    $form.ShowDialog() | Out-Null
    $form.Dispose()
}

# Point d'attention :
# - DOUBLE CONFIRMATION obligatoire avant toute action d'offboarding
# - Les cases à cocher sont pré-cochées selon la configuration du client
# - L'ajout au groupe "Comptes-Désactivés" est automatique si configuré
# - Le suffixe "- DISABLED" est ajouté au jobTitle pour déclencher l'exclusion
#   des groupes dynamiques (règle: user.jobTitle -notContains "- DISABLED")
# - Les groupes dynamiques sont automatiquement ignorés lors du retrait
# - La conversion en boîte partagée se fait AVANT la révocation des licences
# - Si la BAL dépasse 50 Go, un sélecteur de licence Exchange est affiché
# - Le masquage du GAL utilise Exchange Online avec fallback Graph API