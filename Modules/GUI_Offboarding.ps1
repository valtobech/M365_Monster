<#
.FICHIER
    Modules/GUI_Offboarding.ps1

.ROLE
    Formulaire de depart (offboarding) d'un employe.
    Gere la desactivation du compte, la revocation des licences et sessions,
    le retrait des groupes, et la notification.

.DEPENDANCES
    - Core/Functions.ps1 (Write-Log, Show-ConfirmDialog, Show-ResultDialog, Send-Notification)
    - Core/GraphAPI.ps1 (Search-AzUsers, Disable-AzUser, Revoke-AzUserSessions,
      Remove-AzUserGroups, Remove-AzUserLicenses, Add-AzUserToGroup)
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

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Offboarding - Depart employe - $($Config.client_name)"
    $form.Size = New-Object System.Drawing.Size(620, 620)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::WhiteSmoke

    $yPos = 15

    # === Titre ===
    $lblSection = New-Object System.Windows.Forms.Label
    $lblSection.Text = "Recherche de l'employe"
    $lblSection.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblSection.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $lblSection.Location = New-Object System.Drawing.Point(15, $yPos)
    $lblSection.Size = New-Object System.Drawing.Size(560, 25)
    $form.Controls.Add($lblSection)
    $yPos += 35

    # === Barre de recherche ===
    $lblRecherche = New-Object System.Windows.Forms.Label
    $lblRecherche.Text = "Rechercher :"
    $lblRecherche.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblRecherche.Location = New-Object System.Drawing.Point(15, ($yPos + 3))
    $lblRecherche.Size = New-Object System.Drawing.Size(90, 20)
    $form.Controls.Add($lblRecherche)

    $txtRecherche = New-Object System.Windows.Forms.TextBox
    $txtRecherche.Location = New-Object System.Drawing.Point(110, $yPos)
    $txtRecherche.Size = New-Object System.Drawing.Size(380, 25)
    $txtRecherche.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($txtRecherche)

    $btnRecherche = New-Object System.Windows.Forms.Button
    $btnRecherche.Text = "Chercher"
    $btnRecherche.Location = New-Object System.Drawing.Point(500, $yPos)
    $btnRecherche.Size = New-Object System.Drawing.Size(90, 25)
    $form.Controls.Add($btnRecherche)
    $yPos += 33

    # Liste de resultats
    $lstResultats = New-Object System.Windows.Forms.ListBox
    $lstResultats.Location = New-Object System.Drawing.Point(110, $yPos)
    $lstResultats.Size = New-Object System.Drawing.Size(380, 80)
    $lstResultats.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lstResultats.Visible = $false
    $form.Controls.Add($lstResultats)

    # Info utilisateur selectionne
    $lblUserInfo = New-Object System.Windows.Forms.Label
    $lblUserInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblUserInfo.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $lblUserInfo.Location = New-Object System.Drawing.Point(15, ($yPos + 85))
    $lblUserInfo.Size = New-Object System.Drawing.Size(560, 20)
    $lblUserInfo.Visible = $false
    $form.Controls.Add($lblUserInfo)

    $script:SelectedUserId = $null
    $script:SelectedUserUPN = $null
    $script:SelectedUserName = $null
    $script:SearchResults = @()

    # Recherche
    $btnRecherche.Add_Click({
        $terme = $txtRecherche.Text.Trim()
        if ($terme.Length -lt 2) {
            [System.Windows.Forms.MessageBox]::Show("Saisissez au moins 2 caracteres.", "Recherche", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }
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
        else {
            $lstResultats.Visible = $false
            [System.Windows.Forms.MessageBox]::Show("Aucun utilisateur trouve.", "Recherche", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
    })

    $txtRecherche.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $btnRecherche.PerformClick()
            $_.SuppressKeyPress = $true
        }
    })

    # Selection d'un utilisateur
    $lstResultats.Add_SelectedIndexChanged({
        if ($lstResultats.SelectedIndex -ge 0 -and $lstResultats.SelectedIndex -lt $script:SearchResults.Count) {
            $selected = $script:SearchResults[$lstResultats.SelectedIndex]
            $script:SelectedUserId = $selected.Id
            $script:SelectedUserUPN = $selected.UserPrincipalName
            $script:SelectedUserName = $selected.DisplayName
            $lblUserInfo.Text = "Selectionne : $($selected.DisplayName) ($($selected.UserPrincipalName))"
            $lblUserInfo.Visible = $true
            $lstResultats.Visible = $false
        }
    })

    # === Section Details du depart ===
    $ySection2 = 240

    $lblSection2 = New-Object System.Windows.Forms.Label
    $lblSection2.Text = "Details du depart"
    $lblSection2.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblSection2.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $lblSection2.Location = New-Object System.Drawing.Point(15, $ySection2)
    $lblSection2.Size = New-Object System.Drawing.Size(560, 25)
    $form.Controls.Add($lblSection2)
    $ySection2 += 35

    # Date de depart
    $lblDate = New-Object System.Windows.Forms.Label
    $lblDate.Text = "Date de depart :"
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
    $lblRaison.Text = "Raison :"
    $lblRaison.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblRaison.Location = New-Object System.Drawing.Point(15, ($ySection2 + 3))
    $lblRaison.Size = New-Object System.Drawing.Size(120, 20)
    $form.Controls.Add($lblRaison)

    $cboRaison = New-Object System.Windows.Forms.ComboBox
    $cboRaison.Location = New-Object System.Drawing.Point(150, $ySection2)
    $cboRaison.Size = New-Object System.Drawing.Size(200, 25)
    $cboRaison.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cboRaison.DropDownStyle = "DropDownList"
    @("Demission", "Licenciement", "Fin de contrat", "Autre") | ForEach-Object { $cboRaison.Items.Add($_) | Out-Null }
    $cboRaison.SelectedIndex = 0
    $form.Controls.Add($cboRaison)
    $ySection2 += 33

    # Redirection de boite mail
    $lblRedirection = New-Object System.Windows.Forms.Label
    $lblRedirection.Text = "Rediriger mail vers :"
    $lblRedirection.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblRedirection.Location = New-Object System.Drawing.Point(15, ($ySection2 + 3))
    $lblRedirection.Size = New-Object System.Drawing.Size(130, 20)
    $form.Controls.Add($lblRedirection)

    $txtRedirection = New-Object System.Windows.Forms.TextBox
    $txtRedirection.Location = New-Object System.Drawing.Point(150, $ySection2)
    $txtRedirection.Size = New-Object System.Drawing.Size(300, 25)
    $txtRedirection.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($txtRedirection)
    $ySection2 += 40

    # === Cases a cocher des actions ===
    $chkDesactiver = New-Object System.Windows.Forms.CheckBox
    $chkDesactiver.Text = "Desactiver le compte"
    $chkDesactiver.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkDesactiver.Location = New-Object System.Drawing.Point(30, $ySection2)
    $chkDesactiver.Size = New-Object System.Drawing.Size(250, 22)
    $chkDesactiver.Checked = $true
    $form.Controls.Add($chkDesactiver)

    $chkRevoquerLicences = New-Object System.Windows.Forms.CheckBox
    $chkRevoquerLicences.Text = "Revoquer les licences"
    $chkRevoquerLicences.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkRevoquerLicences.Location = New-Object System.Drawing.Point(300, $ySection2)
    $chkRevoquerLicences.Size = New-Object System.Drawing.Size(250, 22)
    $chkRevoquerLicences.Checked = $Config.offboarding.revoke_licenses
    $form.Controls.Add($chkRevoquerLicences)
    $ySection2 += 28

    $chkRetirerGroupes = New-Object System.Windows.Forms.CheckBox
    $chkRetirerGroupes.Text = "Retirer tous les groupes"
    $chkRetirerGroupes.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkRetirerGroupes.Location = New-Object System.Drawing.Point(30, $ySection2)
    $chkRetirerGroupes.Size = New-Object System.Drawing.Size(250, 22)
    $chkRetirerGroupes.Checked = $Config.offboarding.remove_all_groups
    $form.Controls.Add($chkRetirerGroupes)

    $chkRevoquerSessions = New-Object System.Windows.Forms.CheckBox
    $chkRevoquerSessions.Text = "Forcer deconnexion des sessions"
    $chkRevoquerSessions.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkRevoquerSessions.Location = New-Object System.Drawing.Point(300, $ySection2)
    $chkRevoquerSessions.Size = New-Object System.Drawing.Size(280, 22)
    $chkRevoquerSessions.Checked = $true
    $form.Controls.Add($chkRevoquerSessions)
    $ySection2 += 35

    # Label de chargement
    $lblChargement = New-Object System.Windows.Forms.Label
    $lblChargement.Text = "Traitement de l'offboarding en cours..."
    $lblChargement.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $lblChargement.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $lblChargement.Location = New-Object System.Drawing.Point(15, $ySection2)
    $lblChargement.Size = New-Object System.Drawing.Size(350, 20)
    $lblChargement.Visible = $false
    $form.Controls.Add($lblChargement)

    # Label d'erreur
    $lblErreur = New-Object System.Windows.Forms.Label
    $lblErreur.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblErreur.ForeColor = [System.Drawing.Color]::Red
    $lblErreur.Location = New-Object System.Drawing.Point(15, ($ySection2 + 22))
    $lblErreur.Size = New-Object System.Drawing.Size(560, 20)
    $lblErreur.Visible = $false
    $form.Controls.Add($lblErreur)

    # === Boutons ===
    $btnExecuter = New-Object System.Windows.Forms.Button
    $btnExecuter.Text = "Executer l'offboarding"
    $btnExecuter.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnExecuter.Location = New-Object System.Drawing.Point(250, 530)
    $btnExecuter.Size = New-Object System.Drawing.Size(190, 40)
    $btnExecuter.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $btnExecuter.ForeColor = [System.Drawing.Color]::White
    $btnExecuter.FlatStyle = "Flat"
    $form.Controls.Add($btnExecuter)

    $btnAnnuler = New-Object System.Windows.Forms.Button
    $btnAnnuler.Text = "Annuler"
    $btnAnnuler.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnAnnuler.Location = New-Object System.Drawing.Point(450, 530)
    $btnAnnuler.Size = New-Object System.Drawing.Size(140, 40)
    $btnAnnuler.FlatStyle = "Flat"
    $btnAnnuler.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnAnnuler)

    $form.CancelButton = $btnAnnuler

    # === Action d'offboarding ===
    $btnExecuter.Add_Click({
        $lblErreur.Visible = $false

        # Validation : utilisateur selectionne
        if ($null -eq $script:SelectedUserId) {
            $lblErreur.Text = "Veuillez rechercher et selectionner un utilisateur."
            $lblErreur.Visible = $true
            return
        }

        # Verification qu'au moins une action est cochee
        if (-not $chkDesactiver.Checked -and -not $chkRevoquerLicences.Checked -and -not $chkRetirerGroupes.Checked -and -not $chkRevoquerSessions.Checked) {
            $lblErreur.Text = "Veuillez selectionner au moins une action."
            $lblErreur.Visible = $true
            return
        }

        # DOUBLE CONFIRMATION (securite offboarding)
        $confirmMsg = "ATTENTION : Vous etes sur le point d'effectuer l'offboarding de :`n`n"
        $confirmMsg += "Utilisateur : $($script:SelectedUserName)`n"
        $confirmMsg += "UPN : $($script:SelectedUserUPN)`n"
        $confirmMsg += "Raison : $($cboRaison.SelectedItem)`n`n"
        $confirmMsg += "Actions prevues :`n"
        if ($chkDesactiver.Checked) { $confirmMsg += "  - Desactiver le compte`n" }
        if ($chkRevoquerSessions.Checked) { $confirmMsg += "  - Revoquer les sessions actives`n" }
        if ($chkRetirerGroupes.Checked) { $confirmMsg += "  - Retirer tous les groupes`n" }
        if ($chkRevoquerLicences.Checked) { $confirmMsg += "  - Revoquer toutes les licences`n" }
        $confirmMsg += "`nCette action est irreversible. Confirmer ?"

        $confirm1 = Show-ConfirmDialog -Titre "Confirmation de l'offboarding" -Message $confirmMsg -Icon ([System.Windows.Forms.MessageBoxIcon]::Warning)
        if (-not $confirm1) { return }

        # Deuxieme confirmation
        $confirm2 = Show-ConfirmDialog -Titre "SECONDE CONFIRMATION" -Message "Etes-vous ABSOLUMENT certain de vouloir proceder a l'offboarding de $($script:SelectedUserName) ?`n`nCette action ne peut pas etre annulee." -Icon ([System.Windows.Forms.MessageBoxIcon]::Exclamation)
        if (-not $confirm2) { return }

        # Execution
        $btnExecuter.Enabled = $false
        $lblChargement.Visible = $true
        $form.Refresh()

        $userId = $script:SelectedUserId
        $upn = $script:SelectedUserUPN
        $erreurs = @()
        $actionsReussies = @()

        try {
            # 1. Desactivation du compte
            if ($chkDesactiver.Checked) {
                $result = Disable-AzUser -UserId $userId
                if ($result.Success) { $actionsReussies += "Compte desactive" }
                else { $erreurs += "Desactivation : $($result.Error)" }
            }

            # 2. Revocation des sessions
            if ($chkRevoquerSessions.Checked) {
                $result = Revoke-AzUserSessions -UserId $userId
                if ($result.Success) { $actionsReussies += "Sessions revoquees" }
                else { $erreurs += "Sessions : $($result.Error)" }
            }

            # 3. Retrait des groupes
            if ($chkRetirerGroupes.Checked) {
                $result = Remove-AzUserGroups -UserId $userId
                if ($result.Success) { $actionsReussies += "Retire de $($result.RemovedCount) groupe(s)" }
                else { $erreurs += "Groupes : $($result.Error)" }
            }

            # 4. Revocation des licences
            if ($chkRevoquerLicences.Checked) {
                $result = Remove-AzUserLicenses -UserId $userId
                if ($result.Success) { $actionsReussies += "$($result.RemovedCount) licence(s) revoquee(s)" }
                else { $erreurs += "Licences : $($result.Error)" }
            }

            # 5. Ajout au groupe des comptes desactives
            if ($chkDesactiver.Checked -and -not [string]::IsNullOrWhiteSpace($Config.offboarding.disabled_ou_group)) {
                $grpResult = Add-AzUserToGroup -UserId $userId -GroupName $Config.offboarding.disabled_ou_group
                if ($grpResult.Success) { $actionsReussies += "Ajoute au groupe '$($Config.offboarding.disabled_ou_group)'" }
                else { $erreurs += "Groupe desactives : $($grpResult.Error)" }
            }

            # 6. Notification
            if ($Config.notifications.enabled) {
                $sujet = "Offboarding - $($script:SelectedUserName) ($upn)"
                $corps = "<h2>Offboarding effectue</h2>"
                $corps += "<p><strong>Employe :</strong> $($script:SelectedUserName)</p>"
                $corps += "<p><strong>UPN :</strong> $upn</p>"
                $corps += "<p><strong>Raison :</strong> $($cboRaison.SelectedItem)</p>"
                $corps += "<p><strong>Date de depart :</strong> $($dtpDepart.Value.ToString('yyyy-MM-dd'))</p>"
                $corps += "<p><strong>Actions effectuees :</strong></p><ul>"
                foreach ($action in $actionsReussies) { $corps += "<li>$action</li>" }
                $corps += "</ul>"
                if ($erreurs.Count -gt 0) {
                    $corps += "<p style='color:red'><strong>Erreurs :</strong></p><ul>"
                    foreach ($err in $erreurs) { $corps += "<li>$err</li>" }
                    $corps += "</ul>"
                }
                Send-Notification -Sujet $sujet -Corps $corps
            }

            # 7. Resultat
            $lblChargement.Visible = $false

            $recapMsg = "Offboarding de $($script:SelectedUserName) termine.`n`n"
            $recapMsg += "Actions reussies :`n" + ($actionsReussies -join "`n") + "`n"
            if ($erreurs.Count -gt 0) {
                $recapMsg += "`nErreurs :`n" + ($erreurs -join "`n")
            }

            $isSuccess = ($erreurs.Count -eq 0)
            Show-ResultDialog -Titre "Resultat de l'offboarding" -Message $recapMsg -IsSuccess $isSuccess

            Write-Log -Level $(if ($isSuccess) { "SUCCESS" } else { "WARNING" }) -Action "OFFBOARDING" -UPN $upn -Message "Offboarding termine. Reussites: $($actionsReussies.Count), Erreurs: $($erreurs.Count)"

            $form.Close()
        }
        catch {
            $lblChargement.Visible = $false
            $btnExecuter.Enabled = $true
            $errMsg = $_.Exception.Message
            Write-Log -Level "ERROR" -Action "OFFBOARDING" -UPN $upn -Message "Erreur offboarding : $errMsg"
            Show-ResultDialog -Titre "Erreur d'offboarding" -Message "Une erreur est survenue :`n`n$errMsg" -IsSuccess $false
        }
    })

    $form.ShowDialog() | Out-Null
    $form.Dispose()
}

# Point d'attention :
# - DOUBLE CONFIRMATION obligatoire avant toute action d'offboarding
# - Les cases a cocher sont pre-cochees selon la configuration du client
# - L'ajout au groupe "Comptes-Desactives" est automatique si configure
# - La redirection de boite mail est preparee dans le formulaire mais
#   necessite l'API Exchange Online (non implementee ici — CHOIX: a completer)
