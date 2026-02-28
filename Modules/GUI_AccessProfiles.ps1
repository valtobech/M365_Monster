<#
.FICHIER
    Modules/GUI_AccessProfiles.ps1

.ROLE
    Éditeur de profils d'accès (paquets de groupes par poste).
    Permet de créer, modifier, supprimer des profils et d'y associer
    des groupes Entra ID via recherche Graph.

.DEPENDANCES
    - Core/Functions.ps1  (Write-Log, Show-ConfirmDialog, Show-ResultDialog,
                           Get-AccessProfiles, Get-BaselineProfile)
    - Core/GraphAPI.ps1   (Search-AzGroups)
    - Core/Lang.ps1       (Get-Text)
    - Variable globale    $Config, $RootPath

.AUTEUR
    [Equipe IT — M365 Monster]
#>

function Show-AccessProfileEditor {
    <#
    .SYNOPSIS
        Affiche l'éditeur de profils d'accès pour le client courant.
    .OUTPUTS
        [void] — Formulaire modal.
    #>

    # Vérifier que la config supporte les profils
    if (-not $Config.PSObject.Properties["access_profiles"]) {
        Show-ResultDialog -Titre (Get-Text "access_profiles.title") `
            -Message "Aucune section 'access_profiles' dans la configuration client. Ajoutez-la d'abord via le template." `
            -IsSuccess $false
        return
    }

    # Charger les données de travail (copie locale modifiable)
    $script:APData = @{}
    foreach ($key in $Config.access_profiles.PSObject.Properties.Name) {
        $src = $Config.access_profiles.$key
        $script:APData[$key] = @{
            display_name = $src.display_name
            description  = $src.description
            is_baseline  = [bool]$src.is_baseline
            groups       = @($src.groups | ForEach-Object {
                @{ id = $_.id; display_name = $_.display_name }
            })
        }
    }
    $script:APDirty = $false

    # ================================================================
    #  FORMULAIRE PRINCIPAL
    # ================================================================
    $f = New-Object System.Windows.Forms.Form
    $f.Text = Get-Text "access_profiles.title"
    $f.Size = New-Object System.Drawing.Size(820, 620)
    $f.StartPosition = "CenterScreen"
    $f.FormBorderStyle = "FixedDialog"
    $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.BackColor = [System.Drawing.Color]::WhiteSmoke
    $f.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # ---- PANNEAU GAUCHE : liste des profils ----
    $lblList = New-Object System.Windows.Forms.Label
    $lblList.Text = Get-Text "access_profiles.list_label"
    $lblList.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblList.Location = New-Object System.Drawing.Point(10, 10)
    $lblList.Size = New-Object System.Drawing.Size(200, 22)
    $f.Controls.Add($lblList)

    $lstProfiles = New-Object System.Windows.Forms.ListBox
    $lstProfiles.Location = New-Object System.Drawing.Point(10, 35)
    $lstProfiles.Size = New-Object System.Drawing.Size(200, 440)
    $lstProfiles.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $f.Controls.Add($lstProfiles)

    $btnNew = New-Object System.Windows.Forms.Button
    $btnNew.Text = Get-Text "access_profiles.btn_new"
    $btnNew.Location = New-Object System.Drawing.Point(10, 482)
    $btnNew.Size = New-Object System.Drawing.Size(110, 30)
    $btnNew.FlatStyle = "Flat"
    $btnNew.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $btnNew.ForeColor = [System.Drawing.Color]::White
    $f.Controls.Add($btnNew)

    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Text = Get-Text "access_profiles.btn_delete"
    $btnDelete.Location = New-Object System.Drawing.Point(125, 482)
    $btnDelete.Size = New-Object System.Drawing.Size(95, 30)
    $btnDelete.FlatStyle = "Flat"
    $btnDelete.ForeColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $f.Controls.Add($btnDelete)

    # ---- PANNEAU DROIT : édition du profil sélectionné ----
    $pnlEdit = New-Object System.Windows.Forms.Panel
    $pnlEdit.Location = New-Object System.Drawing.Point(220, 10)
    $pnlEdit.Size = New-Object System.Drawing.Size(575, 505)
    $pnlEdit.BorderStyle = "FixedSingle"
    $pnlEdit.BackColor = [System.Drawing.Color]::White
    $f.Controls.Add($pnlEdit)

    # Nom du profil
    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text = Get-Text "access_profiles.edit_name"
    $lblName.Location = New-Object System.Drawing.Point(10, 12)
    $lblName.Size = New-Object System.Drawing.Size(120, 20)
    $pnlEdit.Controls.Add($lblName)

    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(135, 10)
    $txtName.Size = New-Object System.Drawing.Size(250, 25)
    $pnlEdit.Controls.Add($txtName)

    # Description
    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = Get-Text "access_profiles.edit_description"
    $lblDesc.Location = New-Object System.Drawing.Point(10, 42)
    $lblDesc.Size = New-Object System.Drawing.Size(120, 20)
    $pnlEdit.Controls.Add($lblDesc)

    $txtDesc = New-Object System.Windows.Forms.TextBox
    $txtDesc.Location = New-Object System.Drawing.Point(135, 40)
    $txtDesc.Size = New-Object System.Drawing.Size(420, 25)
    $pnlEdit.Controls.Add($txtDesc)

    # Checkbox baseline
    $chkBaseline = New-Object System.Windows.Forms.CheckBox
    $chkBaseline.Text = Get-Text "access_profiles.edit_baseline"
    $chkBaseline.Location = New-Object System.Drawing.Point(135, 70)
    $chkBaseline.Size = New-Object System.Drawing.Size(350, 22)
    $pnlEdit.Controls.Add($chkBaseline)

    # ---- Section groupes assignés ----
    $lblAssigned = New-Object System.Windows.Forms.Label
    $lblAssigned.Text = Get-Text "access_profiles.groups_assigned"
    $lblAssigned.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblAssigned.Location = New-Object System.Drawing.Point(10, 100)
    $lblAssigned.Size = New-Object System.Drawing.Size(350, 20)
    $pnlEdit.Controls.Add($lblAssigned)

    $lstGroups = New-Object System.Windows.Forms.ListBox
    $lstGroups.Location = New-Object System.Drawing.Point(10, 122)
    $lstGroups.Size = New-Object System.Drawing.Size(440, 140)
    $lstGroups.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lstGroups.SelectionMode = "MultiExtended"
    $pnlEdit.Controls.Add($lstGroups)

    $btnRemoveGrp = New-Object System.Windows.Forms.Button
    $btnRemoveGrp.Text = Get-Text "access_profiles.groups_btn_remove"
    $btnRemoveGrp.Location = New-Object System.Drawing.Point(458, 122)
    $btnRemoveGrp.Size = New-Object System.Drawing.Size(100, 30)
    $btnRemoveGrp.FlatStyle = "Flat"
    $btnRemoveGrp.ForeColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $btnRemoveGrp.Enabled = $false
    $pnlEdit.Controls.Add($btnRemoveGrp)

    # ---- Section recherche / ajout de groupes ----
    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text = Get-Text "access_profiles.groups_search"
    $lblSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblSearch.Location = New-Object System.Drawing.Point(10, 272)
    $lblSearch.Size = New-Object System.Drawing.Size(250, 20)
    $pnlEdit.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(10, 294)
    $txtSearch.Size = New-Object System.Drawing.Size(330, 25)
    $txtSearch.PlaceholderText = Get-Text "access_profiles.groups_search_placeholder"
    $pnlEdit.Controls.Add($txtSearch)

    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Text = Get-Text "access_profiles.groups_btn_search"
    $btnSearch.Location = New-Object System.Drawing.Point(345, 294)
    $btnSearch.Size = New-Object System.Drawing.Size(100, 25)
    $btnSearch.FlatStyle = "Flat"
    $btnSearch.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
    $btnSearch.ForeColor = [System.Drawing.Color]::White
    $pnlEdit.Controls.Add($btnSearch)

    $lstSearchResults = New-Object System.Windows.Forms.ListBox
    $lstSearchResults.Location = New-Object System.Drawing.Point(10, 324)
    $lstSearchResults.Size = New-Object System.Drawing.Size(440, 100)
    $lstSearchResults.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lstSearchResults.SelectionMode = "MultiExtended"
    $pnlEdit.Controls.Add($lstSearchResults)

    $btnAddGrp = New-Object System.Windows.Forms.Button
    $btnAddGrp.Text = Get-Text "access_profiles.groups_btn_add"
    $btnAddGrp.Location = New-Object System.Drawing.Point(458, 324)
    $btnAddGrp.Size = New-Object System.Drawing.Size(100, 30)
    $btnAddGrp.FlatStyle = "Flat"
    $btnAddGrp.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $btnAddGrp.ForeColor = [System.Drawing.Color]::White
    $btnAddGrp.Enabled = $false
    $pnlEdit.Controls.Add($btnAddGrp)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = Get-Text "access_profiles.groups_search_hint"
    $lblHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblHint.ForeColor = [System.Drawing.Color]::Gray
    $lblHint.Location = New-Object System.Drawing.Point(10, 428)
    $lblHint.Size = New-Object System.Drawing.Size(400, 18)
    $lblHint.BackColor = [System.Drawing.Color]::Transparent
    $pnlEdit.Controls.Add($lblHint)

    # Bouton Réconcilier (comparer template vs production)
    $btnReconcile = New-Object System.Windows.Forms.Button
    $btnReconcile.Text = Get-Text "access_profiles.btn_reconcile"
    $btnReconcile.Location = New-Object System.Drawing.Point(335, 460)
    $btnReconcile.Size = New-Object System.Drawing.Size(115, 32)
    $btnReconcile.FlatStyle = "Flat"
    $btnReconcile.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
    $btnReconcile.ForeColor = [System.Drawing.Color]::White
    $btnReconcile.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $pnlEdit.Controls.Add($btnReconcile)

    # Bouton Enregistrer (profil courant)
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = Get-Text "access_profiles.btn_save"
    $btnSave.Location = New-Object System.Drawing.Point(458, 460)
    $btnSave.Size = New-Object System.Drawing.Size(100, 32)
    $btnSave.FlatStyle = "Flat"
    $btnSave.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $btnSave.ForeColor = [System.Drawing.Color]::White
    $btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $pnlEdit.Controls.Add($btnSave)

    # ---- BOUTON FERMER (bas du form) ----
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = Get-Text "modification.btn_close"
    $btnClose.Location = New-Object System.Drawing.Point(680, 525)
    $btnClose.Size = New-Object System.Drawing.Size(110, 35)
    $btnClose.FlatStyle = "Flat"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $f.Controls.Add($btnClose)

    # ================================================================
    #  DONNÉES INTERNES
    # ================================================================
    $script:APSearchMap = @{}  # display_name → id (résultats de recherche)
    $script:APCurrentKey = $null  # Clé du profil actuellement sélectionné

    # ================================================================
    #  FONCTIONS INTERNES
    # ================================================================

    # Rafraîchir la ListBox des profils
    $RefreshList = {
        $lstProfiles.Items.Clear()
        foreach ($key in ($script:APData.Keys | Sort-Object)) {
            $item = $script:APData[$key]
            $marker = if ($item.is_baseline) { " [B]" } else { "" }
            $lstProfiles.Items.Add("$($item.display_name)$marker") | Out-Null
        }
    }

    # Charger un profil dans le panneau de droite
    $LoadProfile = {
        param([string]$Key)
        if (-not $script:APData.ContainsKey($Key)) { return }
        $script:APCurrentKey = $Key
        $item = $script:APData[$Key]

        $txtName.Text = $item.display_name
        $txtDesc.Text = $item.description
        $chkBaseline.Checked = $item.is_baseline

        # Charger les groupes assignés
        $lstGroups.Items.Clear()
        foreach ($grp in $item.groups) {
            $lstGroups.Items.Add($grp.display_name) | Out-Null
        }

        # Réinitialiser la recherche
        $lstSearchResults.Items.Clear()
        $script:APSearchMap = @{}
        $btnRemoveGrp.Enabled = $false
        $btnAddGrp.Enabled = $false
    }

    # Sauvegarder le profil courant dans $script:APData
    $SaveCurrentToData = {
        $key = $script:APCurrentKey
        if (-not $key -or -not $script:APData.ContainsKey($key)) { return $false }

        # Gérer le baseline : un seul autorisé
        if ($chkBaseline.Checked) {
            foreach ($k in $script:APData.Keys) {
                if ($k -ne $key) { $script:APData[$k].is_baseline = $false }
            }
        }

        $script:APData[$key].display_name = $txtName.Text.Trim()
        $script:APData[$key].description = $txtDesc.Text.Trim()
        $script:APData[$key].is_baseline = $chkBaseline.Checked

        # Les groupes sont déjà mis à jour en temps réel via ajout/retrait
        $script:APDirty = $true
        return $true
    }

    # Persister $script:APData dans le JSON client
    $PersistToJson = {
        try {
            $configPath = $Config._config_path
            $rawJson = Get-Content -Path $configPath -Raw -Encoding UTF8
            $configObj = $rawJson | ConvertFrom-Json

            # Reconstruire la section access_profiles
            $newProfiles = [ordered]@{}
            foreach ($key in ($script:APData.Keys | Sort-Object)) {
                $item = $script:APData[$key]
                $groupsList = @($item.groups | ForEach-Object {
                    [ordered]@{ id = $_.id; display_name = $_.display_name }
                })
                $newProfiles[$key] = [ordered]@{
                    display_name = $item.display_name
                    description  = $item.description
                    is_baseline  = $item.is_baseline
                    groups       = $groupsList
                }
            }

            # Remplacer la section dans l'objet config
            $configObj.access_profiles = $newProfiles

            # Sauvegarder — Depth 10 pour ne pas tronquer les groupes
            $configObj | ConvertTo-Json -Depth 10 | Out-File -FilePath $configPath -Encoding UTF8 -Force

            # Recharger dans $Config
            $global:Config = Load-ClientConfig -ConfigPath $configPath

            Write-Log -Level "SUCCESS" -Action "ACCESS_PROFILES" -Message "Profils d'accès sauvegardés dans $configPath"
            return $true
        }
        catch {
            Write-Log -Level "ERROR" -Action "ACCESS_PROFILES" -Message "Erreur sauvegarde profils : $($_.Exception.Message)"
            return $false
        }
    }

    # ================================================================
    #  ÉVÉNEMENTS
    # ================================================================

    # Charger la liste initiale
    & $RefreshList

    # Sélection dans la liste
    $lstProfiles.Add_SelectedIndexChanged({
        if ($lstProfiles.SelectedIndex -lt 0) { return }
        $selectedDisplay = $lstProfiles.SelectedItem.ToString() -replace ' \[B\]$', ''
        $foundKey = $script:APData.Keys | Where-Object { $script:APData[$_].display_name -eq $selectedDisplay } | Select-Object -First 1
        if ($foundKey) { & $LoadProfile $foundKey }
    })

    # Activer/désactiver boutons selon sélection
    $lstGroups.Add_SelectedIndexChanged({ $btnRemoveGrp.Enabled = $lstGroups.SelectedItems.Count -gt 0 })
    $lstSearchResults.Add_SelectedIndexChanged({ $btnAddGrp.Enabled = $lstSearchResults.SelectedItems.Count -gt 0 })

    # Recherche de groupes Graph
    $DoSearch = {
        $term = $txtSearch.Text.Trim()
        if ($term.Length -lt 3) { return }

        $lstSearchResults.Items.Clear()
        $script:APSearchMap = @{}

        $result = Search-AzGroups -SearchTerm $term
        if ($result.Success) {
            # Exclure les groupes déjà assignés au profil courant
            $assignedIds = @()
            if ($script:APCurrentKey -and $script:APData.ContainsKey($script:APCurrentKey)) {
                $assignedIds = @($script:APData[$script:APCurrentKey].groups | ForEach-Object { $_.id })
            }
            foreach ($g in $result.Data) {
                if ($g.Id -notin $assignedIds) {
                    $script:APSearchMap[$g.DisplayName] = $g.Id
                    $lstSearchResults.Items.Add($g.DisplayName) | Out-Null
                }
            }
            $lblHint.Text = "$($lstSearchResults.Items.Count) groupe(s) trouvé(s)"
        }
        else {
            $lblHint.Text = "Erreur : $($result.Error)"
        }
    }
    $btnSearch.Add_Click({ & $DoSearch })
    $txtSearch.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { & $DoSearch; $_.SuppressKeyPress = $true }
    })

    # Ajouter des groupes au profil
    $btnAddGrp.Add_Click({
        $key = $script:APCurrentKey
        if (-not $key -or -not $script:APData.ContainsKey($key)) { return }
        $sel = @($lstSearchResults.SelectedItems)
        foreach ($gName in $sel) {
            $gId = $script:APSearchMap[$gName]
            if (-not $gId) { continue }
            # Vérifier doublon
            $exists = $script:APData[$key].groups | Where-Object { $_.id -eq $gId }
            if (-not $exists) {
                $script:APData[$key].groups += @{ id = $gId; display_name = $gName }
                $lstGroups.Items.Add($gName) | Out-Null
            }
        }
        $script:APDirty = $true
        # Relancer la recherche pour exclure les nouvellement ajoutés
        & $DoSearch
    })

    # Retirer des groupes du profil
    $btnRemoveGrp.Add_Click({
        $key = $script:APCurrentKey
        if (-not $key -or -not $script:APData.ContainsKey($key)) { return }
        $sel = @($lstGroups.SelectedItems)
        foreach ($gName in $sel) {
            $script:APData[$key].groups = @($script:APData[$key].groups | Where-Object { $_.display_name -ne $gName })
        }
        $script:APDirty = $true
        # Rafraîchir la liste
        $lstGroups.Items.Clear()
        foreach ($grp in $script:APData[$key].groups) {
            $lstGroups.Items.Add($grp.display_name) | Out-Null
        }
        $btnRemoveGrp.Enabled = $false
    })

    # Nouveau profil
    $btnNew.Add_Click({
        $inputForm = New-Object System.Windows.Forms.Form
        $inputForm.Text = Get-Text "access_profiles.new_prompt_title"
        $inputForm.Size = New-Object System.Drawing.Size(400, 160)
        $inputForm.StartPosition = "CenterParent"
        $inputForm.FormBorderStyle = "FixedDialog"
        $inputForm.MaximizeBox = $false; $inputForm.MinimizeBox = $false

        $lblP = New-Object System.Windows.Forms.Label
        $lblP.Text = Get-Text "access_profiles.new_prompt_msg"
        $lblP.Location = New-Object System.Drawing.Point(10, 15)
        $lblP.Size = New-Object System.Drawing.Size(370, 20)
        $inputForm.Controls.Add($lblP)

        $txtP = New-Object System.Windows.Forms.TextBox
        $txtP.Location = New-Object System.Drawing.Point(10, 40)
        $txtP.Size = New-Object System.Drawing.Size(360, 25)
        $inputForm.Controls.Add($txtP)

        $btnOk = New-Object System.Windows.Forms.Button
        $btnOk.Text = "OK"; $btnOk.Location = New-Object System.Drawing.Point(200, 78)
        $btnOk.Size = New-Object System.Drawing.Size(80, 30)
        $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $inputForm.Controls.Add($btnOk)
        $inputForm.AcceptButton = $btnOk

        if ($inputForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $newKey = $txtP.Text.Trim() -replace '[^a-zA-Z0-9_\-]', ''
            if ($newKey -eq '') { return }
            if ($script:APData.ContainsKey($newKey)) {
                Show-ResultDialog -Titre (Get-Text "access_profiles.new_prompt_title") `
                    -Message (Get-Text "access_profiles.new_duplicate_error") -IsSuccess $false
                return
            }
            $script:APData[$newKey] = @{
                display_name = $newKey
                description  = ""
                is_baseline  = $false
                groups       = @()
            }
            $script:APDirty = $true
            & $RefreshList
        }
        $inputForm.Dispose()
    })

    # Supprimer un profil
    $btnDelete.Add_Click({
        $key = $script:APCurrentKey
        if (-not $key -or -not $script:APData.ContainsKey($key)) { return }
        if ($script:APData[$key].is_baseline) {
            Show-ResultDialog -Titre (Get-Text "access_profiles.delete_confirm_title") `
                -Message (Get-Text "access_profiles.delete_baseline_error") -IsSuccess $false
            return
        }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "access_profiles.delete_confirm_title") `
            -Message (Get-Text "access_profiles.delete_confirm_msg" $script:APData[$key].display_name)
        if ($confirm) {
            $script:APData.Remove($key)
            $script:APCurrentKey = $null
            $script:APDirty = $true
            $txtName.Text = ""; $txtDesc.Text = ""; $chkBaseline.Checked = $false
            $lstGroups.Items.Clear(); $lstSearchResults.Items.Clear()
            & $RefreshList
        }
    })

    # Réconcilier le profil sélectionné
    $btnReconcile.Add_Click({
        $key = $script:APCurrentKey
        if (-not $key) { return }

        # Sauvegarder d'abord si des modifications sont en cours
        if ($script:APDirty) {
            $saved = & $SaveCurrentToData
            if ($saved) { $null = & $PersistToJson }
        }

        # Lancer la réconciliation
        Show-ProfileReconciliation -ProfileKey $key
    })

    # Sauvegarder (profil courant → données → JSON)
    $btnSave.Add_Click({
        $saved = & $SaveCurrentToData
        if ($saved) {
            $persisted = & $PersistToJson
            if ($persisted) {
                $script:APDirty = $false
                Show-ResultDialog -Titre (Get-Text "access_profiles.title") `
                    -Message (Get-Text "access_profiles.save_success") -IsSuccess $true
                & $RefreshList
            }
            else {
                Show-ResultDialog -Titre (Get-Text "access_profiles.title") `
                    -Message (Get-Text "access_profiles.save_error" "Voir les logs.") -IsSuccess $false
            }
        }
    })

    # Avertir si fermeture avec modifications non sauvegardées
    $f.Add_FormClosing({
        if ($script:APDirty) {
            $confirm = Show-ConfirmDialog -Titre (Get-Text "access_profiles.title") `
                -Message (Get-Text "access_profiles.unsaved_warning")
            if (-not $confirm) { $_.Cancel = $true }
        }
    })

    $f.ShowDialog() | Out-Null
    $f.Dispose()
}

<#
.FICHIER
    Fonction Show-ProfileReconciliation — à insérer dans GUI_AccessProfiles.ps1

.ROLE
    Formulaire modal de réconciliation d'un profil d'accès.
    Affiche les utilisateurs avec des écarts par rapport au template,
    permet la réconciliation en masse et l'export CSV.

.DEPENDANCES
    - Get-ProfileReconciliation   (AccessProfiles_Functions.ps1)
    - Invoke-ProfileReconciliation (AccessProfiles_Functions.ps1)
    - Get-Text, Write-Log, Show-ResultDialog
#>

function Show-ProfileReconciliation {
    <#
    .SYNOPSIS
        Affiche le formulaire de réconciliation pour un profil donné.
        Scanne les utilisateurs, affiche les écarts, permet la correction.
    .PARAMETER ProfileKey
        Clé technique du profil à réconcilier.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProfileKey
    )

    # ============================================================
    #  1. Lancer le scan
    # ============================================================
    $scanResult = Get-ProfileReconciliation -ProfileKey $ProfileKey

    # Erreur de scan
    if ($scanResult.Error) {
        Show-ResultDialog -Titre (Get-Text "access_profiles.reconcile_title" $ProfileKey) `
            -Message $scanResult.Error -IsSuccess $false
        return
    }

    # Profil vide
    if ($scanResult.TotalGroups -eq 0) {
        Show-ResultDialog -Titre (Get-Text "access_profiles.reconcile_title" $scanResult.ProfileName) `
            -Message (Get-Text "access_profiles.groups_empty") -IsSuccess $false
        return
    }

    # Aucun écart
    if ($scanResult.Discrepancies.Count -eq 0) {
        Show-ResultDialog -Titre (Get-Text "access_profiles.reconcile_title" $scanResult.ProfileName) `
            -Message (Get-Text "access_profiles.reconcile_none") -IsSuccess $true
        return
    }

    # ============================================================
    #  2. Construire le formulaire
    # ============================================================
    $f = New-Object System.Windows.Forms.Form
    $f.Text = Get-Text "access_profiles.reconcile_title" $scanResult.ProfileName
    $f.Size = New-Object System.Drawing.Size(750, 520)
    $f.StartPosition = "CenterScreen"
    $f.FormBorderStyle = "FixedDialog"
    $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.BackColor = [System.Drawing.Color]::WhiteSmoke
    $f.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # Bandeau d'avertissement
    $pnlWarn = New-Object System.Windows.Forms.Panel
    $pnlWarn.Location = New-Object System.Drawing.Point(10, 8)
    $pnlWarn.Size = New-Object System.Drawing.Size(715, 50)
    $pnlWarn.BackColor = [System.Drawing.Color]::FromArgb(255, 243, 205)
    $pnlWarn.BorderStyle = "FixedSingle"
    $f.Controls.Add($pnlWarn)

    $lblWarn = New-Object System.Windows.Forms.Label
    $lblWarn.Text = Get-Text "access_profiles.reconcile_desc" $scanResult.Discrepancies.Count
    $lblWarn.Location = New-Object System.Drawing.Point(10, 4)
    $lblWarn.Size = New-Object System.Drawing.Size(690, 40)
    $lblWarn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $pnlWarn.Controls.Add($lblWarn)

    # Info complémentaire
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = Get-Text "access_profiles.reconcile_info" $scanResult.TotalUsers $scanResult.TotalGroups
    $lblInfo.Location = New-Object System.Drawing.Point(10, 64)
    $lblInfo.Size = New-Object System.Drawing.Size(715, 18)
    $lblInfo.ForeColor = [System.Drawing.Color]::Gray
    $lblInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $f.Controls.Add($lblInfo)

    # ============================================================
    #  3. DataGridView
    # ============================================================
    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Location = New-Object System.Drawing.Point(10, 86)
    $dgv.Size = New-Object System.Drawing.Size(715, 330)
    $dgv.AllowUserToAddRows = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.AllowUserToResizeRows = $false
    $dgv.ReadOnly = $true
    $dgv.SelectionMode = "FullRowSelect"
    $dgv.MultiSelect = $false
    $dgv.RowHeadersVisible = $false
    $dgv.AutoSizeColumnsMode = "Fill"
    $dgv.BackgroundColor = [System.Drawing.Color]::White
    $dgv.BorderStyle = "FixedSingle"
    $dgv.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dgv.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dgv.EnableHeadersVisualStyles = $false
    $dgv.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(0, 90, 180)
    $dgv.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $dgv.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    $f.Controls.Add($dgv)

    # Colonnes
    $colUPN = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colUPN.Name = "UPN"
    $colUPN.HeaderText = Get-Text "access_profiles.reconcile_col_user"
    $colUPN.FillWeight = 45
    $dgv.Columns.Add($colUPN) | Out-Null

    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.Name = "DisplayName"
    $colName.HeaderText = Get-Text "access_profiles.reconcile_col_name"
    $colName.FillWeight = 20
    $dgv.Columns.Add($colName) | Out-Null

    $colMissing = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMissing.Name = "Missing"
    $colMissing.HeaderText = Get-Text "access_profiles.reconcile_col_missing"
    $colMissing.FillWeight = 35
    $dgv.Columns.Add($colMissing) | Out-Null

    # Remplir les données
    foreach ($disc in $scanResult.Discrepancies) {
        $missingStr = $disc.Missing -join ", "
        $dgv.Rows.Add($disc.UPN, $disc.DisplayName, $missingStr) | Out-Null
    }

    # Coloration des lignes
    foreach ($row in $dgv.Rows) {
        $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 30, 30)
    }

    # ============================================================
    #  4. Boutons d'action
    # ============================================================

    # Bouton Réconcilier tous
    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = Get-Text "access_profiles.reconcile_btn_apply"
    $btnApply.Location = New-Object System.Drawing.Point(10, 425)
    $btnApply.Size = New-Object System.Drawing.Size(180, 35)
    $btnApply.FlatStyle = "Flat"
    $btnApply.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $btnApply.ForeColor = [System.Drawing.Color]::White
    $btnApply.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($btnApply)

    # Bouton Exporter
    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = Get-Text "access_profiles.reconcile_btn_export"
    $btnExport.Location = New-Object System.Drawing.Point(200, 425)
    $btnExport.Size = New-Object System.Drawing.Size(180, 35)
    $btnExport.FlatStyle = "Flat"
    $f.Controls.Add($btnExport)

    # Bouton Fermer
    $btnFermer = New-Object System.Windows.Forms.Button
    $btnFermer.Text = Get-Text "modification.btn_close"
    $btnFermer.Location = New-Object System.Drawing.Point(615, 425)
    $btnFermer.Size = New-Object System.Drawing.Size(110, 35)
    $btnFermer.FlatStyle = "Flat"
    $btnFermer.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $f.Controls.Add($btnFermer)

    # Label de progression
    $lblProgress = New-Object System.Windows.Forms.Label
    $lblProgress.Location = New-Object System.Drawing.Point(10, 465)
    $lblProgress.Size = New-Object System.Drawing.Size(715, 18)
    $lblProgress.ForeColor = [System.Drawing.Color]::Gray
    $lblProgress.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $f.Controls.Add($lblProgress)

    # ============================================================
    #  5. Événements
    # ============================================================

    # Réconcilier tous
    $btnApply.Add_Click({
        $count = $scanResult.Discrepancies.Count
        $confirm = Show-ConfirmDialog `
            -Titre (Get-Text "access_profiles.reconcile_title" $scanResult.ProfileName) `
            -Message (Get-Text "access_profiles.reconcile_confirm" $count)

        if (-not $confirm) { return }

        # Désactiver les boutons pendant l'opération
        $btnApply.Enabled = $false
        $btnExport.Enabled = $false

        $progressCallback = {
            param($current, $total, $upn)
            $lblProgress.Text = Get-Text "access_profiles.reconcile_progress" $current $total $upn
            $f.Refresh()
        }

        $result = Invoke-ProfileReconciliation `
            -Discrepancies $scanResult.Discrepancies `
            -ProfileKey $scanResult.ProfileKey `
            -OnProgress $progressCallback

        # Mettre à jour la grille avec les résultats
        if ($result.Success) {
            $lblProgress.Text = ""
            # Colorer toutes les lignes en vert (réconciliées)
            foreach ($row in $dgv.Rows) {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
            }
            Show-ResultDialog -Titre (Get-Text "access_profiles.reconcile_title" $scanResult.ProfileName) `
                -Message (Get-Text "access_profiles.reconcile_success" $result.Applied) -IsSuccess $true
        }
        else {
            $lblProgress.Text = (Get-Text "access_profiles.reconcile_partial" $result.Applied $result.Failed)
            $errDetail = $result.Errors -join "`n"
            Show-ResultDialog -Titre (Get-Text "access_profiles.reconcile_title" $scanResult.ProfileName) `
                -Message ((Get-Text "access_profiles.reconcile_partial" $result.Applied $result.Failed) + "`n`n$errDetail") `
                -IsSuccess $false
        }

        # Réactiver
        $btnApply.Enabled = $true
        $btnExport.Enabled = $true
    })

    # Exporter CSV
    $btnExport.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV (*.csv)|*.csv"
        $sfd.FileName = "Reconciliation_$($scanResult.ProfileKey)_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $sfd.Title = Get-Text "access_profiles.reconcile_btn_export"

        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $csvData = @()
                foreach ($disc in $scanResult.Discrepancies) {
                    $csvData += [PSCustomObject]@{
                        UPN             = $disc.UPN
                        DisplayName     = $disc.DisplayName
                        MissingGroups   = ($disc.Missing -join "; ")
                        MissingGroupIds = ($disc.MissingIds -join "; ")
                        ProfileKey      = $scanResult.ProfileKey
                        ProfileName     = $scanResult.ProfileName
                    }
                }
                $csvData | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
                Write-Log -Level "SUCCESS" -Action "RECONCILE_EXPORT" `
                    -Message "Export CSV réconciliation : $($sfd.FileName) ($($csvData.Count) lignes)"
                $lblProgress.Text = "Exporté : $($sfd.FileName)"
            }
            catch {
                Write-Log -Level "ERROR" -Action "RECONCILE_EXPORT" `
                    -Message "Erreur export CSV : $($_.Exception.Message)"
                Show-ResultDialog -Titre (Get-Text "access_profiles.reconcile_btn_export") `
                    -Message $_.Exception.Message -IsSuccess $false
            }
        }
        $sfd.Dispose()
    })

    # ============================================================
    #  6. Afficher
    # ============================================================
    $f.ShowDialog() | Out-Null
    $f.Dispose()
}