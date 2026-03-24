<#
.FICHIER
    Modules/GUI_PIMManager.ps1

.ROLE
    Interface WinForms du gestionnaire PIM (Privileged Identity Management).
    Phase 2 : dashboard avec cartes de statut colorees, editeur de groupes PIM,
    selecteur de roles Entra ID avec filtre, sauvegarde dans le JSON client.
    La logique metier est dans Core/PIMFunctions.ps1.

.DEPENDANCES
    - Core/PIMFunctions.ps1 (Initialize-PimConfig, Import-PimData, Load-EntraRoles,
                             Invoke-PimAudit, Invoke-PimUpdate, New-PimGroupEntra,
                             Add-PimRolesToGroup, Export-PimCsvReport, Save-PimConfig,
                             Remove-PimGroupFromConfig, Import-PimGroupsFromTenant)
    - Core/Functions.ps1    (Write-Log, Show-ConfirmDialog, Show-ResultDialog)
    - Core/Lang.ps1         (Get-Text)
    - Variable globale      $Config, $RootPath

.AUTEUR
    [Equipe IT -- M365 Monster]
#>

function Show-PIMManagerForm {
    <#
    .SYNOPSIS
        Affiche le formulaire de gestion PIM pour le client courant.
    #>

    # ================================================================
    #  INITIALISATION
    # ================================================================
    if (-not (Initialize-PimConfig)) {
        Show-ResultDialog -Titre (Get-Text "pim.title") `
            -Message (Get-Text "pim.no_config_section") -IsSuccess $false
        return
    }
    Import-PimData
    $script:PimDirty = $false   # Flag modifications non sauvegardees

    # ================================================================
    #  COULEURS PARTAGEES
    # ================================================================
    $clrOK      = [System.Drawing.Color]::FromArgb(32, 201, 151)
    $clrDrift   = [System.Drawing.Color]::FromArgb(253, 126, 20)
    $clrMissing = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $clrSuccess = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $clrSkipped = [System.Drawing.Color]::FromArgb(108, 117, 125)
    $clrPending = [System.Drawing.Color]::FromArgb(173, 181, 189)
    $clrHeader  = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $clrBg      = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $clrWhite   = [System.Drawing.Color]::White

    function Get-StatusColor {
        param([string]$Status)
        switch ($Status) {
            'OK'      { return $clrOK }
            'Drift'   { return $clrDrift }
            'Missing' { return $clrMissing }
            'Error'   { return $clrMissing }
            'Success' { return $clrSuccess }
            'Skipped' { return $clrSkipped }
            default   { return $clrPending }
        }
    }

    # ================================================================
    #  HELPERS UI
    # ================================================================

    function Build-GroupCard {
        <#
        .SYNOPSIS
            Cree une carte visuelle pour un groupe PIM dans le panneau gauche.
            Bordure coloree a gauche selon le statut. Clic = selection.
        #>
        param([string]$GroupName, [int]$YPos)

        $def    = $script:PimData[$GroupName]
        $status = $script:PimGroupStatus[$GroupName]

        $card = New-Object System.Windows.Forms.Panel
        $card.Location = New-Object System.Drawing.Point(0, $YPos)
        $card.Size = New-Object System.Drawing.Size(310, 62)
        $card.BackColor = $clrWhite
        $card.Tag = $GroupName
        $card.Cursor = [System.Windows.Forms.Cursors]::Hand

        # Bordure coloree a gauche
        $border = New-Object System.Windows.Forms.Panel
        $border.Location = New-Object System.Drawing.Point(0, 0)
        $border.Size = New-Object System.Drawing.Size(5, 62)
        $border.BackColor = Get-StatusColor -Status $status
        $border.Tag = $GroupName
        $card.Controls.Add($border)

        # Pastille de statut
        $dot = New-Object System.Windows.Forms.Panel
        $dot.Location = New-Object System.Drawing.Point(14, 10)
        $dot.Size = New-Object System.Drawing.Size(10, 10)
        $dot.BackColor = Get-StatusColor -Status $status
        $dot.Tag = $GroupName
        $card.Controls.Add($dot)
        # Arrondir la pastille
        $dotPath = New-Object System.Drawing.Drawing2D.GraphicsPath
        $dotPath.AddEllipse(0, 0, 10, 10)
        $dot.Region = New-Object System.Drawing.Region($dotPath)

        # Nom du groupe
        $lblName = New-Object System.Windows.Forms.Label
        $lblName.Text = $GroupName
        $lblName.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $lblName.ForeColor = $clrHeader
        $lblName.Location = New-Object System.Drawing.Point(30, 6)
        $lblName.Size = New-Object System.Drawing.Size(270, 20)
        $lblName.Tag = $GroupName
        $card.Controls.Add($lblName)

        # Sous-titre : type + roles
        $typeLabel = $def.Type
        $rolesCount = $def.Roles.Count
        $statusText = if ($status -ne 'Pending') { " -- $status" } else { '' }
        $lblSub = New-Object System.Windows.Forms.Label
        $lblSub.Text = "$typeLabel | $rolesCount role(s)$statusText"
        $lblSub.Font = New-Object System.Drawing.Font("Segoe UI", 8)
        $lblSub.ForeColor = $clrSkipped
        $lblSub.Location = New-Object System.Drawing.Point(30, 28)
        $lblSub.Size = New-Object System.Drawing.Size(270, 16)
        $lblSub.Tag = $GroupName
        $card.Controls.Add($lblSub)

        # Description tronquee
        $descText = if ($def.Description.Length -gt 50) { $def.Description.Substring(0, 47) + '...' } else { $def.Description }
        $lblDesc = New-Object System.Windows.Forms.Label
        $lblDesc.Text = $descText
        $lblDesc.Font = New-Object System.Drawing.Font("Segoe UI", 7.5)
        $lblDesc.ForeColor = [System.Drawing.Color]::FromArgb(160, 160, 160)
        $lblDesc.Location = New-Object System.Drawing.Point(30, 44)
        $lblDesc.Size = New-Object System.Drawing.Size(270, 14)
        $lblDesc.Tag = $GroupName
        $card.Controls.Add($lblDesc)

        # Clic sur n'importe quel element de la carte
        $clickAction = {
            param($src, $evt)
            $gn = $src.Tag
            if (-not $gn) { return }
            $script:SelectedGroup = $gn
            Show-PimGroupDetail -GroupName $gn
            Update-PimCardSelection
        }
        $card.Add_Click($clickAction)
        foreach ($ctrl in $card.Controls) { $ctrl.Add_Click($clickAction) }

        return $card
    }

    function Update-PimDashboard {
        <#
        .SYNOPSIS
            Reconstruit toutes les cartes du dashboard gauche.
        #>
        $script:DashboardPanel.Controls.Clear()
        $y = 5
        foreach ($gName in $script:PimData.Keys) {
            $card = Build-GroupCard -GroupName $gName -YPos $y
            $script:DashboardPanel.Controls.Add($card)
            $y += 67  # 62 + 5 spacing
        }

        # Compteur resume en haut
        $okC   = @($script:PimGroupStatus.Values | Where-Object { $_ -eq 'OK' }).Count
        $driftC = @($script:PimGroupStatus.Values | Where-Object { $_ -eq 'Drift' }).Count
        $missC = @($script:PimGroupStatus.Values | Where-Object { $_ -eq 'Missing' -or $_ -eq 'Error' }).Count
        $pendC = @($script:PimGroupStatus.Values | Where-Object { $_ -eq 'Pending' }).Count
        $parts = @()
        if ($okC -gt 0)    { $parts += "$okC OK" }
        if ($driftC -gt 0) { $parts += "$driftC Drift" }
        if ($missC -gt 0)  { $parts += "$missC Missing" }
        if ($pendC -gt 0)  { $parts += "$pendC Pending" }
        $script:lblSummary.Text = ($parts -join " | ")
        if (-not $parts) { $script:lblSummary.Text = Get-Text "pim.no_groups" }
    }

    function Update-PimCardSelection {
        <#
        .SYNOPSIS
            Met en surbrillance la carte du groupe selectionne.
        #>
        foreach ($ctrl in $script:DashboardPanel.Controls) {
            if ($ctrl -is [System.Windows.Forms.Panel] -and $ctrl.Tag) {
                if ($ctrl.Tag -eq $script:SelectedGroup) {
                    $ctrl.BackColor = [System.Drawing.Color]::FromArgb(232, 240, 254)
                }
                else {
                    $ctrl.BackColor = $clrWhite
                }
            }
        }
    }

    function Show-PimGroupDetail {
        <#
        .SYNOPSIS
            Affiche le detail d'un groupe PIM dans le panneau droit (mode consultation).
        #>
        param([string]$GroupName)
        $def = $script:PimData[$GroupName]
        if (-not $def) { return }

        # Remplir les champs d'edition
        $script:txtEditName.Text   = $GroupName
        $script:txtEditName.Tag    = $GroupName  # Nom original pour la sauvegarde
        $script:txtEditDesc.Text   = $def.Description
        $script:cboEditType.SelectedItem = $def.Type

        # Remplir la liste des roles assignes
        $script:lstAssignedRoles.Items.Clear()
        foreach ($r in $def.Roles) {
            $item = New-Object System.Windows.Forms.ListViewItem($r)
            $isCustom = ($r -match '\[Custom\]')
            $item.SubItems.Add($(if ($isCustom) { "Custom" } else { "Built-in" }))
            if ($isCustom) { $item.ForeColor = $clrDrift }
            $script:lstAssignedRoles.Items.Add($item) | Out-Null
        }

        # Resultats d'audit
        $script:lstAuditResults.Items.Clear()
        $results = $script:PimAuditResults[$GroupName]
        if ($results -and $results.Count -gt 0) {
            $script:pnlAudit.Visible = $true
            foreach ($item in $results) {
                $lvi = New-Object System.Windows.Forms.ListViewItem($item.Text)
                $lvi.ForeColor = switch ($item.Status) {
                    'OK'    { $clrSuccess }
                    'Error' { $clrMissing }
                    'Warn'  { $clrDrift }
                    'Info'  { [System.Drawing.Color]::FromArgb(0, 123, 255) }
                    default { $clrSkipped }
                }
                $script:lstAuditResults.Items.Add($lvi) | Out-Null
            }
        }
        else {
            $script:pnlAudit.Visible = $false
        }

        # Statut
        $status = $script:PimGroupStatus[$GroupName]
        $script:lblDetailStatus.Text = $status
        $script:lblDetailStatus.ForeColor = Get-StatusColor -Status $status
    }

    function Set-PimButtonsEnabled {
        param([bool]$Enabled)
        $script:btnAudit.Enabled  = $Enabled
        $script:btnUpdate.Enabled = $Enabled
        $script:btnCreate.Enabled = $Enabled
        $script:btnSave.Enabled   = $Enabled
    }

    function Update-PimProgress {
        param([int]$Done, [int]$Total)
        $script:PimProgressBar.Maximum = $Total
        $script:PimProgressBar.Value   = $Done
        $script:PimProgressLabel.Text  = "$Done / $Total"
        [System.Windows.Forms.Application]::DoEvents()
    }

    function Get-CheckedGroupNames {
        <#
        .SYNOPSIS
            Retourne les noms des groupes dont la checkbox est cochee dans le dashboard.
        #>
        $names = @()
        foreach ($ctrl in $script:DashboardPanel.Controls) {
            if ($ctrl -is [System.Windows.Forms.Panel] -and $ctrl.Tag) {
                # Chercher la checkbox dans la carte
                $cb = $ctrl.Controls | Where-Object { $_ -is [System.Windows.Forms.CheckBox] }
                if ($cb -and $cb.Checked) {
                    $names += $ctrl.Tag
                }
            }
        }
        return $names
    }

    # ================================================================
    #  CONSTRUCTION DU FORMULAIRE
    # ================================================================
    $f = New-Object System.Windows.Forms.Form
    $f.Text = "$(Get-Text 'pim.title') -- $($Config.client_name)"
    $f.Size = New-Object System.Drawing.Size(1150, 800)
    $f.StartPosition = "CenterScreen"
    $f.FormBorderStyle = "Sizable"
    $f.MinimumSize = New-Object System.Drawing.Size(950, 650)
    $f.BackColor = $clrBg

    $iconFile = Join-Path -Path $script:RootPath -ChildPath "Assets\M365Monster.ico"
    if (Test-Path $iconFile) { $f.Icon = New-Object System.Drawing.Icon($iconFile) }

    # Avertissement modifications non sauvegardees
    $f.Add_FormClosing({
        if ($script:PimDirty) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                (Get-Text "pim.unsaved_warning"),
                (Get-Text "pim.title"),
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($result -eq [System.Windows.Forms.DialogResult]::No) {
                $_.Cancel = $true
            }
        }
    })

    # ── EN-TETE ─────────────────────────────────────────────────────
    $panelHeader = New-Object System.Windows.Forms.Panel
    $panelHeader.Dock = [System.Windows.Forms.DockStyle]::Top
    $panelHeader.Height = 70
    $panelHeader.BackColor = $clrHeader

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = Get-Text "pim.title"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $clrWhite
    $lblTitle.Location = New-Object System.Drawing.Point(20, 8)
    $lblTitle.Size = New-Object System.Drawing.Size(500, 30)
    $panelHeader.Controls.Add($lblTitle)

    $lblSubtitle = New-Object System.Windows.Forms.Label
    $lblSubtitle.Text = Get-Text "pim.subtitle"
    $lblSubtitle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(173, 181, 189)
    $lblSubtitle.Location = New-Object System.Drawing.Point(20, 38)
    $lblSubtitle.Size = New-Object System.Drawing.Size(700, 20)
    $panelHeader.Controls.Add($lblSubtitle)

    # ── BARRE D'ACTIONS (Dock Bottom) ───────────────────────────────
    $actionPanel = New-Object System.Windows.Forms.Panel
    $actionPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $actionPanel.Height = 100
    $actionPanel.BackColor = $clrBg

    # ── SPLITCONTAINER (Dock Fill) ──────────────────────────────────
    $split = New-Object System.Windows.Forms.SplitContainer
    $split.Dock = [System.Windows.Forms.DockStyle]::Fill
    $split.SplitterDistance = 340
    $split.FixedPanel = [System.Windows.Forms.FixedPanel]::Panel1
    $split.SplitterWidth = 6
    $split.BackColor = $clrBg

    # Ordre d'ajout WinForms : le DERNIER ajoute est resolu EN PREMIER.
    # On veut la resolution : Header Top (1er) -> ActionPanel Bottom (2e) -> Split Fill (3e)
    # Donc on ajoute dans l'ordre inverse : Split, ActionPanel, Header.
    $f.Controls.Add($split)
    $f.Controls.Add($actionPanel)
    $f.Controls.Add($panelHeader)

    # ================================================================
    #  PANNEAU GAUCHE — Dashboard avec cartes de groupes PIM
    #  Dock layout (Top + Fill) : la toolbar reste toujours visible,
    #  le DashboardPanel scroll independamment en dessous.
    # ================================================================
    $leftPanel = $split.Panel1

    # -- Barre d'outils (Dock Top) --
    $leftToolbar = New-Object System.Windows.Forms.Panel
    $leftToolbar.Dock = [System.Windows.Forms.DockStyle]::Top
    $leftToolbar.Height = 92
    $leftToolbar.BackColor = $clrBg

    # Compteur resume
    $script:lblSummary = New-Object System.Windows.Forms.Label
    $script:lblSummary.Text = ""
    $script:lblSummary.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $script:lblSummary.ForeColor = $clrSkipped
    $script:lblSummary.Location = New-Object System.Drawing.Point(10, 4)
    $script:lblSummary.Size = New-Object System.Drawing.Size(310, 20)
    $leftToolbar.Controls.Add($script:lblSummary)

    # Boutons Nouveau / Supprimer (ligne 1)
    $btnNewGroup = New-Object System.Windows.Forms.Button
    $btnNewGroup.Text = Get-Text "pim.btn_new_group"
    $btnNewGroup.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $btnNewGroup.Location = New-Object System.Drawing.Point(10, 28)
    $btnNewGroup.Size = New-Object System.Drawing.Size(148, 26)
    $btnNewGroup.FlatStyle = "Flat"
    $btnNewGroup.BackColor = $clrSuccess
    $btnNewGroup.ForeColor = $clrWhite
    $btnNewGroup.FlatAppearance.BorderSize = 0
    $btnNewGroup.Cursor = [System.Windows.Forms.Cursors]::Hand
    $leftToolbar.Controls.Add($btnNewGroup)

    $btnDeleteGroup = New-Object System.Windows.Forms.Button
    $btnDeleteGroup.Text = Get-Text "pim.btn_delete_group"
    $btnDeleteGroup.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $btnDeleteGroup.Location = New-Object System.Drawing.Point(163, 28)
    $btnDeleteGroup.Size = New-Object System.Drawing.Size(148, 26)
    $btnDeleteGroup.FlatStyle = "Flat"
    $btnDeleteGroup.ForeColor = $clrMissing
    $btnDeleteGroup.Cursor = [System.Windows.Forms.Cursors]::Hand
    $leftToolbar.Controls.Add($btnDeleteGroup)

    # Bouton Importer du tenant (ligne 2)
    $btnImportTenant = New-Object System.Windows.Forms.Button
    $btnImportTenant.Text = Get-Text "pim.btn_import_tenant"
    $btnImportTenant.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $btnImportTenant.Location = New-Object System.Drawing.Point(10, 58)
    $btnImportTenant.Size = New-Object System.Drawing.Size(301, 26)
    $btnImportTenant.FlatStyle = "Flat"
    $btnImportTenant.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $btnImportTenant.ForeColor = $clrWhite
    $btnImportTenant.FlatAppearance.BorderSize = 0
    $btnImportTenant.Cursor = [System.Windows.Forms.Cursors]::Hand
    $leftToolbar.Controls.Add($btnImportTenant)

    # -- ScrollPanel pour les cartes (Dock Fill) --
    $script:DashboardPanel = New-Object System.Windows.Forms.Panel
    $script:DashboardPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:DashboardPanel.AutoScroll = $true
    $script:DashboardPanel.Padding = New-Object System.Windows.Forms.Padding(0, 5, 0, 5)

    # Ordre d'ajout WinForms : le DERNIER ajoute est resolu EN PREMIER pour le Dock.
    # On veut : Toolbar Top resolu d'abord, DashboardPanel Fill ensuite.
    # Donc : DashboardPanel (Fill) en premier, Toolbar (Top) en dernier.
    $leftPanel.Controls.Add($script:DashboardPanel)
    $leftPanel.Controls.Add($leftToolbar)

    # ================================================================
    #  PANNEAU DROIT — Editeur de groupe PIM
    # ================================================================
    $rightPanel = $split.Panel2
    $rightPanel.BackColor = $clrBg

    # Panel scrollable pour tout le contenu droit
    $script:rightContent = New-Object System.Windows.Forms.Panel
    $script:rightContent.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:rightContent.AutoScroll = $true
    $script:rightContent.BackColor = $clrBg
    $rightPanel.Controls.Add($script:rightContent)

    # -- Nom du groupe --
    $lblNameH = New-Object System.Windows.Forms.Label
    $lblNameH.Text = Get-Text "pim.edit_name"
    $lblNameH.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblNameH.ForeColor = $clrSkipped
    $lblNameH.Location = New-Object System.Drawing.Point(10, 10)
    $lblNameH.Size = New-Object System.Drawing.Size(100, 20)
    $script:rightContent.Controls.Add($lblNameH)

    $script:txtEditName = New-Object System.Windows.Forms.TextBox
    $script:txtEditName.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $script:txtEditName.Location = New-Object System.Drawing.Point(115, 6)
    $script:txtEditName.Size = New-Object System.Drawing.Size(400, 28)
    $script:txtEditName.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $script:rightContent.Controls.Add($script:txtEditName)

    # Statut (a droite du nom)
    $script:lblDetailStatus = New-Object System.Windows.Forms.Label
    $script:lblDetailStatus.Text = ""
    $script:lblDetailStatus.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $script:lblDetailStatus.Location = New-Object System.Drawing.Point(525, 10)
    $script:lblDetailStatus.Size = New-Object System.Drawing.Size(150, 22)
    $script:lblDetailStatus.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $script:rightContent.Controls.Add($script:lblDetailStatus)

    # -- Description --
    $lblDescH = New-Object System.Windows.Forms.Label
    $lblDescH.Text = Get-Text "pim.edit_desc"
    $lblDescH.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblDescH.ForeColor = $clrSkipped
    $lblDescH.Location = New-Object System.Drawing.Point(10, 40)
    $lblDescH.Size = New-Object System.Drawing.Size(100, 20)
    $script:rightContent.Controls.Add($lblDescH)

    $script:txtEditDesc = New-Object System.Windows.Forms.TextBox
    $script:txtEditDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:txtEditDesc.Location = New-Object System.Drawing.Point(115, 38)
    $script:txtEditDesc.Size = New-Object System.Drawing.Size(560, 25)
    $script:txtEditDesc.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $script:rightContent.Controls.Add($script:txtEditDesc)

    # -- Type (ComboBox) --
    $lblTypeH = New-Object System.Windows.Forms.Label
    $lblTypeH.Text = "Type"
    $lblTypeH.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblTypeH.ForeColor = $clrSkipped
    $lblTypeH.Location = New-Object System.Drawing.Point(10, 70)
    $lblTypeH.Size = New-Object System.Drawing.Size(100, 20)
    $script:rightContent.Controls.Add($lblTypeH)

    $script:cboEditType = New-Object System.Windows.Forms.ComboBox
    $script:cboEditType.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:cboEditType.Location = New-Object System.Drawing.Point(115, 68)
    $script:cboEditType.Size = New-Object System.Drawing.Size(200, 25)
    $script:cboEditType.DropDownStyle = "DropDownList"
    @('Role_Fixe', 'Groupe', 'Groupe_Critical', 'Role') | ForEach-Object { $script:cboEditType.Items.Add($_) | Out-Null }
    $script:cboEditType.SelectedIndex = 0
    $script:rightContent.Controls.Add($script:cboEditType)

    # -- Roles assignes --
    $lblRolesH = New-Object System.Windows.Forms.Label
    $lblRolesH.Text = Get-Text "pim.roles_assigned"
    $lblRolesH.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblRolesH.Location = New-Object System.Drawing.Point(10, 100)
    $lblRolesH.Size = New-Object System.Drawing.Size(200, 22)
    $script:rightContent.Controls.Add($lblRolesH)

    $script:lstAssignedRoles = New-Object System.Windows.Forms.ListView
    $script:lstAssignedRoles.Location = New-Object System.Drawing.Point(10, 124)
    $script:lstAssignedRoles.Size = New-Object System.Drawing.Size(665, 150)
    $script:lstAssignedRoles.View = [System.Windows.Forms.View]::Details
    $script:lstAssignedRoles.FullRowSelect = $true
    $script:lstAssignedRoles.GridLines = $true
    $script:lstAssignedRoles.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:lstAssignedRoles.Columns.Add((Get-Text "pim.col_role_name"), 500) | Out-Null
    $script:lstAssignedRoles.Columns.Add((Get-Text "pim.col_role_type"), 145) | Out-Null
    $script:lstAssignedRoles.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
        [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $script:rightContent.Controls.Add($script:lstAssignedRoles)

    # Bouton retirer le role selectionne
    $btnRemoveRole = New-Object System.Windows.Forms.Button
    $btnRemoveRole.Text = Get-Text "pim.btn_remove_role"
    $btnRemoveRole.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $btnRemoveRole.Location = New-Object System.Drawing.Point(10, 278)
    $btnRemoveRole.Size = New-Object System.Drawing.Size(130, 26)
    $btnRemoveRole.FlatStyle = "Flat"
    $btnRemoveRole.ForeColor = $clrMissing
    $btnRemoveRole.Cursor = [System.Windows.Forms.Cursors]::Hand
    $script:rightContent.Controls.Add($btnRemoveRole)

    # -- Ajout de role via ComboBox filtree --
    $lblAddRole = New-Object System.Windows.Forms.Label
    $lblAddRole.Text = Get-Text "pim.add_role_label"
    $lblAddRole.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblAddRole.Location = New-Object System.Drawing.Point(10, 312)
    $lblAddRole.Size = New-Object System.Drawing.Size(300, 22)
    $script:rightContent.Controls.Add($lblAddRole)

    # ComboBox avec filtre integre (DropDown mode = editable)
    $script:cboRoleSearch = New-Object System.Windows.Forms.ComboBox
    $script:cboRoleSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:cboRoleSearch.Location = New-Object System.Drawing.Point(10, 336)
    $script:cboRoleSearch.Size = New-Object System.Drawing.Size(530, 25)
    $script:cboRoleSearch.DropDownStyle = "DropDown"
    $script:cboRoleSearch.AutoCompleteMode = "SuggestAppend"
    $script:cboRoleSearch.AutoCompleteSource = "ListItems"
    $script:cboRoleSearch.Sorted = $true
    $script:cboRoleSearch.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
        [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $script:rightContent.Controls.Add($script:cboRoleSearch)

    $btnAddRole = New-Object System.Windows.Forms.Button
    $btnAddRole.Text = Get-Text "pim.btn_add_role"
    $btnAddRole.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $btnAddRole.Location = New-Object System.Drawing.Point(548, 335)
    $btnAddRole.Size = New-Object System.Drawing.Size(127, 27)
    $btnAddRole.FlatStyle = "Flat"
    $btnAddRole.BackColor = $clrSuccess
    $btnAddRole.ForeColor = $clrWhite
    $btnAddRole.FlatAppearance.BorderSize = 0
    $btnAddRole.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnAddRole.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $script:rightContent.Controls.Add($btnAddRole)

    # -- Panneau d'audit --
    $script:pnlAudit = New-Object System.Windows.Forms.Panel
    $script:pnlAudit.Location = New-Object System.Drawing.Point(10, 372)
    $script:pnlAudit.Size = New-Object System.Drawing.Size(665, 190)
    $script:pnlAudit.Visible = $false
    $script:pnlAudit.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
        [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor `
        [System.Windows.Forms.AnchorStyles]::Bottom
    $script:rightContent.Controls.Add($script:pnlAudit)

    $lblAuditH = New-Object System.Windows.Forms.Label
    $lblAuditH.Text = Get-Text "pim.audit_detail_label"
    $lblAuditH.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblAuditH.Location = New-Object System.Drawing.Point(0, 0)
    $lblAuditH.Size = New-Object System.Drawing.Size(300, 22)
    $script:pnlAudit.Controls.Add($lblAuditH)

    $script:lstAuditResults = New-Object System.Windows.Forms.ListView
    $script:lstAuditResults.Location = New-Object System.Drawing.Point(0, 25)
    $script:lstAuditResults.Size = New-Object System.Drawing.Size(665, 160)
    $script:lstAuditResults.View = [System.Windows.Forms.View]::Details
    $script:lstAuditResults.FullRowSelect = $true
    $script:lstAuditResults.Font = New-Object System.Drawing.Font("Consolas", 9)
    $script:lstAuditResults.Columns.Add((Get-Text "pim.col_audit_result"), 645) | Out-Null
    $script:lstAuditResults.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
        [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor `
        [System.Windows.Forms.AnchorStyles]::Bottom
    $script:pnlAudit.Controls.Add($script:lstAuditResults)

    # ================================================================
    #  BARRE D'ACTIONS (bas)
    # ================================================================
    $script:PimProgressBar = New-Object System.Windows.Forms.ProgressBar
    $script:PimProgressBar.Location = New-Object System.Drawing.Point(20, 8)
    $script:PimProgressBar.Size = New-Object System.Drawing.Size(900, 18)
    $script:PimProgressBar.Style = "Continuous"
    $script:PimProgressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
        [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $actionPanel.Controls.Add($script:PimProgressBar)

    $script:PimProgressLabel = New-Object System.Windows.Forms.Label
    $script:PimProgressLabel.Text = "0 / 0"
    $script:PimProgressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $script:PimProgressLabel.ForeColor = [System.Drawing.Color]::Gray
    $script:PimProgressLabel.Location = New-Object System.Drawing.Point(930, 8)
    $script:PimProgressLabel.Size = New-Object System.Drawing.Size(80, 18)
    $script:PimProgressLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $actionPanel.Controls.Add($script:PimProgressLabel)

    # Label de statut d'etape (affiche l'operation en cours)
    $script:lblStepStatus = New-Object System.Windows.Forms.Label
    $script:lblStepStatus.Text = ""
    $script:lblStepStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $script:lblStepStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $script:lblStepStatus.Location = New-Object System.Drawing.Point(20, 28)
    $script:lblStepStatus.Size = New-Object System.Drawing.Size(700, 16)
    $script:lblStepStatus.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
        [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $actionPanel.Controls.Add($script:lblStepStatus)

    # Boutons d'action
    $script:btnAudit = New-Object System.Windows.Forms.Button
    $script:btnAudit.Text = Get-Text "pim.btn_audit"
    $script:btnAudit.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $script:btnAudit.Location = New-Object System.Drawing.Point(20, 50)
    $script:btnAudit.Size = New-Object System.Drawing.Size(140, 35)
    $script:btnAudit.BackColor = $clrOK; $script:btnAudit.ForeColor = $clrWhite
    $script:btnAudit.FlatStyle = "Flat"; $script:btnAudit.FlatAppearance.BorderSize = 0
    $script:btnAudit.Cursor = [System.Windows.Forms.Cursors]::Hand
    $actionPanel.Controls.Add($script:btnAudit)

    $script:btnUpdate = New-Object System.Windows.Forms.Button
    $script:btnUpdate.Text = Get-Text "pim.btn_update"
    $script:btnUpdate.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $script:btnUpdate.Location = New-Object System.Drawing.Point(168, 50)
    $script:btnUpdate.Size = New-Object System.Drawing.Size(170, 35)
    $script:btnUpdate.BackColor = $clrOK; $script:btnUpdate.ForeColor = $clrWhite
    $script:btnUpdate.FlatStyle = "Flat"; $script:btnUpdate.FlatAppearance.BorderSize = 0
    $script:btnUpdate.Cursor = [System.Windows.Forms.Cursors]::Hand
    $actionPanel.Controls.Add($script:btnUpdate)

    $script:btnCreate = New-Object System.Windows.Forms.Button
    $script:btnCreate.Text = Get-Text "pim.btn_create"
    $script:btnCreate.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $script:btnCreate.Location = New-Object System.Drawing.Point(346, 50)
    $script:btnCreate.Size = New-Object System.Drawing.Size(200, 35)
    $script:btnCreate.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $script:btnCreate.ForeColor = $clrWhite
    $script:btnCreate.FlatStyle = "Flat"; $script:btnCreate.FlatAppearance.BorderSize = 0
    $script:btnCreate.Cursor = [System.Windows.Forms.Cursors]::Hand
    $actionPanel.Controls.Add($script:btnCreate)

    $script:btnSave = New-Object System.Windows.Forms.Button
    $script:btnSave.Text = Get-Text "pim.btn_save"
    $script:btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $script:btnSave.Location = New-Object System.Drawing.Point(554, 50)
    $script:btnSave.Size = New-Object System.Drawing.Size(140, 35)
    $script:btnSave.BackColor = [System.Drawing.Color]::FromArgb(111, 66, 193)
    $script:btnSave.ForeColor = $clrWhite
    $script:btnSave.FlatStyle = "Flat"; $script:btnSave.FlatAppearance.BorderSize = 0
    $script:btnSave.Cursor = [System.Windows.Forms.Cursors]::Hand
    $actionPanel.Controls.Add($script:btnSave)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = Get-Text "pim.btn_export"
    $btnExport.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnExport.Location = New-Object System.Drawing.Point(702, 50)
    $btnExport.Size = New-Object System.Drawing.Size(120, 35)
    $btnExport.FlatStyle = "Flat"
    $btnExport.Cursor = [System.Windows.Forms.Cursors]::Hand
    $actionPanel.Controls.Add($btnExport)

    # Callbacks de progression et d'etape
    $progressCallback = { param([int]$Done, [int]$Total); Update-PimProgress -Done $Done -Total $Total }
    $stepCallback = {
        param([string]$StepMsg)
        $script:lblStepStatus.Text = $StepMsg
        [System.Windows.Forms.Application]::DoEvents()
    }

    # ================================================================
    #  EVENTS — Editeur
    # ================================================================

    # Ajouter un role depuis la ComboBox
    $btnAddRole.Add_Click({
        $roleName = $script:cboRoleSearch.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($roleName)) { return }

        # Verifier que le groupe est selectionne
        $gn = $script:SelectedGroup
        if (-not $gn -or -not $script:PimData.Contains($gn)) { return }

        # Verifier que le role n'est pas deja assigne
        $existing = $script:PimData[$gn].Roles | Where-Object { $_ -eq $roleName }
        if ($existing) {
            [System.Windows.Forms.MessageBox]::Show(
                (Get-Text "pim.role_already_assigned"),
                (Get-Text "pim.title"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }

        # Ajouter le role
        $script:PimData[$gn].Roles += $roleName
        $script:PimDirty = $true

        # Rafraichir l'affichage
        Show-PimGroupDetail -GroupName $gn
        Update-PimDashboard
        $script:cboRoleSearch.Text = ''
        Write-Log -Level "INFO" -Action "PIM_EDIT" -Message "[$gn] Role ajoute : $roleName"
    })

    # Retirer le role selectionne
    $btnRemoveRole.Add_Click({
        $gn = $script:SelectedGroup
        if (-not $gn -or -not $script:PimData.Contains($gn)) { return }
        if ($script:lstAssignedRoles.SelectedItems.Count -eq 0) { return }

        $roleName = $script:lstAssignedRoles.SelectedItems[0].Text
        $script:PimData[$gn].Roles = @($script:PimData[$gn].Roles | Where-Object { $_ -ne $roleName })
        $script:PimDirty = $true

        Show-PimGroupDetail -GroupName $gn
        Update-PimDashboard
        Write-Log -Level "INFO" -Action "PIM_EDIT" -Message "[$gn] Role retire : $roleName"
    })

    # Nouveau groupe PIM
    $btnNewGroup.Add_Click({
        # Mini formulaire de saisie du nom
        $inputForm = New-Object System.Windows.Forms.Form
        $inputForm.Text = Get-Text "pim.new_group_title"
        $inputForm.Size = New-Object System.Drawing.Size(500, 180)
        $inputForm.StartPosition = "CenterParent"
        $inputForm.FormBorderStyle = "FixedDialog"
        $inputForm.MaximizeBox = $false; $inputForm.MinimizeBox = $false

        $inputLabel = New-Object System.Windows.Forms.Label
        $inputLabel.Text = Get-Text "pim.new_group_prompt"
        $inputLabel.Location = New-Object System.Drawing.Point(15, 15)
        $inputLabel.Size = New-Object System.Drawing.Size(460, 20)
        $inputForm.Controls.Add($inputLabel)

        $inputBox = New-Object System.Windows.Forms.TextBox
        $inputBox.Text = "PIM_Role_"
        $inputBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $inputBox.Location = New-Object System.Drawing.Point(15, 42)
        $inputBox.Size = New-Object System.Drawing.Size(450, 28)
        $inputForm.Controls.Add($inputBox)

        $inputOK = New-Object System.Windows.Forms.Button
        $inputOK.Text = "OK"
        $inputOK.Location = New-Object System.Drawing.Point(285, 80)
        $inputOK.Size = New-Object System.Drawing.Size(85, 32)
        $inputOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $inputOK.BackColor = $clrSuccess; $inputOK.ForeColor = $clrWhite
        $inputOK.FlatStyle = "Flat"; $inputOK.FlatAppearance.BorderSize = 0
        $inputForm.Controls.Add($inputOK)

        $inputCancel = New-Object System.Windows.Forms.Button
        $inputCancel.Text = Get-Text "pim.btn_cancel"
        $inputCancel.Location = New-Object System.Drawing.Point(380, 80)
        $inputCancel.Size = New-Object System.Drawing.Size(85, 32)
        $inputCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $inputForm.Controls.Add($inputCancel)

        $inputForm.AcceptButton = $inputOK
        $inputForm.CancelButton = $inputCancel

        $result = $inputForm.ShowDialog()
        $newName = $inputBox.Text.Trim()
        $inputForm.Dispose()

        if ($result -ne [System.Windows.Forms.DialogResult]::OK -or [string]::IsNullOrWhiteSpace($newName)) { return }

        if ($script:PimData.Contains($newName)) {
            [System.Windows.Forms.MessageBox]::Show(
                (Get-Text "pim.new_group_duplicate"),
                (Get-Text "pim.title"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $script:PimData[$newName] = @{
            Description = ''
            Type        = 'Role'
            Roles       = @()
        }
        $script:PimGroupStatus[$newName] = 'Pending'
        $script:PimDirty = $true

        Update-PimDashboard
        $script:SelectedGroup = $newName
        Show-PimGroupDetail -GroupName $newName
        Update-PimCardSelection
        Write-Log -Level "INFO" -Action "PIM_EDIT" -Message "Nouveau groupe PIM cree : $newName"
    })

    # Supprimer le groupe selectionne
    $btnDeleteGroup.Add_Click({
        $gn = $script:SelectedGroup
        if (-not $gn -or -not $script:PimData.Contains($gn)) { return }

        $confirm = Show-ConfirmDialog -Titre (Get-Text "pim.delete_confirm_title") `
            -Message ((Get-Text "pim.delete_confirm_msg") -f $gn)
        if (-not $confirm) { return }

        Remove-PimGroupFromConfig -GroupName $gn
        $script:PimDirty = $true
        $script:SelectedGroup = $null

        Update-PimDashboard
        # Vider le panneau droit
        $script:txtEditName.Text = ''; $script:txtEditDesc.Text = ''
        $script:lstAssignedRoles.Items.Clear()
        $script:pnlAudit.Visible = $false
        $script:lblDetailStatus.Text = ''
        Write-Log -Level "INFO" -Action "PIM_EDIT" -Message "Groupe PIM supprime : $gn"
    })

    # Importer les groupes PIM existants depuis le tenant Entra
    $btnImportTenant.Add_Click({
        Set-PimButtonsEnabled -Enabled $false
        $btnImportTenant.Enabled = $false
        $script:lblStepStatus.Text = Get-Text "pim.import_step_searching"
        [System.Windows.Forms.Application]::DoEvents()

        $discovered = Import-PimGroupsFromTenant -StepCallback $stepCallback

        if (-not $discovered -or $discovered.Count -eq 0) {
            Show-ResultDialog -Titre (Get-Text "pim.title") `
                -Message (Get-Text "pim.import_none_found") -IsSuccess $false
            Set-PimButtonsEnabled -Enabled $true
            $btnImportTenant.Enabled = $true
            $script:lblStepStatus.Text = ''
            return
        }

        # --- Dialog de selection des groupes a importer ---
        $importForm = New-Object System.Windows.Forms.Form
        $importForm.Text = Get-Text "pim.import_dialog_title"
        $importForm.Size = New-Object System.Drawing.Size(700, 520)
        $importForm.StartPosition = "CenterParent"
        $importForm.FormBorderStyle = "Sizable"
        $importForm.MinimumSize = New-Object System.Drawing.Size(550, 400)
        $importForm.MaximizeBox = $false

        $iconFile = Join-Path -Path $script:RootPath -ChildPath "Assets\M365Monster.ico"
        if (Test-Path $iconFile) { $importForm.Icon = New-Object System.Drawing.Icon($iconFile) }

        $lblInstr = New-Object System.Windows.Forms.Label
        $lblInstr.Text = Get-Text "pim.import_dialog_instruction"
        $lblInstr.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $lblInstr.Location = New-Object System.Drawing.Point(15, 12)
        $lblInstr.Size = New-Object System.Drawing.Size(660, 36)
        $lblInstr.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        $importForm.Controls.Add($lblInstr)

        # ListView avec checkboxes
        $lstImport = New-Object System.Windows.Forms.ListView
        $lstImport.Location = New-Object System.Drawing.Point(15, 52)
        $lstImport.Size = New-Object System.Drawing.Size(654, 360)
        $lstImport.View = [System.Windows.Forms.View]::Details
        $lstImport.CheckBoxes = $true
        $lstImport.FullRowSelect = $true
        $lstImport.GridLines = $true
        $lstImport.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $lstImport.Columns.Add((Get-Text "pim.import_col_group"), 200) | Out-Null
        $lstImport.Columns.Add((Get-Text "pim.import_col_type"), 100) | Out-Null
        $lstImport.Columns.Add((Get-Text "pim.import_col_roles"), 55) | Out-Null
        $lstImport.Columns.Add((Get-Text "pim.import_col_status"), 100) | Out-Null
        $lstImport.Columns.Add("Description", 180) | Out-Null
        $lstImport.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
            [System.Windows.Forms.AnchorStyles]::Left -bor `
            [System.Windows.Forms.AnchorStyles]::Right -bor `
            [System.Windows.Forms.AnchorStyles]::Bottom
        $importForm.Controls.Add($lstImport)

        # Peupler la liste
        foreach ($gName in $discovered.Keys) {
            $info = $discovered[$gName]
            $item = New-Object System.Windows.Forms.ListViewItem($gName)
            $item.SubItems.Add($info.Type)           | Out-Null
            $item.SubItems.Add("$($info.RoleCount)") | Out-Null
            $statusText = if ($info.IsNew) { Get-Text "pim.import_status_new" } else { Get-Text "pim.import_status_exists" }
            $item.SubItems.Add($statusText) | Out-Null
            $descTrunc = if ($info.Description.Length -gt 40) { $info.Description.Substring(0, 37) + '...' } else { $info.Description }
            $item.SubItems.Add($descTrunc) | Out-Null

            # Cocher par defaut uniquement les nouveaux groupes
            $item.Checked = $info.IsNew
            if (-not $info.IsNew) {
                $item.ForeColor = $clrSkipped
            }
            $item.Tag = $gName
            $lstImport.Items.Add($item) | Out-Null
        }

        # Boutons OK / Annuler
        $btnImportOK = New-Object System.Windows.Forms.Button
        $btnImportOK.Text = Get-Text "pim.import_btn_import"
        $btnImportOK.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $btnImportOK.Size = New-Object System.Drawing.Size(140, 34)
        $btnImportOK.Location = New-Object System.Drawing.Point(390, 420)
        $btnImportOK.BackColor = $clrSuccess
        $btnImportOK.ForeColor = $clrWhite
        $btnImportOK.FlatStyle = "Flat"
        $btnImportOK.FlatAppearance.BorderSize = 0
        $btnImportOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $btnImportOK.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
        $importForm.Controls.Add($btnImportOK)

        $btnImportCancel = New-Object System.Windows.Forms.Button
        $btnImportCancel.Text = Get-Text "pim.btn_cancel"
        $btnImportCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $btnImportCancel.Size = New-Object System.Drawing.Size(120, 34)
        $btnImportCancel.Location = New-Object System.Drawing.Point(540, 420)
        $btnImportCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $btnImportCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
        $importForm.Controls.Add($btnImportCancel)

        $importForm.AcceptButton = $btnImportOK
        $importForm.CancelButton = $btnImportCancel

        $dialogResult = $importForm.ShowDialog()
        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $imported = 0
            $updated  = 0
            foreach ($item in $lstImport.CheckedItems) {
                $gName = $item.Tag
                $info  = $discovered[$gName]
                if (-not $info) { continue }

                if ($info.IsNew) {
                    # Nouveau groupe — ajouter a la config
                    $script:PimData[$gName] = @{
                        Description = $info.Description
                        Type        = $info.Type
                        Roles       = @($info.Roles)
                    }
                    $script:PimGroupStatus[$gName] = 'Pending'
                    $imported++
                    Write-Log -Level "INFO" -Action "PIM_IMPORT" -Message "Importe : $gName ($($info.RoleCount) roles)"
                }
                else {
                    # Groupe existant — fusionner les roles (ajouter les manquants)
                    $existing = $script:PimData[$gName]
                    $added = 0
                    foreach ($role in $info.Roles) {
                        if ($role -notin $existing.Roles) {
                            $existing.Roles += $role
                            $added++
                        }
                    }
                    # Mettre a jour la description si vide localement
                    if ([string]::IsNullOrWhiteSpace($existing.Description) -and $info.Description) {
                        $existing.Description = $info.Description
                    }
                    if ($added -gt 0) {
                        $updated++
                        Write-Log -Level "INFO" -Action "PIM_IMPORT" -Message "Fusionne : $gName (+$added roles)"
                    }
                }
            }

            if (($imported + $updated) -gt 0) {
                $script:PimDirty = $true
                Update-PimDashboard

                # Rafraichir la ComboBox de roles (les custom resolus par l'import
                # ont ete ajoutes a $AllEntraRoles)
                $script:cboRoleSearch.Items.Clear()
                foreach ($roleName in ($script:AllEntraRoles.Keys | Sort-Object)) {
                    $script:cboRoleSearch.Items.Add($roleName) | Out-Null
                }

                $msg = (Get-Text "pim.import_success") -f $imported, $updated
                Show-ResultDialog -Titre (Get-Text "pim.title") -Message $msg -IsSuccess $true
            }
        }
        $importForm.Dispose()

        Set-PimButtonsEnabled -Enabled $true
        $btnImportTenant.Enabled = $true
        $script:lblStepStatus.Text = ''
    })

    # ================================================================
    #  EVENTS — Sauvegarde
    # ================================================================
    $script:btnSave.Add_Click({
        # Sauvegarder les modifications du groupe actuellement affiche
        $gn = $script:SelectedGroup
        if ($gn -and $script:PimData.Contains($gn)) {
            $newName = $script:txtEditName.Text.Trim()
            $newDesc = $script:txtEditDesc.Text.Trim()
            $newType = $script:cboEditType.SelectedItem

            # Si le nom a change, renommer la cle
            $origName = $script:txtEditName.Tag
            if ($origName -and $origName -ne $newName -and -not [string]::IsNullOrWhiteSpace($newName)) {
                if ($script:PimData.Contains($newName)) {
                    [System.Windows.Forms.MessageBox]::Show(
                        (Get-Text "pim.new_group_duplicate"),
                        (Get-Text "pim.title"),
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                    return
                }
                # Copier les donnees vers la nouvelle cle
                $script:PimData[$newName] = $script:PimData[$origName]
                $script:PimGroupStatus[$newName] = $script:PimGroupStatus[$origName]
                if ($script:PimAuditResults[$origName]) {
                    $script:PimAuditResults[$newName] = $script:PimAuditResults[$origName]
                }
                $script:PimData.Remove($origName)
                $script:PimGroupStatus.Remove($origName)
                $script:PimAuditResults.Remove($origName)
                $script:SelectedGroup = $newName
                $gn = $newName
            }

            $script:PimData[$gn].Description = $newDesc
            $script:PimData[$gn].Type = $newType
        }

        # Persister dans le JSON client
        if (Save-PimConfig) {
            $script:PimDirty = $false
            Show-ResultDialog -Titre (Get-Text "pim.title") `
                -Message (Get-Text "pim.save_success") -IsSuccess $true
            Update-PimDashboard
            if ($gn) {
                Show-PimGroupDetail -GroupName $gn
                Update-PimCardSelection
            }
        }
        else {
            Show-ResultDialog -Titre (Get-Text "pim.title") `
                -Message (Get-Text "pim.save_error") -IsSuccess $false
        }
    })

    # ================================================================
    #  EVENTS — Actions PIM (audit, update, create)
    # ================================================================
    $script:btnAudit.Add_Click({
        Set-PimButtonsEnabled -Enabled $false
        $script:lblStepStatus.Text = ''
        if ($script:AllEntraRoles.Count -eq 0 -and -not (Load-EntraRoles)) {
            Show-ResultDialog -Titre (Get-Text "pim.title") -Message (Get-Text "pim.roles_load_error") -IsSuccess $false
            Set-PimButtonsEnabled -Enabled $true; return
        }
        Invoke-PimAudit -ProgressCallback $progressCallback -StepCallback $stepCallback
        Update-PimDashboard
        $gn = $script:SelectedGroup
        if ($gn -and $script:PimData.Contains($gn)) { Show-PimGroupDetail -GroupName $gn }
        Set-PimButtonsEnabled -Enabled $true
    })

    $script:btnUpdate.Add_Click({
        $gn = $script:SelectedGroup
        if (-not $gn) {
            [System.Windows.Forms.MessageBox]::Show(
                (Get-Text "pim.no_selection"), (Get-Text "pim.title"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
        $toUpdate = @($gn) | Where-Object { $script:PimGroupStatus[$_] -ne 'Pending' -and $script:PimGroupStatus[$_] -ne 'Missing' }
        if ($toUpdate.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                (Get-Text "pim.no_existing_selection"), (Get-Text "pim.title"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
        if ($script:AllEntraRoles.Count -eq 0 -and -not (Load-EntraRoles)) {
            Show-ResultDialog -Titre (Get-Text "pim.title") -Message (Get-Text "pim.roles_load_error") -IsSuccess $false; return
        }
        Set-PimButtonsEnabled -Enabled $false
        $script:lblStepStatus.Text = ''

        # Executer la mise a jour
        Invoke-PimUpdate -GroupNames $toUpdate -ProgressCallback $progressCallback -StepCallback $stepCallback

        # Auto-audit apres la mise a jour pour verifier le resultat
        & $stepCallback (Get-Text "pim.step_auto_audit")
        Start-Sleep -Seconds 2
        [System.Windows.Forms.Application]::DoEvents()
        Invoke-PimAudit -ProgressCallback $progressCallback -StepCallback $stepCallback

        Update-PimDashboard
        Show-PimGroupDetail -GroupName $gn
        Update-PimCardSelection
        Set-PimButtonsEnabled -Enabled $true
    })

    $script:btnCreate.Add_Click({
        $gn = $script:SelectedGroup
        if (-not $gn) {
            [System.Windows.Forms.MessageBox]::Show(
                (Get-Text "pim.no_selection"), (Get-Text "pim.title"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
        if ($script:AllEntraRoles.Count -eq 0 -and -not (Load-EntraRoles)) {
            Show-ResultDialog -Titre (Get-Text "pim.title") -Message (Get-Text "pim.roles_load_error") -IsSuccess $false; return
        }
        Set-PimButtonsEnabled -Enabled $false
        $script:lblStepStatus.Text = ''
        $def = $script:PimData[$gn]

        & $stepCallback ("[$gn] " + (Get-Text "pim.step_creating_group"))
        $gid = New-PimGroupEntra -Name $gn -Def $def
        if ($gid -and $script:PimGroupStatus[$gn] -ne 'Skipped') {
            Add-PimRolesToGroup -GroupId $gid -GroupName $gn -Roles $def.Roles -GroupType $def.Type
        }
        Update-PimProgress -Done 1 -Total 1

        # Auto-audit apres la creation
        & $stepCallback (Get-Text "pim.step_auto_audit")
        Start-Sleep -Seconds 2
        [System.Windows.Forms.Application]::DoEvents()
        Invoke-PimAudit -ProgressCallback $progressCallback -StepCallback $stepCallback

        Update-PimDashboard
        Show-PimGroupDetail -GroupName $gn
        Update-PimCardSelection
        Set-PimButtonsEnabled -Enabled $true
    })

    $btnExport.Add_Click({ Export-PimCsvReport })

    # ================================================================
    #  CHARGEMENT INITIAL
    # ================================================================
    Write-Log -Level "INFO" -Action "PIM" -Message "PIM Manager ouvert pour '$($Config.client_name)'"

    # Chargement silencieux des roles Entra
    $rolesLoaded = Load-EntraRoles
    if ($rolesLoaded) {
        # Peupler la ComboBox de recherche de roles
        $script:cboRoleSearch.Items.Clear()
        foreach ($roleName in ($script:AllEntraRoles.Keys | Sort-Object)) {
            $script:cboRoleSearch.Items.Add($roleName) | Out-Null
        }
    }
    else {
        Write-Log -Level "WARNING" -Action "PIM" -Message "Roles Entra non charges -- chargement au premier audit."
    }

    # Construire le dashboard
    Update-PimDashboard
    $script:SelectedGroup = $null

    $f.ShowDialog() | Out-Null
    $f.Dispose()
}