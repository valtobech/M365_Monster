<#
.FICHIER
    Modules/GUI_NestedGroupAudit.ps1

.ROLE
    Module d'audit des groupes Entra ID contenant à la fois des utilisateurs
    et des appareils (groupes "nested" / mixtes).
    Identifie les policies et applications Intune assignées à ces groupes.
    Permet de créer de nouveaux groupes séparés (Users / Devices)
    et d'y transférer les membres correspondants.

    Fonctionnalités :
    - Scanner tous les groupes Entra et identifier les groupes mixtes (Users + Devices)
    - Afficher les détails des membres (Users / Devices) par groupe
    - Scanner toutes les catégories Intune (Policies, Apps, Scripts, Updates...)
      pour identifier celles qui référencent un groupe mixte
    - Créer des groupes séparés {Original}_User / {Original}_Device
    - Transférer les membres vers les nouveaux groupes

.DEPENDANCES
    - Core/Functions.ps1 (Write-Log, Show-ConfirmDialog, Show-ResultDialog)
    - Core/Lang.ps1 (Get-Text)
    - Core/Connect.ps1 (Get-GraphConnectionStatus)
    - Connexion Graph active avec scopes :
        Group.ReadWrite.All, User.Read.All, Device.Read.All,
        DeviceManagementConfiguration.Read.All, DeviceManagementApps.Read.All,
        DeviceManagementManagedDevices.Read.All

.PERMISSIONS_GRAPH
    Delegated :
        - Group.ReadWrite.All (lecture groupes + création + ajout membres)
        - User.Read.All (lecture utilisateurs)
        - Device.Read.All (lecture devices)
        - DeviceManagementConfiguration.Read.All (policies config/compliance)
        - DeviceManagementApps.Read.All (apps Intune)
        - DeviceManagementManagedDevices.Read.All (devices Intune)

.AUTEUR
    [Equipe IT - M365 Monster]
#>

function Show-NestedGroupAuditForm {
    <#
    .SYNOPSIS
        Affiche le formulaire d'audit des groupes nested (mixtes Users + Devices).
        Scanne les groupes Entra, identifie les policies Intune liées,
        et permet la séparation des membres dans de nouveaux groupes.
    #>

    # =================================================================
    # Variables de travail du module
    # =================================================================
    $script:NestedGroupData = @()          # Groupes nested trouvés
    $script:SelectedGroup = $null          # Groupe sélectionné dans l'onglet 1
    $script:SelectedGroupMembers = @{      # Membres du groupe sélectionné
        Users   = @()
        Devices = @()
    }
    $script:IntuneAssignments = @()        # Policies/Apps Intune liées au groupe sélectionné
    $script:ScanCancelled = $false         # Flag d'annulation du scan

    # =================================================================
    # Couleurs cohérentes avec M365 Monster
    # =================================================================
    $colorDark       = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $colorBg         = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $colorWhite      = [System.Drawing.Color]::White
    $colorGray       = [System.Drawing.Color]::FromArgb(108, 117, 125)
    $colorLightGray  = [System.Drawing.Color]::FromArgb(173, 181, 189)
    $colorGreen      = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $colorOrange     = [System.Drawing.Color]::FromArgb(253, 126, 20)
    $colorBlue       = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $colorPurple     = [System.Drawing.Color]::FromArgb(111, 66, 193)

    # =================================================================
    # Fenêtre principale
    # =================================================================
    $form = New-Object System.Windows.Forms.Form
    $form.Text = Get-Text "nested_group_audit.title"
    $form.Size = New-Object System.Drawing.Size(1320, 880)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedSingle"
    $form.MaximizeBox = $false
    $form.BackColor = $colorBg

    # Icône si disponible
    $iconFile = Join-Path -Path $script:RootPath -ChildPath "Assets\M365Monster.ico"
    if (Test-Path $iconFile) { $form.Icon = New-Object System.Drawing.Icon($iconFile) }

    # =================================================================
    # En-tête
    # =================================================================
    $panelHeader = New-Object System.Windows.Forms.Panel
    $panelHeader.Location = New-Object System.Drawing.Point(0, 0)
    $panelHeader.Size = New-Object System.Drawing.Size(1320, 70)
    $panelHeader.BackColor = $colorDark
    $form.Controls.Add($panelHeader)

    $lblTitre = New-Object System.Windows.Forms.Label
    $lblTitre.Text = Get-Text "nested_group_audit.header_title"
    $lblTitre.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $lblTitre.ForeColor = $colorWhite
    $lblTitre.Location = New-Object System.Drawing.Point(20, 8)
    $lblTitre.Size = New-Object System.Drawing.Size(600, 30)
    $panelHeader.Controls.Add($lblTitre)

    $lblSubtitle = New-Object System.Windows.Forms.Label
    $lblSubtitle.Text = Get-Text "nested_group_audit.header_subtitle"
    $lblSubtitle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSubtitle.ForeColor = $colorLightGray
    $lblSubtitle.Location = New-Object System.Drawing.Point(20, 40)
    $lblSubtitle.Size = New-Object System.Drawing.Size(700, 20)
    $panelHeader.Controls.Add($lblSubtitle)

    # =================================================================
    # TabControl — 4 onglets
    # =================================================================
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10, 80)
    $tabControl.Size = New-Object System.Drawing.Size(1290, 720)
    $tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Controls.Add($tabControl)

    # =====================================================================
    # ONGLET 1 — Scan des groupes
    # =====================================================================
    $tab1 = New-Object System.Windows.Forms.TabPage
    $tab1.Text = Get-Text "nested_group_audit.tab_scan"
    $tab1.BackColor = $colorBg
    $tabControl.TabPages.Add($tab1)

    # Panneau de contrôle du scan
    $panelScanCtrl = New-Object System.Windows.Forms.Panel
    $panelScanCtrl.Location = New-Object System.Drawing.Point(10, 10)
    $panelScanCtrl.Size = New-Object System.Drawing.Size(1260, 50)
    $panelScanCtrl.BackColor = $colorWhite
    $tab1.Controls.Add($panelScanCtrl)

    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = Get-Text "nested_group_audit.btn_scan"
    $btnScan.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnScan.Location = New-Object System.Drawing.Point(10, 8)
    $btnScan.Size = New-Object System.Drawing.Size(200, 34)
    $btnScan.BackColor = $colorGreen
    $btnScan.ForeColor = $colorWhite
    $btnScan.FlatStyle = "Flat"
    $btnScan.FlatAppearance.BorderSize = 0
    $panelScanCtrl.Controls.Add($btnScan)

    $btnCancelScan = New-Object System.Windows.Forms.Button
    $btnCancelScan.Text = Get-Text "nested_group_audit.btn_cancel"
    $btnCancelScan.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnCancelScan.Location = New-Object System.Drawing.Point(220, 8)
    $btnCancelScan.Size = New-Object System.Drawing.Size(120, 34)
    $btnCancelScan.BackColor = [System.Drawing.Color]::FromArgb(233, 236, 239)
    $btnCancelScan.ForeColor = $colorDark
    $btnCancelScan.FlatStyle = "Flat"
    $btnCancelScan.FlatAppearance.BorderSize = 0
    $btnCancelScan.Enabled = $false
    $panelScanCtrl.Controls.Add($btnCancelScan)

    # Filtre de recherche
    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = Get-Text "nested_group_audit.lbl_filter"
    $lblFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblFilter.ForeColor = $colorGray
    $lblFilter.Location = New-Object System.Drawing.Point(500, 14)
    $lblFilter.Size = New-Object System.Drawing.Size(80, 20)
    $panelScanCtrl.Controls.Add($lblFilter)

    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtFilter.Location = New-Object System.Drawing.Point(585, 11)
    $txtFilter.Size = New-Object System.Drawing.Size(250, 28)
    $panelScanCtrl.Controls.Add($txtFilter)

    # Bouton Export CSV
    $btnExportScan = New-Object System.Windows.Forms.Button
    $btnExportScan.Text = Get-Text "nested_group_audit.btn_export"
    $btnExportScan.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnExportScan.Location = New-Object System.Drawing.Point(1100, 8)
    $btnExportScan.Size = New-Object System.Drawing.Size(150, 34)
    $btnExportScan.BackColor = $colorBlue
    $btnExportScan.ForeColor = $colorWhite
    $btnExportScan.FlatStyle = "Flat"
    $btnExportScan.FlatAppearance.BorderSize = 0
    $btnExportScan.Enabled = $false
    $panelScanCtrl.Controls.Add($btnExportScan)

    # Statistiques
    $lblScanStats = New-Object System.Windows.Forms.Label
    $lblScanStats.Text = Get-Text "nested_group_audit.stats_empty"
    $lblScanStats.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblScanStats.ForeColor = $colorGray
    $lblScanStats.Location = New-Object System.Drawing.Point(10, 68)
    $lblScanStats.Size = New-Object System.Drawing.Size(1260, 20)
    $tab1.Controls.Add($lblScanStats)

    # Barre de progression
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(10, 92)
    $progressBar.Size = New-Object System.Drawing.Size(1260, 8)
    $progressBar.Style = "Continuous"
    $progressBar.Visible = $false
    $tab1.Controls.Add($progressBar)

    # DataGridView — groupes nested
    $dgvGroups = New-Object System.Windows.Forms.DataGridView
    $dgvGroups.Location = New-Object System.Drawing.Point(10, 108)
    $dgvGroups.Size = New-Object System.Drawing.Size(1260, 560)
    $dgvGroups.BackgroundColor = $colorWhite
    $dgvGroups.BorderStyle = "None"
    $dgvGroups.GridColor = [System.Drawing.Color]::FromArgb(222, 226, 230)
    $dgvGroups.AllowUserToAddRows = $false
    $dgvGroups.AllowUserToDeleteRows = $false
    $dgvGroups.AllowUserToResizeRows = $false
    $dgvGroups.ReadOnly = $true
    $dgvGroups.SelectionMode = "FullRowSelect"
    $dgvGroups.MultiSelect = $false
    $dgvGroups.RowHeadersVisible = $false
    $dgvGroups.AutoSizeColumnsMode = "None"
    $dgvGroups.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dgvGroups.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dgvGroups.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(233, 236, 239)
    $dgvGroups.ColumnHeadersDefaultCellStyle.ForeColor = $colorDark
    $dgvGroups.ColumnHeadersHeight = 35
    $dgvGroups.RowTemplate.Height = 30
    $dgvGroups.EnableHeadersVisualStyles = $false
    $dgvGroups.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(200, 220, 255)
    $dgvGroups.DefaultCellStyle.SelectionForeColor = $colorDark
    $tab1.Controls.Add($dgvGroups)

    # Colonnes du DataGridView groupes
    @(
        @{ Name = "DisplayName";  Header = "nested_group_audit.col_group_name";  Width = 250 },
        @{ Name = "GroupId";      Header = "nested_group_audit.col_group_id";    Width = 280 },
        @{ Name = "GroupType";    Header = "nested_group_audit.col_group_type";  Width = 150 },
        @{ Name = "Membership";   Header = "nested_group_audit.col_membership";  Width = 110 },
        @{ Name = "UserCount";    Header = "nested_group_audit.col_users";       Width = 100 },
        @{ Name = "DeviceCount";  Header = "nested_group_audit.col_devices";     Width = 100 },
        @{ Name = "OtherCount";   Header = "nested_group_audit.col_others";      Width = 80  },
        @{ Name = "TotalCount";   Header = "nested_group_audit.col_total";       Width = 80  }
    ) | ForEach-Object {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $_.Name
        $col.HeaderText = Get-Text $_.Header
        $col.Width = $_.Width
        $dgvGroups.Columns.Add($col) | Out-Null
    }

    # =====================================================================
    # ONGLET 2 — Détails des membres
    # =====================================================================
    $tab2 = New-Object System.Windows.Forms.TabPage
    $tab2.Text = Get-Text "nested_group_audit.tab_members"
    $tab2.BackColor = $colorBg
    $tabControl.TabPages.Add($tab2)

    $lblMembersGroup = New-Object System.Windows.Forms.Label
    $lblMembersGroup.Text = Get-Text "nested_group_audit.no_group_selected"
    $lblMembersGroup.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblMembersGroup.ForeColor = $colorDark
    $lblMembersGroup.Location = New-Object System.Drawing.Point(10, 10)
    $lblMembersGroup.Size = New-Object System.Drawing.Size(1260, 25)
    $tab2.Controls.Add($lblMembersGroup)

    # Titre Users
    $lblUsersTitle = New-Object System.Windows.Forms.Label
    $lblUsersTitle.Text = Get-Text "nested_group_audit.lbl_users"
    $lblUsersTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblUsersTitle.ForeColor = $colorBlue
    $lblUsersTitle.Location = New-Object System.Drawing.Point(10, 42)
    $lblUsersTitle.Size = New-Object System.Drawing.Size(620, 22)
    $tab2.Controls.Add($lblUsersTitle)

    # DataGridView — Users
    $dgvUsers = New-Object System.Windows.Forms.DataGridView
    $dgvUsers.Location = New-Object System.Drawing.Point(10, 68)
    $dgvUsers.Size = New-Object System.Drawing.Size(620, 590)
    $dgvUsers.BackgroundColor = $colorWhite
    $dgvUsers.BorderStyle = "None"
    $dgvUsers.GridColor = [System.Drawing.Color]::FromArgb(222, 226, 230)
    $dgvUsers.AllowUserToAddRows = $false
    $dgvUsers.AllowUserToDeleteRows = $false
    $dgvUsers.ReadOnly = $true
    $dgvUsers.SelectionMode = "FullRowSelect"
    $dgvUsers.MultiSelect = $false
    $dgvUsers.RowHeadersVisible = $false
    $dgvUsers.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dgvUsers.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dgvUsers.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(233, 236, 239)
    $dgvUsers.ColumnHeadersHeight = 32
    $dgvUsers.RowTemplate.Height = 28
    $dgvUsers.EnableHeadersVisualStyles = $false
    $dgvUsers.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(200, 220, 255)
    $dgvUsers.DefaultCellStyle.SelectionForeColor = $colorDark
    $tab2.Controls.Add($dgvUsers)

    @(
        @{ Name = "UserDisplayName"; Header = "nested_group_audit.col_displayname"; Width = 200 },
        @{ Name = "UserUPN";         Header = "nested_group_audit.col_upn";         Width = 250 },
        @{ Name = "UserJobTitle";    Header = "nested_group_audit.col_jobtitle";    Width = 160 }
    ) | ForEach-Object {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $_.Name
        $col.HeaderText = Get-Text $_.Header
        $col.Width = $_.Width
        $dgvUsers.Columns.Add($col) | Out-Null
    }

    # Titre Devices
    $lblDevicesTitle = New-Object System.Windows.Forms.Label
    $lblDevicesTitle.Text = Get-Text "nested_group_audit.lbl_devices"
    $lblDevicesTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblDevicesTitle.ForeColor = $colorOrange
    $lblDevicesTitle.Location = New-Object System.Drawing.Point(645, 42)
    $lblDevicesTitle.Size = New-Object System.Drawing.Size(625, 22)
    $tab2.Controls.Add($lblDevicesTitle)

    # DataGridView — Devices
    $dgvDevices = New-Object System.Windows.Forms.DataGridView
    $dgvDevices.Location = New-Object System.Drawing.Point(645, 68)
    $dgvDevices.Size = New-Object System.Drawing.Size(625, 590)
    $dgvDevices.BackgroundColor = $colorWhite
    $dgvDevices.BorderStyle = "None"
    $dgvDevices.GridColor = [System.Drawing.Color]::FromArgb(222, 226, 230)
    $dgvDevices.AllowUserToAddRows = $false
    $dgvDevices.AllowUserToDeleteRows = $false
    $dgvDevices.ReadOnly = $true
    $dgvDevices.SelectionMode = "FullRowSelect"
    $dgvDevices.MultiSelect = $false
    $dgvDevices.RowHeadersVisible = $false
    $dgvDevices.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dgvDevices.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dgvDevices.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(233, 236, 239)
    $dgvDevices.ColumnHeadersHeight = 32
    $dgvDevices.RowTemplate.Height = 28
    $dgvDevices.EnableHeadersVisualStyles = $false
    $dgvDevices.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(200, 220, 255)
    $dgvDevices.DefaultCellStyle.SelectionForeColor = $colorDark
    $tab2.Controls.Add($dgvDevices)

    @(
        @{ Name = "DeviceDisplayName"; Header = "nested_group_audit.col_device_name"; Width = 220 },
        @{ Name = "DeviceOS";          Header = "nested_group_audit.col_device_os";   Width = 150 },
        @{ Name = "DeviceId";          Header = "nested_group_audit.col_device_id";   Width = 245 }
    ) | ForEach-Object {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $_.Name
        $col.HeaderText = Get-Text $_.Header
        $col.Width = $_.Width
        $dgvDevices.Columns.Add($col) | Out-Null
    }

    # =====================================================================
    # ONGLET 3 — Impact Intune
    # =====================================================================
    $tab3 = New-Object System.Windows.Forms.TabPage
    $tab3.Text = Get-Text "nested_group_audit.tab_intune"
    $tab3.BackColor = $colorBg
    $tabControl.TabPages.Add($tab3)

    # Panneau de contrôle Intune
    $panelIntuneCtrl = New-Object System.Windows.Forms.Panel
    $panelIntuneCtrl.Location = New-Object System.Drawing.Point(10, 10)
    $panelIntuneCtrl.Size = New-Object System.Drawing.Size(1260, 50)
    $panelIntuneCtrl.BackColor = $colorWhite
    $tab3.Controls.Add($panelIntuneCtrl)

    $lblIntuneGroup = New-Object System.Windows.Forms.Label
    $lblIntuneGroup.Text = Get-Text "nested_group_audit.no_group_selected"
    $lblIntuneGroup.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblIntuneGroup.ForeColor = $colorDark
    $lblIntuneGroup.Location = New-Object System.Drawing.Point(10, 14)
    $lblIntuneGroup.Size = New-Object System.Drawing.Size(600, 22)
    $panelIntuneCtrl.Controls.Add($lblIntuneGroup)

    $btnScanIntune = New-Object System.Windows.Forms.Button
    $btnScanIntune.Text = Get-Text "nested_group_audit.btn_scan_intune"
    $btnScanIntune.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnScanIntune.Location = New-Object System.Drawing.Point(900, 8)
    $btnScanIntune.Size = New-Object System.Drawing.Size(200, 34)
    $btnScanIntune.BackColor = $colorPurple
    $btnScanIntune.ForeColor = $colorWhite
    $btnScanIntune.FlatStyle = "Flat"
    $btnScanIntune.FlatAppearance.BorderSize = 0
    $btnScanIntune.Enabled = $false
    $panelIntuneCtrl.Controls.Add($btnScanIntune)

    $btnExportIntune = New-Object System.Windows.Forms.Button
    $btnExportIntune.Text = Get-Text "nested_group_audit.btn_export"
    $btnExportIntune.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnExportIntune.Location = New-Object System.Drawing.Point(1110, 8)
    $btnExportIntune.Size = New-Object System.Drawing.Size(140, 34)
    $btnExportIntune.BackColor = $colorBlue
    $btnExportIntune.ForeColor = $colorWhite
    $btnExportIntune.FlatStyle = "Flat"
    $btnExportIntune.FlatAppearance.BorderSize = 0
    $btnExportIntune.Enabled = $false
    $panelIntuneCtrl.Controls.Add($btnExportIntune)

    # Barre de progression Intune
    $progressBarIntune = New-Object System.Windows.Forms.ProgressBar
    $progressBarIntune.Location = New-Object System.Drawing.Point(10, 65)
    $progressBarIntune.Size = New-Object System.Drawing.Size(1260, 8)
    $progressBarIntune.Style = "Continuous"
    $progressBarIntune.Visible = $false
    $tab3.Controls.Add($progressBarIntune)

    $lblIntuneStats = New-Object System.Windows.Forms.Label
    $lblIntuneStats.Text = ""
    $lblIntuneStats.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblIntuneStats.ForeColor = $colorGray
    $lblIntuneStats.Location = New-Object System.Drawing.Point(10, 78)
    $lblIntuneStats.Size = New-Object System.Drawing.Size(1260, 20)
    $tab3.Controls.Add($lblIntuneStats)

    # DataGridView — Résultats Intune
    $dgvIntune = New-Object System.Windows.Forms.DataGridView
    $dgvIntune.Location = New-Object System.Drawing.Point(10, 102)
    $dgvIntune.Size = New-Object System.Drawing.Size(1260, 560)
    $dgvIntune.BackgroundColor = $colorWhite
    $dgvIntune.BorderStyle = "None"
    $dgvIntune.GridColor = [System.Drawing.Color]::FromArgb(222, 226, 230)
    $dgvIntune.AllowUserToAddRows = $false
    $dgvIntune.AllowUserToDeleteRows = $false
    $dgvIntune.ReadOnly = $true
    $dgvIntune.SelectionMode = "FullRowSelect"
    $dgvIntune.MultiSelect = $false
    $dgvIntune.RowHeadersVisible = $false
    $dgvIntune.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dgvIntune.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dgvIntune.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(233, 236, 239)
    $dgvIntune.ColumnHeadersHeight = 35
    $dgvIntune.RowTemplate.Height = 30
    $dgvIntune.EnableHeadersVisualStyles = $false
    $dgvIntune.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(200, 220, 255)
    $dgvIntune.DefaultCellStyle.SelectionForeColor = $colorDark
    $tab3.Controls.Add($dgvIntune)

    @(
        @{ Name = "IntuneCategory";   Header = "nested_group_audit.col_intune_category";   Width = 220 },
        @{ Name = "IntuneName";       Header = "nested_group_audit.col_intune_name";       Width = 350 },
        @{ Name = "IntunePlatform";   Header = "nested_group_audit.col_intune_platform";   Width = 120 },
        @{ Name = "IntuneIntent";     Header = "nested_group_audit.col_intune_intent";     Width = 120 },
        @{ Name = "IntuneAssignment"; Header = "nested_group_audit.col_intune_assignment"; Width = 120 },
        @{ Name = "IntuneId";         Header = "nested_group_audit.col_intune_id";         Width = 310 }
    ) | ForEach-Object {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $_.Name
        $col.HeaderText = Get-Text $_.Header
        $col.Width = $_.Width
        $dgvIntune.Columns.Add($col) | Out-Null
    }

    # =====================================================================
    # ONGLET 4 — Actions de remédiation
    # =====================================================================
    $tab4 = New-Object System.Windows.Forms.TabPage
    $tab4.Text = Get-Text "nested_group_audit.tab_actions"
    $tab4.BackColor = $colorBg
    $tabControl.TabPages.Add($tab4)

    $lblActionsGroup = New-Object System.Windows.Forms.Label
    $lblActionsGroup.Text = Get-Text "nested_group_audit.no_group_selected"
    $lblActionsGroup.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblActionsGroup.ForeColor = $colorDark
    $lblActionsGroup.Location = New-Object System.Drawing.Point(15, 15)
    $lblActionsGroup.Size = New-Object System.Drawing.Size(1250, 25)
    $tab4.Controls.Add($lblActionsGroup)

    # --- Panneau création groupe Users ---
    $panelCreateUser = New-Object System.Windows.Forms.GroupBox
    $panelCreateUser.Text = Get-Text "nested_group_audit.grp_create_user"
    $panelCreateUser.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $panelCreateUser.ForeColor = $colorBlue
    $panelCreateUser.Location = New-Object System.Drawing.Point(15, 55)
    $panelCreateUser.Size = New-Object System.Drawing.Size(610, 200)
    $tab4.Controls.Add($panelCreateUser)

    $lblUserGroupName = New-Object System.Windows.Forms.Label
    $lblUserGroupName.Text = Get-Text "nested_group_audit.lbl_new_group_name"
    $lblUserGroupName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblUserGroupName.ForeColor = $colorDark
    $lblUserGroupName.Location = New-Object System.Drawing.Point(15, 35)
    $lblUserGroupName.Size = New-Object System.Drawing.Size(150, 20)
    $panelCreateUser.Controls.Add($lblUserGroupName)

    $txtUserGroupName = New-Object System.Windows.Forms.TextBox
    $txtUserGroupName.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $txtUserGroupName.Location = New-Object System.Drawing.Point(170, 32)
    $txtUserGroupName.Size = New-Object System.Drawing.Size(420, 28)
    $panelCreateUser.Controls.Add($txtUserGroupName)

    $lblUserGroupDesc = New-Object System.Windows.Forms.Label
    $lblUserGroupDesc.Text = Get-Text "nested_group_audit.lbl_description"
    $lblUserGroupDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblUserGroupDesc.ForeColor = $colorDark
    $lblUserGroupDesc.Location = New-Object System.Drawing.Point(15, 75)
    $lblUserGroupDesc.Size = New-Object System.Drawing.Size(150, 20)
    $panelCreateUser.Controls.Add($lblUserGroupDesc)

    $txtUserGroupDesc = New-Object System.Windows.Forms.TextBox
    $txtUserGroupDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtUserGroupDesc.Location = New-Object System.Drawing.Point(170, 72)
    $txtUserGroupDesc.Size = New-Object System.Drawing.Size(420, 28)
    $panelCreateUser.Controls.Add($txtUserGroupDesc)

    $lblUserCount = New-Object System.Windows.Forms.Label
    $lblUserCount.Text = ""
    $lblUserCount.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblUserCount.ForeColor = $colorGray
    $lblUserCount.Location = New-Object System.Drawing.Point(15, 115)
    $lblUserCount.Size = New-Object System.Drawing.Size(400, 20)
    $panelCreateUser.Controls.Add($lblUserCount)

    $btnCreateUserGroup = New-Object System.Windows.Forms.Button
    $btnCreateUserGroup.Text = Get-Text "nested_group_audit.btn_create_group"
    $btnCreateUserGroup.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnCreateUserGroup.Location = New-Object System.Drawing.Point(170, 145)
    $btnCreateUserGroup.Size = New-Object System.Drawing.Size(420, 38)
    $btnCreateUserGroup.BackColor = $colorBlue
    $btnCreateUserGroup.ForeColor = $colorWhite
    $btnCreateUserGroup.FlatStyle = "Flat"
    $btnCreateUserGroup.FlatAppearance.BorderSize = 0
    $btnCreateUserGroup.Enabled = $false
    $panelCreateUser.Controls.Add($btnCreateUserGroup)

    # --- Panneau création groupe Devices ---
    $panelCreateDevice = New-Object System.Windows.Forms.GroupBox
    $panelCreateDevice.Text = Get-Text "nested_group_audit.grp_create_device"
    $panelCreateDevice.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $panelCreateDevice.ForeColor = $colorOrange
    $panelCreateDevice.Location = New-Object System.Drawing.Point(645, 55)
    $panelCreateDevice.Size = New-Object System.Drawing.Size(610, 200)
    $tab4.Controls.Add($panelCreateDevice)

    $lblDeviceGroupName = New-Object System.Windows.Forms.Label
    $lblDeviceGroupName.Text = Get-Text "nested_group_audit.lbl_new_group_name"
    $lblDeviceGroupName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblDeviceGroupName.ForeColor = $colorDark
    $lblDeviceGroupName.Location = New-Object System.Drawing.Point(15, 35)
    $lblDeviceGroupName.Size = New-Object System.Drawing.Size(150, 20)
    $panelCreateDevice.Controls.Add($lblDeviceGroupName)

    $txtDeviceGroupName = New-Object System.Windows.Forms.TextBox
    $txtDeviceGroupName.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $txtDeviceGroupName.Location = New-Object System.Drawing.Point(170, 32)
    $txtDeviceGroupName.Size = New-Object System.Drawing.Size(420, 28)
    $panelCreateDevice.Controls.Add($txtDeviceGroupName)

    $lblDeviceGroupDesc = New-Object System.Windows.Forms.Label
    $lblDeviceGroupDesc.Text = Get-Text "nested_group_audit.lbl_description"
    $lblDeviceGroupDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblDeviceGroupDesc.ForeColor = $colorDark
    $lblDeviceGroupDesc.Location = New-Object System.Drawing.Point(15, 75)
    $lblDeviceGroupDesc.Size = New-Object System.Drawing.Size(150, 20)
    $panelCreateDevice.Controls.Add($lblDeviceGroupDesc)

    $txtDeviceGroupDesc = New-Object System.Windows.Forms.TextBox
    $txtDeviceGroupDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtDeviceGroupDesc.Location = New-Object System.Drawing.Point(170, 72)
    $txtDeviceGroupDesc.Size = New-Object System.Drawing.Size(420, 28)
    $panelCreateDevice.Controls.Add($txtDeviceGroupDesc)

    $lblDeviceCount = New-Object System.Windows.Forms.Label
    $lblDeviceCount.Text = ""
    $lblDeviceCount.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblDeviceCount.ForeColor = $colorGray
    $lblDeviceCount.Location = New-Object System.Drawing.Point(15, 115)
    $lblDeviceCount.Size = New-Object System.Drawing.Size(400, 20)
    $panelCreateDevice.Controls.Add($lblDeviceCount)

    $btnCreateDeviceGroup = New-Object System.Windows.Forms.Button
    $btnCreateDeviceGroup.Text = Get-Text "nested_group_audit.btn_create_group"
    $btnCreateDeviceGroup.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnCreateDeviceGroup.Location = New-Object System.Drawing.Point(170, 145)
    $btnCreateDeviceGroup.Size = New-Object System.Drawing.Size(420, 38)
    $btnCreateDeviceGroup.BackColor = $colorOrange
    $btnCreateDeviceGroup.ForeColor = $colorWhite
    $btnCreateDeviceGroup.FlatStyle = "Flat"
    $btnCreateDeviceGroup.FlatAppearance.BorderSize = 0
    $btnCreateDeviceGroup.Enabled = $false
    $panelCreateDevice.Controls.Add($btnCreateDeviceGroup)

    # --- Log des actions ---
    $lblActionsLog = New-Object System.Windows.Forms.Label
    $lblActionsLog.Text = Get-Text "nested_group_audit.lbl_actions_log"
    $lblActionsLog.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblActionsLog.ForeColor = $colorDark
    $lblActionsLog.Location = New-Object System.Drawing.Point(15, 270)
    $lblActionsLog.Size = New-Object System.Drawing.Size(300, 22)
    $tab4.Controls.Add($lblActionsLog)

    $txtActionsLog = New-Object System.Windows.Forms.TextBox
    $txtActionsLog.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtActionsLog.Location = New-Object System.Drawing.Point(15, 298)
    $txtActionsLog.Size = New-Object System.Drawing.Size(1240, 360)
    $txtActionsLog.Multiline = $true
    $txtActionsLog.ScrollBars = "Vertical"
    $txtActionsLog.ReadOnly = $true
    $txtActionsLog.BackColor = $colorWhite
    $tab4.Controls.Add($txtActionsLog)

    # =================================================================
    # Pied de page
    # =================================================================
    $lblFooter = New-Object System.Windows.Forms.Label
    $lblFooter.Text = Get-Text "nested_group_audit.footer"
    $lblFooter.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblFooter.ForeColor = $colorLightGray
    $lblFooter.Location = New-Object System.Drawing.Point(15, 810)
    $lblFooter.Size = New-Object System.Drawing.Size(500, 20)
    $lblFooter.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($lblFooter)

    # =================================================================
    # FONCTIONS INTERNES
    # =================================================================

    # --- Fonction utilitaire : appel Graph paginé via Invoke-MgGraphRequest ---
    function Invoke-GraphRequestPaginated {
        param(
            [string]$Uri,
            [string]$Label = "Graph"
        )
        $allResults = @()
        $currentUri = $Uri
        try {
            while ($null -ne $currentUri) {
                $response = Invoke-MgGraphRequest -Method GET -Uri $currentUri -ErrorAction Stop
                if ($response.value) {
                    $allResults += $response.value
                }
                $currentUri = $response.'@odata.nextLink'
                if ($currentUri) {
                    Start-Sleep -Milliseconds 100  # Anti-throttling
                }
            }
        }
        catch {
            Write-Log -Level "WARNING" -Action "NESTED_AUDIT" -Message "Erreur pagination $Label : $($_.Exception.Message)"
        }
        return $allResults
    }

    # --- Fonction utilitaire : envoie un lot de requêtes Graph via /$batch (max 20) ---
    # Retourne un hashtable { groupId = @(membres) } pour chaque requête du lot.
    function Invoke-GraphBatchMembers {
        param(
            [array]$GroupIds
        )

        $results = @{}

        # Construire le corps du batch — chaque requête GET /groups/{id}/members
        $requests = @()
        for ($i = 0; $i -lt $GroupIds.Count; $i++) {
            $requests += @{
                id     = "$i"
                method = "GET"
                url    = "/groups/$($GroupIds[$i])/members?`$top=500&`$select=id"
                headers = @{ "ConsistencyLevel" = "eventual" }
            }
        }

        $batchBody = @{ requests = $requests } | ConvertTo-Json -Depth 5 -Compress

        try {
            $batchResponse = Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/`$batch" `
                -Body $batchBody `
                -ContentType "application/json" `
                -ErrorAction Stop

            foreach ($resp in $batchResponse.responses) {
                $idx = [int]$resp.id
                $groupId = $GroupIds[$idx]

                if ($resp.status -eq 200 -and $resp.body.value) {
                    $results[$groupId] = $resp.body.value
                }
                else {
                    $results[$groupId] = @()
                }
            }
        }
        catch {
            Write-Log -Level "WARNING" -Action "NESTED_AUDIT" -Message "Erreur batch members : $($_.Exception.Message)"
            # Fallback : marquer tous les groupes du lot comme vides
            foreach ($gid in $GroupIds) {
                if (-not $results.ContainsKey($gid)) {
                    $results[$gid] = @()
                }
            }
        }

        return $results
    }

    # --- Fonction : écrire dans le log GUI de l'onglet 4 ---
    function Write-ActionLog {
        param([string]$Message)
        $timestamp = Get-Date -Format "HH:mm:ss"
        $txtActionsLog.AppendText("[$timestamp] $Message`r`n")
    }

    # --- Fonction : mettre à jour les labels quand un groupe est sélectionné ---
    function Update-SelectedGroupUI {
        param([PSCustomObject]$Group)

        $script:SelectedGroup = $Group

        # Onglet 2 — Titre
        $lblMembersGroup.Text = (Get-Text "nested_group_audit.selected_group") -f $Group.DisplayName

        # Onglet 3 — Titre
        $lblIntuneGroup.Text = (Get-Text "nested_group_audit.selected_group") -f $Group.DisplayName
        $btnScanIntune.Enabled = $true

        # Onglet 4 — Titre + noms de groupes pré-remplis
        $lblActionsGroup.Text = (Get-Text "nested_group_audit.selected_group") -f $Group.DisplayName
        $txtUserGroupName.Text = "$($Group.DisplayName)_User"
        $txtDeviceGroupName.Text = "$($Group.DisplayName)_Device"
        $txtUserGroupDesc.Text = (Get-Text "nested_group_audit.auto_desc_user") -f $Group.DisplayName
        $txtDeviceGroupDesc.Text = (Get-Text "nested_group_audit.auto_desc_device") -f $Group.DisplayName
        $lblUserCount.Text = (Get-Text "nested_group_audit.members_to_transfer") -f $Group.UserCount
        $lblDeviceCount.Text = (Get-Text "nested_group_audit.members_to_transfer") -f $Group.DeviceCount
        $btnCreateUserGroup.Enabled = ($Group.UserCount -gt 0)
        $btnCreateDeviceGroup.Enabled = ($Group.DeviceCount -gt 0)
    }

    # --- Fonction : charger les membres détaillés d'un groupe ---
    function Import-GroupMembers {
        param([string]$GroupId)

        $dgvUsers.Rows.Clear()
        $dgvDevices.Rows.Clear()
        $script:SelectedGroupMembers = @{ Users = @(); Devices = @() }

        try {
            Write-Log -Level "INFO" -Action "NESTED_AUDIT" -Message "Chargement des membres du groupe $GroupId."
            $members = Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop

            foreach ($member in $members) {
                $memberType = $member.AdditionalProperties.'@odata.type'

                switch ($memberType) {
                    '#microsoft.graph.user' {
                        $displayName = $member.AdditionalProperties.displayName
                        $upn = $member.AdditionalProperties.userPrincipalName
                        $jobTitle = $member.AdditionalProperties.jobTitle

                        $script:SelectedGroupMembers.Users += [PSCustomObject]@{
                            Id          = $member.Id
                            DisplayName = $displayName
                            UPN         = $upn
                            JobTitle    = $jobTitle
                        }
                        $dgvUsers.Rows.Add($displayName, $upn, $jobTitle) | Out-Null
                    }
                    '#microsoft.graph.device' {
                        $displayName = $member.AdditionalProperties.displayName
                        $os = $member.AdditionalProperties.operatingSystem
                        $deviceId = $member.AdditionalProperties.deviceId

                        $script:SelectedGroupMembers.Devices += [PSCustomObject]@{
                            Id          = $member.Id
                            DisplayName = $displayName
                            OS          = $os
                            DeviceId    = $deviceId
                        }
                        $dgvDevices.Rows.Add($displayName, $os, $deviceId) | Out-Null
                    }
                }
            }

            $lblUsersTitle.Text = (Get-Text "nested_group_audit.lbl_users_count") -f $script:SelectedGroupMembers.Users.Count
            $lblDevicesTitle.Text = (Get-Text "nested_group_audit.lbl_devices_count") -f $script:SelectedGroupMembers.Devices.Count

            Write-Log -Level "INFO" -Action "NESTED_AUDIT" -Message "Membres chargés : $($script:SelectedGroupMembers.Users.Count) users, $($script:SelectedGroupMembers.Devices.Count) devices."
        }
        catch {
            Write-Log -Level "ERROR" -Action "NESTED_AUDIT" -Message "Erreur chargement membres : $($_.Exception.Message)"
            Show-ResultDialog -Titre (Get-Text "nested_group_audit.error_title") -Message ((Get-Text "nested_group_audit.error_load_members") -f $_.Exception.Message) -IsSuccess $false
        }
    }

    # --- Fonction : scanner les policies/apps Intune pour un groupId ---
    function Invoke-IntuneScan {
        param([string]$GroupId)

        $dgvIntune.Rows.Clear()
        $script:IntuneAssignments = @()
        $progressBarIntune.Visible = $true
        $progressBarIntune.Value = 0

        # Définition des catégories Intune à scanner
        $intuneCategories = @(
            @{ Name = (Get-Text "nested_group_audit.cat_apps");                Uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=(isAssigned eq true)&`$expand=Assignments";  HasExpand = $true; NameProp = "displayName" },
            @{ Name = (Get-Text "nested_group_audit.cat_app_config");          Uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations?`$expand=Assignments";                   HasExpand = $true; NameProp = "displayName" },
            @{ Name = (Get-Text "nested_group_audit.cat_config_policies");     Uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$expand=Assignments";                        HasExpand = $true; NameProp = "name" },
            @{ Name = (Get-Text "nested_group_audit.cat_device_config");       Uri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?`$expand=Assignments";                         HasExpand = $true; NameProp = "displayName" },
            @{ Name = (Get-Text "nested_group_audit.cat_compliance");          Uri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies?`$expand=Assignments";                     HasExpand = $true; NameProp = "displayName" },
            @{ Name = (Get-Text "nested_group_audit.cat_gpo");                 Uri = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations?`$expand=Assignments";                    HasExpand = $true; NameProp = "displayName" },
            @{ Name = (Get-Text "nested_group_audit.cat_autopilot");           Uri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles?`$expand=Assignments";           HasExpand = $true; NameProp = "displayName" },
            @{ Name = (Get-Text "nested_group_audit.cat_feature_updates");     Uri = "https://graph.microsoft.com/beta/deviceManagement/windowsFeatureUpdateProfiles?`$expand=Assignments";                 HasExpand = $true; NameProp = "displayName" },
            @{ Name = (Get-Text "nested_group_audit.cat_quality_updates");     Uri = "https://graph.microsoft.com/beta/deviceManagement/windowsQualityUpdateProfiles?`$expand=Assignments";                 HasExpand = $true; NameProp = "displayName" },
            @{ Name = (Get-Text "nested_group_audit.cat_driver_updates");      Uri = "https://graph.microsoft.com/beta/deviceManagement/windowsDriverUpdateProfiles?`$expand=Assignments";                  HasExpand = $true; NameProp = "displayName" },
            @{ Name = (Get-Text "nested_group_audit.cat_remediation");         Uri = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts";                                               HasExpand = $false; NameProp = "displayName" },
            @{ Name = (Get-Text "nested_group_audit.cat_platform_scripts");    Uri = "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts";                                           HasExpand = $false; NameProp = "displayName" }
        )

        $totalCategories = $intuneCategories.Count
        $catIndex = 0
        $totalFound = 0

        foreach ($category in $intuneCategories) {
            $catIndex++
            $progressBarIntune.Value = [math]::Min(100, [math]::Round(($catIndex / $totalCategories) * 100))
            [System.Windows.Forms.Application]::DoEvents()

            try {
                Write-Log -Level "INFO" -Action "NESTED_AUDIT" -Message "Scan Intune : $($category.Name)..."

                $items = Invoke-GraphRequestPaginated -Uri $category.Uri -Label $category.Name

                foreach ($item in $items) {
                    $assignments = $null

                    if ($category.HasExpand) {
                        # Les assignments sont déjà incluses via $expand
                        $assignments = $item.assignments
                    }
                    else {
                        # Récupérer les assignments via sous-requête
                        try {
                            $assignUri = "https://graph.microsoft.com/beta/deviceManagement/$($category.Uri.Split('/')[-1])/$($item.id)/assignments"
                            # Recomposer l'URI correctement
                            if ($category.Uri -match "deviceHealthScripts") {
                                $assignUri = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts/$($item.id)/assignments"
                            }
                            elseif ($category.Uri -match "deviceManagementScripts") {
                                $assignUri = "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts/$($item.id)/assignments"
                            }
                            $assignResponse = Invoke-MgGraphRequest -Method GET -Uri $assignUri -ErrorAction Stop
                            $assignments = $assignResponse.value
                        }
                        catch {
                            # Certaines policies peuvent ne pas avoir d'assignments accessibles — ignorer
                            continue
                        }
                    }

                    if ($null -eq $assignments) { continue }

                    # Vérifier si le groupId est dans les assignments
                    foreach ($assignment in $assignments) {
                        $targetGroupId = $assignment.target.groupId
                        if ($targetGroupId -eq $GroupId) {
                            # Déterminer le type d'assignment (Inclusion/Exclusion)
                            $assignType = switch ($assignment.target.'@odata.type') {
                                '#microsoft.graph.exclusionGroupAssignmentTarget' { Get-Text "nested_group_audit.assignment_excluded" }
                                '#microsoft.graph.groupAssignmentTarget'          { Get-Text "nested_group_audit.assignment_included" }
                                default                                           { Get-Text "nested_group_audit.assignment_included" }
                            }

                            # Déterminer l'intent (pour les apps)
                            $intent = if ($assignment.intent) { $assignment.intent } else { "-" }

                            # Déterminer la plateforme si disponible
                            $platform = "-"
                            if ($item.platforms)        { $platform = $item.platforms }
                            elseif ($item.'@odata.type') {
                                if ($item.'@odata.type' -match 'windows')  { $platform = "Windows" }
                                elseif ($item.'@odata.type' -match 'ios')  { $platform = "iOS" }
                                elseif ($item.'@odata.type' -match 'android') { $platform = "Android" }
                                elseif ($item.'@odata.type' -match 'macOS')   { $platform = "macOS" }
                            }

                            $itemName = if ($item.($category.NameProp)) { $item.($category.NameProp) } else { $item.displayName }

                            $record = [PSCustomObject]@{
                                Category   = $category.Name
                                Name       = $itemName
                                Platform   = $platform
                                Intent     = $intent
                                Assignment = $assignType
                                Id         = $item.id
                            }

                            $script:IntuneAssignments += $record
                            $dgvIntune.Rows.Add($category.Name, $itemName, $platform, $intent, $assignType, $item.id) | Out-Null
                            $totalFound++
                        }
                    }
                }
            }
            catch {
                Write-Log -Level "WARNING" -Action "NESTED_AUDIT" -Message "Erreur scan catégorie $($category.Name) : $($_.Exception.Message)"
            }
        }

        $progressBarIntune.Visible = $false
        $lblIntuneStats.Text = (Get-Text "nested_group_audit.intune_stats") -f $totalFound, $totalCategories
        $btnExportIntune.Enabled = ($totalFound -gt 0)

        Write-Log -Level "SUCCESS" -Action "NESTED_AUDIT" -Message "Scan Intune terminé : $totalFound résultat(s) dans $totalCategories catégorie(s)."

        if ($totalFound -eq 0) {
            Show-ResultDialog -Titre (Get-Text "nested_group_audit.intune_scan_title") -Message (Get-Text "nested_group_audit.intune_no_results") -IsSuccess $true
        }
    }

    # --- Fonction : créer un groupe de sécurité et y ajouter des membres ---
    function New-SecurityGroupWithMembers {
        param(
            [string]$GroupName,
            [string]$Description,
            [array]$MemberIds,
            [string]$MemberType  # "User" ou "Device"
        )

        Write-ActionLog ((Get-Text "nested_group_audit.log_creating_group") -f $GroupName)

        try {
            # Créer le groupe de sécurité
            $groupBody = @{
                displayName     = $GroupName
                description     = $Description
                mailEnabled     = $false
                mailNickname    = ($GroupName -replace '[^a-zA-Z0-9]', '')
                securityEnabled = $true
                groupTypes      = @()
            }

            $newGroup = New-MgGroup -BodyParameter $groupBody -ErrorAction Stop
            Write-Log -Level "SUCCESS" -Action "NESTED_AUDIT" -Message "Groupe '$GroupName' créé (Id: $($newGroup.Id))."
            Write-ActionLog ((Get-Text "nested_group_audit.log_group_created") -f $GroupName, $newGroup.Id)

            # Ajouter les membres
            $addedCount = 0
            $errorCount = 0

            foreach ($memberId in $MemberIds) {
                try {
                    $memberRef = @{
                        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$memberId"
                    }
                    New-MgGroupMemberByRef -GroupId $newGroup.Id -BodyParameter $memberRef -ErrorAction Stop
                    $addedCount++
                }
                catch {
                    $errorCount++
                    Write-Log -Level "WARNING" -Action "NESTED_AUDIT" -Message "Erreur ajout membre $memberId au groupe $GroupName : $($_.Exception.Message)"
                }

                # Mise à jour visuelle
                if ($addedCount % 10 -eq 0) {
                    Write-ActionLog ((Get-Text "nested_group_audit.log_members_progress") -f $addedCount, $MemberIds.Count)
                    [System.Windows.Forms.Application]::DoEvents()
                }
            }

            Write-ActionLog ((Get-Text "nested_group_audit.log_transfer_complete") -f $addedCount, $MemberIds.Count, $errorCount)
            Write-Log -Level "SUCCESS" -Action "NESTED_AUDIT" -Message "Transfert vers '$GroupName' terminé : $addedCount/$($MemberIds.Count) membres, $errorCount erreur(s)."

            return [PSCustomObject]@{ Success = $true; GroupId = $newGroup.Id; Added = $addedCount; Errors = $errorCount }
        }
        catch {
            $errMsg = $_.Exception.Message
            Write-Log -Level "ERROR" -Action "NESTED_AUDIT" -Message "Erreur création groupe '$GroupName' : $errMsg"
            Write-ActionLog ((Get-Text "nested_group_audit.log_error") -f $errMsg)
            return [PSCustomObject]@{ Success = $false; GroupId = $null; Added = 0; Errors = 0 }
        }
    }

    # --- Fonction : filtrer le DataGridView des groupes ---
    function Update-GroupsFilter {
        $filterText = $txtFilter.Text.Trim().ToLower()
        $dgvGroups.Rows.Clear()

        foreach ($group in $script:NestedGroupData) {
            if ([string]::IsNullOrWhiteSpace($filterText) -or
                $group.DisplayName.ToLower().Contains($filterText) -or
                $group.GroupId.ToLower().Contains($filterText)) {

                $dgvGroups.Rows.Add(
                    $group.DisplayName,
                    $group.GroupId,
                    $group.GroupType,
                    $group.Membership,
                    $group.UserCount,
                    $group.DeviceCount,
                    $group.OtherCount,
                    $group.TotalCount
                ) | Out-Null
            }
        }
    }

    # --- Fonction : export CSV ---
    function Export-DataToCsv {
        param(
            [array]$Data,
            [string]$DefaultFileName
        )

        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV (*.csv)|*.csv"
        $saveDialog.FileName = $DefaultFileName
        $saveDialog.Title = Get-Text "nested_group_audit.export_title"

        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $Data | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
                Show-ResultDialog -Titre (Get-Text "nested_group_audit.export_success_title") -Message ((Get-Text "nested_group_audit.export_success_msg") -f $saveDialog.FileName, $Data.Count) -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "NESTED_AUDIT" -Message "Export CSV : $($saveDialog.FileName) ($($Data.Count) ligne(s))."
            }
            catch {
                Show-ResultDialog -Titre (Get-Text "nested_group_audit.error_title") -Message ((Get-Text "nested_group_audit.export_error_msg") -f $_.Exception.Message) -IsSuccess $false
            }
        }
    }

    # =================================================================
    # ÉVÉNEMENTS
    # =================================================================

    # --- Bouton Scanner les groupes ---
    $btnScan.Add_Click({
        $btnScan.Enabled = $false
        $btnCancelScan.Enabled = $true
        $btnExportScan.Enabled = $false
        $script:ScanCancelled = $false
        $dgvGroups.Rows.Clear()
        $script:NestedGroupData = @()
        $progressBar.Visible = $true
        $progressBar.Value = 0

        Write-Log -Level "INFO" -Action "NESTED_AUDIT" -Message "Démarrage du scan des groupes Entra (mode batch)."

        try {
            # Étape 1 : Récupérer les groupes de sécurité (pré-filtrage)
            $lblScanStats.Text = Get-Text "nested_group_audit.scanning"
            [System.Windows.Forms.Application]::DoEvents()

            $allGroups = Get-MgGroup -All -Property Id, DisplayName, Description, GroupTypes, Mail, MembershipRule -ErrorAction Stop
            $totalGroups = $allGroups.Count

            # Indexer les groupes par ID pour accès rapide
            $groupIndex = @{}
            foreach ($g in $allGroups) {
                $groupIndex[$g.Id] = $g
            }

            $lblScanStats.Text = (Get-Text "nested_group_audit.scanning_progress") -f 0, $totalGroups
            [System.Windows.Forms.Application]::DoEvents()

            # Étape 2 : Scanner les membres par batch de 20
            $batchSize = 20
            $counter = 0
            $nestedCount = 0
            $groupIds = $allGroups | ForEach-Object { $_.Id }

            for ($batchStart = 0; $batchStart -lt $groupIds.Count; $batchStart += $batchSize) {
                if ($script:ScanCancelled) {
                    Write-Log -Level "INFO" -Action "NESTED_AUDIT" -Message "Scan annulé par l'utilisateur."
                    break
                }

                # Découper le lot courant (max 20 groupes)
                $batchEnd = [math]::Min($batchStart + $batchSize - 1, $groupIds.Count - 1)
                $currentBatch = $groupIds[$batchStart..$batchEnd]

                # Appel batch Graph
                $batchResults = Invoke-GraphBatchMembers -GroupIds $currentBatch

                # Classifier les résultats de chaque groupe du lot
                foreach ($groupId in $currentBatch) {
                    $counter++
                    $members = $batchResults[$groupId]

                    if ($null -eq $members -or $members.Count -eq 0) { continue }

                    $userCount = 0
                    $deviceCount = 0
                    $otherCount = 0

                    foreach ($member in $members) {
                        switch ($member.'@odata.type') {
                            '#microsoft.graph.user'   { $userCount++ }
                            '#microsoft.graph.device' { $deviceCount++ }
                            default                   { $otherCount++ }
                        }
                    }

                    # Groupe nested = Users ET Devices
                    if ($userCount -gt 0 -and $deviceCount -gt 0) {
                        $nestedCount++
                        $group = $groupIndex[$groupId]
                        $groupType = if ($group.GroupTypes -contains 'Unified') { 'Microsoft 365' } else { 'Security' }
                        $membership = if ($group.MembershipRule) { Get-Text "nested_group_audit.dynamic" } else { Get-Text "nested_group_audit.assigned" }

                        $record = [PSCustomObject]@{
                            DisplayName = $group.DisplayName
                            GroupId     = $group.Id
                            GroupType   = $groupType
                            Membership  = $membership
                            UserCount   = $userCount
                            DeviceCount = $deviceCount
                            OtherCount  = $otherCount
                            TotalCount  = $members.Count
                        }

                        $script:NestedGroupData += $record
                        $dgvGroups.Rows.Add($group.DisplayName, $group.Id, $groupType, $membership, $userCount, $deviceCount, $otherCount, $members.Count) | Out-Null
                    }
                }

                # Mise à jour progression après chaque lot
                $progressBar.Value = [math]::Min(100, [math]::Round(($counter / $totalGroups) * 100))
                $lblScanStats.Text = (Get-Text "nested_group_audit.scanning_progress") -f $counter, $totalGroups
                [System.Windows.Forms.Application]::DoEvents()

                # Anti-throttling entre les lots
                Start-Sleep -Milliseconds 150
            }

            $progressBar.Visible = $false
            $lblScanStats.Text = (Get-Text "nested_group_audit.scan_complete") -f $totalGroups, $nestedCount
            $btnExportScan.Enabled = ($nestedCount -gt 0)

            Write-Log -Level "SUCCESS" -Action "NESTED_AUDIT" -Message "Scan batch terminé : $totalGroups groupes analysés, $nestedCount nested trouvés."
        }
        catch {
            $progressBar.Visible = $false
            Write-Log -Level "ERROR" -Action "NESTED_AUDIT" -Message "Erreur scan groupes : $($_.Exception.Message)"
            Show-ResultDialog -Titre (Get-Text "nested_group_audit.error_title") -Message ((Get-Text "nested_group_audit.error_scan") -f $_.Exception.Message) -IsSuccess $false
        }
        finally {
            $btnScan.Enabled = $true
            $btnCancelScan.Enabled = $false
        }
    })

    # --- Bouton Annuler le scan ---
    $btnCancelScan.Add_Click({
        $script:ScanCancelled = $true
    })

    # --- Filtre texte ---
    $txtFilter.Add_TextChanged({
        Update-GroupsFilter
    })

    # --- Sélection d'un groupe dans le DataGridView ---
    $dgvGroups.Add_SelectionChanged({
        if ($dgvGroups.SelectedRows.Count -gt 0) {
            $selectedRow = $dgvGroups.SelectedRows[0]
            $groupId = $selectedRow.Cells["GroupId"].Value

            # Retrouver le groupe dans les données
            $group = $script:NestedGroupData | Where-Object { $_.GroupId -eq $groupId }
            if ($group) {
                Update-SelectedGroupUI -Group $group

                # Charger les membres en arrière-plan
                Import-GroupMembers -GroupId $groupId
            }
        }
    })

    # --- Bouton Scanner Intune ---
    $btnScanIntune.Add_Click({
        if ($null -eq $script:SelectedGroup) { return }

        $btnScanIntune.Enabled = $false
        Invoke-IntuneScan -GroupId $script:SelectedGroup.GroupId
        $btnScanIntune.Enabled = $true
    })

    # --- Bouton Export scan groupes ---
    $btnExportScan.Add_Click({
        if ($script:NestedGroupData.Count -gt 0) {
            Export-DataToCsv -Data $script:NestedGroupData -DefaultFileName "NestedGroups_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        }
    })

    # --- Bouton Export Intune ---
    $btnExportIntune.Add_Click({
        if ($script:IntuneAssignments.Count -gt 0) {
            Export-DataToCsv -Data $script:IntuneAssignments -DefaultFileName "IntuneAssignments_$($script:SelectedGroup.DisplayName)_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        }
    })

    # --- Bouton Créer groupe Users ---
    $btnCreateUserGroup.Add_Click({
        if ($null -eq $script:SelectedGroup) { return }
        if ($script:SelectedGroupMembers.Users.Count -eq 0) { return }

        $groupName = $txtUserGroupName.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($groupName)) {
            Show-ResultDialog -Titre (Get-Text "nested_group_audit.error_title") -Message (Get-Text "nested_group_audit.error_empty_name") -IsSuccess $false
            return
        }

        $confirmMsg = (Get-Text "nested_group_audit.confirm_create_user") -f $groupName, $script:SelectedGroupMembers.Users.Count
        $confirmed = Show-ConfirmDialog -Titre (Get-Text "nested_group_audit.confirm_title") -Message $confirmMsg
        if (-not $confirmed) { return }

        $btnCreateUserGroup.Enabled = $false
        $memberIds = $script:SelectedGroupMembers.Users | ForEach-Object { $_.Id }
        $result = New-SecurityGroupWithMembers -GroupName $groupName -Description $txtUserGroupDesc.Text -MemberIds $memberIds -MemberType "User"

        if ($result.Success) {
            Show-ResultDialog -Titre (Get-Text "nested_group_audit.success_title") -Message ((Get-Text "nested_group_audit.success_create_group") -f $groupName, $result.Added, $result.Errors) -IsSuccess $true
        }
        else {
            Show-ResultDialog -Titre (Get-Text "nested_group_audit.error_title") -Message (Get-Text "nested_group_audit.error_create_group") -IsSuccess $false
        }
        $btnCreateUserGroup.Enabled = $true
    })

    # --- Bouton Créer groupe Devices ---
    $btnCreateDeviceGroup.Add_Click({
        if ($null -eq $script:SelectedGroup) { return }
        if ($script:SelectedGroupMembers.Devices.Count -eq 0) { return }

        $groupName = $txtDeviceGroupName.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($groupName)) {
            Show-ResultDialog -Titre (Get-Text "nested_group_audit.error_title") -Message (Get-Text "nested_group_audit.error_empty_name") -IsSuccess $false
            return
        }

        $confirmMsg = (Get-Text "nested_group_audit.confirm_create_device") -f $groupName, $script:SelectedGroupMembers.Devices.Count
        $confirmed = Show-ConfirmDialog -Titre (Get-Text "nested_group_audit.confirm_title") -Message $confirmMsg
        if (-not $confirmed) { return }

        $btnCreateDeviceGroup.Enabled = $false
        $memberIds = $script:SelectedGroupMembers.Devices | ForEach-Object { $_.Id }
        $result = New-SecurityGroupWithMembers -GroupName $groupName -Description $txtDeviceGroupDesc.Text -MemberIds $memberIds -MemberType "Device"

        if ($result.Success) {
            Show-ResultDialog -Titre (Get-Text "nested_group_audit.success_title") -Message ((Get-Text "nested_group_audit.success_create_group") -f $groupName, $result.Added, $result.Errors) -IsSuccess $true
        }
        else {
            Show-ResultDialog -Titre (Get-Text "nested_group_audit.error_title") -Message (Get-Text "nested_group_audit.error_create_group") -IsSuccess $false
        }
        $btnCreateDeviceGroup.Enabled = $true
    })

    # =================================================================
    # Affichage
    # =================================================================
    $form.ShowDialog() | Out-Null
    $form.Dispose()
}