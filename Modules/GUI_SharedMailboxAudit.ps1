<#
.FICHIER
    Modules/GUI_SharedMailboxAudit.ps1

.ROLE
    Module d'audit des boîtes aux lettres partagées (Shared Mailbox).
    Identifie les BAL partagées dont l'objet utilisateur associé est actif,
    vérifie les licences assignées et le dernier sign-in interactif.
    Permet de désactiver un ou plusieurs comptes utilisateurs liés aux BAL
    après validation manuelle.

    Objectifs :
    - Lister toutes les Shared Mailbox du tenant via Graph API
    - Filtrer par statut utilisateur (actif / inactif)
    - Filtrer par licence assignée (avec / sans licence)
    - Afficher le dernier sign-in interactif (lastSuccessfulSignInDateTime)
    - Alerter si un sign-in interactif a été détecté (validation manuelle requise)
    - Permettre la désactivation groupée des comptes utilisateurs liés aux BAL
      via multi-sélection (checkbox + Ctrl/Shift)

.DEPENDANCES
    - Core/Functions.ps1 (Write-Log, Show-ConfirmDialog, Show-ResultDialog)
    - Core/GraphAPI.ps1 (Disable-AzUser)
    - Core/Lang.ps1 (Get-Text)
    - Core/Connect.ps1 (Get-GraphConnectionStatus)
    - Module ExchangeOnlineManagement (Get-EXOMailbox)
    - Connexion Graph active avec scopes : User.Read.All, AuditLog.Read.All, Directory.Read.All
    - Permission API : User.ReadWrite.All (pour désactivation)

.PERMISSIONS_GRAPH
    Delegated :
        - User.Read.All (lecture des utilisateurs et mailboxSettings)
        - AuditLog.Read.All (lecture signInActivity)
        - Directory.Read.All (lecture des licences)
        - User.ReadWrite.All (désactivation du compte)

.AUTEUR
    [Equipe IT - M365 Monster]
#>

function Show-SharedMailboxAuditForm {
    <#
    .SYNOPSIS
        Affiche le formulaire d'audit des boîtes aux lettres partagées.
        Charge les BAL partagées depuis Graph API, permet le filtrage
        et la désactivation groupée des comptes utilisateurs associés.
    #>

    # =================================================================
    # Variables de travail du module
    # =================================================================
    $script:SharedMailboxData = @()          # Données brutes chargées depuis Graph
    $script:FilteredData = @()              # Données après filtrage
    $script:LicenseSkuMap = @{}             # Cache SKU ID → Nom convivial

    # =================================================================
    # Couleurs cohérentes avec M365 Monster
    # =================================================================
    $colorDark       = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $colorBg         = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $colorWhite      = [System.Drawing.Color]::White
    $colorGray       = [System.Drawing.Color]::FromArgb(108, 117, 125)
    $colorLightGray  = [System.Drawing.Color]::FromArgb(173, 181, 189)
    $colorGreen      = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $colorRed        = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $colorOrange     = [System.Drawing.Color]::FromArgb(253, 126, 20)
    $colorBlue       = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $colorPurple     = [System.Drawing.Color]::FromArgb(111, 66, 193)

    # =================================================================
    # Fenêtre principale
    # =================================================================
    $form = New-Object System.Windows.Forms.Form
    $form.Text = Get-Text "shared_mailbox_audit.title"
    $form.Size = New-Object System.Drawing.Size(1280, 820)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedSingle"
    $form.MaximizeBox = $false
    $form.BackColor = $colorBg

    # Icône si disponible
    $iconFile = Join-Path -Path $script:RootPath -ChildPath "Assets\M365Monster.ico"
    if (Test-Path $iconFile) { $form.Icon = New-Object System.Drawing.Icon($iconFile) }

    # =================================================================
    # En-tête (bandeau sombre)
    # =================================================================
    $panelHeader = New-Object System.Windows.Forms.Panel
    $panelHeader.Location = New-Object System.Drawing.Point(0, 0)
    $panelHeader.Size = New-Object System.Drawing.Size(1280, 80)
    $panelHeader.BackColor = $colorDark
    $form.Controls.Add($panelHeader)

    $lblTitre = New-Object System.Windows.Forms.Label
    $lblTitre.Text = Get-Text "shared_mailbox_audit.header_title"
    $lblTitre.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $lblTitre.ForeColor = $colorWhite
    $lblTitre.Location = New-Object System.Drawing.Point(20, 10)
    $lblTitre.Size = New-Object System.Drawing.Size(600, 30)
    $panelHeader.Controls.Add($lblTitre)

    $lblSubtitle = New-Object System.Windows.Forms.Label
    $lblSubtitle.Text = Get-Text "shared_mailbox_audit.header_subtitle"
    $lblSubtitle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSubtitle.ForeColor = $colorLightGray
    $lblSubtitle.Location = New-Object System.Drawing.Point(20, 45)
    $lblSubtitle.Size = New-Object System.Drawing.Size(800, 20)
    $panelHeader.Controls.Add($lblSubtitle)

    # =================================================================
    # Panneau de filtres
    # =================================================================
    $panelFilters = New-Object System.Windows.Forms.Panel
    $panelFilters.Location = New-Object System.Drawing.Point(15, 90)
    $panelFilters.Size = New-Object System.Drawing.Size(1245, 55)
    $panelFilters.BackColor = $colorWhite
    $panelFilters.BorderStyle = "None"
    $form.Controls.Add($panelFilters)

    # Filtre : Statut utilisateur
    $lblFilterStatus = New-Object System.Windows.Forms.Label
    $lblFilterStatus.Text = Get-Text "shared_mailbox_audit.filter_status"
    $lblFilterStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblFilterStatus.Location = New-Object System.Drawing.Point(15, 16)
    $lblFilterStatus.Size = New-Object System.Drawing.Size(120, 20)
    $panelFilters.Controls.Add($lblFilterStatus)

    $cboFilterStatus = New-Object System.Windows.Forms.ComboBox
    $cboFilterStatus.DropDownStyle = "DropDownList"
    $cboFilterStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cboFilterStatus.Location = New-Object System.Drawing.Point(140, 13)
    $cboFilterStatus.Size = New-Object System.Drawing.Size(160, 25)
    $cboFilterStatus.Items.AddRange(@(
        (Get-Text "shared_mailbox_audit.filter_all"),
        (Get-Text "shared_mailbox_audit.filter_active"),
        (Get-Text "shared_mailbox_audit.filter_disabled")
    ))
    $cboFilterStatus.SelectedIndex = 0
    $panelFilters.Controls.Add($cboFilterStatus)

    # Filtre : Licence
    $lblFilterLicense = New-Object System.Windows.Forms.Label
    $lblFilterLicense.Text = Get-Text "shared_mailbox_audit.filter_license"
    $lblFilterLicense.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblFilterLicense.Location = New-Object System.Drawing.Point(320, 16)
    $lblFilterLicense.Size = New-Object System.Drawing.Size(80, 20)
    $panelFilters.Controls.Add($lblFilterLicense)

    $cboFilterLicense = New-Object System.Windows.Forms.ComboBox
    $cboFilterLicense.DropDownStyle = "DropDownList"
    $cboFilterLicense.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cboFilterLicense.Location = New-Object System.Drawing.Point(405, 13)
    $cboFilterLicense.Size = New-Object System.Drawing.Size(180, 25)
    $cboFilterLicense.Items.AddRange(@(
        (Get-Text "shared_mailbox_audit.filter_all"),
        (Get-Text "shared_mailbox_audit.filter_licensed"),
        (Get-Text "shared_mailbox_audit.filter_unlicensed")
    ))
    $cboFilterLicense.SelectedIndex = 0
    $panelFilters.Controls.Add($cboFilterLicense)

    # Filtre : Sign-in interactif
    $lblFilterSignIn = New-Object System.Windows.Forms.Label
    $lblFilterSignIn.Text = Get-Text "shared_mailbox_audit.filter_signin"
    $lblFilterSignIn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblFilterSignIn.Location = New-Object System.Drawing.Point(605, 16)
    $lblFilterSignIn.Size = New-Object System.Drawing.Size(100, 20)
    $panelFilters.Controls.Add($lblFilterSignIn)

    $cboFilterSignIn = New-Object System.Windows.Forms.ComboBox
    $cboFilterSignIn.DropDownStyle = "DropDownList"
    $cboFilterSignIn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cboFilterSignIn.Location = New-Object System.Drawing.Point(710, 13)
    $cboFilterSignIn.Size = New-Object System.Drawing.Size(200, 25)
    $cboFilterSignIn.Items.AddRange(@(
        (Get-Text "shared_mailbox_audit.filter_all"),
        (Get-Text "shared_mailbox_audit.filter_signin_never"),
        (Get-Text "shared_mailbox_audit.filter_signin_detected")
    ))
    $cboFilterSignIn.SelectedIndex = 0
    $panelFilters.Controls.Add($cboFilterSignIn)

    # Bouton Charger / Recharger
    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = Get-Text "shared_mailbox_audit.btn_load"
    $btnLoad.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnLoad.Location = New-Object System.Drawing.Point(940, 10)
    $btnLoad.Size = New-Object System.Drawing.Size(140, 32)
    $btnLoad.BackColor = $colorPurple
    $btnLoad.ForeColor = $colorWhite
    $btnLoad.FlatStyle = "Flat"
    $btnLoad.FlatAppearance.BorderSize = 0
    $panelFilters.Controls.Add($btnLoad)

    # Bouton Exporter CSV
    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = Get-Text "shared_mailbox_audit.btn_export"
    $btnExport.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnExport.Location = New-Object System.Drawing.Point(1090, 10)
    $btnExport.Size = New-Object System.Drawing.Size(140, 32)
    $btnExport.BackColor = $colorBlue
    $btnExport.ForeColor = $colorWhite
    $btnExport.FlatStyle = "Flat"
    $btnExport.FlatAppearance.BorderSize = 0
    $btnExport.Enabled = $false
    $panelFilters.Controls.Add($btnExport)

    # =================================================================
    # Compteurs de résumé
    # =================================================================
    $panelStats = New-Object System.Windows.Forms.Panel
    $panelStats.Location = New-Object System.Drawing.Point(15, 152)
    $panelStats.Size = New-Object System.Drawing.Size(1245, 35)
    $panelStats.BackColor = $colorWhite
    $form.Controls.Add($panelStats)

    $lblStats = New-Object System.Windows.Forms.Label
    $lblStats.Text = Get-Text "shared_mailbox_audit.stats_empty"
    $lblStats.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblStats.ForeColor = $colorGray
    $lblStats.Location = New-Object System.Drawing.Point(15, 8)
    $lblStats.Size = New-Object System.Drawing.Size(1200, 20)
    $panelStats.Controls.Add($lblStats)

    # =================================================================
    # DataGridView — tableau principal (multi-sélection + checkbox)
    # =================================================================
    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Location = New-Object System.Drawing.Point(15, 195)
    $dgv.Size = New-Object System.Drawing.Size(1245, 420)
    $dgv.BackgroundColor = $colorWhite
    $dgv.BorderStyle = "None"
    $dgv.GridColor = [System.Drawing.Color]::FromArgb(222, 226, 230)
    $dgv.AllowUserToAddRows = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.AllowUserToResizeRows = $false
    $dgv.ReadOnly = $false              # Nécessaire pour la colonne checkbox
    $dgv.SelectionMode = "FullRowSelect"
    $dgv.MultiSelect = $true            # ← Activation multi-sélection Ctrl/Shift
    $dgv.RowHeadersVisible = $false
    $dgv.AutoSizeColumnsMode = "None"
    $dgv.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dgv.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dgv.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(233, 236, 239)
    $dgv.ColumnHeadersDefaultCellStyle.ForeColor = $colorDark
    $dgv.ColumnHeadersHeight = 35
    $dgv.RowTemplate.Height = 30
    $dgv.EnableHeadersVisualStyles = $false
    $dgv.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(200, 220, 255)
    $dgv.DefaultCellStyle.SelectionForeColor = $colorDark
    $form.Controls.Add($dgv)

    # --- Colonne checkbox pour la sélection ---
    $colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colCheck.Name = "Select"
    $colCheck.HeaderText = ""
    $colCheck.Width = 35
    $colCheck.FalseValue = $false
    $colCheck.TrueValue = $true
    $colCheck.ReadOnly = $false
    $dgv.Columns.Add($colCheck) | Out-Null

    # Définition des colonnes de données (en ReadOnly)
    $colDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colDisplayName.Name = "DisplayName"
    $colDisplayName.HeaderText = Get-Text "shared_mailbox_audit.col_displayname"
    $colDisplayName.Width = 195
    $colDisplayName.ReadOnly = $true
    $dgv.Columns.Add($colDisplayName) | Out-Null

    $colUPN = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colUPN.Name = "UPN"
    $colUPN.HeaderText = Get-Text "shared_mailbox_audit.col_upn"
    $colUPN.Width = 240
    $colUPN.ReadOnly = $true
    $dgv.Columns.Add($colUPN) | Out-Null

    $colAccountStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAccountStatus.Name = "AccountStatus"
    $colAccountStatus.HeaderText = Get-Text "shared_mailbox_audit.col_account_status"
    $colAccountStatus.Width = 110
    $colAccountStatus.ReadOnly = $true
    $dgv.Columns.Add($colAccountStatus) | Out-Null

    $colLicenses = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colLicenses.Name = "Licenses"
    $colLicenses.HeaderText = Get-Text "shared_mailbox_audit.col_licenses"
    $colLicenses.Width = 230
    $colLicenses.ReadOnly = $true
    $dgv.Columns.Add($colLicenses) | Out-Null

    $colLastSignIn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colLastSignIn.Name = "LastInteractiveSignIn"
    $colLastSignIn.HeaderText = Get-Text "shared_mailbox_audit.col_last_signin"
    $colLastSignIn.Width = 185
    $colLastSignIn.ReadOnly = $true
    $dgv.Columns.Add($colLastSignIn) | Out-Null

    $colAlert = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAlert.Name = "Alert"
    $colAlert.HeaderText = Get-Text "shared_mailbox_audit.col_alert"
    $colAlert.Width = 240
    $colAlert.ReadOnly = $true
    $dgv.Columns.Add($colAlert) | Out-Null

    # Colonne cachée pour l'ID utilisateur
    $colUserId = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colUserId.Name = "UserId"
    $colUserId.Visible = $false
    $colUserId.ReadOnly = $true
    $dgv.Columns.Add($colUserId) | Out-Null

    # =================================================================
    # Panneau de détails + Actions (bas de page)
    # =================================================================
    $panelActions = New-Object System.Windows.Forms.Panel
    $panelActions.Location = New-Object System.Drawing.Point(15, 625)
    $panelActions.Size = New-Object System.Drawing.Size(1245, 115)
    $panelActions.BackColor = $colorWhite
    $form.Controls.Add($panelActions)

    # Infos de la sélection courante
    $lblSelectedInfo = New-Object System.Windows.Forms.Label
    $lblSelectedInfo.Text = Get-Text "shared_mailbox_audit.no_selection"
    $lblSelectedInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSelectedInfo.ForeColor = $colorGray
    $lblSelectedInfo.Location = New-Object System.Drawing.Point(15, 10)
    $lblSelectedInfo.Size = New-Object System.Drawing.Size(800, 40)
    $panelActions.Controls.Add($lblSelectedInfo)

    # Avertissement sign-in détecté
    $lblSignInWarning = New-Object System.Windows.Forms.Label
    $lblSignInWarning.Text = ""
    $lblSignInWarning.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblSignInWarning.ForeColor = $colorOrange
    $lblSignInWarning.Location = New-Object System.Drawing.Point(15, 52)
    $lblSignInWarning.Size = New-Object System.Drawing.Size(800, 20)
    $panelActions.Controls.Add($lblSignInWarning)

    # Bouton Désactiver (multi-sélection)
    $btnDisable = New-Object System.Windows.Forms.Button
    $btnDisable.Text = Get-Text "shared_mailbox_audit.btn_disable"
    $btnDisable.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnDisable.Location = New-Object System.Drawing.Point(870, 15)
    $btnDisable.Size = New-Object System.Drawing.Size(350, 40)
    $btnDisable.BackColor = $colorRed
    $btnDisable.ForeColor = $colorWhite
    $btnDisable.FlatStyle = "Flat"
    $btnDisable.FlatAppearance.BorderSize = 0
    $btnDisable.Enabled = $false
    $panelActions.Controls.Add($btnDisable)

    # Bouton Rafraîchir la sélection
    $btnRefreshRow = New-Object System.Windows.Forms.Button
    $btnRefreshRow.Text = Get-Text "shared_mailbox_audit.btn_refresh_row"
    $btnRefreshRow.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnRefreshRow.Location = New-Object System.Drawing.Point(870, 65)
    $btnRefreshRow.Size = New-Object System.Drawing.Size(350, 32)
    $btnRefreshRow.BackColor = [System.Drawing.Color]::FromArgb(233, 236, 239)
    $btnRefreshRow.ForeColor = $colorDark
    $btnRefreshRow.FlatStyle = "Flat"
    $btnRefreshRow.FlatAppearance.BorderSize = 0
    $btnRefreshRow.Enabled = $false
    $panelActions.Controls.Add($btnRefreshRow)

    # =================================================================
    # Barre de progression (désactivation groupée)
    # =================================================================
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(15, 748)
    $progressBar.Size = New-Object System.Drawing.Size(1245, 8)
    $progressBar.Style = "Marquee"
    $progressBar.MarqueeAnimationSpeed = 30
    $progressBar.Visible = $false
    $form.Controls.Add($progressBar)

    # Pied de page
    $lblFooter = New-Object System.Windows.Forms.Label
    $lblFooter.Text = Get-Text "shared_mailbox_audit.footer"
    $lblFooter.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblFooter.ForeColor = $colorLightGray
    $lblFooter.Location = New-Object System.Drawing.Point(15, 760)
    $lblFooter.Size = New-Object System.Drawing.Size(600, 18)
    $form.Controls.Add($lblFooter)

    # =================================================================
    # FONCTIONS INTERNES
    # =================================================================

    # --- Résolution des noms de licence SKU ---
    function Get-LicenseSkuNames {
        <#
        .SYNOPSIS
            Charge les SKU de licences du tenant et retourne un dictionnaire ID → Nom.
            Utilise un cache pour éviter les appels répétés.
        #>
        if ($script:LicenseSkuMap.Count -gt 0) { return $script:LicenseSkuMap }

        try {
            Write-Log -Level "INFO" -Action "SHARED_AUDIT" -Message "Chargement des SKU de licences du tenant."
            $skus = Get-MgSubscribedSku -Property "skuId,skuPartNumber" -ErrorAction Stop

            foreach ($sku in $skus) {
                $script:LicenseSkuMap[$sku.SkuId] = $sku.SkuPartNumber
            }

            Write-Log -Level "INFO" -Action "SHARED_AUDIT" -Message "$($script:LicenseSkuMap.Count) SKU(s) de licence(s) chargée(s)."
        }
        catch {
            Write-Log -Level "ERROR" -Action "SHARED_AUDIT" -Message "Erreur chargement SKU licences : $($_.Exception.Message)"
        }

        return $script:LicenseSkuMap
    }

    # --- Résolution d'un SKU ID en nom convivial ---
    function Resolve-SkuName {
        param([string]$SkuId)

        $map = Get-LicenseSkuNames
        if ($map.ContainsKey($SkuId)) {
            return $map[$SkuId]
        }
        return $SkuId.Substring(0, [Math]::Min(8, $SkuId.Length)) + "..."
    }

    # --- Connexion Exchange Online ---
    function Connect-ExchangeOnlineSession {
        <#
        .SYNOPSIS
            Connecte à Exchange Online via le module ExchangeOnlineManagement.
            Utilise le même AppId et TenantId que la connexion Graph existante.
            Retourne $true si connecté, $false sinon.
        #>

        # ── Étape 1 : Charger le module en mémoire s'il ne l'est pas déjà ──
        $exoLoaded = Get-Module -Name "ExchangeOnlineManagement" -ErrorAction SilentlyContinue
        if (-not $exoLoaded) {
            try {
                Import-Module ExchangeOnlineManagement -ErrorAction Stop
                Write-Log -Level "INFO" -Action "SHARED_AUDIT" -Message "Module ExchangeOnlineManagement importé."
            }
            catch {
                $exoAvailable = Get-Module -ListAvailable -Name "ExchangeOnlineManagement" -ErrorAction SilentlyContinue
                if (-not $exoAvailable) {
                    Write-Log -Level "ERROR" -Action "SHARED_AUDIT" -Message "Module ExchangeOnlineManagement non installé."
                    Show-ResultDialog -Titre (Get-Text "shared_mailbox_audit.error_title") `
                        -Message (Get-Text "shared_mailbox_audit.error_exo_module") -IsSuccess $false
                    return $false
                }
                else {
                    Write-Log -Level "ERROR" -Action "SHARED_AUDIT" -Message "Module trouvé mais import échoué : $($_.Exception.Message)"
                    Show-ResultDialog -Titre (Get-Text "shared_mailbox_audit.error_title") `
                        -Message (Get-Text "shared_mailbox_audit.error_exo_connect" $_.Exception.Message) -IsSuccess $false
                    return $false
                }
            }
        }

        # ── Étape 2 : Vérifier si une session EXO est déjà active ──
        try {
            $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
            Write-Log -Level "INFO" -Action "SHARED_AUDIT" -Message "Session Exchange Online déjà active."
            return $true
        }
        catch {
            Write-Log -Level "INFO" -Action "SHARED_AUDIT" -Message "Aucune session EXO active, connexion en cours..."
        }

        # ── Étape 3 : Connexion interactive ──
        try {
            $connectParams = @{
                ShowBanner  = $false
                ErrorAction = "Stop"
            }

            if (-not [string]::IsNullOrWhiteSpace($Config.client_id)) {
                $connectParams.AppId = $Config.client_id
                $connectParams.Organization = $Config.tenant_id
            }

            Connect-ExchangeOnline @connectParams

            Write-Log -Level "SUCCESS" -Action "SHARED_AUDIT" -Message "Connexion Exchange Online réussie."
            return $true
        }
        catch {
            Write-Log -Level "ERROR" -Action "SHARED_AUDIT" -Message "Erreur connexion EXO : $($_.Exception.Message)"
            Show-ResultDialog -Titre (Get-Text "shared_mailbox_audit.error_title") `
                -Message (Get-Text "shared_mailbox_audit.error_exo_connect" $_.Exception.Message) -IsSuccess $false
            return $false
        }
    }

    # --- Construction d'un objet résultat enrichi à partir d'une réponse Graph ---
    function Build-MailboxResult {
        <#
        .SYNOPSIS
            Construit un PSCustomObject standardisé à partir des données Graph d'un user.
            Factorise la logique commune entre Get-SharedMailboxes et Update-SingleRowInPlace.
        #>
        param(
            [Parameter(Mandatory)]$GraphUser,
            [string]$FallbackMail = ""
        )

        # Résolution des licences
        $licenseNames = @()
        if ($GraphUser.assignedLicenses -and $GraphUser.assignedLicenses.Count -gt 0) {
            foreach ($lic in $GraphUser.assignedLicenses) {
                $licenseNames += (Resolve-SkuName -SkuId $lic.skuId)
            }
        }

        # Extraction du dernier sign-in interactif
        $lastInteractiveSignIn = $null
        $hasInteractiveSignIn = $false
        $signInActivity = $GraphUser.signInActivity
        if ($signInActivity -and $signInActivity.lastSuccessfulSignInDateTime) {
            $lastInteractiveSignIn = $signInActivity.lastSuccessfulSignInDateTime
            $hasInteractiveSignIn = $true
        }

        # Détermination de l'alerte
        $alertText = ""
        if ($hasInteractiveSignIn) {
            $alertText = Get-Text "shared_mailbox_audit.alert_signin_detected"
        }
        if ($GraphUser.accountEnabled -eq $true -and $licenseNames.Count -gt 0) {
            if ($alertText -ne "") { $alertText += " | " }
            $alertText += Get-Text "shared_mailbox_audit.alert_active_licensed"
        }
        elseif ($GraphUser.accountEnabled -eq $true) {
            if ($alertText -ne "") { $alertText += " | " }
            $alertText += Get-Text "shared_mailbox_audit.alert_active_no_license"
        }
        if ($alertText -eq "") {
            $alertText = Get-Text "shared_mailbox_audit.alert_ok"
        }

        return [PSCustomObject]@{
            UserId                = $GraphUser.id
            DisplayName           = $GraphUser.displayName
            UPN                   = $GraphUser.userPrincipalName
            Mail                  = if ($GraphUser.mail) { $GraphUser.mail } else { $FallbackMail }
            AccountEnabled        = $GraphUser.accountEnabled
            LicenseNames          = ($licenseNames -join ", ")
            HasLicense            = ($licenseNames.Count -gt 0)
            LastInteractiveSignIn = $lastInteractiveSignIn
            HasInteractiveSignIn  = $hasInteractiveSignIn
            Alert                 = $alertText
        }
    }

    # --- Chargement des Shared Mailbox via EXO + enrichissement Graph ---
    function Get-SharedMailboxes {
        <#
        .SYNOPSIS
            Récupère toutes les boîtes aux lettres partagées du tenant.

        .DESCRIPTION
            Stratégie en 3 étapes :
            1. Exchange Online : Get-EXOMailbox -RecipientTypeDetails SharedMailbox (rapide, ~5 sec)
            2. Graph API beta : enrichir avec signInActivity + assignedLicenses par batch
            3. Construction des objets résultat avec alertes via Build-MailboxResult
        #>

        Write-Log -Level "INFO" -Action "SHARED_AUDIT" -Message "Début du chargement des Shared Mailbox."

        # ── ÉTAPE 1 : Connexion EXO ──
        $exoConnected = Connect-ExchangeOnlineSession
        if (-not $exoConnected) { return @() }

        try {
            # ── ÉTAPE 2 : Récupérer les Shared Mailbox via EXO ──
            Write-Log -Level "INFO" -Action "SHARED_AUDIT" -Message "Récupération des Shared Mailbox via Exchange Online..."

            $exoMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited `
                -Properties ExternalDirectoryObjectId, DisplayName, UserPrincipalName, PrimarySmtpAddress `
                -ErrorAction Stop

            $sharedCount = ($exoMailboxes | Measure-Object).Count
            Write-Log -Level "INFO" -Action "SHARED_AUDIT" -Message "$sharedCount Shared Mailbox trouvée(s) via EXO."

            if ($sharedCount -eq 0) {
                Write-Log -Level "INFO" -Action "SHARED_AUDIT" -Message "Aucune Shared Mailbox dans le tenant."
                return @()
            }

            # ── ÉTAPE 3 : Enrichir avec Graph API $batch (signInActivity, licences, accountEnabled) ──
            Write-Log -Level "INFO" -Action "SHARED_AUDIT" -Message "Enrichissement Graph API par batch (signInActivity, licences)..."

            $sharedResults = @()
            $batchSize = 20

            # Préparer la liste des BAL avec ObjectId valide
            $validMailboxes = @()
            foreach ($mbx in $exoMailboxes) {
                if (-not [string]::IsNullOrWhiteSpace($mbx.ExternalDirectoryObjectId)) {
                    $validMailboxes += $mbx
                }
                else {
                    Write-Log -Level "WARNING" -Action "SHARED_AUDIT" -UPN $mbx.UserPrincipalName -Message "Pas d'ExternalDirectoryObjectId — ignorée."
                }
            }

            $totalValid = $validMailboxes.Count
            Write-Log -Level "INFO" -Action "SHARED_AUDIT" -Message "$totalValid BAL à enrichir via Graph ($([Math]::Ceiling($totalValid / $batchSize)) batch(s) de $batchSize)."

            # Traiter par lots de 20
            for ($i = 0; $i -lt $totalValid; $i += $batchSize) {
                $batchEnd = [Math]::Min($i + $batchSize - 1, $totalValid - 1)
                $currentBatch = $validMailboxes[$i..$batchEnd]
                $batchNumber = [Math]::Floor($i / $batchSize) + 1

                # Construire le body $batch
                $batchRequests = @()
                $batchIndex = 0
                foreach ($mbx in $currentBatch) {
                    $batchRequests += @{
                        id     = "$batchIndex"
                        method = "GET"
                        url    = "/users/$($mbx.ExternalDirectoryObjectId)?`$select=id,displayName,userPrincipalName,mail,accountEnabled,assignedLicenses,signInActivity"
                        headers = @{ "ConsistencyLevel" = "eventual" }
                    }
                    $batchIndex++
                }

                $batchBody = @{ requests = $batchRequests } | ConvertTo-Json -Depth 5 -Compress

                try {
                    $batchResponse = Invoke-MgGraphRequest -Method POST `
                        -Uri "https://graph.microsoft.com/beta/`$batch" `
                        -Body $batchBody `
                        -ContentType "application/json" `
                        -ErrorAction Stop

                    # Traiter chaque réponse du batch
                    foreach ($resp in $batchResponse.responses) {
                        $idx = [int]$resp.id
                        $mbx = $currentBatch[$idx]

                        if ($resp.status -eq 200) {
                            $sharedResults += Build-MailboxResult -GraphUser $resp.body -FallbackMail $mbx.PrimarySmtpAddress
                        }
                        else {
                            # Réponse en erreur pour cet utilisateur — fallback EXO
                            Write-Log -Level "WARNING" -Action "SHARED_AUDIT" -UPN $mbx.UserPrincipalName `
                                -Message "Batch réponse $($resp.status) : $($resp.body.error.message)"

                            $sharedResults += [PSCustomObject]@{
                                UserId                = $mbx.ExternalDirectoryObjectId
                                DisplayName           = $mbx.DisplayName
                                UPN                   = $mbx.UserPrincipalName
                                Mail                  = $mbx.PrimarySmtpAddress
                                AccountEnabled        = $null
                                LicenseNames          = ""
                                HasLicense            = $false
                                LastInteractiveSignIn = $null
                                HasInteractiveSignIn  = $false
                                Alert                 = Get-Text "shared_mailbox_audit.alert_graph_error"
                            }
                        }
                    }
                }
                catch {
                    # Échec complet du batch — fallback EXO pour toutes les BAL du lot
                    Write-Log -Level "WARNING" -Action "SHARED_AUDIT" -Message "Échec batch $batchNumber : $($_.Exception.Message)"
                    foreach ($mbx in $currentBatch) {
                        $sharedResults += [PSCustomObject]@{
                            UserId                = $mbx.ExternalDirectoryObjectId
                            DisplayName           = $mbx.DisplayName
                            UPN                   = $mbx.UserPrincipalName
                            Mail                  = $mbx.PrimarySmtpAddress
                            AccountEnabled        = $null
                            LicenseNames          = ""
                            HasLicense            = $false
                            LastInteractiveSignIn = $null
                            HasInteractiveSignIn  = $false
                            Alert                 = Get-Text "shared_mailbox_audit.alert_graph_error"
                        }
                    }
                }

                Write-Log -Level "INFO" -Action "SHARED_AUDIT" -Message "Batch $batchNumber terminé ($($batchEnd + 1) / $totalValid BAL traitées)."
            }

            Write-Log -Level "SUCCESS" -Action "SHARED_AUDIT" -Message "$($sharedResults.Count) Shared Mailbox chargée(s) et enrichie(s)."
            return $sharedResults
        }
        catch {
            $errMsg = $_.Exception.Message
            Write-Log -Level "ERROR" -Action "SHARED_AUDIT" -Message "Erreur lors du chargement : $errMsg"
            Show-ResultDialog -Titre (Get-Text "shared_mailbox_audit.error_title") -Message (Get-Text "shared_mailbox_audit.error_load" $errMsg) -IsSuccess $false
            return @()
        }
    }

    # --- Application des filtres sur les données ---
    function Update-FilteredView {
        $script:FilteredData = $script:SharedMailboxData

        # Filtre statut utilisateur
        $statusIndex = $cboFilterStatus.SelectedIndex
        if ($statusIndex -eq 1) {
            $script:FilteredData = $script:FilteredData | Where-Object { $_.AccountEnabled -eq $true }
        }
        elseif ($statusIndex -eq 2) {
            $script:FilteredData = $script:FilteredData | Where-Object { $_.AccountEnabled -eq $false }
        }

        # Filtre licence
        $licenseIndex = $cboFilterLicense.SelectedIndex
        if ($licenseIndex -eq 1) {
            $script:FilteredData = $script:FilteredData | Where-Object { $_.HasLicense -eq $true }
        }
        elseif ($licenseIndex -eq 2) {
            $script:FilteredData = $script:FilteredData | Where-Object { $_.HasLicense -eq $false }
        }

        # Filtre sign-in interactif
        $signInIndex = $cboFilterSignIn.SelectedIndex
        if ($signInIndex -eq 1) {
            $script:FilteredData = $script:FilteredData | Where-Object { $_.HasInteractiveSignIn -eq $false }
        }
        elseif ($signInIndex -eq 2) {
            $script:FilteredData = $script:FilteredData | Where-Object { $_.HasInteractiveSignIn -eq $true }
        }

        # Mise à jour complète du DataGridView (rebuild)
        Update-DataGridView
        Update-StatsLabel
    }

    # --- Mise à jour complète du DataGridView (rebuild) ---
    function Update-DataGridView {
        $dgv.SuspendLayout()
        $dgv.Rows.Clear()

        foreach ($item in $script:FilteredData) {
            $statusText = if ($item.AccountEnabled) {
                Get-Text "shared_mailbox_audit.status_active"
            }
            else {
                Get-Text "shared_mailbox_audit.status_disabled"
            }

            $signInText = if ($item.HasInteractiveSignIn) {
                try {
                    $dt = [DateTime]::Parse($item.LastInteractiveSignIn)
                    $dt.ToString("yyyy-MM-dd HH:mm")
                }
                catch { $item.LastInteractiveSignIn }
            }
            else {
                Get-Text "shared_mailbox_audit.signin_never"
            }

            $licenseText = if ([string]::IsNullOrWhiteSpace($item.LicenseNames)) {
                Get-Text "shared_mailbox_audit.no_license"
            }
            else {
                $item.LicenseNames
            }

            $rowIndex = $dgv.Rows.Add(
                $false,                 # Checkbox décochée par défaut
                $item.DisplayName,
                $item.UPN,
                $statusText,
                $licenseText,
                $signInText,
                $item.Alert,
                $item.UserId
            )

            # Colorisation conditionnelle
            $row = $dgv.Rows[$rowIndex]

            # Statut actif = orange, inactif = vert
            if ($item.AccountEnabled) {
                $row.Cells["AccountStatus"].Style.ForeColor = $colorOrange
                $row.Cells["AccountStatus"].Style.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            }
            else {
                $row.Cells["AccountStatus"].Style.ForeColor = $colorGreen
            }

            # Sign-in détecté = rouge
            if ($item.HasInteractiveSignIn) {
                $row.Cells["LastInteractiveSignIn"].Style.ForeColor = $colorRed
                $row.Cells["LastInteractiveSignIn"].Style.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            }

            # Alerte non-OK = orange
            if ($item.Alert -ne (Get-Text "shared_mailbox_audit.alert_ok")) {
                $row.Cells["Alert"].Style.ForeColor = $colorOrange
            }
        }

        $dgv.ResumeLayout()
    }

    # --- Mise à jour du label de statistiques ---
    function Update-StatsLabel {
        $total = $script:SharedMailboxData.Count
        $filtered = $script:FilteredData.Count
        $active = ($script:SharedMailboxData | Where-Object { $_.AccountEnabled -eq $true }).Count
        $licensed = ($script:SharedMailboxData | Where-Object { $_.HasLicense -eq $true }).Count
        $withSignIn = ($script:SharedMailboxData | Where-Object { $_.HasInteractiveSignIn -eq $true }).Count

        $lblStats.Text = Get-Text "shared_mailbox_audit.stats_summary" $total $active $licensed $withSignIn $filtered
        $lblStats.ForeColor = $colorDark
    }

    # --- Rafraîchissement IN-PLACE d'une ligne spécifique (sans rebuild de la grille) ---
    function Update-SingleRowInPlace {
        <#
        .SYNOPSIS
            Met à jour les données Graph d'un seul utilisateur et rafraîchit
            uniquement la ligne correspondante dans le DGV — pas de rebuild.
        #>
        param([string]$UserId)

        try {
            Write-Log -Level "INFO" -Action "SHARED_AUDIT" -UPN $UserId -Message "Rafraîchissement des données pour l'utilisateur."

            $uri = "https://graph.microsoft.com/beta/users/$UserId" + '?$select=id,displayName,userPrincipalName,mail,accountEnabled,assignedLicenses,signInActivity'
            $graphHeaders = @{ "ConsistencyLevel" = "eventual" }
            $user = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $graphHeaders -ErrorAction Stop

            # Construire l'objet mis à jour
            $updatedItem = Build-MailboxResult -GraphUser $user

            # Mettre à jour dans les données brutes ($script:SharedMailboxData)
            for ($i = 0; $i -lt $script:SharedMailboxData.Count; $i++) {
                if ($script:SharedMailboxData[$i].UserId -eq $UserId) {
                    $script:SharedMailboxData[$i] = $updatedItem
                    break
                }
            }

            # Mettre à jour dans les données filtrées ($script:FilteredData)
            for ($i = 0; $i -lt $script:FilteredData.Count; $i++) {
                if ($script:FilteredData[$i].UserId -eq $UserId) {
                    $script:FilteredData[$i] = $updatedItem
                    break
                }
            }

            # Mettre à jour la ligne dans le DGV sans rebuild
            foreach ($row in $dgv.Rows) {
                if ($row.Cells["UserId"].Value -eq $UserId) {
                    # Texte du statut
                    $row.Cells["AccountStatus"].Value = if ($updatedItem.AccountEnabled) {
                        Get-Text "shared_mailbox_audit.status_active"
                    } else {
                        Get-Text "shared_mailbox_audit.status_disabled"
                    }

                    # Licences
                    $row.Cells["Licenses"].Value = if ([string]::IsNullOrWhiteSpace($updatedItem.LicenseNames)) {
                        Get-Text "shared_mailbox_audit.no_license"
                    } else {
                        $updatedItem.LicenseNames
                    }

                    # Sign-in
                    $row.Cells["LastInteractiveSignIn"].Value = if ($updatedItem.HasInteractiveSignIn) {
                        try {
                            $dt = [DateTime]::Parse($updatedItem.LastInteractiveSignIn)
                            $dt.ToString("yyyy-MM-dd HH:mm")
                        } catch { $updatedItem.LastInteractiveSignIn }
                    } else {
                        Get-Text "shared_mailbox_audit.signin_never"
                    }

                    # Alerte
                    $row.Cells["Alert"].Value = $updatedItem.Alert

                    # Nom et UPN (peuvent avoir changé)
                    $row.Cells["DisplayName"].Value = $updatedItem.DisplayName
                    $row.Cells["UPN"].Value = $updatedItem.UPN

                    # --- Re-colorisation de la ligne ---
                    # Statut
                    if ($updatedItem.AccountEnabled) {
                        $row.Cells["AccountStatus"].Style.ForeColor = $colorOrange
                        $row.Cells["AccountStatus"].Style.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
                    }
                    else {
                        $row.Cells["AccountStatus"].Style.ForeColor = $colorGreen
                        $row.Cells["AccountStatus"].Style.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                    }

                    # Sign-in
                    if ($updatedItem.HasInteractiveSignIn) {
                        $row.Cells["LastInteractiveSignIn"].Style.ForeColor = $colorRed
                        $row.Cells["LastInteractiveSignIn"].Style.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
                    }
                    else {
                        $row.Cells["LastInteractiveSignIn"].Style.ForeColor = $colorDark
                        $row.Cells["LastInteractiveSignIn"].Style.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                    }

                    # Alerte
                    if ($updatedItem.Alert -ne (Get-Text "shared_mailbox_audit.alert_ok")) {
                        $row.Cells["Alert"].Style.ForeColor = $colorOrange
                    }
                    else {
                        $row.Cells["Alert"].Style.ForeColor = $colorDark
                    }

                    # Décocher la checkbox après action
                    $row.Cells["Select"].Value = $false

                    break
                }
            }

            # Mettre à jour les stats (léger, pas de rebuild)
            Update-StatsLabel

            Write-Log -Level "SUCCESS" -Action "SHARED_AUDIT" -UPN $UserId -Message "Données rafraîchies (in-place)."
        }
        catch {
            Write-Log -Level "ERROR" -Action "SHARED_AUDIT" -UPN $UserId -Message "Erreur rafraîchissement : $($_.Exception.Message)"
        }
    }

    # --- Collecte des BAL cochées ou sélectionnées (union checkbox + sélection DGV) ---
    function Get-SelectedMailboxes {
        <#
        .SYNOPSIS
            Retourne la liste des BAL sélectionnées pour l'action.
            Priorité : les lignes cochées (checkbox). Si aucune cochée,
            fallback sur les lignes sélectionnées (highlight bleu Ctrl/Shift).
            Ne retourne que les comptes actifs (AccountEnabled = $true).
        #>

        $selectedIds = [System.Collections.Generic.HashSet[string]]::new()
        $results = @()

        # 1) Collecter les lignes cochées
        foreach ($row in $dgv.Rows) {
            if ($row.Cells["Select"].Value -eq $true) {
                $uid = $row.Cells["UserId"].Value
                if ($uid -and $selectedIds.Add($uid)) {
                    # Chercher l'objet dans FilteredData
                    $item = $script:FilteredData | Where-Object { $_.UserId -eq $uid } | Select-Object -First 1
                    if ($item) { $results += $item }
                }
            }
        }

        # 2) Si aucune cochée, prendre les lignes sélectionnées (highlight)
        if ($results.Count -eq 0 -and $dgv.SelectedRows.Count -gt 0) {
            foreach ($row in $dgv.SelectedRows) {
                $uid = $row.Cells["UserId"].Value
                if ($uid -and $selectedIds.Add($uid)) {
                    $item = $script:FilteredData | Where-Object { $_.UserId -eq $uid } | Select-Object -First 1
                    if ($item) { $results += $item }
                }
            }
        }

        # Filtrer : ne garder que les comptes actifs
        $activeOnly = $results | Where-Object { $_.AccountEnabled -eq $true }
        return @($activeOnly)
    }

    # --- Mise à jour de l'affichage du panneau d'actions selon la sélection ---
    function Update-SelectionInfo {
        <#
        .SYNOPSIS
            Met à jour le label d'information et les boutons d'actions
            en fonction de la sélection courante (checkbox ou highlight).
        #>

        $selected = Get-SelectedMailboxes
        $totalSelected = $selected.Count

        if ($totalSelected -eq 0) {
            # Vérifier s'il y a au moins une sélection (même inactive)
            $allSelected = @()
            foreach ($row in $dgv.Rows) {
                if ($row.Cells["Select"].Value -eq $true) { $allSelected += $row }
            }
            if ($allSelected.Count -eq 0 -and $dgv.SelectedRows.Count -gt 0) {
                $allSelected = $dgv.SelectedRows
            }

            if ($allSelected.Count -gt 0) {
                # Des lignes sont sélectionnées mais aucune n'est active
                $lblSelectedInfo.Text = Get-Text "shared_mailbox_audit.selected_all_disabled"
                $lblSignInWarning.Text = ""
                $btnDisable.Enabled = $false
                $btnRefreshRow.Enabled = $true
            }
            else {
                $lblSelectedInfo.Text = Get-Text "shared_mailbox_audit.no_selection"
                $lblSignInWarning.Text = ""
                $btnDisable.Enabled = $false
                $btnRefreshRow.Enabled = $false
            }
            return
        }

        $btnRefreshRow.Enabled = $true

        if ($totalSelected -eq 1) {
            # Sélection simple — afficher les détails
            $item = $selected[0]
            $statusStr = if ($item.AccountEnabled) {
                Get-Text "shared_mailbox_audit.status_active"
            } else {
                Get-Text "shared_mailbox_audit.status_disabled"
            }

            $licStr = if ([string]::IsNullOrWhiteSpace($item.LicenseNames)) {
                Get-Text "shared_mailbox_audit.no_license"
            } else {
                $item.LicenseNames
            }

            $lblSelectedInfo.Text = Get-Text "shared_mailbox_audit.selected_info" $item.DisplayName $item.UPN $statusStr $licStr

            # Avertissement sign-in
            if ($item.HasInteractiveSignIn) {
                $lblSignInWarning.Text = Get-Text "shared_mailbox_audit.warning_signin" $item.LastInteractiveSignIn
                $lblSignInWarning.ForeColor = $colorRed
            }
            else {
                $lblSignInWarning.Text = Get-Text "shared_mailbox_audit.warning_no_signin"
                $lblSignInWarning.ForeColor = $colorGreen
            }

            $btnDisable.Text = Get-Text "shared_mailbox_audit.btn_disable"
            $btnDisable.Enabled = $true
        }
        else {
            # Multi-sélection — afficher le compteur
            $withSignIn = ($selected | Where-Object { $_.HasInteractiveSignIn -eq $true }).Count

            $lblSelectedInfo.Text = Get-Text "shared_mailbox_audit.selected_multi" $totalSelected

            if ($withSignIn -gt 0) {
                $lblSignInWarning.Text = Get-Text "shared_mailbox_audit.warning_multi_signin" $withSignIn
                $lblSignInWarning.ForeColor = $colorRed
            }
            else {
                $lblSignInWarning.Text = Get-Text "shared_mailbox_audit.warning_no_signin"
                $lblSignInWarning.ForeColor = $colorGreen
            }

            $btnDisable.Text = Get-Text "shared_mailbox_audit.btn_disable_multi" $totalSelected
            $btnDisable.Enabled = $true
        }
    }

    # =================================================================
    # ÉVÉNEMENTS
    # =================================================================

    # --- Clic sur Charger ---
    $btnLoad.Add_Click({
        $btnLoad.Enabled = $false
        $btnLoad.Text = Get-Text "shared_mailbox_audit.loading"
        $progressBar.Visible = $true
        $progressBar.Style = "Marquee"
        $dgv.Rows.Clear()
        $lblStats.Text = Get-Text "shared_mailbox_audit.loading"
        $lblStats.ForeColor = $colorGray
        $btnExport.Enabled = $false
        $btnDisable.Enabled = $false
        $btnRefreshRow.Enabled = $false
        $lblSelectedInfo.Text = Get-Text "shared_mailbox_audit.no_selection"
        $lblSignInWarning.Text = ""

        # Permettre le rafraîchissement de l'UI
        [System.Windows.Forms.Application]::DoEvents()

        # Réinitialiser le cache de licences pour forcer le rechargement
        $script:LicenseSkuMap = @{}

        # Charger les données
        $script:SharedMailboxData = Get-SharedMailboxes

        # Appliquer les filtres
        Update-FilteredView

        $progressBar.Visible = $false
        $btnLoad.Enabled = $true
        $btnLoad.Text = Get-Text "shared_mailbox_audit.btn_reload"

        if ($script:SharedMailboxData.Count -gt 0) {
            $btnExport.Enabled = $true
        }
    })

    # --- Changement de filtre ---
    $filterAction = {
        if ($script:SharedMailboxData.Count -gt 0) {
            Update-FilteredView
        }
    }
    $cboFilterStatus.Add_SelectedIndexChanged($filterAction)
    $cboFilterLicense.Add_SelectedIndexChanged($filterAction)
    $cboFilterSignIn.Add_SelectedIndexChanged($filterAction)

    # --- Sélection d'une ligne dans le DataGridView (highlight Ctrl/Shift) ---
    $dgv.Add_SelectionChanged({
        Update-SelectionInfo
    })

    # --- Clic sur une checkbox → mettre à jour le panneau d'actions ---
    $dgv.Add_CellContentClick({
        param($sender, $e)
        # Vérifier que c'est la colonne checkbox (index 0)
        if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
            # Forcer le commit de la valeur éditée avant lecture
            $dgv.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
            Update-SelectionInfo
        }
    })

    # --- Clic sur Désactiver (supporte multi-sélection) ---
    $btnDisable.Add_Click({
        $targets = Get-SelectedMailboxes
        if ($targets.Count -eq 0) { return }

        # ── Étape 1 : Construire le message de confirmation ──
        $confirmMsg = ""
        $withSignIn = ($targets | Where-Object { $_.HasInteractiveSignIn -eq $true }).Count

        if ($targets.Count -eq 1) {
            # Désactivation simple — message détaillé
            $item = $targets[0]
            if ($item.HasInteractiveSignIn) {
                $confirmMsg = Get-Text "shared_mailbox_audit.confirm_disable_signin" $item.DisplayName $item.UPN $item.LastInteractiveSignIn
            }
            else {
                $confirmMsg = Get-Text "shared_mailbox_audit.confirm_disable" $item.DisplayName $item.UPN
            }
        }
        else {
            # Désactivation groupée — liste résumée
            $nameList = ($targets | ForEach-Object { "  • $($_.DisplayName) ($($_.UPN))" }) -join "`n"
            if ($withSignIn -gt 0) {
                $confirmMsg = Get-Text "shared_mailbox_audit.confirm_disable_multi_signin" $targets.Count $withSignIn $nameList
            }
            else {
                $confirmMsg = Get-Text "shared_mailbox_audit.confirm_disable_multi" $targets.Count $nameList
            }
        }

        # ── Étape 2 : Dialogue de confirmation ──
        $confirm = Show-ConfirmDialog -Titre (Get-Text "shared_mailbox_audit.confirm_disable_title") -Message $confirmMsg
        if (-not $confirm) { return }

        # ── Étape 3 : Désactivation avec progress bar ──
        $btnDisable.Enabled = $false
        $btnRefreshRow.Enabled = $false
        $btnLoad.Enabled = $false
        $progressBar.Visible = $true
        $progressBar.Style = "Continuous"
        $progressBar.Minimum = 0
        $progressBar.Maximum = $targets.Count
        $progressBar.Value = 0

        $successCount = 0
        $errorCount = 0
        $errorDetails = @()

        for ($idx = 0; $idx -lt $targets.Count; $idx++) {
            $item = $targets[$idx]

            # Mise à jour du texte du bouton avec progression
            $btnDisable.Text = Get-Text "shared_mailbox_audit.disabling_progress" ($idx + 1) $targets.Count
            [System.Windows.Forms.Application]::DoEvents()

            Write-Log -Level "INFO" -Action "SHARED_AUDIT" -UPN $item.UPN -Message "Désactivation du compte utilisateur lié à la BAL partagée."

            try {
                $result = Disable-AzUser -UserId $item.UserId

                if ($result.Success) {
                    $successCount++
                    Write-Log -Level "SUCCESS" -Action "SHARED_AUDIT" -UPN $item.UPN -Message "Compte désactivé avec succès."

                    # Rafraîchir la ligne in-place (pas de rebuild)
                    Update-SingleRowInPlace -UserId $item.UserId
                }
                else {
                    $errorCount++
                    $errorDetails += "$($item.DisplayName) : $($result.Error)"
                    Write-Log -Level "ERROR" -Action "SHARED_AUDIT" -UPN $item.UPN -Message "Échec désactivation : $($result.Error)"
                }
            }
            catch {
                $errorCount++
                $errorDetails += "$($item.DisplayName) : $($_.Exception.Message)"
                Write-Log -Level "ERROR" -Action "SHARED_AUDIT" -UPN $item.UPN -Message "Exception désactivation : $($_.Exception.Message)"
            }

            $progressBar.Value = $idx + 1
            [System.Windows.Forms.Application]::DoEvents()
        }

        # ── Étape 4 : Résultat final ──
        $progressBar.Visible = $false
        $btnLoad.Enabled = $true

        if ($errorCount -eq 0) {
            Show-ResultDialog `
                -Titre (Get-Text "shared_mailbox_audit.disable_success_title") `
                -Message (Get-Text "shared_mailbox_audit.disable_batch_success" $successCount) `
                -IsSuccess $true
        }
        else {
            $errSummary = ($errorDetails -join "`n")
            Show-ResultDialog `
                -Titre (Get-Text "shared_mailbox_audit.error_title") `
                -Message (Get-Text "shared_mailbox_audit.disable_batch_partial" $successCount $errorCount $errSummary) `
                -IsSuccess $false
        }

        # Restaurer le texte du bouton et réévaluer la sélection
        $btnDisable.Text = Get-Text "shared_mailbox_audit.btn_disable"
        Update-SelectionInfo
    })

    # --- Clic sur Rafraîchir la sélection ---
    $btnRefreshRow.Add_Click({
        $targets = Get-SelectedMailboxes
        # Si aucune cible active, rafraîchir toutes les lignes sélectionnées/cochées
        if ($targets.Count -eq 0) {
            $targets = @()
            foreach ($row in $dgv.Rows) {
                if ($row.Cells["Select"].Value -eq $true) {
                    $uid = $row.Cells["UserId"].Value
                    $item = $script:FilteredData | Where-Object { $_.UserId -eq $uid } | Select-Object -First 1
                    if ($item) { $targets += $item }
                }
            }
            if ($targets.Count -eq 0 -and $dgv.SelectedRows.Count -gt 0) {
                foreach ($row in $dgv.SelectedRows) {
                    $uid = $row.Cells["UserId"].Value
                    $item = $script:FilteredData | Where-Object { $_.UserId -eq $uid } | Select-Object -First 1
                    if ($item) { $targets += $item }
                }
            }
        }
        if ($targets.Count -eq 0) { return }

        $btnRefreshRow.Enabled = $false
        $btnRefreshRow.Text = Get-Text "shared_mailbox_audit.refreshing"
        [System.Windows.Forms.Application]::DoEvents()

        foreach ($item in $targets) {
            Update-SingleRowInPlace -UserId $item.UserId
        }

        $btnRefreshRow.Text = Get-Text "shared_mailbox_audit.btn_refresh_row"
        $btnRefreshRow.Enabled = $true
        Update-SelectionInfo
    })

    # --- Clic sur Exporter CSV ---
    $btnExport.Add_Click({
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV (*.csv)|*.csv"
        $saveDialog.Title = Get-Text "shared_mailbox_audit.export_title"
        $saveDialog.FileName = "SharedMailbox_Audit_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"

        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $exportData = $script:FilteredData | Select-Object `
                    @{Name = "DisplayName"; Expression = { $_.DisplayName }},
                    @{Name = "UserPrincipalName"; Expression = { $_.UPN }},
                    @{Name = "Mail"; Expression = { $_.Mail }},
                    @{Name = "AccountEnabled"; Expression = { $_.AccountEnabled }},
                    @{Name = "Licenses"; Expression = { $_.LicenseNames }},
                    @{Name = "LastInteractiveSignIn"; Expression = { $_.LastInteractiveSignIn }},
                    @{Name = "HasInteractiveSignIn"; Expression = { $_.HasInteractiveSignIn }},
                    @{Name = "Alert"; Expression = { $_.Alert }}

                $exportData | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8

                Write-Log -Level "SUCCESS" -Action "SHARED_AUDIT" -Message "Export CSV réussi : $($saveDialog.FileName) ($($exportData.Count) lignes)"
                Show-ResultDialog -Titre (Get-Text "shared_mailbox_audit.export_success_title") -Message (Get-Text "shared_mailbox_audit.export_success_msg" $saveDialog.FileName $exportData.Count) -IsSuccess $true
            }
            catch {
                Write-Log -Level "ERROR" -Action "SHARED_AUDIT" -Message "Erreur export CSV : $($_.Exception.Message)"
                Show-ResultDialog -Titre (Get-Text "shared_mailbox_audit.error_title") -Message (Get-Text "shared_mailbox_audit.export_error_msg" $_.Exception.Message) -IsSuccess $false
            }
        }
    })

    # =================================================================
    # Affichage
    # =================================================================
    $form.ShowDialog() | Out-Null
    $form.Dispose()
}