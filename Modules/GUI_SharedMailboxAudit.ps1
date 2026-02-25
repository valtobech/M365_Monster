<#
.FICHIER
    Modules/GUI_SharedMailboxAudit.ps1

.ROLE
    Module d'audit des boîtes aux lettres partagées (Shared Mailbox).
    Identifie les BAL partagées dont l'objet utilisateur associé est actif,
    vérifie les licences assignées et le dernier sign-in interactif.
    Permet de désactiver le compte utilisateur lié à la BAL après validation manuelle.

    Objectifs :
    - Lister toutes les Shared Mailbox du tenant via Graph API
    - Filtrer par statut utilisateur (actif / inactif)
    - Filtrer par licence assignée (avec / sans licence)
    - Afficher le dernier sign-in interactif (lastSuccessfulSignInDateTime)
    - Alerter si un sign-in interactif a été détecté (validation manuelle requise)
    - Permettre la désactivation du compte utilisateur lié à la BAL

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
        et la désactivation des comptes utilisateurs associés.
    #>

    # =================================================================
    # Variables de travail du module
    # =================================================================
    $script:SharedMailboxData = @()          # Données brutes chargées depuis Graph
    $script:FilteredData = @()              # Données après filtrage
    $script:SelectedMailbox = $null          # BAL sélectionnée dans le DataGridView
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
    # DataGridView — tableau principal
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
    $dgv.ReadOnly = $true
    $dgv.SelectionMode = "FullRowSelect"
    $dgv.MultiSelect = $false
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

    # Définition des colonnes
    $colDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colDisplayName.Name = "DisplayName"
    $colDisplayName.HeaderText = Get-Text "shared_mailbox_audit.col_displayname"
    $colDisplayName.Width = 200
    $dgv.Columns.Add($colDisplayName) | Out-Null

    $colUPN = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colUPN.Name = "UPN"
    $colUPN.HeaderText = Get-Text "shared_mailbox_audit.col_upn"
    $colUPN.Width = 250
    $dgv.Columns.Add($colUPN) | Out-Null

    $colAccountStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAccountStatus.Name = "AccountStatus"
    $colAccountStatus.HeaderText = Get-Text "shared_mailbox_audit.col_account_status"
    $colAccountStatus.Width = 110
    $dgv.Columns.Add($colAccountStatus) | Out-Null

    $colLicenses = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colLicenses.Name = "Licenses"
    $colLicenses.HeaderText = Get-Text "shared_mailbox_audit.col_licenses"
    $colLicenses.Width = 250
    $dgv.Columns.Add($colLicenses) | Out-Null

    $colLastSignIn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colLastSignIn.Name = "LastInteractiveSignIn"
    $colLastSignIn.HeaderText = Get-Text "shared_mailbox_audit.col_last_signin"
    $colLastSignIn.Width = 185
    $dgv.Columns.Add($colLastSignIn) | Out-Null

    $colAlert = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAlert.Name = "Alert"
    $colAlert.HeaderText = Get-Text "shared_mailbox_audit.col_alert"
    $colAlert.Width = 240
    $dgv.Columns.Add($colAlert) | Out-Null

    # Colonne cachée pour l'ID utilisateur
    $colUserId = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colUserId.Name = "UserId"
    $colUserId.Visible = $false
    $dgv.Columns.Add($colUserId) | Out-Null

    # =================================================================
    # Panneau de détails + Actions (bas de page)
    # =================================================================
    $panelActions = New-Object System.Windows.Forms.Panel
    $panelActions.Location = New-Object System.Drawing.Point(15, 625)
    $panelActions.Size = New-Object System.Drawing.Size(1245, 115)
    $panelActions.BackColor = $colorWhite
    $form.Controls.Add($panelActions)

    # Infos de la BAL sélectionnée
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

    # Bouton Désactiver le compte utilisateur
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
    # Barre de progression
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
            # Tenter l'import — cherche dans tous les chemins PSModulePath
            try {
                Import-Module ExchangeOnlineManagement -ErrorAction Stop
                Write-Log -Level "INFO" -Action "SHARED_AUDIT" -Message "Module ExchangeOnlineManagement importé."
            }
            catch {
                # Vérifier s'il est installé mais pas encore importé
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
            # Pas connecté ou session expirée, on continue vers la connexion
            Write-Log -Level "INFO" -Action "SHARED_AUDIT" -Message "Aucune session EXO active, connexion en cours..."
        }

        # ── Étape 3 : Connexion interactive ──
        try {
            $connectParams = @{
                ShowBanner  = $false
                ErrorAction = "Stop"
            }

            # Utiliser le même AppId que Graph pour la cohérence
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

    # --- Chargement des Shared Mailbox via EXO + enrichissement Graph ---
    function Get-SharedMailboxes {
        <#
        .SYNOPSIS
            Récupère toutes les boîtes aux lettres partagées du tenant.

        .DESCRIPTION
            Stratégie en 3 étapes :
            1. Exchange Online : Get-EXOMailbox -RecipientTypeDetails SharedMailbox (rapide, ~5 sec)
            2. Graph API beta : enrichir avec signInActivity + assignedLicenses par batch
            3. Construction des objets résultat avec alertes
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
            # Graph $batch permet d'envoyer jusqu'à 20 requêtes en un seul appel HTTP
            Write-Log -Level "INFO" -Action "SHARED_AUDIT" -Message "Enrichissement Graph API par batch (signInActivity, licences)..."

            $graphHeaders = @{ "ConsistencyLevel" = "eventual" }
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
                            $user = $resp.body

                            # Résolution des licences
                            $licenseNames = @()
                            if ($user.assignedLicenses -and $user.assignedLicenses.Count -gt 0) {
                                foreach ($lic in $user.assignedLicenses) {
                                    $licenseNames += (Resolve-SkuName -SkuId $lic.skuId)
                                }
                            }

                            # Extraction du dernier sign-in interactif
                            $lastInteractiveSignIn = $null
                            $hasInteractiveSignIn = $false
                            $signInActivity = $user.signInActivity
                            if ($signInActivity -and $signInActivity.lastSuccessfulSignInDateTime) {
                                $lastInteractiveSignIn = $signInActivity.lastSuccessfulSignInDateTime
                                $hasInteractiveSignIn = $true
                            }

                            # Détermination de l'alerte
                            $alertText = ""
                            if ($hasInteractiveSignIn) {
                                $alertText = Get-Text "shared_mailbox_audit.alert_signin_detected"
                            }
                            if ($user.accountEnabled -eq $true -and $licenseNames.Count -gt 0) {
                                if ($alertText -ne "") { $alertText += " | " }
                                $alertText += Get-Text "shared_mailbox_audit.alert_active_licensed"
                            }
                            elseif ($user.accountEnabled -eq $true) {
                                if ($alertText -ne "") { $alertText += " | " }
                                $alertText += Get-Text "shared_mailbox_audit.alert_active_no_license"
                            }
                            if ($alertText -eq "") {
                                $alertText = Get-Text "shared_mailbox_audit.alert_ok"
                            }

                            $sharedResults += [PSCustomObject]@{
                                UserId                = $user.id
                                DisplayName           = $user.displayName
                                UPN                   = $user.userPrincipalName
                                Mail                  = if ($user.mail) { $user.mail } else { $mbx.PrimarySmtpAddress }
                                AccountEnabled        = $user.accountEnabled
                                LicenseNames          = ($licenseNames -join ", ")
                                HasLicense            = ($licenseNames.Count -gt 0)
                                LastInteractiveSignIn = $lastInteractiveSignIn
                                HasInteractiveSignIn  = $hasInteractiveSignIn
                                Alert                 = $alertText
                            }
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
            # Actif uniquement
            $script:FilteredData = $script:FilteredData | Where-Object { $_.AccountEnabled -eq $true }
        }
        elseif ($statusIndex -eq 2) {
            # Inactif uniquement
            $script:FilteredData = $script:FilteredData | Where-Object { $_.AccountEnabled -eq $false }
        }

        # Filtre licence
        $licenseIndex = $cboFilterLicense.SelectedIndex
        if ($licenseIndex -eq 1) {
            # Avec licence
            $script:FilteredData = $script:FilteredData | Where-Object { $_.HasLicense -eq $true }
        }
        elseif ($licenseIndex -eq 2) {
            # Sans licence
            $script:FilteredData = $script:FilteredData | Where-Object { $_.HasLicense -eq $false }
        }

        # Filtre sign-in interactif
        $signInIndex = $cboFilterSignIn.SelectedIndex
        if ($signInIndex -eq 1) {
            # Jamais de sign-in
            $script:FilteredData = $script:FilteredData | Where-Object { $_.HasInteractiveSignIn -eq $false }
        }
        elseif ($signInIndex -eq 2) {
            # Sign-in détecté
            $script:FilteredData = $script:FilteredData | Where-Object { $_.HasInteractiveSignIn -eq $true }
        }

        # Mise à jour du DataGridView
        Update-DataGridView
        Update-StatsLabel
    }

    # --- Mise à jour du DataGridView ---
    function Update-DataGridView {
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

            # Alerte non-OK = orange/rouge
            if ($item.Alert -ne (Get-Text "shared_mailbox_audit.alert_ok")) {
                $row.Cells["Alert"].Style.ForeColor = $colorOrange
            }
        }
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

    # --- Rafraîchissement d'une ligne spécifique après action ---
    function Update-SingleRow {
        param([string]$UserId)

        try {
            Write-Log -Level "INFO" -Action "SHARED_AUDIT" -UPN $UserId -Message "Rafraîchissement des données pour l'utilisateur."

            $uri = "https://graph.microsoft.com/beta/users/$UserId" + '?$select=id,displayName,userPrincipalName,mail,accountEnabled,assignedLicenses,signInActivity'
            $graphHeaders = @{ "ConsistencyLevel" = "eventual" }
            $user = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $graphHeaders -ErrorAction Stop

            # Mettre à jour dans les données brutes
            $existingIndex = -1
            for ($i = 0; $i -lt $script:SharedMailboxData.Count; $i++) {
                if ($script:SharedMailboxData[$i].UserId -eq $UserId) {
                    $existingIndex = $i
                    break
                }
            }

            if ($existingIndex -ge 0) {
                $licenseNames = @()
                if ($user.assignedLicenses -and $user.assignedLicenses.Count -gt 0) {
                    foreach ($lic in $user.assignedLicenses) {
                        $licenseNames += (Resolve-SkuName -SkuId $lic.skuId)
                    }
                }

                $lastInteractiveSignIn = $null
                $hasInteractiveSignIn = $false
                if ($user.signInActivity -and $user.signInActivity.lastSuccessfulSignInDateTime) {
                    $lastInteractiveSignIn = $user.signInActivity.lastSuccessfulSignInDateTime
                    $hasInteractiveSignIn = $true
                }

                $alertText = ""
                if ($hasInteractiveSignIn) {
                    $alertText = Get-Text "shared_mailbox_audit.alert_signin_detected"
                }
                if ($user.accountEnabled -eq $true -and $licenseNames.Count -gt 0) {
                    if ($alertText -ne "") { $alertText += " | " }
                    $alertText += Get-Text "shared_mailbox_audit.alert_active_licensed"
                }
                elseif ($user.accountEnabled -eq $true) {
                    if ($alertText -ne "") { $alertText += " | " }
                    $alertText += Get-Text "shared_mailbox_audit.alert_active_no_license"
                }
                if ($alertText -eq "") { $alertText = Get-Text "shared_mailbox_audit.alert_ok" }

                $script:SharedMailboxData[$existingIndex] = [PSCustomObject]@{
                    UserId                = $user.id
                    DisplayName           = $user.displayName
                    UPN                   = $user.userPrincipalName
                    Mail                  = $user.mail
                    AccountEnabled        = $user.accountEnabled
                    LicenseNames          = ($licenseNames -join ", ")
                    HasLicense            = ($licenseNames.Count -gt 0)
                    LastInteractiveSignIn = $lastInteractiveSignIn
                    HasInteractiveSignIn  = $hasInteractiveSignIn
                    Alert                 = $alertText
                }

                # Re-filtrer et rafraîchir la vue
                Update-FilteredView

                Write-Log -Level "SUCCESS" -Action "SHARED_AUDIT" -UPN $UserId -Message "Données rafraîchies."
            }
        }
        catch {
            Write-Log -Level "ERROR" -Action "SHARED_AUDIT" -UPN $UserId -Message "Erreur rafraîchissement : $($_.Exception.Message)"
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

    # --- Sélection d'une ligne dans le DataGridView ---
    $dgv.Add_SelectionChanged({
        if ($dgv.SelectedRows.Count -eq 0) {
            $script:SelectedMailbox = $null
            $btnDisable.Enabled = $false
            $btnRefreshRow.Enabled = $false
            $lblSelectedInfo.Text = Get-Text "shared_mailbox_audit.no_selection"
            $lblSignInWarning.Text = ""
            return
        }

        $row = $dgv.SelectedRows[0]
        $userId = $row.Cells["UserId"].Value
        $displayName = $row.Cells["DisplayName"].Value
        $upn = $row.Cells["UPN"].Value

        # Retrouver l'objet dans les données filtrées
        $script:SelectedMailbox = $script:FilteredData | Where-Object { $_.UserId -eq $userId } | Select-Object -First 1

        if ($null -ne $script:SelectedMailbox) {
            $statusStr = if ($script:SelectedMailbox.AccountEnabled) {
                Get-Text "shared_mailbox_audit.status_active"
            }
            else {
                Get-Text "shared_mailbox_audit.status_disabled"
            }

            $licStr = if ([string]::IsNullOrWhiteSpace($script:SelectedMailbox.LicenseNames)) {
                Get-Text "shared_mailbox_audit.no_license"
            }
            else {
                $script:SelectedMailbox.LicenseNames
            }

            $lblSelectedInfo.Text = Get-Text "shared_mailbox_audit.selected_info" $displayName $upn $statusStr $licStr

            # Avertissement sign-in
            if ($script:SelectedMailbox.HasInteractiveSignIn) {
                $lblSignInWarning.Text = Get-Text "shared_mailbox_audit.warning_signin" $script:SelectedMailbox.LastInteractiveSignIn
                $lblSignInWarning.ForeColor = $colorRed
            }
            else {
                $lblSignInWarning.Text = Get-Text "shared_mailbox_audit.warning_no_signin"
                $lblSignInWarning.ForeColor = $colorGreen
            }

            # Activer le bouton désactiver seulement si le compte est actif
            $btnDisable.Enabled = $script:SelectedMailbox.AccountEnabled
            $btnRefreshRow.Enabled = $true
        }
    })

    # --- Clic sur Désactiver ---
    $btnDisable.Add_Click({
        if ($null -eq $script:SelectedMailbox) { return }

        $displayName = $script:SelectedMailbox.DisplayName
        $upn = $script:SelectedMailbox.UPN
        $userId = $script:SelectedMailbox.UserId

        # Avertissement renforcé si un sign-in interactif a été détecté
        $confirmMsg = ""
        if ($script:SelectedMailbox.HasInteractiveSignIn) {
            $confirmMsg = Get-Text "shared_mailbox_audit.confirm_disable_signin" $displayName $upn $script:SelectedMailbox.LastInteractiveSignIn
        }
        else {
            $confirmMsg = Get-Text "shared_mailbox_audit.confirm_disable" $displayName $upn
        }

        $confirm = Show-ConfirmDialog -Titre (Get-Text "shared_mailbox_audit.confirm_disable_title") -Message $confirmMsg
        if (-not $confirm) { return }

        Write-Log -Level "INFO" -Action "SHARED_AUDIT" -UPN $upn -Message "Désactivation du compte utilisateur lié à la BAL partagée."

        $btnDisable.Enabled = $false
        $btnDisable.Text = Get-Text "shared_mailbox_audit.disabling"
        [System.Windows.Forms.Application]::DoEvents()

        try {
            # Utilisation du wrapper existant de GraphAPI.ps1
            $result = Disable-AzUser -UserId $userId

            if ($result.Success) {
                Write-Log -Level "SUCCESS" -Action "SHARED_AUDIT" -UPN $upn -Message "Compte désactivé avec succès."
                Show-ResultDialog -Titre (Get-Text "shared_mailbox_audit.disable_success_title") -Message (Get-Text "shared_mailbox_audit.disable_success_msg" $displayName) -IsSuccess $true

                # Rafraîchir la ligne
                Update-SingleRow -UserId $userId
            }
            else {
                Write-Log -Level "ERROR" -Action "SHARED_AUDIT" -UPN $upn -Message "Échec désactivation : $($result.Error)"
                Show-ResultDialog -Titre (Get-Text "shared_mailbox_audit.error_title") -Message (Get-Text "shared_mailbox_audit.disable_error_msg" $result.Error) -IsSuccess $false
            }
        }
        catch {
            Write-Log -Level "ERROR" -Action "SHARED_AUDIT" -UPN $upn -Message "Exception désactivation : $($_.Exception.Message)"
            Show-ResultDialog -Titre (Get-Text "shared_mailbox_audit.error_title") -Message (Get-Text "shared_mailbox_audit.disable_error_msg" $_.Exception.Message) -IsSuccess $false
        }
        finally {
            $btnDisable.Text = Get-Text "shared_mailbox_audit.btn_disable"
            if ($null -ne $script:SelectedMailbox -and $script:SelectedMailbox.AccountEnabled) {
                $btnDisable.Enabled = $true
            }
        }
    })

    # --- Clic sur Rafraîchir la sélection ---
    $btnRefreshRow.Add_Click({
        if ($null -eq $script:SelectedMailbox) { return }

        $btnRefreshRow.Enabled = $false
        $btnRefreshRow.Text = Get-Text "shared_mailbox_audit.refreshing"
        [System.Windows.Forms.Application]::DoEvents()

        Update-SingleRow -UserId $script:SelectedMailbox.UserId

        $btnRefreshRow.Text = Get-Text "shared_mailbox_audit.btn_refresh_row"
        $btnRefreshRow.Enabled = $true
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