<#
.FICHIER
    Modules/GUI_Main.ps1

.ROLE
    Fenêtre principale de l'application — menu de navigation.
    Toutes les chaînes affichées passent par Get-Text (i18n).

.DEPENDANCES
    - Core/Lang.ps1 (Get-Text)
    - Core/Functions.ps1 (Write-Log, Show-ConfirmDialog)
    - Core/Connect.ps1 (Disconnect-GraphAPI, Get-GraphConnectionStatus)
    - Variable globale $Config

.AUTEUR
    [Equipe IT - M365 Monster]
#>

function Show-MainWindow {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$(Get-Text 'app.name') — $($Config.client_name)"
    $form.Size = New-Object System.Drawing.Size(820, 948)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedSingle"
    $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)

    # Icône si disponible
    $iconFile = Join-Path -Path $script:RootPath -ChildPath "Assets\M365Monster.ico"
    if (Test-Path $iconFile) { $form.Icon = New-Object System.Drawing.Icon($iconFile) }

    # =================================================================
    # En-tête
    # =================================================================
    $panelHeader = New-Object System.Windows.Forms.Panel
    $panelHeader.Location = New-Object System.Drawing.Point(0, 0)
    $panelHeader.Size = New-Object System.Drawing.Size(820, 100)
    $panelHeader.BackColor = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $form.Controls.Add($panelHeader)

    $lblTitre = New-Object System.Windows.Forms.Label
    $lblTitre.Text = Get-Text "app.name"
    $lblTitre.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $lblTitre.ForeColor = [System.Drawing.Color]::White
    $lblTitre.Location = New-Object System.Drawing.Point(25, 12)
    $lblTitre.Size = New-Object System.Drawing.Size(400, 40)
    $panelHeader.Controls.Add($lblTitre)

    $lblClient = New-Object System.Windows.Forms.Label
    $lblClient.Text = "$(Get-Text 'main_menu.client_label') : $($Config.client_name)  |  $(Get-Text 'main_menu.tenant_label') : $($Config.tenant_id.Substring(0, 8))..."
    $lblClient.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblClient.ForeColor = [System.Drawing.Color]::FromArgb(173, 181, 189)
    $lblClient.Location = New-Object System.Drawing.Point(25, 55)
    $lblClient.Size = New-Object System.Drawing.Size(500, 20)
    $panelHeader.Controls.Add($lblClient)

    $lblConnexion = New-Object System.Windows.Forms.Label
    $lblConnexion.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblConnexion.Location = New-Object System.Drawing.Point(25, 75)
    $lblConnexion.Size = New-Object System.Drawing.Size(200, 20)

    if (Get-GraphConnectionStatus) {
        $lblConnexion.Text = Get-Text "main_menu.connected"
        $lblConnexion.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    }
    else {
        $lblConnexion.Text = Get-Text "main_menu.disconnected"
        $lblConnexion.ForeColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    }
    $panelHeader.Controls.Add($lblConnexion)

    $btnDeconnexion = New-Object System.Windows.Forms.Button
    $btnDeconnexion.Text = Get-Text "main_menu.btn_change_client"
    $btnDeconnexion.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $btnDeconnexion.Location = New-Object System.Drawing.Point(640, 55)
    $btnDeconnexion.Size = New-Object System.Drawing.Size(150, 30)
    $btnDeconnexion.FlatStyle = "Flat"
    $btnDeconnexion.ForeColor = [System.Drawing.Color]::White
    $btnDeconnexion.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
    $btnDeconnexion.Add_Click({
        $confirm = Show-ConfirmDialog -Titre (Get-Text "main_menu.change_client_title") -Message (Get-Text "main_menu.change_client_msg")
        if ($confirm) {
            Write-Log -Level "INFO" -Action "GUI" -Message "Changement de client demandé."
            Disconnect-GraphAPI
            $form.Close()
        }
    })
    $panelHeader.Controls.Add($btnDeconnexion)

    # =================================================================
    # Grille de tuiles — 2 colonnes × 4 lignes
    # =================================================================
    $yStart  = 130
    $btnW    = 350
    $btnH    = 90
    $spacingV = 18
    $leftCol = 30
    $rightCol = 420

    function New-MenuTile {
        param([string]$Text, [string]$SubText, [int]$X, [int]$Y, $Color)

        $panel = New-Object System.Windows.Forms.Panel
        $panel.Location = New-Object System.Drawing.Point($X, $Y)
        $panel.Size = New-Object System.Drawing.Size($btnW, $btnH)
        $panel.BackColor = [System.Drawing.Color]::White
        $panel.BorderStyle = "None"
        $panel.Cursor = [System.Windows.Forms.Cursors]::Hand

        $border = New-Object System.Windows.Forms.Panel
        $border.Location = New-Object System.Drawing.Point(0, 0)
        $border.Size = New-Object System.Drawing.Size(5, $btnH)
        $border.BackColor = $Color
        $panel.Controls.Add($border)

        $lblTitle = New-Object System.Windows.Forms.Label
        $lblTitle.Text = $Text
        $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
        $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)
        $lblTitle.Location = New-Object System.Drawing.Point(20, 12)
        $lblTitle.Size = New-Object System.Drawing.Size(310, 28)
        $panel.Controls.Add($lblTitle)

        $lblSub = New-Object System.Windows.Forms.Label
        $lblSub.Text = $SubText
        $lblSub.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $lblSub.ForeColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
        $lblSub.Location = New-Object System.Drawing.Point(20, 44)
        $lblSub.Size = New-Object System.Drawing.Size(310, 35)
        $panel.Controls.Add($lblSub)

        return $panel
    }

    function Set-TileAction {
        param($Panel, [scriptblock]$Action)
        $Panel.Add_Click($Action)
        foreach ($ctrl in $Panel.Controls) {
            if ($ctrl -is [System.Windows.Forms.Label]) {
                $ctrl.Add_Click($Action)
            }
        }
    }

    # LIGNE 1
    $row1Y = $yStart

    $tileOnboarding = New-MenuTile -Text (Get-Text "main_menu.tile_onboarding") -SubText (Get-Text "main_menu.tile_onboarding_desc") `
        -X $leftCol -Y $row1Y -Color ([System.Drawing.Color]::FromArgb(40, 167, 69))
    Set-TileAction -Panel $tileOnboarding -Action {
        Write-Log -Level "INFO" -Action "GUI" -Message "Ouverture du formulaire Onboarding."
        Show-OnboardingForm
    }
    $form.Controls.Add($tileOnboarding)

    $tileOffboarding = New-MenuTile -Text (Get-Text "main_menu.tile_offboarding") -SubText (Get-Text "main_menu.tile_offboarding_desc") `
        -X $rightCol -Y $row1Y -Color ([System.Drawing.Color]::FromArgb(220, 53, 69))
    Set-TileAction -Panel $tileOffboarding -Action {
        Write-Log -Level "INFO" -Action "GUI" -Message "Ouverture du formulaire Offboarding."
        Show-OffboardingForm
    }
    $form.Controls.Add($tileOffboarding)

    # LIGNE 2
    $row2Y = $row1Y + $btnH + $spacingV

    $tileModification = New-MenuTile -Text (Get-Text "main_menu.tile_modification") -SubText (Get-Text "main_menu.tile_modification_desc") `
        -X $leftCol -Y $row2Y -Color ([System.Drawing.Color]::FromArgb(0, 123, 255))
    Set-TileAction -Panel $tileModification -Action {
        Write-Log -Level "INFO" -Action "GUI" -Message "Ouverture du formulaire Modification."
        Show-ModificationForm
    }
    $form.Controls.Add($tileModification)

    $tileEmployeeType = New-MenuTile -Text (Get-Text "main_menu.tile_employee_type") -SubText (Get-Text "main_menu.tile_employee_type_desc") `
        -X $rightCol -Y $row2Y -Color ([System.Drawing.Color]::FromArgb(111, 66, 193))
    Set-TileAction -Panel $tileEmployeeType -Action {
        Write-Log -Level "INFO" -Action "GUI" -Message "Lancement du script Employee Type Manager."
        $scriptPath = Join-Path -Path $script:RootPath -ChildPath "Scripts\AzureAD_EmployeeTypeManageGUI.ps1"
        if (Test-Path -Path $scriptPath) {
            . $scriptPath
        }
        else {
            Show-ResultDialog -Titre (Get-Text "main_menu.script_not_found_title") -Message (Get-Text "main_menu.script_not_found_msg" $scriptPath) -IsSuccess $false
        }
    }
    $form.Controls.Add($tileEmployeeType)

    # LIGNE 3
    $row3Y = $row2Y + $btnH + $spacingV

    $tileStaleDevices = New-MenuTile -Text (Get-Text "main_menu.tile_stale_devices") -SubText (Get-Text "main_menu.tile_stale_devices_desc") `
        -X $leftCol -Y $row3Y -Color ([System.Drawing.Color]::FromArgb(253, 126, 20))
    Set-TileAction -Panel $tileStaleDevices -Action {
        Write-Log -Level "INFO" -Action "GUI" -Message "Lancement du script Stale Devices."
        $scriptPath = Join-Path -Path $script:RootPath -ChildPath "Scripts\AzureAD_CleanStaleDeviceGUI.ps1"
        if (Test-Path -Path $scriptPath) {
            . $scriptPath
        }
        else {
            Show-ResultDialog -Titre (Get-Text "main_menu.script_not_found_title") -Message (Get-Text "main_menu.script_not_found_msg" $scriptPath) -IsSuccess $false
        }
    }
    $form.Controls.Add($tileStaleDevices)

    $tileSettings = New-MenuTile -Text (Get-Text "main_menu.tile_settings") -SubText (Get-Text "main_menu.tile_settings_desc") `
        -X $rightCol -Y $row3Y -Color ([System.Drawing.Color]::FromArgb(255, 193, 7))
    Set-TileAction -Panel $tileSettings -Action {
        Write-Log -Level "INFO" -Action "GUI" -Message "Ouverture des paramètres."
        Show-SettingsForm
    }
    $form.Controls.Add($tileSettings)

    # LIGNE 4
    $row4Y = $row3Y + $btnH + $spacingV

    $tileSharedMailbox = New-MenuTile -Text (Get-Text "main_menu.tile_shared_mailbox") -SubText (Get-Text "main_menu.tile_shared_mailbox_desc") `
        -X $leftCol -Y $row4Y -Color ([System.Drawing.Color]::FromArgb(23, 162, 184))
    Set-TileAction -Panel $tileSharedMailbox -Action {
        Write-Log -Level "INFO" -Action "GUI" -Message "Ouverture de l'audit Shared Mailbox."
        Show-SharedMailboxAuditForm
    }
    $form.Controls.Add($tileSharedMailbox)

    $tileNestedAudit = New-MenuTile -Text (Get-Text "main_menu.tile_nested_audit") -SubText (Get-Text "main_menu.tile_nested_audit_desc") `
        -X $rightCol -Y $row4Y -Color ([System.Drawing.Color]::FromArgb(32, 201, 151))
    Set-TileAction -Panel $tileNestedAudit -Action {
        Write-Log -Level "INFO" -Action "GUI" -Message "Ouverture de l'audit Nested Groups."
        Show-NestedGroupAuditForm
    }
    $form.Controls.Add($tileNestedAudit)

    # =================================================================
    # Barre de statut et pied de page
    # =================================================================
    $statusBar = New-Object System.Windows.Forms.StatusStrip
    $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $statusLabel.Text = "$(Get-Text 'app.ready') — $($Config.client_name) — $(Get-Text 'app.domain_label') : $($Config.smtp_domain)"
    $statusBar.Items.Add($statusLabel) | Out-Null
    $form.Controls.Add($statusBar)

    $lblFooter = New-Object System.Windows.Forms.Label
    $lblFooter.Text = Get-Text "app.footer"
    $lblFooter.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblFooter.ForeColor = [System.Drawing.Color]::FromArgb(173, 181, 189)
    $lblFooter.Location = New-Object System.Drawing.Point(25, 868)
    $lblFooter.Size = New-Object System.Drawing.Size(450, 20)
    $lblFooter.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($lblFooter)

    [System.Windows.Forms.Application]::Run($form)
}