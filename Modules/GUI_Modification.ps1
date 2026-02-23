<#
.FICHIER
    Modules/GUI_Modification.ps1

.ROLE
    Formulaire de modification des attributs d'un employé existant.
    Menu organisé en sections : Profil, Contact, Messagerie/UPN, Licences,
    Groupes, Sécurité, Audit.

.DEPENDANCES
    - Core/Functions.ps1  (Write-Log, Show-ConfirmDialog, Show-ResultDialog,
                           New-SecurePassword, Show-PasswordDialog)
    - Core/GraphAPI.ps1   (Search-AzUsers, Get-AzUser, Set-AzUser,
                           Set-AzUserManager, Set-AzUserLicense,
                           Reset-AzUserPassword, Disable-AzUser, Enable-AzUser,
                           Get-AzUserManager, Revoke-AzUserSessions,
                           Get-AzUserGroups, Remove-AzUserFromGroup,
                           Get-AzUserSignInLogs, Get-AzTenantDomains,
                           Set-AzUserUPN, Add-AzUserProxyAddress,
                           Remove-AzUserProxyAddress, Get-AzUserProxyAddresses,
                           Remove-AzUserFromGroups, Set-AzUserLicenses,
                           Reset-AzUserMfaMethods)
    - Core/Lang.ps1       (Get-Text)
    - Variable globale    $Config

.AUTEUR
    [Equipe IT — M365 Monster]
#>

# ============================================================
#  CONSTANTES DE COULEURS (cohérence visuelle)
# ============================================================
$script:COLOR_BLUE     = [System.Drawing.Color]::FromArgb(0, 123, 255)
$script:COLOR_ORANGE   = [System.Drawing.Color]::FromArgb(255, 140, 0)
$script:COLOR_RED      = [System.Drawing.Color]::FromArgb(220, 53, 69)
$script:COLOR_GREEN    = [System.Drawing.Color]::FromArgb(40, 167, 69)
$script:COLOR_BG       = [System.Drawing.Color]::WhiteSmoke
$script:COLOR_SECTION  = [System.Drawing.Color]::FromArgb(240, 242, 245)
$script:COLOR_WHITE    = [System.Drawing.Color]::White

# ============================================================
#  HELPERS INTERNES
# ============================================================

function New-SubForm {
    <#
    .SYNOPSIS
        Crée un sous-formulaire modal standardisé.
    #>
    param(
        [string]$Titre,
        [int]$Largeur  = 480,
        [int]$Hauteur  = 300
    )

    $f = New-Object System.Windows.Forms.Form
    $f.Text            = $Titre
    $f.Size            = New-Object System.Drawing.Size($Largeur, $Hauteur)
    $f.StartPosition   = "CenterScreen"
    $f.FormBorderStyle = "FixedDialog"
    $f.MaximizeBox     = $false
    $f.MinimizeBox     = $false
    $f.BackColor       = $script:COLOR_BG
    return $f
}

function New-BtnAppliquer {
    <#
    .SYNOPSIS
        Crée un bouton "Appliquer" standardisé (bleu).
    #>
    param([int]$X, [int]$Y)

    $b = New-Object System.Windows.Forms.Button
    $b.Text      = Get-Text "modification.btn_apply"
    $b.Location  = New-Object System.Drawing.Point($X, $Y)
    $b.Size      = New-Object System.Drawing.Size(110, 35)
    $b.BackColor = $script:COLOR_BLUE
    $b.ForeColor = $script:COLOR_WHITE
    $b.FlatStyle = "Flat"
    $b.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    return $b
}

function New-BtnAnnuler {
    <#
    .SYNOPSIS
        Crée un bouton "Annuler" standardisé.
    #>
    param([int]$X, [int]$Y)

    $b = New-Object System.Windows.Forms.Button
    $b.Text        = Get-Text "modification.btn_cancel"
    $b.Location    = New-Object System.Drawing.Point($X, $Y)
    $b.Size        = New-Object System.Drawing.Size(110, 35)
    $b.FlatStyle   = "Flat"
    $b.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    return $b
}

function New-LabelCurrent {
    <#
    .SYNOPSIS
        Crée un label "Valeur actuelle : X" grisé.
    #>
    param([string]$Texte, [int]$X = 15, [int]$Y = 15, [int]$Largeur = 430)

    $l = New-Object System.Windows.Forms.Label
    $l.Text      = $Texte
    $l.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $l.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $l.Location  = New-Object System.Drawing.Point($X, $Y)
    $l.Size      = New-Object System.Drawing.Size($Largeur, 20)
    return $l
}

function New-SectionLabel {
    <#
    .SYNOPSIS
        Crée un label de section dans le panneau d'actions.
    #>
    param([string]$Texte, [int]$Y)

    $l = New-Object System.Windows.Forms.Label
    $l.Text      = "  $Texte"
    $l.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $l.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $l.BackColor = $script:COLOR_SECTION
    $l.Location  = New-Object System.Drawing.Point(0, $Y)
    $l.Size      = New-Object System.Drawing.Size(200, 22)
    $l.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    return $l
}

function New-MenuButton {
    <#
    .SYNOPSIS
        Crée un bouton de menu latéral standardisé.
    #>
    param([string]$Texte, [string]$Tag, [int]$Y)

    $b = New-Object System.Windows.Forms.Button
    $b.Text      = "  $Texte"
    $b.Tag       = $Tag
    $b.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
    $b.Location  = New-Object System.Drawing.Point(0, $Y)
    $b.Size      = New-Object System.Drawing.Size(200, 30)
    $b.FlatStyle = "Flat"
    $b.BackColor = $script:COLOR_WHITE
    $b.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $b.Padding   = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
    return $b
}

# ============================================================
#  FORMULAIRE PRINCIPAL
# ============================================================

function Show-ModificationForm {
    <#
    .SYNOPSIS
        Formulaire principal de modification — recherche + menu latéral.
    .OUTPUTS
        [void] — Formulaire modal.
    #>

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = Get-Text "modification.title" $Config.client_name
    $form.Size            = New-Object System.Drawing.Size(860, 780)
    $form.StartPosition   = "CenterScreen"
    $form.FormBorderStyle = "Sizable"
    $form.MinimumSize     = New-Object System.Drawing.Size(760, 600)
    $form.MaximizeBox     = $true
    $form.MinimizeBox     = $false
    $form.BackColor       = $script:COLOR_BG

    # --- En-tête ---
    $lblSection = New-Object System.Windows.Forms.Label
    $lblSection.Text     = Get-Text "modification.search_label"
    $lblSection.Font     = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblSection.Location = New-Object System.Drawing.Point(15, 15)
    $lblSection.Size     = New-Object System.Drawing.Size(770, 25)
    $form.Controls.Add($lblSection)

    # --- Champ de recherche ---
    $txtRecherche = New-Object System.Windows.Forms.TextBox
    $txtRecherche.Location = New-Object System.Drawing.Point(15, 50)
    $txtRecherche.Size     = New-Object System.Drawing.Size(570, 25)
    $txtRecherche.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($txtRecherche)

    $btnRecherche = New-Object System.Windows.Forms.Button
    $btnRecherche.Text     = Get-Text "modification.btn_search"
    $btnRecherche.Location = New-Object System.Drawing.Point(595, 50)
    $btnRecherche.Size     = New-Object System.Drawing.Size(90, 25)
    $btnRecherche.FlatStyle = "Flat"
    $form.Controls.Add($btnRecherche)

    # --- Liste des résultats ---
    $lstResultats = New-Object System.Windows.Forms.ListBox
    $lstResultats.Location = New-Object System.Drawing.Point(15, 80)
    $lstResultats.Size     = New-Object System.Drawing.Size(770, 75)
    $lstResultats.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
    $lstResultats.Visible  = $false
    $form.Controls.Add($lstResultats)

    # --- Bandeau utilisateur sélectionné ---
    $pnlUser = New-Object System.Windows.Forms.Panel
    $pnlUser.Location  = New-Object System.Drawing.Point(15, 80)
    $pnlUser.Size      = New-Object System.Drawing.Size(770, 55)
    $pnlUser.BackColor = [System.Drawing.Color]::FromArgb(230, 240, 255)
    $pnlUser.Visible   = $false
    $pnlUser.BorderStyle = "FixedSingle"
    $form.Controls.Add($pnlUser)

    $lblUserInfo = New-Object System.Windows.Forms.Label
    $lblUserInfo.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblUserInfo.ForeColor = [System.Drawing.Color]::FromArgb(0, 70, 150)
    $lblUserInfo.Location  = New-Object System.Drawing.Point(8, 5)
    $lblUserInfo.Size      = New-Object System.Drawing.Size(650, 20)
    $pnlUser.Controls.Add($lblUserInfo)

    $lblUserInfo2 = New-Object System.Windows.Forms.Label
    $lblUserInfo2.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblUserInfo2.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $lblUserInfo2.Location  = New-Object System.Drawing.Point(8, 28)
    $lblUserInfo2.Size      = New-Object System.Drawing.Size(650, 18)
    $pnlUser.Controls.Add($lblUserInfo2)

    $btnChanger = New-Object System.Windows.Forms.Button
    $btnChanger.Text      = Get-Text "modification.btn_change_user"
    $btnChanger.Location  = New-Object System.Drawing.Point(668, 13)
    $btnChanger.Size      = New-Object System.Drawing.Size(90, 28)
    $btnChanger.FlatStyle = "Flat"
    $btnChanger.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
    $pnlUser.Controls.Add($btnChanger)

    # --- Zone principale : menu gauche + zone de travail droite ---
    $pnlMain = New-Object System.Windows.Forms.Panel
    $pnlMain.Location   = New-Object System.Drawing.Point(15, 145)
    $pnlMain.Size       = New-Object System.Drawing.Size(810, 570)
    $pnlMain.Visible    = $false
    $pnlMain.Anchor     = ([System.Windows.Forms.AnchorStyles]::Top -bor `
                           [System.Windows.Forms.AnchorStyles]::Left -bor `
                           [System.Windows.Forms.AnchorStyles]::Right)
    $form.Controls.Add($pnlMain)

    # Menu lateral gauche -- scrollable si contenu depasse la hauteur
    $pnlMenu = New-Object System.Windows.Forms.Panel
    $pnlMenu.Location    = New-Object System.Drawing.Point(0, 0)
    $pnlMenu.Size        = New-Object System.Drawing.Size(210, 570)
    $pnlMenu.BackColor   = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $pnlMenu.BorderStyle = "FixedSingle"
    $pnlMenu.AutoScroll  = $true
    $pnlMenu.Anchor      = ([System.Windows.Forms.AnchorStyles]::Top -bor `
                            [System.Windows.Forms.AnchorStyles]::Bottom -bor `
                            [System.Windows.Forms.AnchorStyles]::Left)
    $pnlMain.Controls.Add($pnlMenu)

    # Zone droite : panneau profil + connexions recentes
    $pnlRight = New-Object System.Windows.Forms.Panel
    $pnlRight.Location  = New-Object System.Drawing.Point(215, 0)
    $pnlRight.Size      = New-Object System.Drawing.Size(590, 570)
    $pnlRight.BackColor = [System.Drawing.Color]::White
    $pnlRight.BorderStyle = "FixedSingle"
    $pnlRight.Anchor    = ([System.Windows.Forms.AnchorStyles]::Top -bor `
                           [System.Windows.Forms.AnchorStyles]::Bottom -bor `
                           [System.Windows.Forms.AnchorStyles]::Left -bor `
                           [System.Windows.Forms.AnchorStyles]::Right)
    $pnlMain.Controls.Add($pnlRight)

    # Message par defaut (aucun utilisateur sélectionné)
    $lblInstruction = New-Object System.Windows.Forms.Label
    $lblInstruction.Text      = Get-Text "modification.select_action"
    $lblInstruction.Font      = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblInstruction.ForeColor = [System.Drawing.Color]::FromArgb(160, 160, 160)
    $lblInstruction.Location  = New-Object System.Drawing.Point(50, 220)
    $lblInstruction.Size      = New-Object System.Drawing.Size(490, 80)
    $lblInstruction.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $pnlRight.Controls.Add($lblInstruction)

    # ---- Fiche profil (visible après sélection) ----
    $pnlProfile = New-Object System.Windows.Forms.Panel
    $pnlProfile.Location  = New-Object System.Drawing.Point(5, 5)
    $pnlProfile.Size      = New-Object System.Drawing.Size(578, 560)
    $pnlProfile.Visible   = $false
    $pnlProfile.AutoScroll = $true
    $pnlRight.Controls.Add($pnlProfile)

    # En-tete bleu profil
    $pnlProfileHeader = New-Object System.Windows.Forms.Panel
    $pnlProfileHeader.Location  = New-Object System.Drawing.Point(0, 0)
    $pnlProfileHeader.Size      = New-Object System.Drawing.Size(578, 55)
    $pnlProfileHeader.BackColor = [System.Drawing.Color]::FromArgb(0, 90, 180)
    $pnlProfile.Controls.Add($pnlProfileHeader)

    $lblProfileName = New-Object System.Windows.Forms.Label
    $lblProfileName.Font      = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblProfileName.ForeColor = [System.Drawing.Color]::White
    $lblProfileName.Location  = New-Object System.Drawing.Point(12, 8)
    $lblProfileName.Size      = New-Object System.Drawing.Size(554, 24)
    $pnlProfileHeader.Controls.Add($lblProfileName)

    $lblProfileUpn = New-Object System.Windows.Forms.Label
    $lblProfileUpn.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblProfileUpn.ForeColor = [System.Drawing.Color]::FromArgb(200, 225, 255)
    $lblProfileUpn.Location  = New-Object System.Drawing.Point(12, 33)
    $lblProfileUpn.Size      = New-Object System.Drawing.Size(554, 18)
    $pnlProfileHeader.Controls.Add($lblProfileUpn)

    # Grille d informations
    $script:ProfileLabels = @{}
    $profileFields = @(
        @{ Key = "dept";     Label = "Département" },
        @{ Key = "title";    Label = "Titre" },
        @{ Key = "emptype";  Label = "Employee Type" },
        @{ Key = "country";  Label = "Pays" },
        @{ Key = "office";   Label = "Bureau" },
        @{ Key = "mobile";   Label = "Mobile" },
        @{ Key = "phone";    Label = "Poste fixe" },
        @{ Key = "status";   Label = "Statut" }
    )
    $yField = 65
    foreach ($pf in $profileFields) {
        $lLbl = New-Object System.Windows.Forms.Label
        $lLbl.Text      = $pf.Label + " :"
        $lLbl.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $lLbl.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
        $lLbl.Location  = New-Object System.Drawing.Point(12, $yField)
        $lLbl.Size      = New-Object System.Drawing.Size(120, 20)
        $pnlProfile.Controls.Add($lLbl)

        $lVal = New-Object System.Windows.Forms.Label
        $lVal.Text      = ""
        $lVal.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
        $lVal.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
        $lVal.Location  = New-Object System.Drawing.Point(135, $yField)
        $lVal.Size      = New-Object System.Drawing.Size(430, 20)
        $pnlProfile.Controls.Add($lVal)
        $script:ProfileLabels[$pf.Key] = $lVal
        $yField += 26
    }

    # Separateur + titre dernières connexions
    $lblSepLine = New-Object System.Windows.Forms.Label
    $lblSepLine.Text      = "─── Dernières connexions ───────────────────────────────────────"
    $lblSepLine.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblSepLine.ForeColor = [System.Drawing.Color]::FromArgb(150, 150, 150)
    $lblSepLine.Location  = New-Object System.Drawing.Point(12, $yField)
    $lblSepLine.Size      = New-Object System.Drawing.Size(554, 18)
    $pnlProfile.Controls.Add($lblSepLine)
    $yField += 22

    $dgvSignIn = New-Object System.Windows.Forms.DataGridView
    $dgvSignIn.Location           = New-Object System.Drawing.Point(5, $yField)
    $dgvSignIn.Size               = New-Object System.Drawing.Size(564, 155)
    $dgvSignIn.ReadOnly           = $true
    $dgvSignIn.AutoSizeColumnsMode = "Fill"
    $dgvSignIn.RowHeadersVisible  = $false
    $dgvSignIn.AllowUserToAddRows = $false
    $dgvSignIn.BackgroundColor    = [System.Drawing.Color]::White
    $dgvSignIn.BorderStyle        = "None"
    $dgvSignIn.Font               = New-Object System.Drawing.Font("Segoe UI", 8)
    $dgvSignIn.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $dgvSignIn.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 242, 245)
    $dgvSignIn.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
    foreach ($col in @("Date", "Application", "IP", "Statut")) {
        $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c.HeaderText = $col
        $c.Name       = $col
        $dgvSignIn.Columns.Add($c) | Out-Null
    }
    $pnlProfile.Controls.Add($dgvSignIn)

    # ============================================================
    #  CONSTRUCTION DU MENU LATÉRAL
    # ============================================================
    $menuItems = @(
        # Section Profil
        @{ Type = "section"; Text = Get-Text "modification.section_profile" },
        @{ Type = "btn"; Text = Get-Text "modification.action_name";          Tag = "name" },
        @{ Type = "btn"; Text = Get-Text "modification.action_title";         Tag = "title" },
        @{ Type = "btn"; Text = Get-Text "modification.action_department";    Tag = "dept" },
        @{ Type = "btn"; Text = Get-Text "modification.action_employee_type"; Tag = "emptype" },
        @{ Type = "btn"; Text = Get-Text "modification.action_country";       Tag = "country" },
        @{ Type = "btn"; Text = Get-Text "modification.action_address";       Tag = "address" },
        @{ Type = "btn"; Text = Get-Text "modification.action_office";        Tag = "office" },
        @{ Type = "btn"; Text = Get-Text "modification.action_manager";       Tag = "manager" },

        # Section Contact
        @{ Type = "section"; Text = Get-Text "modification.section_contact" },
        @{ Type = "btn"; Text = Get-Text "modification.action_mobile";        Tag = "mobile" },
        @{ Type = "btn"; Text = Get-Text "modification.action_phone";         Tag = "phone" },

        # Section Messagerie
        @{ Type = "section"; Text = Get-Text "modification.section_messaging" },
        @{ Type = "btn"; Text = Get-Text "modification.action_upn";           Tag = "upn" },
        @{ Type = "btn"; Text = Get-Text "modification.action_aliases";       Tag = "aliases" },

        # Section Licences
        @{ Type = "section"; Text = Get-Text "modification.section_licenses" },
        @{ Type = "btn"; Text = Get-Text "modification.action_licenses";      Tag = "licenses" },

        # Section Groupes
        @{ Type = "section"; Text = Get-Text "modification.section_groups" },
        @{ Type = "btn"; Text = Get-Text "modification.action_groups_manage"; Tag = "groups" },

        # Section Sécurité
        @{ Type = "section"; Text = Get-Text "modification.section_security" },
        @{ Type = "btn"; Text = Get-Text "modification.action_password";      Tag = "password" },
        @{ Type = "btn"; Text = Get-Text "modification.action_revoke";        Tag = "revoke" },
        @{ Type = "btn"; Text = Get-Text "modification.action_mfa_reset";     Tag = "mfa" },
        @{ Type = "btn"; Text = Get-Text "modification.action_toggle";        Tag = "toggle" },

        # Section Audit
        @{ Type = "section"; Text = Get-Text "modification.section_audit" },
        @{ Type = "btn"; Text = Get-Text "modification.action_signin_logs";   Tag = "logs" }
    )

    $yMenu = 5
    foreach ($item in $menuItems) {
        if ($item.Type -eq "section") {
            $lbl = New-SectionLabel -Texte $item.Text -Y $yMenu
            $pnlMenu.Controls.Add($lbl)
            $yMenu += 24
        }
        else {
            $btn = New-MenuButton -Texte $item.Text -Tag $item.Tag -Y $yMenu
            $btn.Add_Click({
                $tag      = $this.Tag
                $userId   = $script:SelectedUserId
                $upn      = $script:SelectedUserUPN
                $userData = $script:SelectedUserData

                switch ($tag) {
                    "name"     { Show-ModifyName        -UserId $userId -UPN $upn -CurrentGiven $userData.GivenName -CurrentSurname $userData.Surname -CurrentDisplay $userData.DisplayName }
                    "title"    { Show-ModifyComboField   -UserId $userId -UPN $upn -Field "JobTitle"         -Label (Get-Text "modification.action_title")         -CurrentValue $userData.JobTitle         -Items @() -GraphProperty "jobTitle" }
                    "dept"     { Show-ModifyComboField   -UserId $userId -UPN $upn -Field "Department"       -Label (Get-Text "modification.action_department")    -CurrentValue $userData.Department       -Items $Config.departments -GraphProperty "department" }
                    "emptype"  { Show-ModifyComboField   -UserId $userId -UPN $upn -Field "EmployeeType"     -Label (Get-Text "modification.action_employee_type") -CurrentValue $userData.EmployeeType     -Items $Config.employee_types -GraphProperty "employeeType" }
                    "country"  { Show-ModifyCountry      -UserId $userId -UPN $upn -CurrentValue $userData.Country }
                    "address"  { Show-ModifyAddress      -UserId $userId -UPN $upn -CurrentData $userData }
                    "office"   { Show-ModifyComboField   -UserId $userId -UPN $upn -Field "OfficeLocation"   -Label (Get-Text "modification.action_office")        -CurrentValue $userData.OfficeLocation   -Items @() -GraphProperty "officeLocation" }
                    "manager"  { Show-ModifyManager      -UserId $userId -UPN $upn }
                    "mobile"   { Show-ModifySimpleField  -UserId $userId -UPN $upn -Field "MobilePhone"      -Label (Get-Text "modification.action_mobile")        -CurrentValue $userData.MobilePhone }
                    "phone"    { Show-ModifyBusinessPhone -UserId $userId -UPN $upn -CurrentValue ($userData.BusinessPhones | Select-Object -First 1) }
                    "upn"      { Show-ModifyUPN          -UserId $userId -UPN $upn }
                    "aliases"  { Show-ManageAliases      -UserId $userId -UPN $upn }
                    "licenses" { Show-ManageLicenses     -UserId $userId -UPN $upn }
                    "groups"   { Show-ManageGroups       -UserId $userId -UPN $upn }
                    "password" { Show-ResetPassword      -UserId $userId -UPN $upn }
                    "revoke"   { Invoke-RevokeSession    -UserId $userId -UPN $upn }
                    "mfa"      { Invoke-MfaReset         -UserId $userId -UPN $upn }
                    "toggle"   { Show-ToggleAccount      -UserId $userId -UPN $upn -IsEnabled $userData.AccountEnabled }
                    "logs"     { Show-SignInLogs         -UserId $userId -UPN $upn }
                }
            })
            $pnlMenu.Controls.Add($btn)
            $yMenu += 31
        }
    }

    # Forcer la hauteur de contenu du menu pour que le scroll soit actif
    $pnlMenu.AutoScrollMinSize = New-Object System.Drawing.Size(190, ($yMenu + 10))

    # ============================================================
    #  LOGIQUE DE RECHERCHE
    # ============================================================
    $script:SelectedUserId   = $null
    $script:SelectedUserUPN  = $null
    $script:SelectedUserData = $null
    $script:SearchResults    = @()

    $btnRecherche.Add_Click({
        $terme = $txtRecherche.Text.Trim()
        if ($terme.Length -lt 2) { return }
        $lstResultats.Items.Clear()
        $result = Search-AzUsers -SearchTerm $terme -MaxResults 15
        if ($result.Success -and $result.Data) {
            $script:SearchResults = @($result.Data)
            foreach ($user in $script:SearchResults) {
                $statut = if ($user.AccountEnabled) { Get-Text "modification.status_active" } else { Get-Text "modification.status_disabled" }
                $lstResultats.Items.Add("$($user.DisplayName) — $($user.UserPrincipalName) [$statut]") | Out-Null
            }
            $pnlUser.Visible  = $false
            $pnlMain.Visible  = $false
            $lstResultats.Visible = $true
        }
    })

    $txtRecherche.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $btnRecherche.PerformClick()
            $_.SuppressKeyPress = $true
        }
    })

    $lstResultats.Add_SelectedIndexChanged({
        if ($lstResultats.SelectedIndex -ge 0 -and $lstResultats.SelectedIndex -lt $script:SearchResults.Count) {
            Invoke-SelectUser -Selected $script:SearchResults[$lstResultats.SelectedIndex]
            $lstResultats.Visible = $false
        }
    })

    $btnChanger.Add_Click({
        $pnlUser.Visible      = $false
        $pnlMain.Visible      = $false
        $txtRecherche.Text    = ""
        $lstResultats.Items.Clear()
        $lstResultats.Visible = $false
        $txtRecherche.Focus()
    })

    # Panel footer ancre en bas — garantit la visibilite des boutons
    $pnlFooter = New-Object System.Windows.Forms.Panel
    $pnlFooter.Dock      = [System.Windows.Forms.DockStyle]::Bottom
    $pnlFooter.Height    = 44
    $pnlFooter.BackColor = [System.Drawing.Color]::FromArgb(235, 237, 240)
    $form.Controls.Add($pnlFooter)

    # Recalcul dynamique de la hauteur de pnlMain pour qu'il s'arrête toujours
    # au-dessus du pnlFooter — nécessaire car Anchor::Bottom et Dock::Bottom entrent en conflit
    $script:AjusterLayout = {
        $pnlMain.Height = $form.ClientSize.Height - $pnlMain.Top - $pnlFooter.Height - 4
    }
    $form.Add_Shown({ & $script:AjusterLayout })
    $form.Add_Resize({ & $script:AjusterLayout })

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text      = Get-Text "modification.btn_refresh"
    $btnRefresh.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnRefresh.Location  = New-Object System.Drawing.Point(8, 7)
    $btnRefresh.Size      = New-Object System.Drawing.Size(150, 30)
    $btnRefresh.FlatStyle = "Flat"
    $btnRefresh.Anchor    = ([System.Windows.Forms.AnchorStyles]::Bottom -bor `
                             [System.Windows.Forms.AnchorStyles]::Left)
    $btnRefresh.Add_Click({
        if ($null -ne $script:SelectedUserId) {
            $sel = [PSCustomObject]@{ Id = $script:SelectedUserId; UserPrincipalName = $script:SelectedUserUPN }
            Invoke-SelectUser -Selected $sel
        }
    })
    $pnlFooter.Controls.Add($btnRefresh)

    $btnFermer = New-Object System.Windows.Forms.Button
    $btnFermer.Text        = Get-Text "modification.btn_close"
    $btnFermer.Font        = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnFermer.Location    = New-Object System.Drawing.Point(690, 7)
    $btnFermer.Size        = New-Object System.Drawing.Size(150, 30)
    $btnFermer.FlatStyle   = "Flat"
    $btnFermer.BackColor   = $script:COLOR_RED
    $btnFermer.ForeColor   = $script:COLOR_WHITE
    $btnFermer.Anchor      = ([System.Windows.Forms.AnchorStyles]::Bottom -bor `
                              [System.Windows.Forms.AnchorStyles]::Right)
    $btnFermer.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $pnlFooter.Controls.Add($btnFermer)
    $form.CancelButton = $btnFermer

    $form.ShowDialog() | Out-Null
    $form.Dispose()
}

function Invoke-SelectUser {
    <#
    .SYNOPSIS
        Charge les détails d'un utilisateur sélectionné et affiche le menu.
    #>
    param($Selected)

    # Récupération des propriétés étendues nécessaires
    $detailResult = Get-MgUser -UserId $Selected.Id -Property `
        "id,displayName,givenName,surname,userPrincipalName,accountEnabled,department,
         jobTitle,mobilePhone,businessPhones,mail,employeeType,assignedLicenses,
         country,streetAddress,city,postalCode,state,officeLocation,proxyAddresses,
         usageLocation" -ErrorAction SilentlyContinue

    if ($null -ne $detailResult) {
        $script:SelectedUserId   = $Selected.Id
        $script:SelectedUserUPN  = $Selected.UserPrincipalName
        $script:SelectedUserData = $detailResult

        $statut = if ($detailResult.AccountEnabled) { Get-Text "modification.status_active" } else { Get-Text "modification.status_disabled" }

        # Bandeau principal
        $lblUserInfo.Text  = "$($detailResult.DisplayName) — $($detailResult.UserPrincipalName)"
        $lblUserInfo2.Text = "$($detailResult.Department) | $($detailResult.JobTitle) | $statut"

        # Remplir le panneau profil droit
        $lblProfileName.Text = $detailResult.DisplayName
        $lblProfileUpn.Text  = $detailResult.UserPrincipalName
        $script:ProfileLabels["dept"].Text    = if ($detailResult.Department)    { $detailResult.Department }    else { "—" }
        $script:ProfileLabels["title"].Text   = if ($detailResult.JobTitle)      { $detailResult.JobTitle }      else { "—" }
        $script:ProfileLabels["emptype"].Text = if ($detailResult.EmployeeType)  { $detailResult.EmployeeType }  else { "—" }
        $script:ProfileLabels["country"].Text = if ($detailResult.Country)       { $detailResult.Country }       else { "—" }
        $script:ProfileLabels["office"].Text  = if ($detailResult.OfficeLocation){ $detailResult.OfficeLocation }else { "—" }
        $script:ProfileLabels["mobile"].Text  = if ($detailResult.MobilePhone)   { $detailResult.MobilePhone }   else { "—" }
        $script:ProfileLabels["phone"].Text   = if ($detailResult.BusinessPhones -and $detailResult.BusinessPhones.Count -gt 0) { $detailResult.BusinessPhones[0] } else { "—" }
        $script:ProfileLabels["status"].Text  = $statut

        # Charger les dernières connexions (5 dernières)
        $dgvSignIn.Rows.Clear()
        try {
            $upnForQuery = $detailResult.UserPrincipalName
            $signIns = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=userPrincipalName eq '$upnForQuery'&`$top=5&`$orderby=createdDateTime desc" `
                -ErrorAction Stop
            if ($signIns.value) {
                foreach ($log in $signIns.value) {
                    $date   = if ($log.createdDateTime) { [datetime]$log.createdDateTime | Get-Date -Format "MM-dd HH:mm" } else { "-" }
                    $app    = if ($log.appDisplayName)  { $log.appDisplayName }  else { "-" }
                    $ip     = if ($log.ipAddress)       { $log.ipAddress }       else { "-" }
                    $ok     = if ($log.status -and $log.status.errorCode -eq 0)  { "✓" } else { "✗" }
                    $row    = $dgvSignIn.Rows.Add($date, $app, $ip, $ok)
                    if ($ok -eq "✗") {
                        $dgvSignIn.Rows[$row].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Crimson
                    }
                }
            }
            else {
                $dgvSignIn.Rows.Add("-", "Aucune connexion récente", "-", "-") | Out-Null
            }
        }
        catch {
            $dgvSignIn.Rows.Add("-", "AuditLog.Read.All requis", "-", "-") | Out-Null
        }

        $lblInstruction.Visible = $false
        $pnlProfile.Visible     = $true
    }
    else {
        $script:SelectedUserId   = $Selected.Id
        $script:SelectedUserUPN  = $Selected.UserPrincipalName
        $script:SelectedUserData = $Selected
        $lblUserInfo.Text  = "$($Selected.DisplayName) — $($Selected.UserPrincipalName)"
        $lblUserInfo2.Text = ""
    }

    $pnlUser.Visible = $true
    $pnlMain.Visible = $true
}

# ============================================================
#  SOUS-FORMULAIRES — PROFIL & IDENTITÉ
# ============================================================

function Show-ModifyName {
    <#
    .SYNOPSIS
        Modification du prénom, nom de famille et nom d'affichage.
    #>
    param(
        [string]$UserId,
        [string]$UPN,
        [string]$CurrentGiven,
        [string]$CurrentSurname,
        [string]$CurrentDisplay
    )

    $f = New-SubForm -Titre (Get-Text "modification.action_name") -Largeur 460 -Hauteur 310
    $f.Controls.Add((New-LabelCurrent -Texte (Get-Text "modification.current_value" $CurrentDisplay)))

    $y = 45
    foreach ($pair in @(
        @{ Label = Get-Text "modification.field_givenname";   Var = "txtGiven";   Val = $CurrentGiven },
        @{ Label = Get-Text "modification.field_surname";     Var = "txtSurname"; Val = $CurrentSurname },
        @{ Label = Get-Text "modification.field_displayname"; Var = "txtDisplay"; Val = $CurrentDisplay }
    )) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text     = $pair.Label
        $lbl.Location = New-Object System.Drawing.Point(15, ($y + 3))
        $lbl.Size     = New-Object System.Drawing.Size(130, 20)
        $f.Controls.Add($lbl)

        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Text     = $pair.Val
        $txt.Location = New-Object System.Drawing.Point(150, $y)
        $txt.Size     = New-Object System.Drawing.Size(285, 25)
        $txt.Name     = $pair.Var
        $f.Controls.Add($txt)
        $y += 40
    }

    $btnA = New-BtnAppliquer -X 120 -Y ($y + 10)
    $btnA.Add_Click({
        $txts   = $f.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] }
        $given   = ($txts | Where-Object { $_.Name -eq "txtGiven" }).Text.Trim()
        $surname = ($txts | Where-Object { $_.Name -eq "txtSurname" }).Text.Trim()
        $display = ($txts | Where-Object { $_.Name -eq "txtDisplay" }).Text.Trim()

        if ([string]::IsNullOrWhiteSpace($given) -or [string]::IsNullOrWhiteSpace($surname) -or [string]::IsNullOrWhiteSpace($display)) {
            [System.Windows.Forms.MessageBox]::Show(
                (Get-Text "modification.error_required_fields"),
                (Get-Text "modification.error_title"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") `
            -Message (Get-Text "modification.confirm_name" $display $UPN)

        if ($confirm) {
            $result = Set-AzUser -UserId $UserId -Properties @{
                GivenName   = $given
                Surname     = $surname
                DisplayName = $display
            }
            if ($result.Success) {
                Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.success_name") -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "MODIFY_NAME" -UPN $UPN -Message "Nom modifié : '$CurrentDisplay' -> '$display'"
                $f.Close()
            }
            else {
                Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $result.Error -IsSuccess $false
            }
        }
    })
    $f.Controls.Add($btnA)

    $btnC = New-BtnAnnuler -X 240 -Y ($y + 10)
    $f.Controls.Add($btnC)
    $f.CancelButton = $btnC

    $f.ShowDialog() | Out-Null
    $f.Dispose()
}

function Format-GraphErrorMessage {
    <#
    .SYNOPSIS
        Formate un message d erreur Graph avec conseils contextuels.
        Détecte : Forbidden (403), Authorization_RequestDenied, BadRequest (400), read-only.
    #>
    param([string]$ErrorMessage, [string]$Field = "")

    # 403 Forbidden — scope absent du token actuel, reconnexion nécessaire
    if ($ErrorMessage -like "*Forbidden*" -or $ErrorMessage -like "*403*") {
        if ($Field -in @("MobilePhone", "BusinessPhones")) {
            return (Get-Text "modification.error_phone_forbidden")
        }
        return (Get-Text "modification.error_forbidden_reconnect")
    }

    # Authorization_RequestDenied / Insufficient privileges (rôle insuffisant)
    if ($ErrorMessage -like "*Authorization_RequestDenied*" -or $ErrorMessage -like "*Insufficient privileges*") {
        if ($Field -in @("MobilePhone", "BusinessPhones")) {
            return (Get-Text "modification.error_phone_permission_hint")
        }
        return (Get-Text "modification.error_permission_hint")
    }

    # 400 BadRequest — proxyAddresses ou autre contrainte Exchange
    if ($ErrorMessage -like "*BadRequest*" -or $ErrorMessage -like "*Bad Request*") {
        if ($Field -in @("proxyAddresses", "ProxyAddresses")) {
            return (Get-Text "modification.error_proxy_badrequest")
        }
        return "$ErrorMessage"
    }

    # Propriété en lecture seule (Exchange Online)
    if ($ErrorMessage -like "*read-only*") {
        return (Get-Text "modification.error_proxy_readonly")
    }

    return $ErrorMessage
}

function Show-ModifySimpleField {
    <#
    .SYNOPSIS
        Sous-formulaire generique pour un champ texte libre.
    .PARAMETER Field
        Nom de la propriete Graph (ex: "MobilePhone").
    .PARAMETER Label
        Libelle affiche.
    #>
    param(
        [string]$UserId,
        [string]$UPN,
        [string]$Field,
        [string]$Label,
        [string]$CurrentValue
    )

    $f = New-SubForm -Titre $Label -Largeur 480 -Hauteur 220
    $f.Controls.Add((New-LabelCurrent -Texte (Get-Text "modification.current_value" $CurrentValue)))

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text     = $Label
    $lbl.Location = New-Object System.Drawing.Point(15, 50)
    $lbl.Size     = New-Object System.Drawing.Size(130, 20)
    $f.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Text     = $CurrentValue
    $txt.Location = New-Object System.Drawing.Point(150, 47)
    $txt.Size     = New-Object System.Drawing.Size(305, 25)
    $f.Controls.Add($txt)

    # Hint permission pour les champs telephone
    if ($Field -in @("MobilePhone", "BusinessPhones")) {
        $lblHint = New-Object System.Windows.Forms.Label
        $lblHint.Text      = Get-Text "modification.phone_scope_hint"
        $lblHint.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
        $lblHint.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
        $lblHint.Location  = New-Object System.Drawing.Point(15, 78)
        $lblHint.Size      = New-Object System.Drawing.Size(440, 18)
        $f.Controls.Add($lblHint)
    }

    $btnA = New-BtnAppliquer -X 120 -Y 125
    $btnA.Add_Click({
        $newVal = $txt.Text.Trim()
        if ($newVal -eq $CurrentValue) {
            Show-ResultDialog -Titre (Get-Text "modification.info_title") -Message (Get-Text "modification.info_no_change") -IsSuccess $true
            return
        }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") `
            -Message (Get-Text "modification.confirm_field" $Label $CurrentValue $newVal $UPN)
        if ($confirm) {
            # Pour MobilePhone/BusinessPhones : PATCH direct via REST (Update-MgUser bloque sur Exchange Online)
            if ($Field -eq "MobilePhone") {
                $body = @{ mobilePhone = $newVal } | ConvertTo-Json
            }
            elseif ($Field -eq "BusinessPhones") {
                $phones = if ([string]::IsNullOrWhiteSpace($newVal)) { @() } else { @($newVal) }
                $body = @{ businessPhones = $phones } | ConvertTo-Json
            }
            else {
                $body = @{ $Field = $newVal } | ConvertTo-Json
            }
            try {
                Invoke-MgGraphRequest -Method PATCH `
                    -Uri "https://graph.microsoft.com/v1.0/users/$UserId" `
                    -Body $body -ContentType "application/json" -ErrorAction Stop
                Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.success_field" $Label) -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "MODIFY_FIELD" -UPN $UPN -Message "$Field : '$CurrentValue' -> '$newVal'"
                $f.Close()
            }
            catch {
                $errMsg = Format-GraphErrorMessage -ErrorMessage $_.Exception.Message -Field $Field
                Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $errMsg -IsSuccess $false
            }
        }
    })
    $f.Controls.Add($btnA)

    $btnC = New-BtnAnnuler -X 250 -Y 125
    $f.Controls.Add($btnC)
    $f.CancelButton = $btnC

    $f.ShowDialog() | Out-Null
    $f.Dispose()
}

function Show-ModifyComboField {
    <#
    .SYNOPSIS
        Sous-formulaire pour un champ combo + saisie libre.
        Charge dynamiquement les valeurs existantes depuis Graph si $Items est vide.
        Le ComboBox est en mode DropDown (saisie libre autorisee).
    .PARAMETER GraphProperty
        Nom de la propriete Graph pour le chargement dynamique
        ("department", "jobTitle", "employeeType", "officeLocation").
        Optionnel -- si absent, seule la liste $Items est utilisee.
    #>
    param(
        [string]$UserId,
        [string]$UPN,
        [string]$Field,
        [string]$Label,
        [string]$CurrentValue,
        [array]$Items,
        [string]$GraphProperty = ""
    )

    $f = New-SubForm -Titre $Label -Largeur 480 -Hauteur 230
    $f.Controls.Add((New-LabelCurrent -Texte (Get-Text "modification.current_value" $CurrentValue)))

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text     = $Label
    $lbl.Location = New-Object System.Drawing.Point(15, 50)
    $lbl.Size     = New-Object System.Drawing.Size(130, 20)
    $f.Controls.Add($lbl)

    # ComboBox en mode DropDown : liste + saisie libre
    $cbo = New-Object System.Windows.Forms.ComboBox
    $cbo.Location      = New-Object System.Drawing.Point(150, 47)
    $cbo.Size          = New-Object System.Drawing.Size(295, 25)
    $cbo.DropDownStyle = "DropDown"   # saisie libre autorisee
    $cbo.AutoCompleteMode   = "SuggestAppend"
    $cbo.AutoCompleteSource = "ListItems"

    # Remplir depuis la config si disponible
    if ($null -ne $Items -and $Items.Count -gt 0) {
        foreach ($item in $Items) { $cbo.Items.Add($item) | Out-Null }
    }
    $f.Controls.Add($cbo)

    # Label de statut de chargement
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text      = ""
    $lblStatus.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $lblStatus.Location  = New-Object System.Drawing.Point(15, 78)
    $lblStatus.Size      = New-Object System.Drawing.Size(440, 18)
    $f.Controls.Add($lblStatus)

    # Note saisie libre
    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text      = Get-Text "modification.combo_free_entry_note"
    $lblNote.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblNote.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $lblNote.Location  = New-Object System.Drawing.Point(15, 98)
    $lblNote.Size      = New-Object System.Drawing.Size(440, 18)
    $f.Controls.Add($lblNote)

    # Pre-selection
    $idx = $cbo.Items.IndexOf($CurrentValue)
    if ($idx -ge 0) {
        $cbo.SelectedIndex = $idx
    }
    elseif (-not [string]::IsNullOrWhiteSpace($CurrentValue)) {
        $cbo.Text = $CurrentValue
    }

    $btnA = New-BtnAppliquer -X 110 -Y 130
    $f.Controls.Add($btnA)

    $btnC = New-BtnAnnuler -X 235 -Y 130
    $f.Controls.Add($btnC)
    $f.CancelButton = $btnC

    # Chargement dynamique depuis Graph apres affichage
    $f.Add_Shown({
        if (-not [string]::IsNullOrWhiteSpace($GraphProperty)) {
            $lblStatus.Text = Get-Text "modification.combo_loading"
            try {
                $existing = Get-AzDistinctValues -Property $GraphProperty -ErrorAction Stop
                if ($existing.Success -and $existing.Data.Count -gt 0) {
                    foreach ($val in $existing.Data) {
                        if ($val -notin $cbo.Items) {
                            $cbo.Items.Add($val) | Out-Null
                        }
                    }
                    # Re-selectionner apres chargement
                    $idx2 = $cbo.Items.IndexOf($CurrentValue)
                    if ($idx2 -ge 0) { $cbo.SelectedIndex = $idx2 }
                    $lblStatus.Text = Get-Text "modification.combo_loaded" $cbo.Items.Count
                }
                else {
                    $lblStatus.Text = Get-Text "modification.combo_no_existing"
                }
            }
            catch {
                $lblStatus.Text = Get-Text "modification.combo_load_error"
            }
        }
    })

    $btnA.Add_Click({
        $newVal = $cbo.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($newVal)) {
            $msgReq = Get-Text "modification.error_required_fields"
            [System.Windows.Forms.MessageBox]::Show($msgReq, (Get-Text "modification.error_title"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
        if ($newVal -eq $CurrentValue) {
            Show-ResultDialog -Titre (Get-Text "modification.info_title") -Message (Get-Text "modification.info_no_change") -IsSuccess $true
            return
        }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") `
            -Message (Get-Text "modification.confirm_field" $Label $CurrentValue $newVal $UPN)
        if ($confirm) {
            $result = Set-AzUser -UserId $UserId -Properties @{ $Field = $newVal }
            if ($result.Success) {
                Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.success_field" $Label) -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "MODIFY_FIELD" -UPN $UPN -Message "$Field : '$CurrentValue' -> '$newVal'"
                $f.Close()
            }
            else {
                Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $result.Error -IsSuccess $false
            }
        }
    })

    $f.ShowDialog() | Out-Null
    $f.Dispose()
}

function Show-ModifyCountry {
    <#
    .SYNOPSIS
        Modification du pays (country + usageLocation Graph).
    #>
    param([string]$UserId, [string]$UPN, [string]$CurrentValue)

    # Codes ISO 3166-1 alpha-2 courants
    $countries = @(
        "AD","AE","AF","AG","AL","AM","AO","AR","AT","AU","AZ","BA","BB","BD","BE","BF","BG","BH","BI","BJ",
        "BN","BO","BR","BS","BT","BW","BY","BZ","CA","CD","CF","CG","CH","CI","CL","CM","CN","CO","CR","CU",
        "CV","CY","CZ","DE","DJ","DK","DM","DO","DZ","EC","EE","EG","ER","ES","ET","FI","FJ","FR","GA","GB",
        "GD","GE","GH","GM","GN","GQ","GR","GT","GW","GY","HN","HR","HT","HU","ID","IE","IL","IN","IQ","IR",
        "IS","IT","JM","JO","JP","KE","KG","KH","KI","KM","KN","KP","KR","KW","KZ","LA","LB","LC","LI","LK",
        "LR","LS","LT","LU","LV","LY","MA","MC","MD","ME","MG","MH","MK","ML","MM","MN","MR","MT","MU","MV",
        "MW","MX","MY","MZ","NA","NE","NG","NI","NL","NO","NP","NR","NZ","OM","PA","PE","PG","PH","PK","PL",
        "PT","PW","PY","QA","RO","RS","RU","RW","SA","SB","SC","SD","SE","SG","SI","SK","SL","SM","SN","SO",
        "SR","SS","ST","SV","SY","SZ","TD","TG","TH","TJ","TL","TM","TN","TO","TR","TT","TV","TZ","UA","UG",
        "US","UY","UZ","VA","VC","VE","VN","VU","WS","XK","YE","ZA","ZM","ZW"
    ) | Sort-Object

    $f = New-SubForm -Titre (Get-Text "modification.action_country") -Largeur 430 -Hauteur 200
    $f.Controls.Add((New-LabelCurrent -Texte (Get-Text "modification.current_value" $CurrentValue)))

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text     = Get-Text "modification.field_country"
    $lbl.Location = New-Object System.Drawing.Point(15, 50)
    $lbl.Size     = New-Object System.Drawing.Size(130, 20)
    $f.Controls.Add($lbl)

    $cbo = New-Object System.Windows.Forms.ComboBox
    $cbo.Location      = New-Object System.Drawing.Point(150, 47)
    $cbo.Size          = New-Object System.Drawing.Size(120, 25)
    $cbo.DropDownStyle = "DropDownList"
    foreach ($c in $countries) { $cbo.Items.Add($c) | Out-Null }
    $idx = $cbo.Items.IndexOf($CurrentValue.ToUpper())
    if ($idx -ge 0) { $cbo.SelectedIndex = $idx }
    $f.Controls.Add($cbo)

    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text      = Get-Text "modification.country_note"
    $lblNote.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblNote.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $lblNote.Location  = New-Object System.Drawing.Point(15, 78)
    $lblNote.Size      = New-Object System.Drawing.Size(390, 18)
    $f.Controls.Add($lblNote)

    $btnA = New-BtnAppliquer -X 100 -Y 110
    $btnA.Add_Click({
        if ($null -eq $cbo.SelectedItem) { return }
        $newVal = $cbo.SelectedItem.ToString()
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") `
            -Message (Get-Text "modification.confirm_country" $newVal $UPN)
        if ($confirm) {
            # UsageLocation = même code ISO (obligatoire pour les licences)
            $result = Set-AzUser -UserId $UserId -Properties @{
                Country         = $newVal
                UsageLocation   = $newVal
            }
            if ($result.Success) {
                Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.success_country") -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "MODIFY_COUNTRY" -UPN $UPN -Message "Pays : '$CurrentValue' -> '$newVal'"
                $f.Close()
            }
            else {
                Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $result.Error -IsSuccess $false
            }
        }
    })
    $f.Controls.Add($btnA)

    $btnC = New-BtnAnnuler -X 220 -Y 110
    $f.Controls.Add($btnC)
    $f.CancelButton = $btnC

    $f.ShowDialog() | Out-Null
    $f.Dispose()
}

function Show-ModifyAddress {
    <#
    .SYNOPSIS
        Modification complète de l'adresse postale.
    #>
    param([string]$UserId, [string]$UPN, $CurrentData)

    $f = New-SubForm -Titre (Get-Text "modification.action_address") -Largeur 480 -Hauteur 340

    $fields = @(
        @{ Label = Get-Text "modification.field_street"; Field = "StreetAddress"; Val = $CurrentData.StreetAddress },
        @{ Label = Get-Text "modification.field_city";   Field = "City";          Val = $CurrentData.City },
        @{ Label = Get-Text "modification.field_state";  Field = "State";         Val = $CurrentData.State },
        @{ Label = Get-Text "modification.field_postal"; Field = "PostalCode";    Val = $CurrentData.PostalCode }
    )

    $y = 15
    $textBoxes = @{}
    foreach ($fld in $fields) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text     = $fld.Label
        $lbl.Location = New-Object System.Drawing.Point(15, ($y + 3))
        $lbl.Size     = New-Object System.Drawing.Size(130, 20)
        $f.Controls.Add($lbl)

        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Text     = if ($null -eq $fld.Val) { "" } else { $fld.Val }
        $txt.Location = New-Object System.Drawing.Point(150, $y)
        $txt.Size     = New-Object System.Drawing.Size(300, 25)
        $txt.Name     = $fld.Field
        $f.Controls.Add($txt)
        $textBoxes[$fld.Field] = $txt
        $y += 40
    }

    $btnA = New-BtnAppliquer -X 120 -Y ($y + 10)
    $btnA.Add_Click({
        $props = @{}
        foreach ($fld in $fields) {
            $newVal = $textBoxes[$fld.Field].Text.Trim()
            $props[$fld.Field] = $newVal
        }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") `
            -Message (Get-Text "modification.confirm_address" $UPN)
        if ($confirm) {
            $result = Set-AzUser -UserId $UserId -Properties $props
            if ($result.Success) {
                Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.success_address") -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "MODIFY_ADDRESS" -UPN $UPN -Message "Adresse mise à jour."
                $f.Close()
            }
            else {
                Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $result.Error -IsSuccess $false
            }
        }
    })
    $f.Controls.Add($btnA)

    $btnC = New-BtnAnnuler -X 240 -Y ($y + 10)
    $f.Controls.Add($btnC)
    $f.CancelButton = $btnC

    $f.ShowDialog() | Out-Null
    $f.Dispose()
}

function Show-ModifyManager {
    <#
    .SYNOPSIS
        Recherche et changement de manager.
    #>
    param([string]$UserId, [string]$UPN)

    $f = New-SubForm -Titre (Get-Text "modification.action_manager") -Largeur 480 -Hauteur 290

    $mgrResult  = Get-AzUserManager -UserId $UserId
    $currentMgr = if ($mgrResult.Success -and $mgrResult.Data) { $mgrResult.Data.AdditionalProperties.displayName } else { Get-Text "modification.no_manager" }
    $f.Controls.Add((New-LabelCurrent -Texte (Get-Text "modification.current_value" $currentMgr)))

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text     = Get-Text "modification.field_new_manager"
    $lblSearch.Location = New-Object System.Drawing.Point(15, 50)
    $lblSearch.Size     = New-Object System.Drawing.Size(130, 20)
    $f.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(150, 47)
    $txtSearch.Size     = New-Object System.Drawing.Size(210, 25)
    $f.Controls.Add($txtSearch)

    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Text     = Get-Text "modification.btn_search"
    $btnSearch.Location = New-Object System.Drawing.Point(365, 47)
    $btnSearch.Size     = New-Object System.Drawing.Size(80, 25)
    $btnSearch.FlatStyle = "Flat"
    $f.Controls.Add($btnSearch)

    $lstMgr = New-Object System.Windows.Forms.ListBox
    $lstMgr.Location = New-Object System.Drawing.Point(150, 78)
    $lstMgr.Size     = New-Object System.Drawing.Size(295, 80)
    $lstMgr.Visible  = $false
    $f.Controls.Add($lstMgr)

    $script:MgrSearchResults = @()
    $script:NewManagerId     = $null

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
            $txtSearch.Text      = $script:MgrSearchResults[$lstMgr.SelectedIndex].DisplayName
            $lstMgr.Visible      = $false
        }
    })

    $btnA = New-BtnAppliquer -X 120 -Y 200
    $btnA.Add_Click({
        if ($null -eq $script:NewManagerId) { return }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") `
            -Message (Get-Text "modification.confirm_manager" $txtSearch.Text $UPN)
        if ($confirm) {
            $result = Set-AzUserManager -UserId $UserId -ManagerId $script:NewManagerId
            if ($result.Success) {
                Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.success_manager") -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "MODIFY_MANAGER" -UPN $UPN -Message "Manager -> $($txtSearch.Text)"
                $f.Close()
            }
            else {
                Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $result.Error -IsSuccess $false
            }
        }
    })
    $f.Controls.Add($btnA)

    $btnC = New-BtnAnnuler -X 240 -Y 200
    $f.Controls.Add($btnC)
    $f.CancelButton = $btnC

    $f.ShowDialog() | Out-Null
    $f.Dispose()
}

# ============================================================
#  SOUS-FORMULAIRES — CONTACT
# ============================================================

function Show-ModifyBusinessPhone {
    <#
    .SYNOPSIS
        Modification du numéro de poste fixe (businessPhones — tableau dans Graph).
    #>
    param([string]$UserId, [string]$UPN, [string]$CurrentValue)

    $f = New-SubForm -Titre (Get-Text "modification.action_phone") -Largeur 480 -Hauteur 220
    $f.Controls.Add((New-LabelCurrent -Texte (Get-Text "modification.current_value" $CurrentValue)))

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text     = Get-Text "modification.field_phone"
    $lbl.Location = New-Object System.Drawing.Point(15, 50)
    $lbl.Size     = New-Object System.Drawing.Size(130, 20)
    $f.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Text     = $CurrentValue
    $txt.Location = New-Object System.Drawing.Point(150, 47)
    $txt.Size     = New-Object System.Drawing.Size(305, 25)
    $f.Controls.Add($txt)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text      = Get-Text "modification.phone_scope_hint"
    $lblHint.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblHint.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $lblHint.Location  = New-Object System.Drawing.Point(15, 78)
    $lblHint.Size      = New-Object System.Drawing.Size(440, 18)
    $f.Controls.Add($lblHint)

    $btnA = New-BtnAppliquer -X 120 -Y 125
    $btnA.Add_Click({
        $newVal = $txt.Text.Trim()
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") `
            -Message (Get-Text "modification.confirm_field" (Get-Text "modification.action_phone") $CurrentValue $newVal $UPN)
        if ($confirm) {
            # businessPhones doit être un tableau JSON
            $phoneArray = if ([string]::IsNullOrWhiteSpace($newVal)) { @() } else { @($newVal) }
            try {
                $body = @{ businessPhones = $phoneArray } | ConvertTo-Json
                Invoke-MgGraphRequest -Method PATCH `
                    -Uri "https://graph.microsoft.com/v1.0/users/$UserId" `
                    -Body $body -ContentType "application/json" -ErrorAction Stop
                Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.success_field" (Get-Text "modification.action_phone")) -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "MODIFY_PHONE" -UPN $UPN -Message "Poste fixe : '$CurrentValue' -> '$newVal'"
                $f.Close()
            }
            catch {
                $errMsg = Format-GraphErrorMessage -ErrorMessage $_.Exception.Message -Field "BusinessPhones"
                Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $errMsg -IsSuccess $false
            }
        }
    })
    $f.Controls.Add($btnA)

    $btnC = New-BtnAnnuler -X 240 -Y 105
    $f.Controls.Add($btnC)
    $f.CancelButton = $btnC

    $f.ShowDialog() | Out-Null
    $f.Dispose()
}

# ============================================================
#  SOUS-FORMULAIRES — MESSAGERIE & UPN
# ============================================================

function Show-ModifyUPN {
    <#
    .SYNOPSIS
        Modification de l'UPN avec liste des domaines disponibles.
        Propose d'ajouter l'ancien UPN comme alias Exchange (proxyAddress SMTP).
    #>
    param([string]$UserId, [string]$UPN)

    $f = New-SubForm -Titre (Get-Text "modification.action_upn") -Largeur 560 -Hauteur 320

    # Chargement des domaines vérifiés du tenant
    $lblLoad = New-Object System.Windows.Forms.Label
    $lblLoad.Text     = Get-Text "modification.loading_domains"
    $lblLoad.Location = New-Object System.Drawing.Point(15, 15)
    $lblLoad.Size     = New-Object System.Drawing.Size(510, 20)
    $f.Controls.Add($lblLoad)

    # Décomposition de l'UPN actuel
    $upnParts    = $UPN.Split("@")
    $localPart   = $upnParts[0]
    $currentDomain = if ($upnParts.Count -gt 1) { "@" + $upnParts[1] } else { "" }

    $lblCurrent = New-LabelCurrent -Texte (Get-Text "modification.current_value" $UPN) -Y 40
    $f.Controls.Add($lblCurrent)

    $lblLocal = New-Object System.Windows.Forms.Label
    $lblLocal.Text     = Get-Text "modification.field_upn_local"
    $lblLocal.Location = New-Object System.Drawing.Point(15, 75)
    $lblLocal.Size     = New-Object System.Drawing.Size(100, 20)
    $f.Controls.Add($lblLocal)

    $txtLocal = New-Object System.Windows.Forms.TextBox
    $txtLocal.Text     = $localPart
    $txtLocal.Location = New-Object System.Drawing.Point(120, 72)
    $txtLocal.Size     = New-Object System.Drawing.Size(180, 25)
    $f.Controls.Add($txtLocal)

    $lblAt = New-Object System.Windows.Forms.Label
    $lblAt.Text     = "@"
    $lblAt.Font     = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblAt.Location = New-Object System.Drawing.Point(305, 72)
    $lblAt.Size     = New-Object System.Drawing.Size(15, 25)
    $f.Controls.Add($lblAt)

    $cboDomain = New-Object System.Windows.Forms.ComboBox
    $cboDomain.Location      = New-Object System.Drawing.Point(323, 72)
    $cboDomain.Size          = New-Object System.Drawing.Size(210, 25)
    $cboDomain.DropDownStyle = "DropDownList"
    $f.Controls.Add($cboDomain)

    # Case à cocher : ajouter l'ancien UPN comme alias Exchange
    $chkAlias = New-Object System.Windows.Forms.CheckBox
    $chkAlias.Text     = Get-Text "modification.upn_keep_alias"
    $chkAlias.Location = New-Object System.Drawing.Point(15, 115)
    $chkAlias.Size     = New-Object System.Drawing.Size(510, 35)
    $chkAlias.Checked  = $true
    $chkAlias.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
    $f.Controls.Add($chkAlias)

    $lblAliasNote = New-Object System.Windows.Forms.Label
    $lblAliasNote.Text      = Get-Text "modification.upn_alias_note"
    $lblAliasNote.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblAliasNote.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $lblAliasNote.Location  = New-Object System.Drawing.Point(30, 150)
    $lblAliasNote.Size      = New-Object System.Drawing.Size(500, 30)
    $f.Controls.Add($lblAliasNote)

    $btnA = New-BtnAppliquer -X 140 -Y 200
    $btnA.Enabled = $false
    $f.Controls.Add($btnA)

    $btnC = New-BtnAnnuler -X 260 -Y 200
    $f.Controls.Add($btnC)
    $f.CancelButton = $btnC

    # Chargement asynchrone des domaines après affichage du formulaire
    $f.Add_Shown({
        try {
            $domains = Get-MgDomain -ErrorAction Stop | Where-Object { $_.IsVerified } | Select-Object -ExpandProperty Id | Sort-Object
            $cboDomain.Items.Clear()
            foreach ($d in $domains) { $cboDomain.Items.Add($d) | Out-Null }

            # Pré-sélectionner le domaine actuel
            $domainOnly = $currentDomain.TrimStart("@")
            $idx = $cboDomain.Items.IndexOf($domainOnly)
            if ($idx -ge 0) { $cboDomain.SelectedIndex = $idx }

            $lblLoad.Text    = Get-Text "modification.domains_loaded" $cboDomain.Items.Count
            $btnA.Enabled    = $true
        }
        catch {
            $lblLoad.Text    = Get-Text "modification.error_domains" $_.Exception.Message
        }
    })

    $btnA.Add_Click({
        $newLocal  = $txtLocal.Text.Trim()
        $newDomain = $cboDomain.SelectedItem
        if ([string]::IsNullOrWhiteSpace($newLocal) -or $null -eq $newDomain) { return }

        $newUPN = "$newLocal@$newDomain"
        if ($newUPN -eq $UPN) {
            Show-ResultDialog -Titre (Get-Text "modification.info_title") -Message (Get-Text "modification.info_no_change") -IsSuccess $true
            return
        }

        $msg = Get-Text "modification.confirm_upn" $UPN $newUPN
        if ($chkAlias.Checked) {
            $msg += "`n`n" + (Get-Text "modification.confirm_upn_alias" $UPN)
        }

        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") -Message $msg
        if ($confirm) {
            # 1. Mise à jour de l'UPN
            $result = Set-AzUser -UserId $UserId -Properties @{ UserPrincipalName = $newUPN }
            if (-not $result.Success) {
                Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $result.Error -IsSuccess $false
                return
            }
            Write-Log -Level "SUCCESS" -Action "MODIFY_UPN" -UPN $UPN -Message "UPN : '$UPN' -> '$newUPN'"

            # 2. Ajout de l'ancien UPN comme alias smtp (minuscule) si coché
            if ($chkAlias.Checked) {
                try {
                    $user = Get-MgUser -UserId $UserId -Property "proxyAddresses" -ErrorAction Stop
                    $newProxies = [System.Collections.Generic.List[string]]($user.ProxyAddresses)
                    $aliasSmtp  = "smtp:$UPN"
                    if ($aliasSmtp -notin $newProxies) {
                        $newProxies.Add($aliasSmtp)
                        $bodyObj2 = [ordered]@{ proxyAddresses = [string[]]$newProxies }
                        $body2 = $bodyObj2 | ConvertTo-Json -Depth 3 -Compress
                        Invoke-MgGraphRequest -Method PATCH `
                            -Uri "https://graph.microsoft.com/v1.0/users/$UserId" `
                            -Body $body2 -ContentType "application/json" -ErrorAction Stop
                        Write-Log -Level "SUCCESS" -Action "ADD_ALIAS" -UPN $newUPN -Message "Alias ajouté : $aliasSmtp"
                    }
                }
                catch {
                    Write-Log -Level "WARNING" -Action "ADD_ALIAS" -UPN $newUPN -Message "Impossible d'ajouter l'alias : $($_.Exception.Message)"
                }
            }

            Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.success_upn" $newUPN) -IsSuccess $true
            $f.Close()
        }
    })

    $f.ShowDialog() | Out-Null
    $f.Dispose()
}

function Show-ManageAliases {
    <#
    .SYNOPSIS
        Gestion des alias email (proxyAddresses) : visualiser, ajouter, supprimer.
        SMTP: = adresse principale (lecture seule ici), smtp: = alias secondaires.
    #>
    param([string]$UserId, [string]$UPN)

    $f = New-SubForm -Titre (Get-Text "modification.action_aliases") -Largeur 580 -Hauteur 440

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text     = Get-Text "modification.aliases_for" $UPN
    $lblTitle.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblTitle.Location = New-Object System.Drawing.Point(15, 10)
    $lblTitle.Size     = New-Object System.Drawing.Size(540, 20)
    $f.Controls.Add($lblTitle)

    # Liste des alias actuels
    $lstAliases = New-Object System.Windows.Forms.ListBox
    $lstAliases.Location  = New-Object System.Drawing.Point(15, 38)
    $lstAliases.Size      = New-Object System.Drawing.Size(540, 180)
    $lstAliases.Font      = New-Object System.Drawing.Font("Consolas", 9)
    $f.Controls.Add($lstAliases)

    $lblLegend = New-Object System.Windows.Forms.Label
    $lblLegend.Text      = Get-Text "modification.aliases_legend"
    $lblLegend.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblLegend.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $lblLegend.Location  = New-Object System.Drawing.Point(15, 222)
    $lblLegend.Size      = New-Object System.Drawing.Size(540, 18)
    $f.Controls.Add($lblLegend)

    # Zone d'ajout d'un alias
    $lblAdd = New-Object System.Windows.Forms.Label
    $lblAdd.Text     = Get-Text "modification.alias_add_label"
    $lblAdd.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblAdd.Location = New-Object System.Drawing.Point(15, 252)
    $lblAdd.Size     = New-Object System.Drawing.Size(540, 18)
    $f.Controls.Add($lblAdd)

    $txtNewAlias = New-Object System.Windows.Forms.TextBox
    $txtNewAlias.Location    = New-Object System.Drawing.Point(15, 275)
    $txtNewAlias.Size        = New-Object System.Drawing.Size(380, 25)
    $txtNewAlias.PlaceholderText = "alias@domaine.com"
    $f.Controls.Add($txtNewAlias)

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text      = Get-Text "modification.alias_btn_add"
    $btnAdd.Location  = New-Object System.Drawing.Point(400, 275)
    $btnAdd.Size      = New-Object System.Drawing.Size(155, 25)
    $btnAdd.BackColor = $script:COLOR_GREEN
    $btnAdd.ForeColor = $script:COLOR_WHITE
    $btnAdd.FlatStyle = "Flat"
    $f.Controls.Add($btnAdd)

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text      = Get-Text "modification.alias_btn_remove"
    $btnRemove.Location  = New-Object System.Drawing.Point(15, 315)
    $btnRemove.Size      = New-Object System.Drawing.Size(180, 30)
    $btnRemove.BackColor = $script:COLOR_RED
    $btnRemove.ForeColor = $script:COLOR_WHITE
    $btnRemove.FlatStyle = "Flat"
    $btnRemove.Enabled   = $false
    $f.Controls.Add($btnRemove)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text        = Get-Text "modification.btn_close"
    $btnClose.Location    = New-Object System.Drawing.Point(440, 365)
    $btnClose.Size        = New-Object System.Drawing.Size(120, 35)
    $btnClose.FlatStyle   = "Flat"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $f.Controls.Add($btnClose)

    # --- Chargement et rafraîchissement de la liste ---
    $script:ProxyAddresses = @()

    $Refresh = {
        try {
            $user = Get-MgUser -UserId $UserId -Property "proxyAddresses" -ErrorAction Stop
            $script:ProxyAddresses = @($user.ProxyAddresses)
            $lstAliases.Items.Clear()
            foreach ($addr in $script:ProxyAddresses | Sort-Object) {
                $lstAliases.Items.Add($addr) | Out-Null
            }
        }
        catch {
            Write-Log -Level "ERROR" -Action "GET_ALIASES" -UPN $UPN -Message $_.Exception.Message
        }
    }
    & $Refresh

    $lstAliases.Add_SelectedIndexChanged({
        # Désactiver la suppression sur l'adresse principale (SMTP: en majuscule)
        $sel = $lstAliases.SelectedItem
        $btnRemove.Enabled = ($null -ne $sel -and -not $sel.StartsWith("SMTP:"))
    })

    $btnAdd.Add_Click({
        $newAddr = $txtNewAlias.Text.Trim().ToLower()
        if (-not $newAddr.Contains("@")) {
            [System.Windows.Forms.MessageBox]::Show(
                (Get-Text "modification.alias_invalid"),
                (Get-Text "modification.error_title"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
        $aliasEntry = "smtp:$newAddr"
        if ($aliasEntry -in $script:ProxyAddresses) {
            [System.Windows.Forms.MessageBox]::Show(
                (Get-Text "modification.alias_exists"),
                (Get-Text "modification.info_title"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") `
            -Message (Get-Text "modification.alias_confirm_add" $newAddr $UPN)
        if ($confirm) {
            try {
                $newList = [System.Collections.Generic.List[string]]($script:ProxyAddresses)
                $newList.Add($aliasEntry)
                # proxyAddresses : PATCH direct via REST (Update-MgUser bloque sur boites Exchange Online)
                # Forcer un tableau JSON valide même pour un seul élément
                $proxyArray = [System.Collections.Generic.List[string]]$newList
                $bodyObj = [ordered]@{ proxyAddresses = [string[]]$proxyArray }
                $body = $bodyObj | ConvertTo-Json -Depth 3 -Compress
                Invoke-MgGraphRequest -Method PATCH `
                    -Uri "https://graph.microsoft.com/v1.0/users/$UserId" `
                    -Body $body -ContentType "application/json" -ErrorAction Stop
                Write-Log -Level "SUCCESS" -Action "ADD_ALIAS" -UPN $UPN -Message "Alias ajouté : $aliasEntry"
                $txtNewAlias.Text = ""
                & $Refresh
            }
            catch {
                $errMsg = Format-GraphErrorMessage -ErrorMessage $_.Exception.Message -Field "proxyAddresses"
                Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $errMsg -IsSuccess $false
            }
        }
    })

    $btnRemove.Add_Click({
        $selected = $lstAliases.SelectedItem
        if ($null -eq $selected -or $selected.StartsWith("SMTP:")) { return }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") `
            -Message (Get-Text "modification.alias_confirm_remove" $selected $UPN)
        if ($confirm) {
            try {
                $newList = [System.Collections.Generic.List[string]]($script:ProxyAddresses)
                $newList.Remove($selected) | Out-Null
                $proxyArray = [System.Collections.Generic.List[string]]$newList
                $bodyObj = [ordered]@{ proxyAddresses = [string[]]$proxyArray }
                $body = $bodyObj | ConvertTo-Json -Depth 3 -Compress
                Invoke-MgGraphRequest -Method PATCH `
                    -Uri "https://graph.microsoft.com/v1.0/users/$UserId" `
                    -Body $body -ContentType "application/json" -ErrorAction Stop
                Write-Log -Level "SUCCESS" -Action "REMOVE_ALIAS" -UPN $UPN -Message "Alias supprimé : $selected"
                & $Refresh
            }
            catch {
                $errMsg = Format-GraphErrorMessage -ErrorMessage $_.Exception.Message -Field "proxyAddresses"
                Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $errMsg -IsSuccess $false
            }
        }
    })

    $f.ShowDialog() | Out-Null
    $f.Dispose()
}

# ============================================================
#  SOUS-FORMULAIRES — LICENCES
# ============================================================

function Show-ManageLicenses {
    <#
    .SYNOPSIS
        Gestion des licences : deux colonnes (assignées / disponibles via config).
        Basé sur la clé license_groups du fichier de configuration client.
    #>
    param([string]$UserId, [string]$UPN)

    $f = New-SubForm -Titre (Get-Text "modification.action_licenses") -Largeur 680 -Hauteur 480

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text     = Get-Text "modification.licenses_for" $UPN
    $lblTitle.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblTitle.Location = New-Object System.Drawing.Point(15, 10)
    $lblTitle.Size     = New-Object System.Drawing.Size(640, 20)
    $f.Controls.Add($lblTitle)

    # --- Colonne gauche : licences assignées ---
    $lblAssigned = New-Object System.Windows.Forms.Label
    $lblAssigned.Text     = Get-Text "modification.licenses_assigned"
    $lblAssigned.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblAssigned.ForeColor = $script:COLOR_GREEN
    $lblAssigned.Location = New-Object System.Drawing.Point(15, 38)
    $lblAssigned.Size     = New-Object System.Drawing.Size(300, 20)
    $f.Controls.Add($lblAssigned)

    $lstAssigned = New-Object System.Windows.Forms.CheckedListBox
    $lstAssigned.Location     = New-Object System.Drawing.Point(15, 60)
    $lstAssigned.Size         = New-Object System.Drawing.Size(295, 280)
    $lstAssigned.CheckOnClick = $true
    $lstAssigned.Font         = New-Object System.Drawing.Font("Segoe UI", 9)
    $f.Controls.Add($lstAssigned)

    # --- Colonne droite : licences disponibles (config) non encore assignées ---
    $lblAvail = New-Object System.Windows.Forms.Label
    $lblAvail.Text      = Get-Text "modification.licenses_available"
    $lblAvail.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblAvail.ForeColor = $script:COLOR_BLUE
    $lblAvail.Location  = New-Object System.Drawing.Point(355, 38)
    $lblAvail.Size      = New-Object System.Drawing.Size(300, 20)
    $f.Controls.Add($lblAvail)

    $lstAvail = New-Object System.Windows.Forms.CheckedListBox
    $lstAvail.Location     = New-Object System.Drawing.Point(355, 60)
    $lstAvail.Size         = New-Object System.Drawing.Size(295, 280)
    $lstAvail.CheckOnClick = $true
    $lstAvail.Font         = New-Object System.Drawing.Font("Segoe UI", 9)
    $f.Controls.Add($lstAvail)

    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text      = Get-Text "modification.licenses_note"
    $lblNote.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblNote.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $lblNote.Location  = New-Object System.Drawing.Point(15, 347)
    $lblNote.Size      = New-Object System.Drawing.Size(640, 18)
    $f.Controls.Add($lblNote)

    # Boutons d'action
    $btnRevoke = New-Object System.Windows.Forms.Button
    $btnRevoke.Text      = Get-Text "modification.licenses_btn_revoke"
    $btnRevoke.Location  = New-Object System.Drawing.Point(15, 375)
    $btnRevoke.Size      = New-Object System.Drawing.Size(180, 35)
    $btnRevoke.BackColor = $script:COLOR_RED
    $btnRevoke.ForeColor = $script:COLOR_WHITE
    $btnRevoke.FlatStyle = "Flat"
    $btnRevoke.Enabled   = $false
    $f.Controls.Add($btnRevoke)

    $btnAssign = New-Object System.Windows.Forms.Button
    $btnAssign.Text      = Get-Text "modification.licenses_btn_assign"
    $btnAssign.Location  = New-Object System.Drawing.Point(355, 375)
    $btnAssign.Size      = New-Object System.Drawing.Size(180, 35)
    $btnAssign.BackColor = $script:COLOR_BLUE
    $btnAssign.ForeColor = $script:COLOR_WHITE
    $btnAssign.FlatStyle = "Flat"
    $btnAssign.Enabled   = $false
    $f.Controls.Add($btnAssign)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text        = Get-Text "modification.btn_close"
    $btnClose.Location    = New-Object System.Drawing.Point(545, 415)
    $btnClose.Size        = New-Object System.Drawing.Size(110, 30)
    $btnClose.FlatStyle   = "Flat"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $f.Controls.Add($btnClose)

    # --- Chargement ---
    $script:AssignedSkuIds = @{}   # groupName -> skuId (groupes déjà assignés)
    $configGroups = @()
    if ($Config.PSObject.Properties["license_groups"]) {
        $configGroups = @($Config.license_groups)
    }

    $f.Add_Shown({
        try {
            # Récupérer les groupes d'appartenance pour détecter les groupes de licence actifs
            $memberOf = Get-MgUserMemberOf -UserId $UserId -ErrorAction Stop
            $memberGroupNames = @($memberOf | ForEach-Object {
                if ($_.AdditionalProperties.ContainsKey("displayName")) { $_.AdditionalProperties.displayName }
            })

            $lstAssigned.Items.Clear()
            $lstAvail.Items.Clear()

            foreach ($grp in $configGroups) {
                if ($grp -in $memberGroupNames) {
                    $lstAssigned.Items.Add($grp, $false) | Out-Null
                }
                else {
                    $lstAvail.Items.Add($grp, $false) | Out-Null
                }
            }

            # Groupes assignés hors config (lecture seule, informatif)
            foreach ($n in $memberGroupNames) {
                if ($n -notin $configGroups -and $n -like "LIC-*") {
                    $lstAssigned.Items.Add("* $n", $false) | Out-Null
                }
            }
        }
        catch {
            Write-Log -Level "ERROR" -Action "LOAD_LICENSES" -UPN $UPN -Message $_.Exception.Message
        }
    })

    # Activer les boutons selon la sélection — on utilise ItemCheck avec NewValue
    # car CheckedItems.Count n'est pas encore mis à jour au moment de l'événement
    $script:licBtnRevoke = $btnRevoke
    $script:licBtnAssign = $btnAssign
    $script:licLstAssigned = $lstAssigned
    $script:licLstAvail = $lstAvail

    $lstAssigned.Add_ItemCheck({
        param($s, $e)
        # NewValue = Checked signifie qu'on coche => +1, Unchecked => -1
        $delta = if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) { 1 } else { -1 }
        $script:licBtnRevoke.Enabled = (($script:licLstAssigned.CheckedItems.Count + $delta) -gt 0)
    })

    $lstAvail.Add_ItemCheck({
        param($s, $e)
        $delta = if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) { 1 } else { -1 }
        $script:licBtnAssign.Enabled = (($script:licLstAvail.CheckedItems.Count + $delta) -gt 0)
    })

    $btnRevoke.Add_Click({
        $toRevoke = @($lstAssigned.CheckedItems | Where-Object { -not $_.StartsWith("*") })
        if ($toRevoke.Count -eq 0) { return }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") `
            -Message (Get-Text "modification.licenses_confirm_revoke" ($toRevoke -join ", ") $UPN)
        if ($confirm) {
            $errors = @()
            foreach ($grp in $toRevoke) {
                try {
                    $groupObj = Get-MgGroup -Filter "displayName eq '$grp'" -ErrorAction Stop
                    if ($groupObj) {
                        Remove-MgGroupMemberByRef -GroupId $groupObj.Id -DirectoryObjectId $UserId -ErrorAction Stop
                        Write-Log -Level "SUCCESS" -Action "REVOKE_LICENSE_GROUP" -UPN $UPN -Message "Retiré du groupe licence : $grp"
                    }
                }
                catch {
                    $errors += "$grp : $($_.Exception.Message)"
                    Write-Log -Level "ERROR" -Action "REVOKE_LICENSE_GROUP" -UPN $UPN -Message "Erreur retrait $grp : $($_.Exception.Message)"
                }
            }

            if ($errors.Count -eq 0) {
                Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.licenses_success_revoke") -IsSuccess $true
            }
            else {
                Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message ($errors -join "`n") -IsSuccess $false
            }

            # Rafraichir les listes directement
            $script:licLstAssigned.Items.Clear()
            $script:licLstAvail.Items.Clear()
            $script:licBtnRevoke.Enabled = $false
            $script:licBtnAssign.Enabled = $false
            try {
                $memberOf2 = Get-MgUserMemberOf -UserId $UserId -ErrorAction Stop
                $memberGroupNames2 = @($memberOf2 | ForEach-Object {
                    if ($_.AdditionalProperties.ContainsKey("displayName")) { $_.AdditionalProperties.displayName }
                })
                foreach ($grp in $configGroups) {
                    if ($grp -in $memberGroupNames2) {
                        $script:licLstAssigned.Items.Add($grp, $false) | Out-Null
                    }
                    else {
                        $script:licLstAvail.Items.Add($grp, $false) | Out-Null
                    }
                }
            }
            catch { }
        }
    })

    $btnAssign.Add_Click({
        $toAssign = @($lstAvail.CheckedItems)
        if ($toAssign.Count -eq 0) { return }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") `
            -Message (Get-Text "modification.licenses_confirm_assign" ($toAssign -join ", ") $UPN)
        if ($confirm) {
            $errors = @()
            foreach ($grp in $toAssign) {
                try {
                    $groupObj = Get-MgGroup -Filter "displayName eq '$grp'" -ErrorAction Stop
                    if ($groupObj) {
                        $bodyParam = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId" }
                        New-MgGroupMemberByRef -GroupId $groupObj.Id -BodyParameter $bodyParam -ErrorAction Stop
                        Write-Log -Level "SUCCESS" -Action "ASSIGN_LICENSE_GROUP" -UPN $UPN -Message "Ajouté au groupe licence : $grp"
                    }
                }
                catch {
                    if ($_.Exception.Message -notlike "*already exist*") {
                        $errors += "$grp : $($_.Exception.Message)"
                        Write-Log -Level "ERROR" -Action "ASSIGN_LICENSE_GROUP" -UPN $UPN -Message "Erreur ajout $grp : $($_.Exception.Message)"
                    }
                }
            }

            if ($errors.Count -eq 0) {
                Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.licenses_success_assign") -IsSuccess $true
            }
            else {
                Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message ($errors -join "`n") -IsSuccess $false
            }
        }
    })

    $f.ShowDialog() | Out-Null
    $f.Dispose()
}

# ============================================================
#  SOUS-FORMULAIRES — GROUPES
# ============================================================

function Show-ManageGroups {
    <#
    .SYNOPSIS
        Gestion des groupes : deux colonnes (assignés / disponibles dans le tenant).
        Permet de retirer ET d assigner des groupes.
    #>
    param([string]$UserId, [string]$UPN)

    $f = New-SubForm -Titre (Get-Text "modification.action_groups_manage") -Largeur 700 -Hauteur 580

    # Bandeau avertissement
    $pnlWarn = New-Object System.Windows.Forms.Panel
    $pnlWarn.Location    = New-Object System.Drawing.Point(10, 10)
    $pnlWarn.Size        = New-Object System.Drawing.Size(660, 50)
    $pnlWarn.BackColor   = [System.Drawing.Color]::FromArgb(255, 243, 205)
    $pnlWarn.BorderStyle = "FixedSingle"
    $f.Controls.Add($pnlWarn)

    $lblWarnIcon = New-Object System.Windows.Forms.Label
    $lblWarnIcon.Text      = "⚠"
    $lblWarnIcon.Font      = New-Object System.Drawing.Font("Segoe UI", 14)
    $lblWarnIcon.ForeColor = [System.Drawing.Color]::FromArgb(133, 77, 14)
    $lblWarnIcon.Location  = New-Object System.Drawing.Point(8, 5)
    $lblWarnIcon.Size      = New-Object System.Drawing.Size(30, 38)
    $pnlWarn.Controls.Add($lblWarnIcon)

    $lblWarn = New-Object System.Windows.Forms.Label
    $lblWarn.Text      = Get-Text "modification.groups_warning"
    $lblWarn.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblWarn.ForeColor = [System.Drawing.Color]::FromArgb(92, 45, 3)
    $lblWarn.Location  = New-Object System.Drawing.Point(44, 6)
    $lblWarn.Size      = New-Object System.Drawing.Size(608, 36)
    $pnlWarn.Controls.Add($lblWarn)

    # --- Colonne gauche : groupes assignés ---
    $lblAssigned = New-Object System.Windows.Forms.Label
    $lblAssigned.Text      = Get-Text "modification.groups_assigned_label"
    $lblAssigned.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblAssigned.ForeColor = [System.Drawing.Color]::FromArgb(40, 120, 40)
    $lblAssigned.Location  = New-Object System.Drawing.Point(10, 68)
    $lblAssigned.Size      = New-Object System.Drawing.Size(320, 20)
    $f.Controls.Add($lblAssigned)

    $lstAssigned = New-Object System.Windows.Forms.ListBox
    $lstAssigned.Location      = New-Object System.Drawing.Point(10, 90)
    $lstAssigned.Size          = New-Object System.Drawing.Size(320, 370)
    $lstAssigned.SelectionMode = "MultiExtended"
    $lstAssigned.Font          = New-Object System.Drawing.Font("Segoe UI", 9)
    $lstAssigned.Sorted        = $true
    $f.Controls.Add($lstAssigned)

    # --- Colonne droite : groupes disponibles (tenant) ---
    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text     = Get-Text "modification.groups_search_label"
    $lblSearch.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblSearch.Location = New-Object System.Drawing.Point(345, 68)
    $lblSearch.Size     = New-Object System.Drawing.Size(200, 20)
    $f.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location      = New-Object System.Drawing.Point(345, 90)
    $txtSearch.Size          = New-Object System.Drawing.Size(240, 25)
    $txtSearch.PlaceholderText = Get-Text "modification.groups_search_placeholder"
    $f.Controls.Add($txtSearch)

    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Text      = Get-Text "modification.btn_search"
    $btnSearch.Location  = New-Object System.Drawing.Point(590, 90)
    $btnSearch.Size      = New-Object System.Drawing.Size(80, 25)
    $btnSearch.FlatStyle = "Flat"
    $f.Controls.Add($btnSearch)

    $lstAvail = New-Object System.Windows.Forms.ListBox
    $lstAvail.Location      = New-Object System.Drawing.Point(345, 120)
    $lstAvail.Size          = New-Object System.Drawing.Size(325, 340)
    $lstAvail.SelectionMode = "MultiExtended"
    $lstAvail.Font          = New-Object System.Drawing.Font("Segoe UI", 9)
    $lstAvail.Sorted        = $true
    $f.Controls.Add($lstAvail)

    # Compteurs
    $lblCountAssigned = New-Object System.Windows.Forms.Label
    $lblCountAssigned.Text     = ""
    $lblCountAssigned.Font     = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblCountAssigned.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $lblCountAssigned.Location = New-Object System.Drawing.Point(10, 464)
    $lblCountAssigned.Size     = New-Object System.Drawing.Size(320, 16)
    $f.Controls.Add($lblCountAssigned)

    $lblCountAvail = New-Object System.Windows.Forms.Label
    $lblCountAvail.Text      = Get-Text "modification.groups_search_hint"
    $lblCountAvail.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblCountAvail.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $lblCountAvail.Location  = New-Object System.Drawing.Point(345, 464)
    $lblCountAvail.Size      = New-Object System.Drawing.Size(325, 16)
    $f.Controls.Add($lblCountAvail)

    # Boutons d action
    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text      = Get-Text "modification.groups_btn_remove"
    $btnRemove.Location  = New-Object System.Drawing.Point(10, 490)
    $btnRemove.Size      = New-Object System.Drawing.Size(220, 32)
    $btnRemove.BackColor = $script:COLOR_RED
    $btnRemove.ForeColor = $script:COLOR_WHITE
    $btnRemove.FlatStyle = "Flat"
    $btnRemove.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnRemove.Enabled   = $false
    $f.Controls.Add($btnRemove)

    $btnAssign = New-Object System.Windows.Forms.Button
    $btnAssign.Text      = Get-Text "modification.groups_btn_assign"
    $btnAssign.Location  = New-Object System.Drawing.Point(345, 490)
    $btnAssign.Size      = New-Object System.Drawing.Size(220, 32)
    $btnAssign.BackColor = $script:COLOR_BLUE
    $btnAssign.ForeColor = $script:COLOR_WHITE
    $btnAssign.FlatStyle = "Flat"
    $btnAssign.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnAssign.Enabled   = $false
    $f.Controls.Add($btnAssign)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text         = Get-Text "modification.btn_close"
    $btnClose.Location     = New-Object System.Drawing.Point(570, 490)
    $btnClose.Size         = New-Object System.Drawing.Size(100, 32)
    $btnClose.FlatStyle    = "Flat"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $f.Controls.Add($btnClose)

    # Map groupId <-> displayName
    $script:GrpUserMap = @{}     # nom -> id (groupes assignés)
    $script:GrpAvailMap = @{}    # nom -> id (groupes disponibles)

    # Chargement des groupes assignés
    $LoadAssigned = {
        $lstAssigned.Items.Clear()
        $script:GrpUserMap = @{}
        try {
            $memberOf = Get-MgUserMemberOf -UserId $UserId -ErrorAction Stop
            foreach ($m in $memberOf) {
                if ($m.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group') {
                    $name = $m.AdditionalProperties.displayName
                    $script:GrpUserMap[$name] = $m.Id
                    $lstAssigned.Items.Add($name) | Out-Null
                }
            }
            $lblCountAssigned.Text = Get-Text "modification.groups_count" $lstAssigned.Items.Count
        }
        catch {
            Write-Log -Level "ERROR" -Action "LOAD_GROUPS" -UPN $UPN -Message $_.Exception.Message
        }
    }
    & $LoadAssigned

    # Recherche de groupes dans le tenant
    $SearchGroups = {
        $terme = $txtSearch.Text.Trim()
        if ($terme.Length -lt 2) { return }
        $lstAvail.Items.Clear()
        $script:GrpAvailMap = @{}
        try {
            $groups = Get-MgGroup -Filter "startsWith(displayName,'$terme')" `
                -Top 50 -Property "id,displayName" `
                -ConsistencyLevel "eventual" -CountVariable c -ErrorAction Stop
            foreach ($g in $groups) {
                if ($g.DisplayName -notin $script:GrpUserMap.Keys) {
                    $script:GrpAvailMap[$g.DisplayName] = $g.Id
                    $lstAvail.Items.Add($g.DisplayName) | Out-Null
                }
            }
            $lblCountAvail.Text = Get-Text "modification.groups_count" $lstAvail.Items.Count
        }
        catch {
            $lblCountAvail.Text = "Erreur : $($_.Exception.Message)"
        }
    }

    $btnSearch.Add_Click({ & $SearchGroups })
    $txtSearch.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            & $SearchGroups
            $_.SuppressKeyPress = $true
        }
    })

    $lstAssigned.Add_SelectedIndexChanged({
        $btnRemove.Enabled = $lstAssigned.SelectedItems.Count -gt 0
    })

    $lstAvail.Add_SelectedIndexChanged({
        $btnAssign.Enabled = $lstAvail.SelectedItems.Count -gt 0
    })

    $btnRemove.Add_Click({
        $selected = @($lstAssigned.SelectedItems)
        if ($selected.Count -eq 0) { return }

        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") `
            -Message (Get-Text "modification.groups_confirm_remove" ($selected -join ", ") $UPN)

        if ($confirm) {
            $errors  = @()
            $removed = 0
            foreach ($groupName in $selected) {
                $groupId = $script:GrpUserMap[$groupName]
                if (-not $groupId) { continue }
                try {
                    Remove-MgGroupMemberByRef -GroupId $groupId -DirectoryObjectId $UserId -ErrorAction Stop
                    $removed++
                    Write-Log -Level "SUCCESS" -Action "REMOVE_FROM_GROUP" -UPN $UPN -Message "Retiré du groupe : $groupName"
                }
                catch {
                    $errors += "$groupName : $($_.Exception.Message)"
                    Write-Log -Level "ERROR" -Action "REMOVE_FROM_GROUP" -UPN $UPN -Message "Erreur : $($_.Exception.Message)"
                }
            }
            if ($errors.Count -eq 0) {
                Show-ResultDialog -Titre (Get-Text "modification.success_title") `
                    -Message (Get-Text "modification.groups_success_remove" $removed) -IsSuccess $true
            }
            else {
                Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message ($errors -join "`n") -IsSuccess $false
            }
            & $LoadAssigned
            if ($txtSearch.Text.Length -ge 2) { & $SearchGroups }
        }
    })

    $btnAssign.Add_Click({
        $selected = @($lstAvail.SelectedItems)
        if ($selected.Count -eq 0) { return }

        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") `
            -Message (Get-Text "modification.groups_confirm_assign" ($selected -join ", ") $UPN)

        if ($confirm) {
            $errors   = @()
            $assigned = 0
            foreach ($groupName in $selected) {
                $groupId = $script:GrpAvailMap[$groupName]
                if (-not $groupId) { continue }
                try {
                    $bodyParam = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId" }
                    New-MgGroupMemberByRef -GroupId $groupId -BodyParameter $bodyParam -ErrorAction Stop
                    $assigned++
                    Write-Log -Level "SUCCESS" -Action "ADD_TO_GROUP" -UPN $UPN -Message "Ajouté au groupe : $groupName"
                }
                catch {
                    if ($_.Exception.Message -notlike "*already exist*") {
                        $errors += "$groupName : $($_.Exception.Message)"
                        Write-Log -Level "ERROR" -Action "ADD_TO_GROUP" -UPN $UPN -Message "Erreur : $($_.Exception.Message)"
                    }
                    else { $assigned++ }
                }
            }
            if ($errors.Count -eq 0) {
                Show-ResultDialog -Titre (Get-Text "modification.success_title") `
                    -Message (Get-Text "modification.groups_success_assign" $assigned) -IsSuccess $true
            }
            else {
                Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message ($errors -join "`n") -IsSuccess $false
            }
            & $LoadAssigned
            if ($txtSearch.Text.Length -ge 2) { & $SearchGroups }
        }
    })

    $f.ShowDialog() | Out-Null
    $f.Dispose()
}


# ============================================================
#  SOUS-FORMULAIRES — SÉCURITÉ
# ============================================================

function Show-ResetPassword {
    <#
    .SYNOPSIS
        Réinitialisation du mot de passe avec génération d'un mot de passe sécurisé.
    #>
    param([string]$UserId, [string]$UPN)

    $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.password_title") `
        -Message (Get-Text "modification.password_confirm" $UPN)
    if (-not $confirm) { return }

    $newPasswordPlain  = New-SecurePassword
    # Conversion en SecureString — Reset-AzUserPassword n'accepte plus le plaintext
    $newPassword = ConvertTo-SecureString -String $newPasswordPlain -AsPlainText -Force
    $result = Reset-AzUserPassword -UserId $UserId -NewPassword $newPassword `
        -ForceChange $Config.password_policy.force_change_at_login

    if ($result.Success) {
        Write-Log -Level "SUCCESS" -Action "RESET_PASSWORD" -UPN $UPN -Message "Mot de passe réinitialisé."
        Show-PasswordDialog -UPN $UPN -Password $newPasswordPlain
    }
    else {
        Show-ResultDialog -Titre (Get-Text "modification.error_title") `
            -Message (Get-Text "modification.password_error" $result.Error) -IsSuccess $false
    }
}

function Invoke-RevokeSession {
    <#
    .SYNOPSIS
        Révoque toutes les sessions actives de l'utilisateur (Revoke-MgUserSignInSession).
    #>
    param([string]$UserId, [string]$UPN)

    $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.revoke_title") `
        -Message (Get-Text "modification.revoke_confirm" $UPN)
    if (-not $confirm) { return }

    try {
        Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/users/$UserId/revokeSignInSessions" `
            -ErrorAction Stop | Out-Null
        Write-Log -Level "SUCCESS" -Action "REVOKE_SESSIONS" -UPN $UPN -Message "Sessions révoquées."
        Show-ResultDialog -Titre (Get-Text "modification.success_title") `
            -Message (Get-Text "modification.revoke_success") -IsSuccess $true
    }
    catch {
        Write-Log -Level "ERROR" -Action "REVOKE_SESSIONS" -UPN $UPN -Message $_.Exception.Message
        Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $_.Exception.Message -IsSuccess $false
    }
}

function Invoke-MfaReset {
    <#
    .SYNOPSIS
        Force la ré-inscription MFA en supprimant les méthodes d'authentification fortes.
        Nécessite la permission UserAuthenticationMethod.ReadWrite.All.
    #>
    param([string]$UserId, [string]$UPN)

    $f = New-SubForm -Titre (Get-Text "modification.mfa_title") -Largeur 540 -Hauteur 340

    $pnlWarn = New-Object System.Windows.Forms.Panel
    $pnlWarn.Location    = New-Object System.Drawing.Point(10, 10)
    $pnlWarn.Size        = New-Object System.Drawing.Size(500, 60)
    $pnlWarn.BackColor   = [System.Drawing.Color]::FromArgb(255, 235, 235)
    $pnlWarn.BorderStyle = "FixedSingle"
    $f.Controls.Add($pnlWarn)

    $lblWarn = New-Object System.Windows.Forms.Label
    $lblWarn.Text      = Get-Text "modification.mfa_warning"
    $lblWarn.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblWarn.ForeColor = [System.Drawing.Color]::FromArgb(150, 0, 0)
    $lblWarn.Location  = New-Object System.Drawing.Point(10, 8)
    $lblWarn.Size      = New-Object System.Drawing.Size(478, 44)
    $pnlWarn.Controls.Add($lblWarn)

    $lblMethods = New-Object System.Windows.Forms.Label
    $lblMethods.Text     = Get-Text "modification.mfa_methods_label"
    $lblMethods.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblMethods.Location = New-Object System.Drawing.Point(10, 82)
    $lblMethods.Size     = New-Object System.Drawing.Size(500, 20)
    $f.Controls.Add($lblMethods)

    $lstMethods = New-Object System.Windows.Forms.ListBox
    $lstMethods.Location = New-Object System.Drawing.Point(10, 106)
    $lstMethods.Size     = New-Object System.Drawing.Size(500, 130)
    $lstMethods.Font     = New-Object System.Drawing.Font("Consolas", 9)
    $f.Controls.Add($lstMethods)

    # Chargement des méthodes MFA existantes
    $script:MfaMethods = @()
    try {
        $methods = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/users/$UserId/authentication/methods" `
            -ErrorAction Stop
        if ($methods.value) {
            $script:MfaMethods = @($methods.value)
            foreach ($m in $script:MfaMethods) {
                $type = $m.'@odata.type'.Split(".")[-1]
                $lstMethods.Items.Add("$type — $($m.id)") | Out-Null
            }
        }
        else {
            $noMethodsText = Get-Text "modification.mfa_no_methods"
            $lstMethods.Items.Add($noMethodsText) | Out-Null
        }
    }
    catch {
        $errText = Get-Text "modification.mfa_load_error" $_.Exception.Message
        $lstMethods.Items.Add($errText) | Out-Null
    }

    $btnReset = New-Object System.Windows.Forms.Button
    $btnReset.Text      = Get-Text "modification.mfa_btn_reset"
    $btnReset.Location  = New-Object System.Drawing.Point(10, 255)
    $btnReset.Size      = New-Object System.Drawing.Size(220, 35)
    $btnReset.BackColor = $script:COLOR_ORANGE
    $btnReset.ForeColor = $script:COLOR_WHITE
    $btnReset.FlatStyle = "Flat"
    $btnReset.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($btnReset)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text        = Get-Text "modification.btn_close"
    $btnClose.Location    = New-Object System.Drawing.Point(390, 255)
    $btnClose.Size        = New-Object System.Drawing.Size(120, 35)
    $btnClose.FlatStyle   = "Flat"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $f.Controls.Add($btnClose)

    $btnReset.Add_Click({
        if ($script:MfaMethods.Count -eq 0) { return }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") `
            -Message (Get-Text "modification.mfa_confirm_reset" $UPN $script:MfaMethods.Count)
        if ($confirm) {
            $errors  = @()
            $deleted = 0
            foreach ($m in $script:MfaMethods) {
                $type = $m.'@odata.type'.Split(".")[-1]
                # passwordAuthenticationMethod ne peut pas être supprimé
                if ($type -eq "passwordAuthenticationMethod") { continue }
                try {
                    Invoke-MgGraphRequest -Method DELETE `
                        -Uri "https://graph.microsoft.com/v1.0/users/$UserId/authentication/$($type)s/$($m.id)" `
                        -ErrorAction Stop | Out-Null
                    $deleted++
                    Write-Log -Level "SUCCESS" -Action "MFA_RESET" -UPN $UPN -Message "Méthode MFA supprimée : $type ($($m.id))"
                }
                catch {
                    $errors += "$type : $($_.Exception.Message)"
                    Write-Log -Level "WARNING" -Action "MFA_RESET" -UPN $UPN -Message "Impossible de supprimer $type : $($_.Exception.Message)"
                }
            }
            if ($errors.Count -eq 0) {
                Show-ResultDialog -Titre (Get-Text "modification.success_title") `
                    -Message (Get-Text "modification.mfa_success" $deleted) -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "MFA_RESET" -UPN $UPN -Message "$deleted méthode(s) MFA supprimée(s)."
                $f.Close()
            }
            else {
                Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message ($errors -join "`n") -IsSuccess $false
            }
        }
    })

    $f.ShowDialog() | Out-Null
    $f.Dispose()
}

function Show-ToggleAccount {
    <#
    .SYNOPSIS
        Activation / désactivation du compte utilisateur.
    #>
    param([string]$UserId, [string]$UPN, [bool]$IsEnabled)

    $action  = if ($IsEnabled) { Get-Text "modification.toggle_disable" } else { Get-Text "modification.toggle_enable" }
    $confirm = Show-ConfirmDialog -Titre $action `
        -Message (Get-Text "modification.toggle_confirm" $action $UPN)

    if (-not $confirm) { return }

    if ($IsEnabled) {
        $result = Disable-AzUser -UserId $UserId
    }
    else {
        $result = Enable-AzUser -UserId $UserId
    }

    if ($result.Success) {
        Show-ResultDialog -Titre (Get-Text "modification.success_title") `
            -Message (Get-Text "modification.toggle_success" $action) -IsSuccess $true
        Write-Log -Level "SUCCESS" -Action "TOGGLE_ACCOUNT" -UPN $UPN -Message "Compte $action"
    }
    else {
        Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $result.Error -IsSuccess $false
    }
}

# ============================================================
#  SOUS-FORMULAIRES — AUDIT
# ============================================================

function Show-SignInLogs {
    <#
    .SYNOPSIS
        Affiche les 20 dernières connexions de l'utilisateur (signInActivity).
        Nécessite AuditLog.Read.All ou Reports.Read.All.
    #>
    param([string]$UserId, [string]$UPN)

    $f = New-SubForm -Titre (Get-Text "modification.action_signin_logs") -Largeur 780 -Hauteur 500

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text     = Get-Text "modification.signin_logs_for" $UPN
    $lblTitle.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblTitle.Location = New-Object System.Drawing.Point(10, 10)
    $lblTitle.Size     = New-Object System.Drawing.Size(740, 20)
    $f.Controls.Add($lblTitle)

    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Location              = New-Object System.Drawing.Point(10, 38)
    $dgv.Size                  = New-Object System.Drawing.Size(745, 380)
    $dgv.ReadOnly              = $true
    $dgv.AllowUserToAddRows    = $false
    $dgv.AutoSizeColumnsMode   = "Fill"
    $dgv.SelectionMode         = "FullRowSelect"
    $dgv.Font                  = New-Object System.Drawing.Font("Segoe UI", 8)
    $dgv.RowHeadersVisible     = $false
    $dgv.BackgroundColor       = $script:COLOR_WHITE
    $f.Controls.Add($dgv)

    # Colonnes
    $cols = @(
        @{ Name = "Date";         Header = Get-Text "modification.signin_col_date";    Width = 160 },
        @{ Name = "App";          Header = Get-Text "modification.signin_col_app";     Width = 180 },
        @{ Name = "IP";           Header = Get-Text "modification.signin_col_ip";      Width = 120 },
        @{ Name = "Location";     Header = Get-Text "modification.signin_col_location"; Width = 140 },
        @{ Name = "Status";       Header = Get-Text "modification.signin_col_status";  Width = 80 },
        @{ Name = "Conditional";  Header = Get-Text "modification.signin_col_ca";      Width = 60 }
    )
    foreach ($col in $cols) {
        $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c.Name       = $col.Name
        $c.HeaderText = $col.Header
        $dgv.Columns.Add($c) | Out-Null
    }

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text     = Get-Text "modification.signin_loading"
    $lblStatus.Font     = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $lblStatus.Location = New-Object System.Drawing.Point(10, 423)
    $lblStatus.Size     = New-Object System.Drawing.Size(640, 18)
    $f.Controls.Add($lblStatus)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text        = Get-Text "modification.btn_close"
    $btnClose.Location    = New-Object System.Drawing.Point(640, 420)
    $btnClose.Size        = New-Object System.Drawing.Size(115, 30)
    $btnClose.FlatStyle   = "Flat"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $f.Controls.Add($btnClose)

    $f.Add_Shown({
        try {
            # L'API signIns nécessite AuditLog.Read.All — top 20, filtré sur l'UPN
            $response = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=userPrincipalName eq '$UPN'&`$top=20&`$orderby=createdDateTime desc" `
                -ErrorAction Stop

            if ($response.value -and $response.value.Count -gt 0) {
                foreach ($log in $response.value) {
                    $date     = if ($log.createdDateTime) { [datetime]$log.createdDateTime | Get-Date -Format "yyyy-MM-dd HH:mm" } else { "-" }
                    $app      = if ($log.appDisplayName) { $log.appDisplayName } else { "-" }
                    $ip       = if ($log.ipAddress)     { $log.ipAddress }      else { "-" }
                    $country  = if ($log.location -and $log.location.city) { "$($log.location.city), $($log.location.countryOrRegion)" } else { "-" }
                    $status   = if ($log.status -and $log.status.errorCode -eq 0) { "✓" } else { "✗" }
                    $ca       = if ($log.appliedConditionalAccessPolicies -and $log.appliedConditionalAccessPolicies.Count -gt 0) { "Oui" } else { "Non" }

                    $row = $dgv.Rows.Add($date, $app, $ip, $country, $status, $ca)
                    if ($status -eq "✗") {
                        $dgv.Rows[$row].DefaultCellStyle.ForeColor = $script:COLOR_RED
                    }
                }
                $lblStatus.Text = Get-Text "modification.signin_loaded" $response.value.Count
            }
            else {
                $lblStatus.Text = Get-Text "modification.signin_no_data"
            }
        }
        catch {
            $lblStatus.Text = Get-Text "modification.signin_error" $_.Exception.Message
            Write-Log -Level "ERROR" -Action "SIGNIN_LOGS" -UPN $UPN -Message $_.Exception.Message
        }
    })

    $f.ShowDialog() | Out-Null
    $f.Dispose()
}