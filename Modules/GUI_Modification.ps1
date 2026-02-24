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
                           Set-AzUserManager, Reset-AzUserPassword,
                           Disable-AzUser, Enable-AzUser,
                           Get-AzUserManager, Get-AzDistinctValues)
    - Core/Lang.ps1       (Get-Text)
    - Variable globale    $Config

.AUTEUR
    [Equipe IT — M365 Monster]
#>

# ============================================================
#  CONSTANTES
# ============================================================
$script:COLOR_BLUE    = [System.Drawing.Color]::FromArgb(0, 123, 255)
$script:COLOR_ORANGE  = [System.Drawing.Color]::FromArgb(255, 140, 0)
$script:COLOR_RED     = [System.Drawing.Color]::FromArgb(220, 53, 69)
$script:COLOR_GREEN   = [System.Drawing.Color]::FromArgb(40, 167, 69)
$script:COLOR_BG      = [System.Drawing.Color]::WhiteSmoke
$script:COLOR_SECTION = [System.Drawing.Color]::FromArgb(240, 242, 245)
$script:COLOR_WHITE   = [System.Drawing.Color]::White
$script:COLOR_GRAY    = [System.Drawing.Color]::FromArgb(100, 100, 100)
$script:COLOR_HINT    = [System.Drawing.Color]::FromArgb(120, 120, 120)

# Codes ISO 3166-1 alpha-2
$script:ISO_COUNTRIES = @(
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

# Propriétés étendues pour la fiche utilisateur
$script:USER_PROPS = @(
    "id","displayName","givenName","surname","userPrincipalName","accountEnabled",
    "department","jobTitle","mobilePhone","businessPhones","mail","employeeType",
    "assignedLicenses","country","streetAddress","city","postalCode","state",
    "officeLocation","proxyAddresses","usageLocation"
) -join ","

# ============================================================
#  HELPERS GUI — Création de contrôles standardisés
# ============================================================

function New-SubForm {
    param([string]$Titre, [int]$Largeur = 480, [int]$Hauteur = 300)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Titre; $f.Size = New-Object System.Drawing.Size($Largeur, $Hauteur)
    $f.StartPosition = "CenterScreen"; $f.FormBorderStyle = "FixedDialog"
    $f.MaximizeBox = $false; $f.MinimizeBox = $false; $f.BackColor = $script:COLOR_BG
    return $f
}

function New-BtnAppliquer {
    param([int]$X, [int]$Y)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = Get-Text "modification.btn_apply"; $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size(110, 35); $b.BackColor = $script:COLOR_BLUE
    $b.ForeColor = $script:COLOR_WHITE; $b.FlatStyle = "Flat"
    $b.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    return $b
}

function New-BtnAnnuler {
    param([int]$X, [int]$Y)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = Get-Text "modification.btn_cancel"; $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size(110, 35); $b.FlatStyle = "Flat"
    $b.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    return $b
}

function New-ActionButton {
    <#  Bouton d'action coloré (assign, revoke, reset…).  #>
    param([string]$Texte, [int]$X, [int]$Y, [int]$W = 180, [int]$H = 35,
          [System.Drawing.Color]$Color = $script:COLOR_BLUE, [bool]$Enabled = $true)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Texte; $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size($W, $H); $b.BackColor = $Color
    $b.ForeColor = $script:COLOR_WHITE; $b.FlatStyle = "Flat"; $b.Enabled = $Enabled
    $b.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    return $b
}

function New-CloseButton {
    param([int]$X, [int]$Y, [int]$W = 120, [int]$H = 35)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = Get-Text "modification.btn_close"; $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size($W, $H); $b.FlatStyle = "Flat"
    $b.DialogResult = [System.Windows.Forms.DialogResult]::OK
    return $b
}

function New-LabelCurrent {
    param([string]$Texte, [int]$X = 15, [int]$Y = 15, [int]$W = 430)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Texte; $l.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $l.ForeColor = $script:COLOR_GRAY; $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size = New-Object System.Drawing.Size($W, 20)
    return $l
}

function New-HintLabel {
    param([string]$Texte, [int]$X = 15, [int]$Y = 78, [int]$W = 440)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Texte; $l.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $l.ForeColor = $script:COLOR_HINT; $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size = New-Object System.Drawing.Size($W, 18)
    return $l
}

function New-SectionLabel {
    param([string]$Texte, [int]$Y)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = "  $Texte"; $l.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $l.ForeColor = $script:COLOR_GRAY; $l.BackColor = $script:COLOR_SECTION
    $l.Location = New-Object System.Drawing.Point(0, $Y); $l.Size = New-Object System.Drawing.Size(200, 22)
    $l.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    return $l
}

function New-MenuButton {
    param([string]$Texte, [string]$Tag, [int]$Y)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = "  $Texte"; $b.Tag = $Tag; $b.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $b.Location = New-Object System.Drawing.Point(0, $Y); $b.Size = New-Object System.Drawing.Size(200, 30)
    $b.FlatStyle = "Flat"; $b.BackColor = $script:COLOR_WHITE
    $b.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $b.Padding = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
    return $b
}

function New-FormField {
    <#  Crée Label + TextBox, retourne le TextBox.  #>
    param([System.Windows.Forms.Control]$Parent, [string]$Label, [string]$Value = "",
          [int]$Y, [string]$Name = "", [int]$LW = 130, [int]$FW = 300)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Label; $lbl.Location = New-Object System.Drawing.Point(15, ($Y + 3))
    $lbl.Size = New-Object System.Drawing.Size($LW, 20)
    $Parent.Controls.Add($lbl)
    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Text = $Value; $txt.Name = $Name
    $txt.Location = New-Object System.Drawing.Point(($LW + 20), $Y)
    $txt.Size = New-Object System.Drawing.Size($FW, 25)
    $Parent.Controls.Add($txt)
    return $txt
}

function New-WarningBanner {
    <#  Bandeau d'avertissement (icône ⚠ + message).  #>
    param([System.Windows.Forms.Control]$Parent, [string]$Message,
          [int]$X = 10, [int]$Y = 10, [int]$W = 500, [int]$H = 50,
          [System.Drawing.Color]$Bg  = [System.Drawing.Color]::FromArgb(255, 243, 205),
          [System.Drawing.Color]$Fg  = [System.Drawing.Color]::FromArgb(92, 45, 3),
          [System.Drawing.Color]$Ico = [System.Drawing.Color]::FromArgb(133, 77, 14))
    $pnl = New-Object System.Windows.Forms.Panel
    $pnl.Location = New-Object System.Drawing.Point($X, $Y)
    $pnl.Size = New-Object System.Drawing.Size($W, $H); $pnl.BackColor = $Bg; $pnl.BorderStyle = "FixedSingle"
    $li = New-Object System.Windows.Forms.Label
    $li.Text = [char]0x26A0; $li.Font = New-Object System.Drawing.Font("Segoe UI", 14)
    $li.ForeColor = $Ico; $li.Location = New-Object System.Drawing.Point(8, 5)
    $li.Size = New-Object System.Drawing.Size(30, 38); $pnl.Controls.Add($li)
    $lt = New-Object System.Windows.Forms.Label
    $lt.Text = $Message; $lt.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $lt.ForeColor = $Fg; $lt.Location = New-Object System.Drawing.Point(44, 6)
    $lt.Size = New-Object System.Drawing.Size(($W - 52), ($H - 14)); $pnl.Controls.Add($lt)
    $Parent.Controls.Add($pnl)
    return $pnl
}

# ============================================================
#  HELPERS MÉTIER
# ============================================================

function Format-GraphErrorMessage {
    param([string]$ErrorMessage, [string]$Field = "")
    if ($ErrorMessage -like "*Forbidden*" -or $ErrorMessage -like "*403*") {
        if ($Field -in @("MobilePhone","BusinessPhones")) { return (Get-Text "modification.error_phone_forbidden") }
        return (Get-Text "modification.error_forbidden_reconnect")
    }
    if ($ErrorMessage -like "*Authorization_RequestDenied*" -or $ErrorMessage -like "*Insufficient privileges*") {
        if ($Field -in @("MobilePhone","BusinessPhones")) { return (Get-Text "modification.error_phone_permission_hint") }
        return (Get-Text "modification.error_permission_hint")
    }
    if ($ErrorMessage -like "*BadRequest*" -or $ErrorMessage -like "*Bad Request*") {
        if ($Field -in @("proxyAddresses","ProxyAddresses")) { return (Get-Text "modification.error_proxy_badrequest") }
        return "$ErrorMessage"
    }
    if ($ErrorMessage -like "*read-only*") { return (Get-Text "modification.error_proxy_readonly") }
    return $ErrorMessage
}

function Invoke-ProxyAddressChange {
    <#
    .SYNOPSIS  Ajoute ou supprime une proxyAddress avec fallback EXO → Graph.
    .OUTPUTS   [bool] — $true si l'opération a réussi.
    #>
    param([string]$UserId, [string]$UPN,
          [ValidateSet("Add","Remove")][string]$Action,
          [string]$Address, [string[]]$CurrentList)

    # Tentative 1 : Set-Mailbox via Exchange Online
    if (Get-Command Set-Mailbox -ErrorAction SilentlyContinue) {
        try {
            Set-Mailbox -Identity $UPN -EmailAddresses @{ $Action = $Address } -ErrorAction Stop
            Write-Log -Level "SUCCESS" -Action "${Action}_ALIAS" -UPN $UPN -Message "Alias via EXO : $Address"
            return $true
        }
        catch {
            Write-Log -Level "WARNING" -Action "${Action}_ALIAS" -UPN $UPN `
                -Message "Échec EXO : $($_.Exception.Message) — tentative Graph"
        }
    }
    # Tentative 2 : PATCH Graph
    try {
        $newList = [System.Collections.Generic.List[string]]($CurrentList)
        if ($Action -eq "Add") { $newList.Add($Address) } else { $newList.Remove($Address) | Out-Null }
        $body = [ordered]@{ proxyAddresses = [string[]]$newList } | ConvertTo-Json -Depth 3 -Compress
        Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$UserId" `
            -Body $body -ContentType "application/json" -ErrorAction Stop
        Write-Log -Level "SUCCESS" -Action "${Action}_ALIAS" -UPN $UPN -Message "Alias via Graph : $Address"
        return $true
    }
    catch {
        $errMsg = Format-GraphErrorMessage -ErrorMessage $_.Exception.Message -Field "proxyAddresses"
        Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $errMsg -IsSuccess $false
        return $false
    }
}

# ============================================================
#  FORMULAIRE PRINCIPAL
# ============================================================

function Show-ModificationForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = Get-Text "modification.title" $Config.client_name
    $form.Size = New-Object System.Drawing.Size(860, 780); $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "Sizable"; $form.MinimumSize = New-Object System.Drawing.Size(760, 600)
    $form.MaximizeBox = $true; $form.MinimizeBox = $false; $form.BackColor = $script:COLOR_BG

    $lblSection = New-Object System.Windows.Forms.Label
    $lblSection.Text = Get-Text "modification.search_label"
    $lblSection.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblSection.Location = New-Object System.Drawing.Point(15, 15); $lblSection.Size = New-Object System.Drawing.Size(770, 25)
    $form.Controls.Add($lblSection)

    $txtRecherche = New-Object System.Windows.Forms.TextBox
    $txtRecherche.Location = New-Object System.Drawing.Point(15, 50)
    $txtRecherche.Size = New-Object System.Drawing.Size(570, 25); $txtRecherche.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($txtRecherche)

    $btnRecherche = New-Object System.Windows.Forms.Button
    $btnRecherche.Text = Get-Text "modification.btn_search"
    $btnRecherche.Location = New-Object System.Drawing.Point(595, 50); $btnRecherche.Size = New-Object System.Drawing.Size(90, 25)
    $btnRecherche.FlatStyle = "Flat"; $form.Controls.Add($btnRecherche)

    $lstResultats = New-Object System.Windows.Forms.ListBox
    $lstResultats.Location = New-Object System.Drawing.Point(15, 80); $lstResultats.Size = New-Object System.Drawing.Size(770, 75)
    $lstResultats.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lstResultats.Visible = $false
    $form.Controls.Add($lstResultats)

    # Bandeau utilisateur sélectionné
    $pnlUser = New-Object System.Windows.Forms.Panel
    $pnlUser.Location = New-Object System.Drawing.Point(15, 80); $pnlUser.Size = New-Object System.Drawing.Size(770, 55)
    $pnlUser.BackColor = [System.Drawing.Color]::FromArgb(230, 240, 255)
    $pnlUser.Visible = $false; $pnlUser.BorderStyle = "FixedSingle"; $form.Controls.Add($pnlUser)

    $lblUserInfo = New-Object System.Windows.Forms.Label
    $lblUserInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblUserInfo.ForeColor = [System.Drawing.Color]::FromArgb(0, 70, 150)
    $lblUserInfo.Location = New-Object System.Drawing.Point(8, 5); $lblUserInfo.Size = New-Object System.Drawing.Size(650, 20)
    $pnlUser.Controls.Add($lblUserInfo)

    $lblUserInfo2 = New-Object System.Windows.Forms.Label
    $lblUserInfo2.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblUserInfo2.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $lblUserInfo2.Location = New-Object System.Drawing.Point(8, 28); $lblUserInfo2.Size = New-Object System.Drawing.Size(650, 18)
    $pnlUser.Controls.Add($lblUserInfo2)

    $btnChanger = New-Object System.Windows.Forms.Button
    $btnChanger.Text = Get-Text "modification.btn_change_user"
    $btnChanger.Location = New-Object System.Drawing.Point(668, 13); $btnChanger.Size = New-Object System.Drawing.Size(90, 28)
    $btnChanger.FlatStyle = "Flat"; $btnChanger.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $pnlUser.Controls.Add($btnChanger)

    # Zone principale : menu gauche + zone droite
    $pnlMain = New-Object System.Windows.Forms.Panel
    $pnlMain.Location = New-Object System.Drawing.Point(15, 145); $pnlMain.Size = New-Object System.Drawing.Size(810, 570)
    $pnlMain.Visible = $false
    $pnlMain.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
    $form.Controls.Add($pnlMain)

    $pnlMenu = New-Object System.Windows.Forms.Panel
    $pnlMenu.Location = New-Object System.Drawing.Point(0, 0); $pnlMenu.Size = New-Object System.Drawing.Size(210, 570)
    $pnlMenu.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $pnlMenu.BorderStyle = "FixedSingle"; $pnlMenu.AutoScroll = $true
    $pnlMenu.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left)
    $pnlMain.Controls.Add($pnlMenu)

    $pnlRight = New-Object System.Windows.Forms.Panel
    $pnlRight.Location = New-Object System.Drawing.Point(215, 0); $pnlRight.Size = New-Object System.Drawing.Size(590, 570)
    $pnlRight.BackColor = [System.Drawing.Color]::White; $pnlRight.BorderStyle = "FixedSingle"
    $pnlRight.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
    $pnlMain.Controls.Add($pnlRight)

    $lblInstruction = New-Object System.Windows.Forms.Label
    $lblInstruction.Text = Get-Text "modification.select_action"
    $lblInstruction.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblInstruction.ForeColor = [System.Drawing.Color]::FromArgb(160, 160, 160)
    $lblInstruction.Location = New-Object System.Drawing.Point(50, 220); $lblInstruction.Size = New-Object System.Drawing.Size(490, 80)
    $lblInstruction.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $pnlRight.Controls.Add($lblInstruction)

    # Fiche profil
    $pnlProfile = New-Object System.Windows.Forms.Panel
    $pnlProfile.Location = New-Object System.Drawing.Point(5, 5); $pnlProfile.Size = New-Object System.Drawing.Size(578, 560)
    $pnlProfile.Visible = $false; $pnlProfile.AutoScroll = $true; $pnlRight.Controls.Add($pnlProfile)

    $pnlPH = New-Object System.Windows.Forms.Panel
    $pnlPH.Location = New-Object System.Drawing.Point(0, 0); $pnlPH.Size = New-Object System.Drawing.Size(578, 55)
    $pnlPH.BackColor = [System.Drawing.Color]::FromArgb(0, 90, 180); $pnlProfile.Controls.Add($pnlPH)

    $lblPName = New-Object System.Windows.Forms.Label
    $lblPName.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblPName.ForeColor = $script:COLOR_WHITE; $lblPName.Location = New-Object System.Drawing.Point(12, 8)
    $lblPName.Size = New-Object System.Drawing.Size(554, 24); $pnlPH.Controls.Add($lblPName)

    $lblPUpn = New-Object System.Windows.Forms.Label
    $lblPUpn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblPUpn.ForeColor = [System.Drawing.Color]::FromArgb(200, 225, 255)
    $lblPUpn.Location = New-Object System.Drawing.Point(12, 33); $lblPUpn.Size = New-Object System.Drawing.Size(554, 18)
    $pnlPH.Controls.Add($lblPUpn)

    # Grille profil — labels i18n via les clés action_* existantes
    $script:ProfileLabels = @{}
    $profileFields = @(
        @{ Key = "dept";    Lang = "modification.action_department" },
        @{ Key = "title";   Lang = "modification.action_title" },
        @{ Key = "emptype"; Lang = "modification.action_employee_type" },
        @{ Key = "country"; Lang = "modification.action_country" },
        @{ Key = "office";  Lang = "modification.action_office" },
        @{ Key = "mobile";  Lang = "modification.action_mobile" },
        @{ Key = "phone";   Lang = "modification.action_phone" },
        @{ Key = "status";  Lang = "modification.status_active" }
    )
    $yF = 65
    foreach ($pf in $profileFields) {
        $ll = New-Object System.Windows.Forms.Label
        $ll.Text = (Get-Text $pf.Lang) + " :"; $ll.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $ll.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
        $ll.Location = New-Object System.Drawing.Point(12, $yF); $ll.Size = New-Object System.Drawing.Size(120, 20)
        $pnlProfile.Controls.Add($ll)
        $lv = New-Object System.Windows.Forms.Label
        $lv.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lv.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
        $lv.Location = New-Object System.Drawing.Point(135, $yF); $lv.Size = New-Object System.Drawing.Size(430, 20)
        $pnlProfile.Controls.Add($lv)
        $script:ProfileLabels[$pf.Key] = $lv
        $yF += 26
    }

    # Séparateur connexions
    $lblSep = New-Object System.Windows.Forms.Label
    $lblSep.Text = "--- " + (Get-Text "modification.section_audit") + " " + ("-" * 40)
    $lblSep.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblSep.ForeColor = [System.Drawing.Color]::FromArgb(150, 150, 150)
    $lblSep.Location = New-Object System.Drawing.Point(12, $yF); $lblSep.Size = New-Object System.Drawing.Size(554, 18)
    $pnlProfile.Controls.Add($lblSep); $yF += 22

    $dgvSignIn = New-Object System.Windows.Forms.DataGridView
    $dgvSignIn.Location = New-Object System.Drawing.Point(5, $yF); $dgvSignIn.Size = New-Object System.Drawing.Size(564, 155)
    $dgvSignIn.ReadOnly = $true; $dgvSignIn.AutoSizeColumnsMode = "Fill"; $dgvSignIn.RowHeadersVisible = $false
    $dgvSignIn.AllowUserToAddRows = $false; $dgvSignIn.BackgroundColor = $script:COLOR_WHITE
    $dgvSignIn.BorderStyle = "None"; $dgvSignIn.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $dgvSignIn.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $dgvSignIn.ColumnHeadersDefaultCellStyle.BackColor = $script:COLOR_SECTION
    $dgvSignIn.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
    foreach ($cn in @((Get-Text "modification.signin_col_date"), (Get-Text "modification.signin_col_app"),
                      (Get-Text "modification.signin_col_ip"), (Get-Text "modification.signin_col_status"))) {
        $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $c.HeaderText = $cn; $c.Name = $cn
        $dgvSignIn.Columns.Add($c) | Out-Null
    }
    $pnlProfile.Controls.Add($dgvSignIn)

    # Menu latéral
    $menuItems = @(
        @{ Type="section"; Text=Get-Text "modification.section_profile" },
        @{ Type="btn"; Text=Get-Text "modification.action_name";          Tag="name" },
        @{ Type="btn"; Text=Get-Text "modification.action_title";         Tag="title" },
        @{ Type="btn"; Text=Get-Text "modification.action_department";    Tag="dept" },
        @{ Type="btn"; Text=Get-Text "modification.action_employee_type"; Tag="emptype" },
        @{ Type="btn"; Text=Get-Text "modification.action_country";       Tag="country" },
        @{ Type="btn"; Text=Get-Text "modification.action_address";       Tag="address" },
        @{ Type="btn"; Text=Get-Text "modification.action_office";        Tag="office" },
        @{ Type="btn"; Text=Get-Text "modification.action_manager";       Tag="manager" },
        @{ Type="section"; Text=Get-Text "modification.section_contact" },
        @{ Type="btn"; Text=Get-Text "modification.action_mobile"; Tag="mobile" },
        @{ Type="btn"; Text=Get-Text "modification.action_phone";  Tag="phone" },
        @{ Type="section"; Text=Get-Text "modification.section_messaging" },
        @{ Type="btn"; Text=Get-Text "modification.action_upn";     Tag="upn" },
        @{ Type="btn"; Text=Get-Text "modification.action_aliases"; Tag="aliases" },
        @{ Type="section"; Text=Get-Text "modification.section_licenses" },
        @{ Type="btn"; Text=Get-Text "modification.action_licenses"; Tag="licenses" },
        @{ Type="section"; Text=Get-Text "modification.section_groups" },
        @{ Type="btn"; Text=Get-Text "modification.action_groups_manage"; Tag="groups" },
        @{ Type="section"; Text=Get-Text "modification.section_security" },
        @{ Type="btn"; Text=Get-Text "modification.action_password";  Tag="password" },
        @{ Type="btn"; Text=Get-Text "modification.action_revoke";    Tag="revoke" },
        @{ Type="btn"; Text=Get-Text "modification.action_mfa_reset"; Tag="mfa" },
        @{ Type="btn"; Text=Get-Text "modification.action_toggle";    Tag="toggle" },
        @{ Type="section"; Text=Get-Text "modification.section_audit" },
        @{ Type="btn"; Text=Get-Text "modification.action_signin_logs"; Tag="logs" }
    )
    $yM = 5
    foreach ($item in $menuItems) {
        if ($item.Type -eq "section") {
            $pnlMenu.Controls.Add((New-SectionLabel -Texte $item.Text -Y $yM)); $yM += 24
        } else {
            $btn = New-MenuButton -Texte $item.Text -Tag $item.Tag -Y $yM
            $btn.Add_Click({
                $t = $this.Tag; $uid = $script:SelectedUserId; $u = $script:SelectedUserUPN; $ud = $script:SelectedUserData
                switch ($t) {
                    "name"     { Show-ModifyName       -UserId $uid -UPN $u -CurrentGiven $ud.GivenName -CurrentSurname $ud.Surname -CurrentDisplay $ud.DisplayName }
                    "title"    { Show-ModifyComboField  -UserId $uid -UPN $u -Field "JobTitle"       -Label (Get-Text "modification.action_title")         -CurrentValue $ud.JobTitle       -Items @() -GraphProperty "jobTitle" }
                    "dept"     { Show-ModifyComboField  -UserId $uid -UPN $u -Field "Department"     -Label (Get-Text "modification.action_department")    -CurrentValue $ud.Department     -Items $Config.departments -GraphProperty "department" }
                    "emptype"  { Show-ModifyComboField  -UserId $uid -UPN $u -Field "EmployeeType"   -Label (Get-Text "modification.action_employee_type") -CurrentValue $ud.EmployeeType   -Items $Config.employee_types -GraphProperty "employeeType" }
                    "country"  { Show-ModifyCountry     -UserId $uid -UPN $u -CurrentValue $ud.Country }
                    "address"  { Show-ModifyAddress     -UserId $uid -UPN $u -CurrentData $ud }
                    "office"   { Show-ModifyComboField  -UserId $uid -UPN $u -Field "OfficeLocation" -Label (Get-Text "modification.action_office")        -CurrentValue $ud.OfficeLocation -Items @() -GraphProperty "officeLocation" }
                    "manager"  { Show-ModifyManager     -UserId $uid -UPN $u }
                    "mobile"   { Show-ModifySimpleField -UserId $uid -UPN $u -Field "MobilePhone"    -Label (Get-Text "modification.action_mobile") -CurrentValue $ud.MobilePhone }
                    "phone"    { Show-ModifySimpleField -UserId $uid -UPN $u -Field "BusinessPhones" -Label (Get-Text "modification.action_phone")  -CurrentValue ($ud.BusinessPhones | Select-Object -First 1) }
                    "upn"      { Show-ModifyUPN         -UserId $uid -UPN $u }
                    "aliases"  { Show-ManageAliases     -UserId $uid -UPN $u }
                    "licenses" { Show-ManageLicenses    -UserId $uid -UPN $u }
                    "groups"   { Show-ManageGroups      -UserId $uid -UPN $u }
                    "password" { Show-ResetPassword     -UserId $uid -UPN $u }
                    "revoke"   { Invoke-RevokeSession   -UserId $uid -UPN $u }
                    "mfa"      { Invoke-MfaReset        -UserId $uid -UPN $u }
                    "toggle"   { Show-ToggleAccount     -UserId $uid -UPN $u -IsEnabled $ud.AccountEnabled }
                    "logs"     { Show-SignInLogs        -UserId $uid -UPN $u }
                }
            })
            $pnlMenu.Controls.Add($btn); $yM += 31
        }
    }
    $pnlMenu.AutoScrollMinSize = New-Object System.Drawing.Size(190, ($yM + 10))

    # Logique de recherche
    $script:SelectedUserId = $null; $script:SelectedUserUPN = $null
    $script:SelectedUserData = $null; $script:SearchResults = @()

    $btnRecherche.Add_Click({
        $terme = $txtRecherche.Text.Trim()
        if ($terme.Length -lt 2) { return }
        $lstResultats.Items.Clear()
        $result = Search-AzUsers -SearchTerm $terme -MaxResults 15
        if ($result.Success -and $result.Data) {
            $script:SearchResults = @($result.Data)
            foreach ($user in $script:SearchResults) {
                $st = if ($user.AccountEnabled) { Get-Text "modification.status_active" } else { Get-Text "modification.status_disabled" }
                $lstResultats.Items.Add("$($user.DisplayName) — $($user.UserPrincipalName) [$st]") | Out-Null
            }
            $pnlUser.Visible = $false; $pnlMain.Visible = $false; $lstResultats.Visible = $true
        }
    })
    $txtRecherche.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $btnRecherche.PerformClick(); $_.SuppressKeyPress = $true }
    })
    $lstResultats.Add_SelectedIndexChanged({
        if ($lstResultats.SelectedIndex -ge 0 -and $lstResultats.SelectedIndex -lt $script:SearchResults.Count) {
            Invoke-SelectUser -Selected $script:SearchResults[$lstResultats.SelectedIndex]
            $lstResultats.Visible = $false
        }
    })
    $btnChanger.Add_Click({
        $pnlUser.Visible = $false; $pnlMain.Visible = $false
        $txtRecherche.Text = ""; $lstResultats.Items.Clear(); $lstResultats.Visible = $false; $txtRecherche.Focus()
    })

    # Footer
    $pnlFooter = New-Object System.Windows.Forms.Panel
    $pnlFooter.Dock = [System.Windows.Forms.DockStyle]::Bottom; $pnlFooter.Height = 44
    $pnlFooter.BackColor = [System.Drawing.Color]::FromArgb(235, 237, 240); $form.Controls.Add($pnlFooter)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = Get-Text "modification.btn_refresh"; $btnRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnRefresh.Location = New-Object System.Drawing.Point(8, 7); $btnRefresh.Size = New-Object System.Drawing.Size(150, 30)
    $btnRefresh.FlatStyle = "Flat"
    $btnRefresh.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left)
    $btnRefresh.Add_Click({
        if ($null -ne $script:SelectedUserId) {
            Invoke-SelectUser -Selected ([PSCustomObject]@{ Id = $script:SelectedUserId; UserPrincipalName = $script:SelectedUserUPN })
        }
    })
    $pnlFooter.Controls.Add($btnRefresh)

    $btnFermer = New-Object System.Windows.Forms.Button
    $btnFermer.Text = Get-Text "modification.btn_close"; $btnFermer.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnFermer.Location = New-Object System.Drawing.Point(690, 7); $btnFermer.Size = New-Object System.Drawing.Size(150, 30)
    $btnFermer.FlatStyle = "Flat"; $btnFermer.BackColor = $script:COLOR_RED; $btnFermer.ForeColor = $script:COLOR_WHITE
    $btnFermer.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
    $btnFermer.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $pnlFooter.Controls.Add($btnFermer); $form.CancelButton = $btnFermer

    $form.ShowDialog() | Out-Null; $form.Dispose()
}

# ============================================================
#  SÉLECTION UTILISATEUR
# ============================================================

function Invoke-SelectUser {
    param($Selected)
    try {
        Write-Log -Level "INFO" -Action "GET_USER_DETAIL" -UPN $Selected.UserPrincipalName -Message "Chargement propriétés étendues."
        $detail = Get-MgUser -UserId $Selected.Id -Property $script:USER_PROPS -ErrorAction Stop
    }
    catch {
        Write-Log -Level "ERROR" -Action "GET_USER_DETAIL" -UPN $Selected.UserPrincipalName -Message $_.Exception.Message
        $detail = $null
    }

    if ($null -ne $detail) {
        $script:SelectedUserId = $Selected.Id; $script:SelectedUserUPN = $Selected.UserPrincipalName
        $script:SelectedUserData = $detail
        $st = if ($detail.AccountEnabled) { Get-Text "modification.status_active" } else { Get-Text "modification.status_disabled" }
        $lblUserInfo.Text  = "$($detail.DisplayName) — $($detail.UserPrincipalName)"
        $lblUserInfo2.Text = "$($detail.Department) | $($detail.JobTitle) | $st"
        $lblPName.Text = $detail.DisplayName; $lblPUpn.Text = $detail.UserPrincipalName
        $script:ProfileLabels["dept"].Text    = if ($detail.Department)     { $detail.Department }     else { "—" }
        $script:ProfileLabels["title"].Text   = if ($detail.JobTitle)       { $detail.JobTitle }       else { "—" }
        $script:ProfileLabels["emptype"].Text = if ($detail.EmployeeType)   { $detail.EmployeeType }   else { "—" }
        $script:ProfileLabels["country"].Text = if ($detail.Country)        { $detail.Country }        else { "—" }
        $script:ProfileLabels["office"].Text  = if ($detail.OfficeLocation) { $detail.OfficeLocation } else { "—" }
        $script:ProfileLabels["mobile"].Text  = if ($detail.MobilePhone)    { $detail.MobilePhone }    else { "—" }
        $script:ProfileLabels["phone"].Text   = if ($detail.BusinessPhones -and $detail.BusinessPhones.Count -gt 0) { $detail.BusinessPhones[0] } else { "—" }
        $script:ProfileLabels["status"].Text  = $st

        $dgvSignIn.Rows.Clear()
        try {
            $upnQ = $detail.UserPrincipalName
            $si = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=userPrincipalName eq '$upnQ'&`$top=5&`$orderby=createdDateTime desc" -ErrorAction Stop
            if ($si.value) {
                foreach ($log in $si.value) {
                    $d = if ($log.createdDateTime) { [datetime]$log.createdDateTime | Get-Date -Format "MM-dd HH:mm" } else { "-" }
                    $a = if ($log.appDisplayName)  { $log.appDisplayName }  else { "-" }
                    $i = if ($log.ipAddress)       { $log.ipAddress }       else { "-" }
                    $ok = if ($log.status -and $log.status.errorCode -eq 0) { [char]0x2713 } else { [char]0x2717 }
                    $r = $dgvSignIn.Rows.Add($d, $a, $i, $ok)
                    if ($ok -eq [char]0x2717) { $dgvSignIn.Rows[$r].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Crimson }
                }
            } else { $dgvSignIn.Rows.Add("-", (Get-Text "modification.signin_no_data"), "-", "-") | Out-Null }
        }
        catch { $dgvSignIn.Rows.Add("-", (Get-Text "modification.signin_error" "AuditLog.Read.All"), "-", "-") | Out-Null }
        $lblInstruction.Visible = $false; $pnlProfile.Visible = $true
    } else {
        $script:SelectedUserId = $Selected.Id; $script:SelectedUserUPN = $Selected.UserPrincipalName
        $script:SelectedUserData = $Selected
        $lblUserInfo.Text = "$($Selected.DisplayName) — $($Selected.UserPrincipalName)"; $lblUserInfo2.Text = ""
    }
    $pnlUser.Visible = $true; $pnlMain.Visible = $true
}

# ============================================================
#  SOUS-FORMULAIRES — PROFIL & IDENTITÉ
# ============================================================

function Show-ModifyName {
    param([string]$UserId, [string]$UPN, [string]$CurrentGiven, [string]$CurrentSurname, [string]$CurrentDisplay)
    $f = New-SubForm -Titre (Get-Text "modification.action_name") -Largeur 460 -Hauteur 310
    $f.Controls.Add((New-LabelCurrent -Texte (Get-Text "modification.current_value" $CurrentDisplay)))
    $y = 45
    $tG = New-FormField -Parent $f -Label (Get-Text "modification.field_givenname")   -Value $CurrentGiven   -Y $y -Name "tG"; $y += 40
    $tS = New-FormField -Parent $f -Label (Get-Text "modification.field_surname")     -Value $CurrentSurname -Y $y -Name "tS"; $y += 40
    $tD = New-FormField -Parent $f -Label (Get-Text "modification.field_displayname") -Value $CurrentDisplay -Y $y -Name "tD"; $y += 40

    $btnA = New-BtnAppliquer -X 120 -Y ($y + 10)
    $btnA.Add_Click({
        $g = $tG.Text.Trim(); $s = $tS.Text.Trim(); $d = $tD.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($g) -or [string]::IsNullOrWhiteSpace($s) -or [string]::IsNullOrWhiteSpace($d)) {
            [System.Windows.Forms.MessageBox]::Show((Get-Text "modification.error_required_fields"), (Get-Text "modification.error_title"),
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") -Message (Get-Text "modification.confirm_name" $d $UPN)
        if ($confirm) {
            $result = Set-AzUser -UserId $UserId -Properties @{ GivenName=$g; Surname=$s; DisplayName=$d }
            if ($result.Success) {
                Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.success_name") -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "MODIFY_NAME" -UPN $UPN -Message "Nom modifié : '$CurrentDisplay' -> '$d'"
                $f.Close()
            } else { Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $result.Error -IsSuccess $false }
        }
    })
    $f.Controls.Add($btnA)
    $btnC = New-BtnAnnuler -X 240 -Y ($y + 10); $f.Controls.Add($btnC); $f.CancelButton = $btnC
    $f.ShowDialog() | Out-Null; $f.Dispose()
}

function Show-ModifySimpleField {
    <#  Sous-formulaire générique pour un champ texte. Gère MobilePhone et BusinessPhones via PATCH REST.  #>
    param([string]$UserId, [string]$UPN, [string]$Field, [string]$Label, [string]$CurrentValue)
    $f = New-SubForm -Titre $Label -Largeur 480 -Hauteur 220
    $f.Controls.Add((New-LabelCurrent -Texte (Get-Text "modification.current_value" $CurrentValue)))
    $txt = New-FormField -Parent $f -Label $Label -Value $CurrentValue -Y 50 -Name "txtVal"
    if ($Field -in @("MobilePhone","BusinessPhones")) {
        $f.Controls.Add((New-HintLabel -Texte (Get-Text "modification.phone_scope_hint")))
    }
    $btnA = New-BtnAppliquer -X 120 -Y 125
    $btnA.Add_Click({
        $nv = $txt.Text.Trim()
        if ($nv -eq $CurrentValue) {
            Show-ResultDialog -Titre (Get-Text "modification.info_title") -Message (Get-Text "modification.info_no_change") -IsSuccess $true; return
        }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") -Message (Get-Text "modification.confirm_field" $Label $CurrentValue $nv $UPN)
        if ($confirm) {
            if ($Field -eq "MobilePhone")    { $body = @{ mobilePhone = $nv } | ConvertTo-Json }
            elseif ($Field -eq "BusinessPhones") {
                $ph = if ([string]::IsNullOrWhiteSpace($nv)) { @() } else { @($nv) }
                $body = @{ businessPhones = $ph } | ConvertTo-Json
            } else { $body = @{ $Field = $nv } | ConvertTo-Json }
            try {
                Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$UserId" `
                    -Body $body -ContentType "application/json" -ErrorAction Stop
                Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.success_field" $Label) -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "MODIFY_FIELD" -UPN $UPN -Message "$Field : '$CurrentValue' -> '$nv'"
                $f.Close()
            } catch {
                $errMsg = Format-GraphErrorMessage -ErrorMessage $_.Exception.Message -Field $Field
                Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $errMsg -IsSuccess $false
            }
        }
    })
    $f.Controls.Add($btnA)
    $btnC = New-BtnAnnuler -X 250 -Y 125; $f.Controls.Add($btnC); $f.CancelButton = $btnC
    $f.ShowDialog() | Out-Null; $f.Dispose()
}

function Show-ModifyComboField {
    <#  Sous-formulaire combo + saisie libre avec chargement dynamique Graph.  #>
    param([string]$UserId, [string]$UPN, [string]$Field, [string]$Label,
          [string]$CurrentValue, [array]$Items, [string]$GraphProperty = "")
    $f = New-SubForm -Titre $Label -Largeur 480 -Hauteur 230
    $f.Controls.Add((New-LabelCurrent -Texte (Get-Text "modification.current_value" $CurrentValue)))
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Label; $lbl.Location = New-Object System.Drawing.Point(15, 50); $lbl.Size = New-Object System.Drawing.Size(130, 20)
    $f.Controls.Add($lbl)
    $cbo = New-Object System.Windows.Forms.ComboBox
    $cbo.Location = New-Object System.Drawing.Point(150, 47); $cbo.Size = New-Object System.Drawing.Size(295, 25)
    $cbo.DropDownStyle = "DropDown"; $cbo.AutoCompleteMode = "SuggestAppend"; $cbo.AutoCompleteSource = "ListItems"
    if ($null -ne $Items -and $Items.Count -gt 0) { foreach ($i in $Items) { $cbo.Items.Add($i) | Out-Null } }
    $f.Controls.Add($cbo)
    $lblSt = New-HintLabel -Texte "" -Y 78; $f.Controls.Add($lblSt)
    $f.Controls.Add((New-HintLabel -Texte (Get-Text "modification.combo_free_entry_note") -Y 98))
    $idx = $cbo.Items.IndexOf($CurrentValue)
    if ($idx -ge 0) { $cbo.SelectedIndex = $idx } elseif (-not [string]::IsNullOrWhiteSpace($CurrentValue)) { $cbo.Text = $CurrentValue }
    $btnA = New-BtnAppliquer -X 110 -Y 130; $f.Controls.Add($btnA)
    $btnC = New-BtnAnnuler -X 235 -Y 130; $f.Controls.Add($btnC); $f.CancelButton = $btnC

    $f.Add_Shown({
        if (-not [string]::IsNullOrWhiteSpace($GraphProperty)) {
            $lblSt.Text = Get-Text "modification.combo_loading"
            try {
                $ex = Get-AzDistinctValues -Property $GraphProperty -ErrorAction Stop
                if ($ex.Success -and $ex.Data.Count -gt 0) {
                    foreach ($v in $ex.Data) { if ($v -notin $cbo.Items) { $cbo.Items.Add($v) | Out-Null } }
                    $i2 = $cbo.Items.IndexOf($CurrentValue); if ($i2 -ge 0) { $cbo.SelectedIndex = $i2 }
                    $lblSt.Text = Get-Text "modification.combo_loaded" $cbo.Items.Count
                } else { $lblSt.Text = Get-Text "modification.combo_no_existing" }
            } catch { $lblSt.Text = Get-Text "modification.combo_load_error" }
        }
    })
    $btnA.Add_Click({
        $nv = $cbo.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($nv)) {
            [System.Windows.Forms.MessageBox]::Show((Get-Text "modification.error_required_fields"), (Get-Text "modification.error_title"),
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null; return
        }
        if ($nv -eq $CurrentValue) { Show-ResultDialog -Titre (Get-Text "modification.info_title") -Message (Get-Text "modification.info_no_change") -IsSuccess $true; return }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") -Message (Get-Text "modification.confirm_field" $Label $CurrentValue $nv $UPN)
        if ($confirm) {
            $result = Set-AzUser -UserId $UserId -Properties @{ $Field = $nv }
            if ($result.Success) {
                Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.success_field" $Label) -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "MODIFY_FIELD" -UPN $UPN -Message "$Field : '$CurrentValue' -> '$nv'"; $f.Close()
            } else { Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $result.Error -IsSuccess $false }
        }
    })
    $f.ShowDialog() | Out-Null; $f.Dispose()
}

function Show-ModifyCountry {
    param([string]$UserId, [string]$UPN, [string]$CurrentValue)
    $f = New-SubForm -Titre (Get-Text "modification.action_country") -Largeur 430 -Hauteur 200
    $f.Controls.Add((New-LabelCurrent -Texte (Get-Text "modification.current_value" $CurrentValue)))
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = Get-Text "modification.field_country"; $lbl.Location = New-Object System.Drawing.Point(15, 50)
    $lbl.Size = New-Object System.Drawing.Size(130, 20); $f.Controls.Add($lbl)
    $cbo = New-Object System.Windows.Forms.ComboBox
    $cbo.Location = New-Object System.Drawing.Point(150, 47); $cbo.Size = New-Object System.Drawing.Size(120, 25)
    $cbo.DropDownStyle = "DropDownList"
    foreach ($c in $script:ISO_COUNTRIES) { $cbo.Items.Add($c) | Out-Null }
    $idx = $cbo.Items.IndexOf($CurrentValue.ToUpper()); if ($idx -ge 0) { $cbo.SelectedIndex = $idx }
    $f.Controls.Add($cbo)
    $f.Controls.Add((New-HintLabel -Texte (Get-Text "modification.country_note")))
    $btnA = New-BtnAppliquer -X 100 -Y 110
    $btnA.Add_Click({
        if ($null -eq $cbo.SelectedItem) { return }
        $nv = $cbo.SelectedItem.ToString()
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") -Message (Get-Text "modification.confirm_country" $nv $UPN)
        if ($confirm) {
            $result = Set-AzUser -UserId $UserId -Properties @{ Country=$nv; UsageLocation=$nv }
            if ($result.Success) {
                Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.success_country") -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "MODIFY_COUNTRY" -UPN $UPN -Message "Pays : '$CurrentValue' -> '$nv'"; $f.Close()
            } else { Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $result.Error -IsSuccess $false }
        }
    })
    $f.Controls.Add($btnA)
    $btnC = New-BtnAnnuler -X 220 -Y 110; $f.Controls.Add($btnC); $f.CancelButton = $btnC
    $f.ShowDialog() | Out-Null; $f.Dispose()
}

function Show-ModifyAddress {
    param([string]$UserId, [string]$UPN, $CurrentData)
    $f = New-SubForm -Titre (Get-Text "modification.action_address") -Largeur 480 -Hauteur 340
    $fields = @(
        @{ Label=Get-Text "modification.field_street"; Field="StreetAddress"; Val=$CurrentData.StreetAddress },
        @{ Label=Get-Text "modification.field_city";   Field="City";          Val=$CurrentData.City },
        @{ Label=Get-Text "modification.field_state";  Field="State";         Val=$CurrentData.State },
        @{ Label=Get-Text "modification.field_postal"; Field="PostalCode";    Val=$CurrentData.PostalCode }
    )
    $y = 15; $tbs = @{}
    foreach ($fl in $fields) {
        $v = if ($null -eq $fl.Val) { "" } else { $fl.Val }
        $tbs[$fl.Field] = New-FormField -Parent $f -Label $fl.Label -Value $v -Y $y -Name $fl.Field; $y += 40
    }
    $btnA = New-BtnAppliquer -X 120 -Y ($y + 10)
    $btnA.Add_Click({
        $props = @{}; foreach ($fl in $fields) { $props[$fl.Field] = $tbs[$fl.Field].Text.Trim() }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") -Message (Get-Text "modification.confirm_address" $UPN)
        if ($confirm) {
            $result = Set-AzUser -UserId $UserId -Properties $props
            if ($result.Success) {
                Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.success_address") -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "MODIFY_ADDRESS" -UPN $UPN -Message "Adresse mise à jour."; $f.Close()
            } else { Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $result.Error -IsSuccess $false }
        }
    })
    $f.Controls.Add($btnA)
    $btnC = New-BtnAnnuler -X 240 -Y ($y + 10); $f.Controls.Add($btnC); $f.CancelButton = $btnC
    $f.ShowDialog() | Out-Null; $f.Dispose()
}

function Show-ModifyManager {
    param([string]$UserId, [string]$UPN)
    $f = New-SubForm -Titre (Get-Text "modification.action_manager") -Largeur 480 -Hauteur 290
    $mr = Get-AzUserManager -UserId $UserId
    $cm = if ($mr.Success -and $mr.Data) { $mr.Data.AdditionalProperties.displayName } else { Get-Text "modification.no_manager" }
    $f.Controls.Add((New-LabelCurrent -Texte (Get-Text "modification.current_value" $cm)))
    $tS = New-FormField -Parent $f -Label (Get-Text "modification.field_new_manager") -Y 50 -Name "tMgr" -FW 210
    $bS = New-Object System.Windows.Forms.Button
    $bS.Text = Get-Text "modification.btn_search"; $bS.Location = New-Object System.Drawing.Point(365, 47)
    $bS.Size = New-Object System.Drawing.Size(80, 25); $bS.FlatStyle = "Flat"; $f.Controls.Add($bS)
    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Location = New-Object System.Drawing.Point(150, 78); $lst.Size = New-Object System.Drawing.Size(295, 80); $lst.Visible = $false
    $f.Controls.Add($lst)
    $f.Tag = @{ MgrResults = @(); NewMgrId = $null }
    $bS.Add_Click({
        $t = $tS.Text.Trim(); if ($t.Length -lt 2) { return }
        $lst.Items.Clear()
        $res = Search-AzUsers -SearchTerm $t -MaxResults 10
        if ($res.Success -and $res.Data) {
            $f.Tag.MgrResults = @($res.Data)
            foreach ($u in $f.Tag.MgrResults) { $lst.Items.Add("$($u.DisplayName) ($($u.UserPrincipalName))") | Out-Null }
            $lst.Visible = $true
        }
    })
    $lst.Add_SelectedIndexChanged({
        if ($lst.SelectedIndex -ge 0) {
            $f.Tag.NewMgrId = $f.Tag.MgrResults[$lst.SelectedIndex].Id
            $tS.Text = $f.Tag.MgrResults[$lst.SelectedIndex].DisplayName; $lst.Visible = $false
        }
    })
    $btnA = New-BtnAppliquer -X 120 -Y 200
    $btnA.Add_Click({
        if ($null -eq $f.Tag.NewMgrId) { return }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") -Message (Get-Text "modification.confirm_manager" $tS.Text $UPN)
        if ($confirm) {
            $result = Set-AzUserManager -UserId $UserId -ManagerId $f.Tag.NewMgrId
            if ($result.Success) {
                Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.success_manager") -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "MODIFY_MANAGER" -UPN $UPN -Message "Manager -> $($tS.Text)"; $f.Close()
            } else { Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $result.Error -IsSuccess $false }
        }
    })
    $f.Controls.Add($btnA)
    $btnC = New-BtnAnnuler -X 240 -Y 200; $f.Controls.Add($btnC); $f.CancelButton = $btnC
    $f.ShowDialog() | Out-Null; $f.Dispose()
}

# ============================================================
#  SOUS-FORMULAIRES — MESSAGERIE & UPN
# ============================================================

function Show-ModifyUPN {
    param([string]$UserId, [string]$UPN)
    $f = New-SubForm -Titre (Get-Text "modification.action_upn") -Largeur 560 -Hauteur 320
    $lblLoad = New-Object System.Windows.Forms.Label
    $lblLoad.Text = Get-Text "modification.loading_domains"
    $lblLoad.Location = New-Object System.Drawing.Point(15, 15); $lblLoad.Size = New-Object System.Drawing.Size(510, 20)
    $f.Controls.Add($lblLoad)
    $parts = $UPN.Split("@"); $localPart = $parts[0]
    $curDomain = if ($parts.Count -gt 1) { "@" + $parts[1] } else { "" }
    $f.Controls.Add((New-LabelCurrent -Texte (Get-Text "modification.current_value" $UPN) -Y 40))

    $lL = New-Object System.Windows.Forms.Label
    $lL.Text = Get-Text "modification.field_upn_local"; $lL.Location = New-Object System.Drawing.Point(15, 75)
    $lL.Size = New-Object System.Drawing.Size(100, 20); $f.Controls.Add($lL)
    $tL = New-Object System.Windows.Forms.TextBox
    $tL.Text = $localPart; $tL.Location = New-Object System.Drawing.Point(120, 72); $tL.Size = New-Object System.Drawing.Size(180, 25)
    $f.Controls.Add($tL)
    $lAt = New-Object System.Windows.Forms.Label
    $lAt.Text = "@"; $lAt.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lAt.Location = New-Object System.Drawing.Point(305, 72); $lAt.Size = New-Object System.Drawing.Size(15, 25); $f.Controls.Add($lAt)
    $cbD = New-Object System.Windows.Forms.ComboBox
    $cbD.Location = New-Object System.Drawing.Point(323, 72); $cbD.Size = New-Object System.Drawing.Size(210, 25)
    $cbD.DropDownStyle = "DropDownList"; $f.Controls.Add($cbD)

    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = Get-Text "modification.upn_keep_alias"; $chk.Location = New-Object System.Drawing.Point(15, 115)
    $chk.Size = New-Object System.Drawing.Size(510, 35); $chk.Checked = $true
    $chk.Font = New-Object System.Drawing.Font("Segoe UI", 9); $f.Controls.Add($chk)
    $f.Controls.Add((New-HintLabel -Texte (Get-Text "modification.upn_alias_note") -X 30 -Y 150 -W 500))

    $btnA = New-BtnAppliquer -X 140 -Y 200; $btnA.Enabled = $false; $f.Controls.Add($btnA)
    $btnC = New-BtnAnnuler -X 260 -Y 200; $f.Controls.Add($btnC); $f.CancelButton = $btnC

    $f.Add_Shown({
        try {
            $doms = Get-MgDomain -ErrorAction Stop | Where-Object { $_.IsVerified } | Select-Object -ExpandProperty Id | Sort-Object
            $cbD.Items.Clear(); foreach ($d in $doms) { $cbD.Items.Add($d) | Out-Null }
            $dOnly = $curDomain.TrimStart("@"); $ix = $cbD.Items.IndexOf($dOnly)
            if ($ix -ge 0) { $cbD.SelectedIndex = $ix }
            $lblLoad.Text = Get-Text "modification.domains_loaded" $cbD.Items.Count; $btnA.Enabled = $true
        } catch { $lblLoad.Text = Get-Text "modification.error_domains" $_.Exception.Message }
    })
    $btnA.Add_Click({
        $nl = $tL.Text.Trim(); $nd = $cbD.SelectedItem
        if ([string]::IsNullOrWhiteSpace($nl) -or $null -eq $nd) { return }
        $newUPN = "$nl@$nd"
        if ($newUPN -eq $UPN) { Show-ResultDialog -Titre (Get-Text "modification.info_title") -Message (Get-Text "modification.info_no_change") -IsSuccess $true; return }
        $msg = Get-Text "modification.confirm_upn" $UPN $newUPN
        if ($chk.Checked) { $msg += "`n`n" + (Get-Text "modification.confirm_upn_alias" $UPN) }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") -Message $msg
        if ($confirm) {
            $result = Set-AzUser -UserId $UserId -Properties @{ UserPrincipalName = $newUPN }
            if (-not $result.Success) { Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $result.Error -IsSuccess $false; return }
            Write-Log -Level "SUCCESS" -Action "MODIFY_UPN" -UPN $UPN -Message "UPN : '$UPN' -> '$newUPN'"
            if ($chk.Checked) {
                try {
                    $usr = Get-MgUser -UserId $UserId -Property "proxyAddresses" -ErrorAction Stop
                    $alias = "smtp:$UPN"
                    if ($alias -notin $usr.ProxyAddresses) {
                        Invoke-ProxyAddressChange -UserId $UserId -UPN $newUPN -Action "Add" -Address $alias -CurrentList $usr.ProxyAddresses
                    }
                } catch { Write-Log -Level "WARNING" -Action "ADD_ALIAS" -UPN $newUPN -Message "Impossible d'ajouter l'alias : $($_.Exception.Message)" }
            }
            Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.success_upn" $newUPN) -IsSuccess $true
            $f.Close()
        }
    })
    $f.ShowDialog() | Out-Null; $f.Dispose()
}

function Show-ManageAliases {
    <#  Gestion des alias email (proxyAddresses) via Invoke-ProxyAddressChange.  #>
    param([string]$UserId, [string]$UPN)
    $f = New-SubForm -Titre (Get-Text "modification.action_aliases") -Largeur 580 -Hauteur 440
    $lT = New-Object System.Windows.Forms.Label
    $lT.Text = Get-Text "modification.aliases_for" $UPN
    $lT.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lT.Location = New-Object System.Drawing.Point(15, 10); $lT.Size = New-Object System.Drawing.Size(540, 20); $f.Controls.Add($lT)
    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Location = New-Object System.Drawing.Point(15, 38); $lst.Size = New-Object System.Drawing.Size(540, 180)
    $lst.Font = New-Object System.Drawing.Font("Consolas", 9); $f.Controls.Add($lst)
    $f.Controls.Add((New-HintLabel -Texte (Get-Text "modification.aliases_legend") -Y 222 -W 540))
    $lA = New-Object System.Windows.Forms.Label
    $lA.Text = Get-Text "modification.alias_add_label"
    $lA.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lA.Location = New-Object System.Drawing.Point(15, 252); $lA.Size = New-Object System.Drawing.Size(540, 18); $f.Controls.Add($lA)
    $tNA = New-Object System.Windows.Forms.TextBox
    $tNA.Location = New-Object System.Drawing.Point(15, 275); $tNA.Size = New-Object System.Drawing.Size(380, 25)
    $tNA.PlaceholderText = "alias@domaine.com"; $f.Controls.Add($tNA)

    $bAdd = New-ActionButton -Texte (Get-Text "modification.alias_btn_add")    -X 400 -Y 275 -W 155 -H 25 -Color $script:COLOR_GREEN
    $bRem = New-ActionButton -Texte (Get-Text "modification.alias_btn_remove") -X 15  -Y 315 -W 180 -H 30 -Color $script:COLOR_RED -Enabled $false
    $bCl  = New-CloseButton -X 440 -Y 365
    $f.Controls.Add($bAdd); $f.Controls.Add($bRem); $f.Controls.Add($bCl)

    $script:ProxyAddresses = @()
    $Refresh = {
        try {
            $u = Get-MgUser -UserId $UserId -Property "proxyAddresses" -ErrorAction Stop
            $script:ProxyAddresses = @($u.ProxyAddresses); $lst.Items.Clear()
            foreach ($a in $script:ProxyAddresses | Sort-Object) { $lst.Items.Add($a) | Out-Null }
        } catch { Write-Log -Level "ERROR" -Action "GET_ALIASES" -UPN $UPN -Message $_.Exception.Message }
    }
    & $Refresh

    $lst.Add_SelectedIndexChanged({ $bRem.Enabled = ($null -ne $lst.SelectedItem -and -not $lst.SelectedItem.StartsWith("SMTP:")) })

    $bAdd.Add_Click({
        $na = $tNA.Text.Trim().ToLower()
        if (-not $na.Contains("@")) {
            [System.Windows.Forms.MessageBox]::Show((Get-Text "modification.alias_invalid"), (Get-Text "modification.error_title"),
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null; return
        }
        $ae = "smtp:$na"
        if ($ae -in $script:ProxyAddresses) {
            [System.Windows.Forms.MessageBox]::Show((Get-Text "modification.alias_exists"), (Get-Text "modification.info_title"),
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null; return
        }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") -Message (Get-Text "modification.alias_confirm_add" $na $UPN)
        if ($confirm) {
            $ok = Invoke-ProxyAddressChange -UserId $UserId -UPN $UPN -Action "Add" -Address $ae -CurrentList $script:ProxyAddresses
            if ($ok) { $tNA.Text = ""; Start-Sleep -Milliseconds 500; & $Refresh }
        }
    })
    $bRem.Add_Click({
        $sel = $lst.SelectedItem; if ($null -eq $sel -or $sel.StartsWith("SMTP:")) { return }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") -Message (Get-Text "modification.alias_confirm_remove" $sel $UPN)
        if ($confirm) {
            $ok = Invoke-ProxyAddressChange -UserId $UserId -UPN $UPN -Action "Remove" -Address $sel -CurrentList $script:ProxyAddresses
            if ($ok) { Start-Sleep -Milliseconds 500; & $Refresh }
        }
    })
    $f.ShowDialog() | Out-Null; $f.Dispose()
}

# ============================================================
#  SOUS-FORMULAIRES — LICENCES
# ============================================================

function Show-ManageLicenses {
    param([string]$UserId, [string]$UPN)
    $f = New-SubForm -Titre (Get-Text "modification.action_licenses") -Largeur 680 -Hauteur 480
    $lT = New-Object System.Windows.Forms.Label
    $lT.Text = Get-Text "modification.licenses_for" $UPN
    $lT.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lT.Location = New-Object System.Drawing.Point(15, 10); $lT.Size = New-Object System.Drawing.Size(640, 20); $f.Controls.Add($lT)

    # Colonne gauche : assignées
    $lAs = New-Object System.Windows.Forms.Label
    $lAs.Text = Get-Text "modification.licenses_assigned"
    $lAs.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $lAs.ForeColor = $script:COLOR_GREEN
    $lAs.Location = New-Object System.Drawing.Point(15, 38); $lAs.Size = New-Object System.Drawing.Size(300, 20); $f.Controls.Add($lAs)
    $lstA = New-Object System.Windows.Forms.CheckedListBox
    $lstA.Location = New-Object System.Drawing.Point(15, 60); $lstA.Size = New-Object System.Drawing.Size(295, 280)
    $lstA.CheckOnClick = $true; $lstA.Font = New-Object System.Drawing.Font("Segoe UI", 9); $f.Controls.Add($lstA)

    # Colonne droite : disponibles
    $lAv = New-Object System.Windows.Forms.Label
    $lAv.Text = Get-Text "modification.licenses_available"
    $lAv.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $lAv.ForeColor = $script:COLOR_BLUE
    $lAv.Location = New-Object System.Drawing.Point(355, 38); $lAv.Size = New-Object System.Drawing.Size(300, 20); $f.Controls.Add($lAv)
    $lstV = New-Object System.Windows.Forms.CheckedListBox
    $lstV.Location = New-Object System.Drawing.Point(355, 60); $lstV.Size = New-Object System.Drawing.Size(295, 280)
    $lstV.CheckOnClick = $true; $lstV.Font = New-Object System.Drawing.Font("Segoe UI", 9); $f.Controls.Add($lstV)

    $f.Controls.Add((New-HintLabel -Texte (Get-Text "modification.licenses_note") -Y 347 -W 640))
    $bRev = New-ActionButton -Texte (Get-Text "modification.licenses_btn_revoke") -X 15  -Y 375 -Color $script:COLOR_RED  -Enabled $false
    $bAss = New-ActionButton -Texte (Get-Text "modification.licenses_btn_assign") -X 355 -Y 375 -Color $script:COLOR_BLUE -Enabled $false
    $bCl  = New-CloseButton -X 545 -Y 415 -W 110 -H 30
    $f.Controls.Add($bRev); $f.Controls.Add($bAss); $f.Controls.Add($bCl)

    $cfgGrp = @()
    if ($Config.PSObject.Properties["license_groups"]) { $cfgGrp = @($Config.license_groups) }

    # Scriptblock de rafraîchissement réutilisable
    $RefreshLic = {
        $lstA.Items.Clear(); $lstV.Items.Clear(); $bRev.Enabled = $false; $bAss.Enabled = $false
        try {
            $mo = Get-MgUserMemberOf -UserId $UserId -ErrorAction Stop
            $mgn = @($mo | ForEach-Object { if ($_.AdditionalProperties.ContainsKey("displayName")) { $_.AdditionalProperties.displayName } })
            foreach ($g in $cfgGrp) {
                if ($g -in $mgn) { $lstA.Items.Add($g, $false) | Out-Null } else { $lstV.Items.Add($g, $false) | Out-Null }
            }
            foreach ($n in $mgn) { if ($n -notin $cfgGrp -and $n -like "LIC-*") { $lstA.Items.Add("* $n", $false) | Out-Null } }
        } catch { Write-Log -Level "ERROR" -Action "LOAD_LICENSES" -UPN $UPN -Message $_.Exception.Message }
    }
    $f.Add_Shown({ & $RefreshLic })

    $lstA.Add_ItemCheck({ param($s,$e)
        $d = if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) { 1 } else { -1 }
        $bRev.Enabled = (($lstA.CheckedItems.Count + $d) -gt 0)
    })
    $lstV.Add_ItemCheck({ param($s,$e)
        $d = if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) { 1 } else { -1 }
        $bAss.Enabled = (($lstV.CheckedItems.Count + $d) -gt 0)
    })

    $bRev.Add_Click({
        $sel = @($lstA.CheckedItems | Where-Object { -not $_.StartsWith("*") })
        if ($sel.Count -eq 0) { return }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") -Message (Get-Text "modification.licenses_confirm_revoke" ($sel -join ", ") $UPN)
        if ($confirm) {
            $errs = @()
            foreach ($g in $sel) {
                try {
                    $go = Get-MgGroup -Filter "displayName eq '$g'" -ErrorAction Stop
                    if ($go) { Remove-MgGroupMemberByRef -GroupId $go.Id -DirectoryObjectId $UserId -ErrorAction Stop
                        Write-Log -Level "SUCCESS" -Action "REVOKE_LICENSE_GROUP" -UPN $UPN -Message "Retiré : $g" }
                } catch { $errs += "$g : $($_.Exception.Message)"; Write-Log -Level "ERROR" -Action "REVOKE_LICENSE_GROUP" -UPN $UPN -Message "Erreur $g : $($_.Exception.Message)" }
            }
            if ($errs.Count -eq 0) { Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.licenses_success_revoke") -IsSuccess $true }
            else { Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message ($errs -join "`n") -IsSuccess $false }
            & $RefreshLic
        }
    })
    $bAss.Add_Click({
        $sel = @($lstV.CheckedItems); if ($sel.Count -eq 0) { return }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") -Message (Get-Text "modification.licenses_confirm_assign" ($sel -join ", ") $UPN)
        if ($confirm) {
            $errs = @()
            foreach ($g in $sel) {
                try {
                    $go = Get-MgGroup -Filter "displayName eq '$g'" -ErrorAction Stop
                    if ($go) {
                        $bp = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId" }
                        New-MgGroupMemberByRef -GroupId $go.Id -BodyParameter $bp -ErrorAction Stop
                        Write-Log -Level "SUCCESS" -Action "ASSIGN_LICENSE_GROUP" -UPN $UPN -Message "Ajouté : $g"
                    }
                } catch {
                    if ($_.Exception.Message -notlike "*already exist*") {
                        $errs += "$g : $($_.Exception.Message)"; Write-Log -Level "ERROR" -Action "ASSIGN_LICENSE_GROUP" -UPN $UPN -Message "Erreur $g : $($_.Exception.Message)"
                    }
                }
            }
            if ($errs.Count -eq 0) { Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.licenses_success_assign") -IsSuccess $true }
            else { Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message ($errs -join "`n") -IsSuccess $false }
            & $RefreshLic
        }
    })
    $f.ShowDialog() | Out-Null; $f.Dispose()
}

# ============================================================
#  SOUS-FORMULAIRES — GROUPES
# ============================================================

function Show-ManageGroups {
    param([string]$UserId, [string]$UPN)
    $f = New-SubForm -Titre (Get-Text "modification.action_groups_manage") -Largeur 700 -Hauteur 580
    New-WarningBanner -Parent $f -Message (Get-Text "modification.groups_warning") -W 660

    $lAs = New-Object System.Windows.Forms.Label
    $lAs.Text = Get-Text "modification.groups_assigned_label"
    $lAs.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lAs.ForeColor = [System.Drawing.Color]::FromArgb(40, 120, 40)
    $lAs.Location = New-Object System.Drawing.Point(10, 68); $lAs.Size = New-Object System.Drawing.Size(320, 20); $f.Controls.Add($lAs)
    $lstA = New-Object System.Windows.Forms.ListBox
    $lstA.Location = New-Object System.Drawing.Point(10, 90); $lstA.Size = New-Object System.Drawing.Size(320, 370)
    $lstA.SelectionMode = "MultiExtended"; $lstA.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lstA.Sorted = $true
    $f.Controls.Add($lstA)

    $lSr = New-Object System.Windows.Forms.Label
    $lSr.Text = Get-Text "modification.groups_search_label"
    $lSr.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lSr.Location = New-Object System.Drawing.Point(345, 68); $lSr.Size = New-Object System.Drawing.Size(200, 20); $f.Controls.Add($lSr)
    $tSr = New-Object System.Windows.Forms.TextBox
    $tSr.Location = New-Object System.Drawing.Point(345, 90); $tSr.Size = New-Object System.Drawing.Size(240, 25)
    $tSr.PlaceholderText = Get-Text "modification.groups_search_placeholder"; $f.Controls.Add($tSr)
    $bSr = New-Object System.Windows.Forms.Button
    $bSr.Text = Get-Text "modification.btn_search"; $bSr.Location = New-Object System.Drawing.Point(590, 90)
    $bSr.Size = New-Object System.Drawing.Size(80, 25); $bSr.FlatStyle = "Flat"; $f.Controls.Add($bSr)
    $lstV = New-Object System.Windows.Forms.ListBox
    $lstV.Location = New-Object System.Drawing.Point(345, 120); $lstV.Size = New-Object System.Drawing.Size(325, 340)
    $lstV.SelectionMode = "MultiExtended"; $lstV.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lstV.Sorted = $true
    $f.Controls.Add($lstV)

    $lcA = New-HintLabel -Texte "" -X 10 -Y 464 -W 320; $f.Controls.Add($lcA)
    $lcV = New-HintLabel -Texte (Get-Text "modification.groups_search_hint") -X 345 -Y 464 -W 325; $f.Controls.Add($lcV)

    $bRem = New-ActionButton -Texte (Get-Text "modification.groups_btn_remove") -X 10  -Y 490 -W 220 -H 32 -Color $script:COLOR_RED  -Enabled $false
    $bAss = New-ActionButton -Texte (Get-Text "modification.groups_btn_assign") -X 345 -Y 490 -W 220 -H 32 -Color $script:COLOR_BLUE -Enabled $false
    $bCl  = New-CloseButton -X 570 -Y 490 -W 100 -H 32
    $f.Controls.Add($bRem); $f.Controls.Add($bAss); $f.Controls.Add($bCl)

    $script:GrpUserMap = @{}; $script:GrpAvailMap = @{}
    $LoadA = {
        $lstA.Items.Clear(); $script:GrpUserMap = @{}
        try {
            $mo = Get-MgUserMemberOf -UserId $UserId -ErrorAction Stop
            foreach ($m in $mo) {
                if ($m.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group') {
                    $n = $m.AdditionalProperties.displayName; $script:GrpUserMap[$n] = $m.Id
                    $lstA.Items.Add($n) | Out-Null
                }
            }
            $lcA.Text = Get-Text "modification.groups_count" $lstA.Items.Count
        } catch { Write-Log -Level "ERROR" -Action "LOAD_GROUPS" -UPN $UPN -Message $_.Exception.Message }
    }
    & $LoadA
    $SrchG = {
        $t = $tSr.Text.Trim(); if ($t.Length -lt 2) { return }
        $lstV.Items.Clear(); $script:GrpAvailMap = @{}
        try {
            $gs = Get-MgGroup -Filter "startsWith(displayName,'$t')" -Top 50 -Property "id,displayName" -ConsistencyLevel "eventual" -CountVariable c -ErrorAction Stop
            foreach ($g in $gs) {
                if ($g.DisplayName -notin $script:GrpUserMap.Keys) {
                    $script:GrpAvailMap[$g.DisplayName] = $g.Id; $lstV.Items.Add($g.DisplayName) | Out-Null
                }
            }
            $lcV.Text = Get-Text "modification.groups_count" $lstV.Items.Count
        } catch { $lcV.Text = "Erreur : $($_.Exception.Message)" }
    }
    $bSr.Add_Click({ & $SrchG })
    $tSr.Add_KeyDown({ if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { & $SrchG; $_.SuppressKeyPress = $true } })
    $lstA.Add_SelectedIndexChanged({ $bRem.Enabled = $lstA.SelectedItems.Count -gt 0 })
    $lstV.Add_SelectedIndexChanged({ $bAss.Enabled = $lstV.SelectedItems.Count -gt 0 })

    $bRem.Add_Click({
        $sel = @($lstA.SelectedItems); if ($sel.Count -eq 0) { return }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") -Message (Get-Text "modification.groups_confirm_remove" ($sel -join ", ") $UPN)
        if ($confirm) {
            $errs = @(); $rm = 0
            foreach ($gn in $sel) {
                $gid = $script:GrpUserMap[$gn]; if (-not $gid) { continue }
                try { Remove-MgGroupMemberByRef -GroupId $gid -DirectoryObjectId $UserId -ErrorAction Stop; $rm++
                    Write-Log -Level "SUCCESS" -Action "REMOVE_FROM_GROUP" -UPN $UPN -Message "Retiré : $gn"
                } catch { $errs += "$gn : $($_.Exception.Message)"; Write-Log -Level "ERROR" -Action "REMOVE_FROM_GROUP" -UPN $UPN -Message $_.Exception.Message }
            }
            if ($errs.Count -eq 0) { Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.groups_success_remove" $rm) -IsSuccess $true }
            else { Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message ($errs -join "`n") -IsSuccess $false }
            & $LoadA; if ($tSr.Text.Length -ge 2) { & $SrchG }
        }
    })
    $bAss.Add_Click({
        $sel = @($lstV.SelectedItems); if ($sel.Count -eq 0) { return }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") -Message (Get-Text "modification.groups_confirm_assign" ($sel -join ", ") $UPN)
        if ($confirm) {
            $errs = @(); $as = 0
            foreach ($gn in $sel) {
                $gid = $script:GrpAvailMap[$gn]; if (-not $gid) { continue }
                try {
                    $bp = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId" }
                    New-MgGroupMemberByRef -GroupId $gid -BodyParameter $bp -ErrorAction Stop; $as++
                    Write-Log -Level "SUCCESS" -Action "ADD_TO_GROUP" -UPN $UPN -Message "Ajouté : $gn"
                } catch {
                    if ($_.Exception.Message -notlike "*already exist*") {
                        $errs += "$gn : $($_.Exception.Message)"; Write-Log -Level "ERROR" -Action "ADD_TO_GROUP" -UPN $UPN -Message $_.Exception.Message
                    } else { $as++ }
                }
            }
            if ($errs.Count -eq 0) { Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.groups_success_assign" $as) -IsSuccess $true }
            else { Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message ($errs -join "`n") -IsSuccess $false }
            & $LoadA; if ($tSr.Text.Length -ge 2) { & $SrchG }
        }
    })
    $f.ShowDialog() | Out-Null; $f.Dispose()
}

# ============================================================
#  SOUS-FORMULAIRES — SÉCURITÉ
# ============================================================

function Show-ResetPassword {
    param([string]$UserId, [string]$UPN)
    $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.password_title") -Message (Get-Text "modification.password_confirm" $UPN)
    if (-not $confirm) { return }
    $pwPlain = New-SecurePassword
    $pw = ConvertTo-SecureString -String $pwPlain -AsPlainText -Force
    $result = Reset-AzUserPassword -UserId $UserId -NewPassword $pw -ForceChange $Config.password_policy.force_change_at_login
    if ($result.Success) {
        Write-Log -Level "SUCCESS" -Action "RESET_PASSWORD" -UPN $UPN -Message "Mot de passe réinitialisé."
        Show-PasswordDialog -UPN $UPN -Password $pwPlain
    } else {
        Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message (Get-Text "modification.password_error" $result.Error) -IsSuccess $false
    }
}

function Invoke-RevokeSession {
    param([string]$UserId, [string]$UPN)
    $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.revoke_title") -Message (Get-Text "modification.revoke_confirm" $UPN)
    if (-not $confirm) { return }
    try {
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$UserId/revokeSignInSessions" -ErrorAction Stop | Out-Null
        Write-Log -Level "SUCCESS" -Action "REVOKE_SESSIONS" -UPN $UPN -Message "Sessions révoquées."
        Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.revoke_success") -IsSuccess $true
    } catch {
        Write-Log -Level "ERROR" -Action "REVOKE_SESSIONS" -UPN $UPN -Message $_.Exception.Message
        Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $_.Exception.Message -IsSuccess $false
    }
}

function Invoke-MfaReset {
    param([string]$UserId, [string]$UPN)
    $f = New-SubForm -Titre (Get-Text "modification.mfa_title") -Largeur 540 -Hauteur 340
    New-WarningBanner -Parent $f -Message (Get-Text "modification.mfa_warning") -W 500 -H 60 `
        -Bg ([System.Drawing.Color]::FromArgb(255, 235, 235)) `
        -Fg ([System.Drawing.Color]::FromArgb(150, 0, 0)) `
        -Ico ([System.Drawing.Color]::FromArgb(150, 0, 0))

    $lM = New-Object System.Windows.Forms.Label
    $lM.Text = Get-Text "modification.mfa_methods_label"
    $lM.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lM.Location = New-Object System.Drawing.Point(10, 82); $lM.Size = New-Object System.Drawing.Size(500, 20); $f.Controls.Add($lM)
    $lstM = New-Object System.Windows.Forms.ListBox
    $lstM.Location = New-Object System.Drawing.Point(10, 106); $lstM.Size = New-Object System.Drawing.Size(500, 130)
    $lstM.Font = New-Object System.Drawing.Font("Consolas", 9); $f.Controls.Add($lstM)

    $script:MfaMethods = @()
    try {
        $mths = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UserId/authentication/methods" -ErrorAction Stop
        if ($mths.value) {
            $script:MfaMethods = @($mths.value)
            foreach ($m in $script:MfaMethods) { $t = $m.'@odata.type'.Split(".")[-1]; $lstM.Items.Add("$t — $($m.id)") | Out-Null }
        } else { $lstM.Items.Add((Get-Text "modification.mfa_no_methods")) | Out-Null }
    } catch { $lstM.Items.Add((Get-Text "modification.mfa_load_error" $_.Exception.Message)) | Out-Null }

    $bRes = New-ActionButton -Texte (Get-Text "modification.mfa_btn_reset") -X 10 -Y 255 -W 220 -Color $script:COLOR_ORANGE
    $bCl  = New-CloseButton -X 390 -Y 255
    $f.Controls.Add($bRes); $f.Controls.Add($bCl)

    $bRes.Add_Click({
        if ($script:MfaMethods.Count -eq 0) { return }
        $confirm = Show-ConfirmDialog -Titre (Get-Text "modification.confirm_title") -Message (Get-Text "modification.mfa_confirm_reset" $UPN $script:MfaMethods.Count)
        if ($confirm) {
            $errs = @(); $del = 0
            foreach ($m in $script:MfaMethods) {
                $t = $m.'@odata.type'.Split(".")[-1]
                if ($t -eq "passwordAuthenticationMethod") { continue }
                try {
                    Invoke-MgGraphRequest -Method DELETE `
                        -Uri "https://graph.microsoft.com/v1.0/users/$UserId/authentication/$($t)s/$($m.id)" -ErrorAction Stop | Out-Null
                    $del++; Write-Log -Level "SUCCESS" -Action "MFA_RESET" -UPN $UPN -Message "MFA supprimée : $t ($($m.id))"
                } catch { $errs += "$t : $($_.Exception.Message)"; Write-Log -Level "WARNING" -Action "MFA_RESET" -UPN $UPN -Message "Impossible supprimer $t : $($_.Exception.Message)" }
            }
            if ($errs.Count -eq 0) {
                Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.mfa_success" $del) -IsSuccess $true
                Write-Log -Level "SUCCESS" -Action "MFA_RESET" -UPN $UPN -Message "$del méthode(s) MFA supprimée(s)."; $f.Close()
            } else { Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message ($errs -join "`n") -IsSuccess $false }
        }
    })
    $f.ShowDialog() | Out-Null; $f.Dispose()
}

function Show-ToggleAccount {
    param([string]$UserId, [string]$UPN, [bool]$IsEnabled)
    $action = if ($IsEnabled) { Get-Text "modification.toggle_disable" } else { Get-Text "modification.toggle_enable" }
    $confirm = Show-ConfirmDialog -Titre $action -Message (Get-Text "modification.toggle_confirm" $action $UPN)
    if (-not $confirm) { return }
    $result = if ($IsEnabled) { Disable-AzUser -UserId $UserId } else { Enable-AzUser -UserId $UserId }
    if ($result.Success) {
        Show-ResultDialog -Titre (Get-Text "modification.success_title") -Message (Get-Text "modification.toggle_success" $action) -IsSuccess $true
        Write-Log -Level "SUCCESS" -Action "TOGGLE_ACCOUNT" -UPN $UPN -Message "Compte $action"
    } else { Show-ResultDialog -Titre (Get-Text "modification.error_title") -Message $result.Error -IsSuccess $false }
}

# ============================================================
#  SOUS-FORMULAIRES — AUDIT
# ============================================================

function Show-SignInLogs {
    param([string]$UserId, [string]$UPN)
    $f = New-SubForm -Titre (Get-Text "modification.action_signin_logs") -Largeur 780 -Hauteur 500
    $lT = New-Object System.Windows.Forms.Label
    $lT.Text = Get-Text "modification.signin_logs_for" $UPN
    $lT.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lT.Location = New-Object System.Drawing.Point(10, 10); $lT.Size = New-Object System.Drawing.Size(740, 20); $f.Controls.Add($lT)

    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Location = New-Object System.Drawing.Point(10, 38); $dgv.Size = New-Object System.Drawing.Size(745, 380)
    $dgv.ReadOnly = $true; $dgv.AllowUserToAddRows = $false; $dgv.AutoSizeColumnsMode = "Fill"
    $dgv.SelectionMode = "FullRowSelect"; $dgv.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $dgv.RowHeadersVisible = $false; $dgv.BackgroundColor = $script:COLOR_WHITE; $f.Controls.Add($dgv)
    foreach ($col in @(
        @{ N="Date";     H=Get-Text "modification.signin_col_date";     W=160 },
        @{ N="App";      H=Get-Text "modification.signin_col_app";      W=180 },
        @{ N="IP";       H=Get-Text "modification.signin_col_ip";       W=120 },
        @{ N="Location"; H=Get-Text "modification.signin_col_location"; W=140 },
        @{ N="Status";   H=Get-Text "modification.signin_col_status";   W=80 },
        @{ N="CA";       H=Get-Text "modification.signin_col_ca";       W=60 }
    )) {
        $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $c.Name = $col.N; $c.HeaderText = $col.H
        $dgv.Columns.Add($c) | Out-Null
    }

    $lblSt = New-HintLabel -Texte (Get-Text "modification.signin_loading") -X 10 -Y 423 -W 640; $f.Controls.Add($lblSt)
    $bCl = New-CloseButton -X 640 -Y 420 -W 115 -H 30; $f.Controls.Add($bCl)

    $f.Add_Shown({
        try {
            $resp = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=userPrincipalName eq '$UPN'&`$top=20&`$orderby=createdDateTime desc" -ErrorAction Stop
            if ($resp.value -and $resp.value.Count -gt 0) {
                foreach ($log in $resp.value) {
                    $date = if ($log.createdDateTime) { [datetime]$log.createdDateTime | Get-Date -Format "yyyy-MM-dd HH:mm" } else { "-" }
                    $app  = if ($log.appDisplayName) { $log.appDisplayName } else { "-" }
                    $ip   = if ($log.ipAddress)      { $log.ipAddress }      else { "-" }
                    $loc  = if ($log.location -and $log.location.city) { "$($log.location.city), $($log.location.countryOrRegion)" } else { "-" }
                    $ok   = if ($log.status -and $log.status.errorCode -eq 0) { [char]0x2713 } else { [char]0x2717 }
                    $ca   = if ($log.appliedConditionalAccessPolicies -and $log.appliedConditionalAccessPolicies.Count -gt 0) { "+" } else { "-" }
                    $r    = $dgv.Rows.Add($date, $app, $ip, $loc, $ok, $ca)
                    if ($ok -eq [char]0x2717) { $dgv.Rows[$r].DefaultCellStyle.ForeColor = $script:COLOR_RED }
                }
                $lblSt.Text = Get-Text "modification.signin_loaded" $resp.value.Count
            } else { $lblSt.Text = Get-Text "modification.signin_no_data" }
        } catch {
            $lblSt.Text = Get-Text "modification.signin_error" $_.Exception.Message
            Write-Log -Level "ERROR" -Action "SIGNIN_LOGS" -UPN $UPN -Message $_.Exception.Message
        }
    })
    $f.ShowDialog() | Out-Null; $f.Dispose()
}