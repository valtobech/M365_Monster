<#
.FICHIER
    Core/Functions.ps1

.ROLE
    Fonctions utilitaires réutilisables par tous les modules :
    journalisation, génération de mot de passe, génération de nom d'utilisateur,
    dialogs WinForms, envoi de notifications.

.DEPENDANCES
    - [System.Windows.Forms] (chargé dans Main.ps1)
    - Core/Config.ps1 (variable $Config)
    - Core/GraphAPI.ps1 (pour Send-Notification)

.AUTEUR
    [Équipe IT — GestionRH-AzureAD]
#>

# === Variable de chemin du fichier log pour la session courante ===
$script:LogFilePath = $null

function Initialize-LogFile {
    <#
    .SYNOPSIS
        Initialise le fichier de log pour la session courante.

    .PARAMETER LogFolder
        Chemin vers le dossier Logs/.

    .OUTPUTS
        [string] — Chemin complet du fichier de log créé.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogFolder
    )

    # Création du dossier si inexistant
    if (-not (Test-Path -Path $LogFolder)) {
        New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
    $script:LogFilePath = Join-Path -Path $LogFolder -ChildPath "session_$timestamp.log"

    # Écriture de l'en-tête
    $header = "====== Session M365 Monster démarrée le $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ======"
    $header | Out-File -FilePath $script:LogFilePath -Encoding UTF8 -Append

    return $script:LogFilePath
}

function Write-Log {
    <#
    .SYNOPSIS
        Écrit une entrée dans le fichier de log horodaté.

    .PARAMETER Level
        Niveau du message : INFO, WARNING, ERROR, SUCCESS.

    .PARAMETER Action
        Identifiant de l'action en cours (ex: CREATE_USER, CONNEXION).

    .PARAMETER UPN
        UPN de l'utilisateur cible (optionnel).

    .PARAMETER Message
        Contenu du message à logger.

    .OUTPUTS
        [void]
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level,

        [Parameter(Mandatory = $true)]
        [string]$Action,

        [string]$UPN = "-",

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $ligne = "[$timestamp] [$Level] [$Action] [$UPN] $Message"

    # Écriture dans le fichier log
    if ($script:LogFilePath -and (Test-Path -Path (Split-Path $script:LogFilePath -Parent))) {
        $ligne | Out-File -FilePath $script:LogFilePath -Encoding UTF8 -Append
    }

    # Écriture aussi dans la console pour le débogage
    switch ($Level) {
        "ERROR"   { Write-Host $ligne -ForegroundColor Red }
        "WARNING" { Write-Host $ligne -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $ligne -ForegroundColor Green }
        default   { Write-Host $ligne -ForegroundColor Gray }
    }
}

function New-UsernameVariants {
    <#
    .SYNOPSIS
        Génère les 3 variantes de UserPrincipalName pour un employé,
        vérifie l'unicité dans Azure AD, et incrémente si nécessaire.

    .PARAMETER Prenom
        Prénom de l'employé.

    .PARAMETER Nom
        Nom de famille de l'employé.

    .DESCRIPTION
        Formats générés :
        1. Première lettre du prénom + nom       → lbechetoille@domaine.com
        2. Prénom.Nom                             → leo.bechetoille@domaine.com
        3. Prénom + Nom (collés)                  → leobechetoille@domaine.com

        Si un UPN existe déjà, on ajoute un chiffre incrémental :
        lbechetoille1, lbechetoille2, etc.

    .OUTPUTS
        [array] — Liste d'objets {Label: string, UPN: string, MailNickname: string}
                  Chaque UPN est garanti unique dans Azure AD.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prenom,

        [Parameter(Mandatory = $true)]
        [string]$Nom
    )

    $domain = $Config.smtp_domain

    # Nettoyage : accents, espaces, minuscules
    $prenomClean = (Remove-Diacritics -Text $Prenom.Trim()).ToLower() -replace '[^a-z\-]', ''
    $nomClean = (Remove-Diacritics -Text $Nom.Trim()).ToLower() -replace '[^a-z\-]', ''

    # Retirer les tirets pour les calculs de base
    $prenomSafe = $prenomClean -replace '-', ''
    $nomSafe = $nomClean -replace '-', ''

    # Les 3 bases de nommage
    $bases = @(
        @{ Label = "Initiale + Nom";    Base = "$($prenomSafe[0])$nomSafe" }
        @{ Label = "Prénom.Nom";        Base = "$prenomClean.$nomClean" }
        @{ Label = "Prénom + Nom";      Base = "$prenomSafe$nomSafe" }
    )

    $resultats = @()

    foreach ($item in $bases) {
        $baseUsername = $item.Base
        $label = $item.Label
        $candidat = $baseUsername
        $compteur = 0

        # Vérifier l'unicité — incrémenter si nécessaire
        while ($true) {
            $upnTest = "$candidat$domain"
            $existe = Test-AzUserExists -UPN $upnTest

            if (-not $existe) {
                break
            }

            $compteur++
            $candidat = "$baseUsername$compteur"

            # Sécurité : éviter boucle infinie
            if ($compteur -gt 99) {
                Write-Log -Level "WARNING" -Action "USERNAME" -Message "Plus de 99 doublons pour '$baseUsername' — abandon."
                $candidat = "$baseUsername$compteur"
                break
            }
        }

        $upnFinal = "$candidat$domain"
        $mailNickname = $candidat

        # Si une incrémentation a été nécessaire, l'indiquer dans le label
        if ($compteur -gt 0) {
            $label = "$label (incrémenté +$compteur)"
        }

        $resultats += [PSCustomObject]@{
            Label        = $label
            UPN          = $upnFinal
            MailNickname = $mailNickname
            Username     = $candidat
        }
    }

    Write-Log -Level "INFO" -Action "USERNAME" -Message "Variantes générées pour '$Prenom $Nom' : $($resultats | ForEach-Object { $_.UPN }) "

    return $resultats
}

function Remove-Diacritics {
    <#
    .SYNOPSIS
        Supprime les accents et caractères diacritiques d'une chaîne.

    .PARAMETER Text
        Chaîne à nettoyer.

    .OUTPUTS
        [string] — Chaîne sans accents.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    # CHOIX: Utilisation de la normalisation Unicode pour retirer les diacritiques
    $normalized = $Text.Normalize([System.Text.NormalizationForm]::FormD)
    $builder = New-Object System.Text.StringBuilder

    foreach ($char in $normalized.ToCharArray()) {
        $category = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($char)
        if ($category -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($char)
        }
    }

    return $builder.ToString().Normalize([System.Text.NormalizationForm]::FormC)
}

function New-SecurePassword {
    <#
    .SYNOPSIS
        Génère un mot de passe sécurisé selon la politique définie dans la configuration client.

    .OUTPUTS
        [string] — Mot de passe généré.
    #>
    [CmdletBinding()]
    param()

    $longueur = $Config.password_policy.length
    $inclureSpeciaux = $Config.password_policy.include_special_chars

    # Jeux de caractères
    $minuscules = 'abcdefghijkmnopqrstuvwxyz'  # Sans l (confusion avec 1)
    $majuscules = 'ABCDEFGHJKLMNPQRSTUVWXYZ'    # Sans I, O (confusion avec 1, 0)
    $chiffres = '23456789'                       # Sans 0, 1 (confusion)
    $speciaux = '!@#$%&*-_=+'

    # Garantir au moins un caractère de chaque type
    $password = @()
    $password += $minuscules[(Get-Random -Maximum $minuscules.Length)]
    $password += $majuscules[(Get-Random -Maximum $majuscules.Length)]
    $password += $chiffres[(Get-Random -Maximum $chiffres.Length)]

    if ($inclureSpeciaux) {
        $password += $speciaux[(Get-Random -Maximum $speciaux.Length)]
    }

    # Compléter la longueur restante
    $tousCaracteres = $minuscules + $majuscules + $chiffres
    if ($inclureSpeciaux) {
        $tousCaracteres += $speciaux
    }

    $restant = $longueur - $password.Count
    for ($i = 0; $i -lt $restant; $i++) {
        $password += $tousCaracteres[(Get-Random -Maximum $tousCaracteres.Length)]
    }

    # Mélanger les caractères (Fisher-Yates shuffle)
    # $password est déjà un tableau PowerShell @() de caractères
    for ($i = $password.Count - 1; $i -gt 0; $i--) {
        $j = Get-Random -Maximum ($i + 1)
        $temp = $password[$i]
        $password[$i] = $password[$j]
        $password[$j] = $temp
    }

    return -join $password
}

function Show-ConfirmDialog {
    <#
    .SYNOPSIS
        Affiche une boîte de dialogue de confirmation WinForms.

    .PARAMETER Titre
        Titre de la fenêtre.

    .PARAMETER Message
        Message à afficher.

    .PARAMETER Icon
        Icône à afficher (défaut : Question).

    .OUTPUTS
        [bool] — $true si l'utilisateur confirme (Yes), $false sinon.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Titre,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Question
    )

    $result = [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Titre,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        $Icon
    )

    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Show-ResultDialog {
    <#
    .SYNOPSIS
        Affiche une boîte de dialogue de résultat (succès ou erreur).

    .PARAMETER Titre
        Titre de la fenêtre.

    .PARAMETER Message
        Message à afficher.

    .PARAMETER IsSuccess
        Si $true, affiche une icône de succès. Sinon, une icône d'erreur.

    .OUTPUTS
        [void]
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Titre,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [bool]$IsSuccess = $true
    )

    $icon = if ($IsSuccess) {
        [System.Windows.Forms.MessageBoxIcon]::Information
    }
    else {
        [System.Windows.Forms.MessageBoxIcon]::Error
    }

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Titre,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $icon
    ) | Out-Null
}

function Show-PasswordDialog {
    <#
    .SYNOPSIS
        Affiche le mot de passe généré dans une fenêtre avec bouton "Copier".
        Le mot de passe n'est affiché qu'une seule fois.

    .PARAMETER UPN
        Nom d'utilisateur associé.

    .PARAMETER InitialToken
        Mot de passe temporaire à afficher.

    .OUTPUTS
        [void]
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UPN,

        [Parameter(Mandatory = $true)]
        [string]$InitialToken
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = Get-Text "password_dialog.title"
    $form.Size = New-Object System.Drawing.Size(450, 280)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    # Icône d'avertissement
    $lblWarning = New-Object System.Windows.Forms.Label
    $lblWarning.Text = Get-Text "password_dialog.warning"
    $lblWarning.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblWarning.ForeColor = [System.Drawing.Color]::DarkRed
    $lblWarning.Location = New-Object System.Drawing.Point(20, 15)
    $lblWarning.Size = New-Object System.Drawing.Size(400, 25)
    $form.Controls.Add($lblWarning)

    # UPN
    $lblUpnTitle = New-Object System.Windows.Forms.Label
    $lblUpnTitle.Text = Get-Text "password_dialog.username_label"
    $lblUpnTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblUpnTitle.Location = New-Object System.Drawing.Point(20, 55)
    $lblUpnTitle.Size = New-Object System.Drawing.Size(130, 20)
    $form.Controls.Add($lblUpnTitle)

    $txtUpn = New-Object System.Windows.Forms.TextBox
    $txtUpn.Text = $UPN
    $txtUpn.ReadOnly = $true
    $txtUpn.Font = New-Object System.Drawing.Font("Consolas", 10)
    $txtUpn.Location = New-Object System.Drawing.Point(155, 52)
    $txtUpn.Size = New-Object System.Drawing.Size(260, 25)
    $form.Controls.Add($txtUpn)

    # Mot de passe
    $lblPwdTitle = New-Object System.Windows.Forms.Label
    $lblPwdTitle.Text = Get-Text "password_dialog.password_label"
    $lblPwdTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblPwdTitle.Location = New-Object System.Drawing.Point(20, 90)
    $lblPwdTitle.Size = New-Object System.Drawing.Size(130, 20)
    $form.Controls.Add($lblPwdTitle)

    $txtPwd = New-Object System.Windows.Forms.TextBox
    $txtPwd.Text = $InitialToken
    $txtPwd.ReadOnly = $true
    $txtPwd.Font = New-Object System.Drawing.Font("Consolas", 10)
    $txtPwd.Location = New-Object System.Drawing.Point(155, 87)
    $txtPwd.Size = New-Object System.Drawing.Size(260, 25)
    $form.Controls.Add($txtPwd)

    # Bouton Copier UPN + MDP
    $btnCopier = New-Object System.Windows.Forms.Button
    $btnCopier.Text = Get-Text "password_dialog.btn_copy"
    $btnCopier.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnCopier.Location = New-Object System.Drawing.Point(20, 135)
    $btnCopier.Size = New-Object System.Drawing.Size(395, 35)
    $btnCopier.Add_Click({
        $clipText = "$(Get-Text 'password_dialog.clipboard_user'): $UPN`n$(Get-Text 'password_dialog.clipboard_pass'): $InitialToken"
        [System.Windows.Forms.Clipboard]::SetText($clipText)
        [System.Windows.Forms.MessageBox]::Show(    (Get-Text "password_dialog.copied_msg"), (Get-Text "password_dialog.copied_title"), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    })
    $form.Controls.Add($btnCopier)

    # Bouton Fermer
    $btnFermer = New-Object System.Windows.Forms.Button
    $btnFermer.Text = Get-Text "password_dialog.btn_close"
    $btnFermer.Location = New-Object System.Drawing.Point(155, 190)
    $btnFermer.Size = New-Object System.Drawing.Size(120, 35)
    $btnFermer.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnFermer)

    $form.AcceptButton = $btnFermer
    $form.ShowDialog() | Out-Null
    $form.Dispose()
}

function Send-Notification {
    <#
    .SYNOPSIS
        Envoie une notification par email via Microsoft Graph.

    .PARAMETER Sujet
        Sujet de l'email.

    .PARAMETER Corps
        Corps du message (HTML supporté).

    .PARAMETER Destinataires
        Liste des adresses email destinataires.
        Si non spécifié, utilise les destinataires de $Config.notifications.recipients.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Sujet,

        [Parameter(Mandatory = $true)]
        [string]$Corps,

        [string[]]$Destinataires
    )

    # Vérification que les notifications sont activées
    if (-not $Config.notifications.enabled) {
        Write-Log -Level "INFO" -Action "NOTIFICATION" -Message "Notifications désactivées dans la configuration."
        return [PSCustomObject]@{ Success = $true; Error = $null }
    }

    if (-not $Destinataires -or $Destinataires.Count -eq 0) {
        $Destinataires = $Config.notifications.recipients
    }

    if (-not $Destinataires -or $Destinataires.Count -eq 0) {
        Write-Log -Level "WARNING" -Action "NOTIFICATION" -Message "Aucun destinataire configuré."
        return [PSCustomObject]@{ Success = $true; Error = $null }
    }

    try {
        Write-Log -Level "INFO" -Action "NOTIFICATION" -Message "Envoi de notification : '$Sujet' à $($Destinataires -join ', ')"

        $toRecipients = $Destinataires | ForEach-Object {
            @{
                EmailAddress = @{ Address = $_ }
            }
        }

        # CHOIX: On utilise l'utilisateur connecté comme expéditeur via /me/sendMail
        $params = @{
            Message = @{
                Subject      = $Sujet
                Body         = @{
                    ContentType = "HTML"
                    Content     = $Corps
                }
                ToRecipients = @($toRecipients)
            }
            SaveToSentItems = $false
        }

        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/me/sendMail" -Body ($params | ConvertTo-Json -Depth 10) -ContentType "application/json" -ErrorAction Stop

        Write-Log -Level "SUCCESS" -Action "NOTIFICATION" -Message "Notification envoyée."
        return [PSCustomObject]@{ Success = $true; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "NOTIFICATION" -Message "Erreur envoi notification : $errMsg"
        return [PSCustomObject]@{ Success = $false; Error = $errMsg }
    }
}

function Test-GraphConnection {
    <#
    .SYNOPSIS
        Vérifie que le token Graph est valide en effectuant un appel simple.

    .OUTPUTS
        [bool] — $true si la connexion est fonctionnelle.
    #>
    [CmdletBinding()]
    param()

    try {
        $context = Get-MgContext
        if ($null -eq $context) { return $false }

        # Test avec un appel léger
        Get-MgOrganization -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Get-MailNickname {
    <#
    .SYNOPSIS
        Génère le MailNickname à partir du UPN (partie avant le @).

    .PARAMETER UPN
        UserPrincipalName complet.

    .OUTPUTS
        [string] — MailNickname.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UPN
    )

    return $UPN.Split('@')[0]
}

# Point d'attention :
# - Write-Log écrit à la fois dans le fichier ET dans la console
# - New-SecurePassword exclut les caractères ambigus (0/O, 1/l/I)
# - Remove-Diacritics est essentiel pour les noms francophones (é, è, ê, ç, etc.)
# - Show-PasswordDialog est modal et force l'attention de l'utilisateur
# - Send-Notification utilise /me/sendMail — l'utilisateur connecté doit avoir une boîte mail
