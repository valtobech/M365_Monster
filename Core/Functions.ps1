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

function Get-AccessProfiles {
    <#
    .SYNOPSIS
        Retourne la liste des profils d'accès du client courant.

    .PARAMETER ExcludeBaseline
        Exclure le profil baseline (pratique pour les dropdowns de sélection).

    .OUTPUTS
        [PSCustomObject[]] — Liste des profils avec Key, DisplayName, Description, IsBaseline, Groups.
    #>
    [CmdletBinding()]
    param(
        [switch]$ExcludeBaseline
    )

    # Vérifier que la section existe dans la config client
    if (-not $Config.PSObject.Properties["access_profiles"]) {
        Write-Log -Level "WARNING" -Action "ACCESS_PROFILES" -Message "Aucun profil d'accès configuré pour ce client."
        return @()
    }

    $profiles = @()
    foreach ($key in $Config.access_profiles.PSObject.Properties.Name) {
        $ap = $Config.access_profiles.$key
        if ($ExcludeBaseline -and $ap.is_baseline) { continue }

        $profiles += [PSCustomObject]@{
            Key         = $key
            DisplayName = $ap.display_name
            Description = $ap.description
            IsBaseline  = [bool]$ap.is_baseline
            Groups      = @($ap.groups)
}
    }

    return $profiles
}

function Get-BaselineProfile {
    <#
    .SYNOPSIS
        Retourne le profil baseline (is_baseline = true), ou $null si inexistant.

    .OUTPUTS
        [PSCustomObject] — Profil baseline avec Key, DisplayName, Description, IsBaseline, Groups.
                           $null si aucun profil baseline n'est défini.
    #>
    [CmdletBinding()]
    param()

    $all = Get-AccessProfiles
    return $all | Where-Object { $_.IsBaseline -eq $true } | Select-Object -First 1
}

function Compare-AccessProfileGroups {
    <#
    .SYNOPSIS
        Compare deux ensembles de profils et retourne le diff de groupes.
        Gère intelligemment les intersections pour éviter retrait + re-ajout inutile.

    .PARAMETER OldProfileKeys
        Clés des profils actuellement assignés à l'utilisateur.

    .PARAMETER NewProfileKeys
        Clés des profils cibles à appliquer.

    .PARAMETER IncludeBaseline
        Inclure automatiquement le profil baseline dans les deux ensembles.

    .OUTPUTS
        [PSCustomObject] — {
            ToAdd    : [array] — Groupes à ajouter   (objets {id, display_name})
            ToRemove : [array] — Groupes à retirer    (objets {id, display_name})
            ToKeep   : [array] — Groupes inchangés    (objets {id, display_name})
        }
    #>
    [CmdletBinding()]
    param(
        [string[]]$OldProfileKeys = @(),
        [string[]]$NewProfileKeys = @(),
        [switch]$IncludeBaseline
    )

    # Collecter tous les groupes des anciens profils (hashtable id → objet)
    $oldGroups = @{}
    foreach ($key in $OldProfileKeys) {
        if ($Config.access_profiles.PSObject.Properties[$key]) {
            foreach ($grp in $Config.access_profiles.$key.groups) {
                if (-not $oldGroups.ContainsKey($grp.id)) {
                    $oldGroups[$grp.id] = $grp
                }
            }
        }
    }

    # Collecter tous les groupes des nouveaux profils
    $newGroups = @{}
    foreach ($key in $NewProfileKeys) {
        if ($Config.access_profiles.PSObject.Properties[$key]) {
            foreach ($grp in $Config.access_profiles.$key.groups) {
                if (-not $newGroups.ContainsKey($grp.id)) {
                    $newGroups[$grp.id] = $grp
                }
            }
        }
    }

    # Ajouter le baseline dans les deux ensembles si demandé
    if ($IncludeBaseline) {
        $baseline = Get-BaselineProfile
        if ($baseline) {
            foreach ($grp in $baseline.Groups) {
                if ($OldProfileKeys.Count -gt 0 -and -not $oldGroups.ContainsKey($grp.id)) { $oldGroups[$grp.id] = $grp }
                if (-not $newGroups.ContainsKey($grp.id)) { $newGroups[$grp.id] = $grp }
            }
        }
    }

    # Calculer le diff
    $toAdd    = @()
    $toRemove = @()
    $toKeep   = @()

    # Groupes dans le nouveau set : soit à ajouter, soit à conserver
    foreach ($id in $newGroups.Keys) {
        if ($oldGroups.ContainsKey($id)) {
            $toKeep += $newGroups[$id]
        }
        else {
            $toAdd += $newGroups[$id]
        }
    }

    # Groupes dans l'ancien set mais pas dans le nouveau → à retirer
    foreach ($id in $oldGroups.Keys) {
        if (-not $newGroups.ContainsKey($id)) {
            $toRemove += $oldGroups[$id]
        }
    }

    return [PSCustomObject]@{
        ToAdd    = $toAdd
        ToRemove = $toRemove
        ToKeep   = $toKeep
    }
}

function Invoke-AccessProfileChange {
    <#
    .SYNOPSIS
        Applique un changement de profils d'accès à un utilisateur Entra ID.
        Retire chirurgicalement les groupes obsolètes, ajoute les nouveaux,
        et conserve les groupes à l'intersection sans les toucher.

    .PARAMETER UserId
        ID Entra de l'utilisateur cible.

    .PARAMETER UPN
        UPN de l'utilisateur (pour le logging).

    .PARAMETER OldProfileKeys
        Clés des profils actuels à retirer. Vide pour un onboarding.

    .PARAMETER NewProfileKeys
        Clés des profils cibles à appliquer.

    .OUTPUTS
        [PSCustomObject] — {
            Success : bool
            Added   : int — Nombre de groupes ajoutés
            Removed : int — Nombre de groupes retirés
            Kept    : int — Nombre de groupes conservés (intersection)
            Errors  : string[] — Messages d'erreur éventuels
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$UPN,

        [string[]]$OldProfileKeys = @(),
        [string[]]$NewProfileKeys = @()
    )

    # 1. Calculer le diff intelligent
    $diff = Compare-AccessProfileGroups `
        -OldProfileKeys $OldProfileKeys `
        -NewProfileKeys $NewProfileKeys `
        -IncludeBaseline

    Write-Log -Level "INFO" -Action "PROFILE_CHANGE" -UPN $UPN `
        -Message "Changement de profil : [$($OldProfileKeys -join ', ')] → [$($NewProfileKeys -join ', ')] | +$($diff.ToAdd.Count) -$($diff.ToRemove.Count) =$($diff.ToKeep.Count)"

    $errors = @()

    # 2. Retirer les groupes obsolètes (ceux de l'ancien profil qui ne sont pas dans le nouveau)
    foreach ($grp in $diff.ToRemove) {
        try {
            Remove-MgGroupMemberByRef -GroupId $grp.id -DirectoryObjectId $UserId -ErrorAction Stop
            Write-Log -Level "SUCCESS" -Action "PROFILE_REMOVE_GROUP" -UPN $UPN `
                -Message "Retiré du groupe '$($grp.display_name)' ($($grp.id))"
        }
        catch {
            $errMsg = $_.Exception.Message
            # "Not found" ou "does not exist" = déjà retiré, pas une erreur bloquante
            if ($errMsg -like "*not found*" -or $errMsg -like "*does not exist*") {
                Write-Log -Level "WARNING" -Action "PROFILE_REMOVE_GROUP" -UPN $UPN `
                    -Message "Déjà absent du groupe '$($grp.display_name)' — ignoré."
            }
            else {
                $errors += "Retrait '$($grp.display_name)' : $errMsg"
                Write-Log -Level "ERROR" -Action "PROFILE_REMOVE_GROUP" -UPN $UPN `
                    -Message "Erreur retrait '$($grp.display_name)' : $errMsg"
            }
        }
    }

    # 3. Ajouter les nouveaux groupes
    foreach ($grp in $diff.ToAdd) {
        # Utiliser Add-AzUserToGroup qui gère déjà le "already exists" comme un succès
        $result = Add-AzUserToGroup -UserId $UserId -GroupName $grp.display_name
        if (-not $result.Success) {
            $errors += "Ajout '$($grp.display_name)' : $($result.Error)"
        }
    }

    # 4. Log récapitulatif
    $level = if ($errors.Count -eq 0) { "SUCCESS" } else { "WARNING" }
    Write-Log -Level $level -Action "PROFILE_CHANGE" -UPN $UPN `
        -Message "Résultat : +$($diff.ToAdd.Count) ajouté(s), -$($diff.ToRemove.Count) retiré(s), $($diff.ToKeep.Count) conservé(s). Erreurs: $($errors.Count)"

    return [PSCustomObject]@{
        Success = ($errors.Count -eq 0)
        Added   = $diff.ToAdd.Count
        Removed = $diff.ToRemove.Count
        Kept    = $diff.ToKeep.Count
        Errors  = $errors
    }
}

function Get-UserActiveProfiles {
    <#
    .SYNOPSIS
        Détecte quels profils d'accès sont actuellement actifs pour un utilisateur
        en comparant ses appartenances de groupes avec les définitions de profils.

    .DESCRIPTION
        Un profil est considéré "actif" si TOUS ses groupes sont présents dans
        les appartenances de l'utilisateur. Un profil sans groupes n'est jamais actif.

    .PARAMETER UserId
        ID Entra de l'utilisateur à analyser.

    .OUTPUTS
        [string[]] — Clés des profils actifs (ex: @("Finance", "TI")).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    try {
        # Récupérer toutes les appartenances de groupes de l'utilisateur
        $memberships = Get-MgUserMemberOf -UserId $UserId -ErrorAction Stop
        $userGroupIds = @(
            $memberships |
            Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' } |
            ForEach-Object { $_.Id }
        )

        $activeProfiles = @()
        $profiles = Get-AccessProfiles -ExcludeBaseline

        foreach ($ap in $profiles) {
                if ($ap.Groups.Count -eq 0) { continue }

            # Vérifier que TOUS les groupes du profil sont présents
            $allPresent = $true
            foreach ($grp in $ap.Groups) {
                if ($grp.id -notin $userGroupIds) {
                    $allPresent = $false
                    break
                }
            }

            if ($allPresent) {
                $activeProfiles += $ap.Key
            }
        }

        Write-Log -Level "INFO" -Action "DETECT_PROFILES" -UPN $UserId `
            -Message "Profils actifs détectés : [$($activeProfiles -join ', ')]"

        return $activeProfiles
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "DETECT_PROFILES" -UPN $UserId `
            -Message "Erreur détection profils : $errMsg"
        return @()
    }
}

# =============================================================================
#  RÉCONCILIATION — Fonctions à AJOUTER à la fin de AccessProfiles_Functions.ps1
#  (avant le commentaire "Point d'attention")
# =============================================================================

function Get-ProfileReconciliation {
    <#
    .SYNOPSIS
        Compare les utilisateurs associés à un profil avec sa définition actuelle.
        Identifie les écarts : groupes du template manquants chez certains users.

    .DESCRIPTION
        Stratégie batch (O(N) appels Graph, N = nb de groupes dans le profil) :
        1. Pour chaque groupe du profil, récupère ses membres via Get-MgGroupMember.
        2. Construit un dictionnaire userId → { UPN, DisplayName, groupes présents }.
        3. Retourne les users qui sont dans AU MOINS 1 groupe du profil
           mais à qui il MANQUE au moins 1 groupe.

        Cela détecte les users existants qui n'ont pas reçu un groupe
        récemment ajouté au template.

    .PARAMETER ProfileKey
        Clé technique du profil à analyser (ex: "Finance").

    .PARAMETER MinGroupThreshold
        Nombre minimum de groupes du profil que l'utilisateur doit posséder
        pour être considéré comme "associé" au profil. Par défaut : -1
        (= totalGroupes - 1, soit tolérance d'un seul groupe manquant).
        Mettre 1 pour attraper tous les users qui ont au moins 1 groupe du profil.

    .OUTPUTS
        [PSCustomObject] — {
            ProfileKey:      string
            ProfileName:     string
            TotalGroups:     int
            Discrepancies:   [PSCustomObject[]] — { UserId, UPN, DisplayName, Missing[] }
            TotalUsers:      int  (nombre d'utilisateurs analysés)
            Error:           string | $null
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProfileKey,

        [int]$MinGroupThreshold = -1
    )

    # Validation
    if (-not $Config.PSObject.Properties["access_profiles"]) {
        return [PSCustomObject]@{
            ProfileKey    = $ProfileKey
            ProfileName   = ""
            TotalGroups   = 0
            Discrepancies = @()
            TotalUsers    = 0
            Error         = "Aucune section access_profiles dans la configuration."
        }
    }

    if (-not $Config.access_profiles.PSObject.Properties[$ProfileKey]) {
        return [PSCustomObject]@{
            ProfileKey    = $ProfileKey
            ProfileName   = ""
            TotalGroups   = 0
            Discrepancies = @()
            TotalUsers    = 0
            Error         = "Profil '$ProfileKey' introuvable."
        }
    }

    $ap = $Config.access_profiles.$ProfileKey
    $profileGroups = @($ap.groups)
    $totalGroups   = $profileGroups.Count

    if ($totalGroups -eq 0) {
        return [PSCustomObject]@{
            ProfileKey    = $ProfileKey
            ProfileName   = $ap.display_name
            TotalGroups   = 0
            Discrepancies = @()
            TotalUsers    = 0
            Error         = $null
        }
    }

    # Seuil par défaut : l'utilisateur doit avoir au moins (N-1) groupes
    if ($MinGroupThreshold -le 0) {
        $MinGroupThreshold = [Math]::Max(1, $totalGroups - 1)
    }

    Write-Log -Level "INFO" -Action "RECONCILE_SCAN" `
        -Message "Scan réconciliation profil '$ProfileKey' ($totalGroups groupes, seuil=$MinGroupThreshold)"

    try {
        # Dictionnaire : userId → { UPN, DisplayName, PresentGroupIds = HashSet }
        $userMap = @{}

        foreach ($grp in $profileGroups) {
            try {
                $members = Get-MgGroupMember -GroupId $grp.id -All -ErrorAction Stop
                foreach ($member in $members) {
                    # Filtrer : uniquement les utilisateurs (pas les devices, groupes imbriqués, etc.)
                    if ($member.AdditionalProperties.'@odata.type' -ne '#microsoft.graph.user') { continue }

                    $uid = $member.Id
                    if (-not $userMap.ContainsKey($uid)) {
                        $userMap[$uid] = @{
                            UPN             = $member.AdditionalProperties.userPrincipalName
                            DisplayName     = $member.AdditionalProperties.displayName
                            PresentGroupIds = [System.Collections.Generic.HashSet[string]]::new()
                        }
                    }
                    [void]$userMap[$uid].PresentGroupIds.Add($grp.id)
                }
            }
            catch {
                Write-Log -Level "WARNING" -Action "RECONCILE_SCAN" `
                    -Message "Impossible de lire les membres du groupe '$($grp.display_name)' ($($grp.id)) : $($_.Exception.Message)"
            }
        }

        # Identifier les écarts
        $discrepancies = @()
        $profileGroupIds = @($profileGroups | ForEach-Object { $_.id })

        foreach ($uid in $userMap.Keys) {
            $entry = $userMap[$uid]
            $presentCount = $entry.PresentGroupIds.Count

            # L'utilisateur doit être dans au moins $MinGroupThreshold groupes
            # ET il doit lui manquer au moins 1 groupe
            if ($presentCount -ge $MinGroupThreshold -and $presentCount -lt $totalGroups) {
                $missingIds = @($profileGroupIds | Where-Object { -not $entry.PresentGroupIds.Contains($_) })
                $missingNames = @($missingIds | ForEach-Object {
                    $gid = $_
                    ($profileGroups | Where-Object { $_.id -eq $gid }).display_name
                })

                $discrepancies += [PSCustomObject]@{
                    UserId      = $uid
                    UPN         = $entry.UPN
                    DisplayName = $entry.DisplayName
                    Missing     = $missingNames
                    MissingIds  = $missingIds
                }
            }
        }

        # Trier par UPN pour lisibilité
        $discrepancies = @($discrepancies | Sort-Object -Property UPN)

        Write-Log -Level "INFO" -Action "RECONCILE_SCAN" `
            -Message "Résultat : $($userMap.Count) utilisateurs analysés, $($discrepancies.Count) écart(s) détecté(s)"

        return [PSCustomObject]@{
            ProfileKey    = $ProfileKey
            ProfileName   = $ap.display_name
            TotalGroups   = $totalGroups
            Discrepancies = $discrepancies
            TotalUsers    = $userMap.Count
            Error         = $null
        }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "RECONCILE_SCAN" -Message "Erreur réconciliation : $errMsg"
        return [PSCustomObject]@{
            ProfileKey    = $ProfileKey
            ProfileName   = $ap.display_name
            TotalGroups   = $totalGroups
            Discrepancies = @()
            TotalUsers    = 0
            Error         = $errMsg
        }
    }
}


function Invoke-ProfileReconciliation {
    <#
    .SYNOPSIS
        Applique la réconciliation : ajoute les groupes manquants aux utilisateurs
        identifiés par Get-ProfileReconciliation.

    .PARAMETER Discrepancies
        Tableau de PSCustomObject tel que retourné par Get-ProfileReconciliation.Discrepancies.
        Chaque objet contient : UserId, UPN, MissingIds, Missing (display_names).

    .PARAMETER ProfileKey
        Clé du profil (pour le logging).

    .PARAMETER OnProgress
        ScriptBlock optionnel appelé à chaque itération avec ($current, $total, $upn).

    .OUTPUTS
        [PSCustomObject] — {
            Success:   bool
            Applied:   int
            Failed:    int
            Errors:    string[]
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Discrepancies,

        [string]$ProfileKey = "?",

        [scriptblock]$OnProgress = $null
    )

    $errors  = @()
    $applied = 0
    $failed  = 0
    $total   = $Discrepancies.Count

    Write-Log -Level "INFO" -Action "RECONCILE_APPLY" `
        -Message "Début réconciliation profil '$ProfileKey' pour $total utilisateur(s)"

    for ($i = 0; $i -lt $total; $i++) {
        $disc = $Discrepancies[$i]

        # Callback de progression
        if ($OnProgress) {
            try { & $OnProgress ($i + 1) $total $disc.UPN } catch {}
        }

        $userErrors = @()
        foreach ($gid in $disc.MissingIds) {
            $gName = ($disc.Missing | Select-Object -Index ([Array]::IndexOf($disc.MissingIds, $gid)))
            if (-not $gName) { $gName = $gid }

            try {
                $body = @{
                    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($disc.UserId)"
                }
                New-MgGroupMemberByRef -GroupId $gid -BodyParameter $body -ErrorAction Stop

                Write-Log -Level "SUCCESS" -Action "RECONCILE_ADD_GROUP" -UPN $disc.UPN `
                    -Message "Ajouté au groupe '$gName' ($gid)"
            }
            catch {
                $errMsg = $_.Exception.Message
                # "already a member" n'est pas une vraie erreur
                if ($errMsg -like "*already exist*" -or $errMsg -like "*already a member*") {
                    Write-Log -Level "INFO" -Action "RECONCILE_ADD_GROUP" -UPN $disc.UPN `
                        -Message "Déjà membre du groupe '$gName' — ignoré"
                }
                else {
                    $userErrors += "$gName : $errMsg"
                    Write-Log -Level "ERROR" -Action "RECONCILE_ADD_GROUP" -UPN $disc.UPN `
                        -Message "Erreur ajout '$gName' : $errMsg"
                }
            }
        }

        if ($userErrors.Count -eq 0) {
            $applied++
        }
        else {
            $failed++
            $errors += "$($disc.UPN) : $($userErrors -join ' | ')"
        }
    }

    $level = if ($failed -eq 0) { "SUCCESS" } else { "WARNING" }
    Write-Log -Level $level -Action "RECONCILE_APPLY" `
        -Message "Réconciliation terminée : $applied réussi(s), $failed échoué(s) sur $total"

    return [PSCustomObject]@{
        Success = ($failed -eq 0)
        Applied = $applied
        Failed  = $failed
        Errors  = $errors
    }
}

# Point d'attention :
# - Write-Log écrit à la fois dans le fichier ET dans la console
# - New-SecurePassword exclut les caractères ambigus (0/O, 1/l/I)
# - Remove-Diacritics est essentiel pour les noms francophones (é, è, ê, ç, etc.)
# - Show-PasswordDialog est modal et force l'attention de l'utilisateur
# - Send-Notification utilise /me/sendMail — l'utilisateur connecté doit avoir une boîte mail
