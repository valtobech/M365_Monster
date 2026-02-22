<#
.FICHIER
    Core/Lang.ps1

.ROLE
    Système d'internationalisation (i18n) de M365 Monster.
    Charge le fichier de langue JSON correspondant au choix de l'utilisateur.
    Fournit la fonction Get-Text pour accéder aux chaînes traduites.

    Le choix de langue est stocké dans settings.json à la racine du projet.
    Au premier lancement, une popup permet de choisir la langue.

.DEPENDANCES
    - Lang/fr.json, Lang/en.json
    - settings.json (créé automatiquement)

.AUTEUR
    [Equipe IT - M365 Monster]
#>

# Variable globale contenant toutes les chaînes de la langue active
$global:M365Strings = $null
$global:M365CurrentLang = $null

function Get-AvailableLanguages {
    <#
    .SYNOPSIS
        Liste les langues disponibles dans le dossier Lang/.
    .OUTPUTS
        [Array] — Objets {Code, Name, FilePath}
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $langFolder = Join-Path -Path $RootPath -ChildPath "Lang"
    if (-not (Test-Path $langFolder)) { return @() }

    $languages = @()
    Get-ChildItem -Path $langFolder -Filter "*.json" | ForEach-Object {
        try {
            # Essai 1 : UTF8 explicite
            $raw = Get-Content -Path $_.FullName -Raw -Encoding UTF8
            # Nettoyer le BOM s'il est présent en tant que caractère
            if ($raw.Length -gt 0 -and [int]$raw[0] -eq 65279) {
                $raw = $raw.Substring(1)
            }
            $content = $raw | ConvertFrom-Json
            if ($content._code -and $content._language) {
                $languages += [PSCustomObject]@{
                    Code     = $content._code
                    Name     = $content._language
                    FilePath = $_.FullName
                }
            }
        }
        catch {
            # Essai 2 : encodage par défaut
            try {
                $raw = Get-Content -Path $_.FullName -Raw
                $content = $raw | ConvertFrom-Json
                if ($content._code -and $content._language) {
                    $languages += [PSCustomObject]@{
                        Code     = $content._code
                        Name     = $content._language
                        FilePath = $_.FullName
                    }
                }
            }
            catch {
                Write-Warning "Cannot parse language file: $($_.Exception.Message)"
            }
        }
    }
    return $languages
}

function Get-SavedLanguage {
    <#
    .SYNOPSIS
        Lit le code langue sauvegardé dans settings.json.
    .OUTPUTS
        [string] — Code langue (ex: "fr", "en") ou $null si pas encore choisi.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $settingsFile = Join-Path -Path $RootPath -ChildPath "settings.json"
    if (-not (Test-Path $settingsFile)) { return $null }

    try {
        $settings = Get-Content -Path $settingsFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if (-not [string]::IsNullOrWhiteSpace($settings.language)) {
            return $settings.language
        }
    }
    catch {}

    return $null
}

function Save-LanguageChoice {
    <#
    .SYNOPSIS
        Sauvegarde le code langue dans settings.json.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $true)]
        [string]$LangCode
    )

    $settingsFile = Join-Path -Path $RootPath -ChildPath "settings.json"

    # Charger le fichier existant ou créer un nouvel objet
    $settings = $null
    if (Test-Path $settingsFile) {
        try {
            $settings = Get-Content -Path $settingsFile -Raw -Encoding UTF8 | ConvertFrom-Json
        }
        catch {}
    }

    if ($null -eq $settings) {
        $settings = [PSCustomObject]@{ language = $LangCode }
    }
    else {
        # Ajouter ou mettre à jour la propriété language
        if ($settings.PSObject.Properties["language"]) {
            $settings.language = $LangCode
        }
        else {
            $settings | Add-Member -NotePropertyName "language" -NotePropertyValue $LangCode
        }
    }

    $settings | ConvertTo-Json -Depth 3 | Out-File -FilePath $settingsFile -Encoding UTF8 -Force
}

function Show-LanguageSelector {
    <#
    .SYNOPSIS
        Affiche une popup de sélection de langue au premier lancement.
        Texte bilingue (neutre) pour que l'utilisateur puisse comprendre
        quelle que soit sa langue.
    .OUTPUTS
        [string] — Code langue sélectionné, ou $null si annulé.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $languages = Get-AvailableLanguages -RootPath $RootPath
    if ($languages.Count -eq 0) { return $null }
    if ($languages.Count -eq 1) { return $languages[0].Code }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "M365 Monster — Langue / Language"
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $form.TopMost = $true

    # Icône si disponible
    $iconFile = Join-Path -Path $RootPath -ChildPath "Assets\M365Monster.ico"
    if (Test-Path $iconFile) {
        $form.Icon = New-Object System.Drawing.Icon($iconFile)
    }

    # Titre bilingue
    $lblTitre = New-Object System.Windows.Forms.Label
    $lblTitre.Text = "Choisissez votre langue`nChoose your language"
    $lblTitre.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitre.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $lblTitre.Location = New-Object System.Drawing.Point(20, 20)
    $lblTitre.Size = New-Object System.Drawing.Size(370, 60)
    $lblTitre.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($lblTitre)

    # Boutons radio pour chaque langue
    $yPos = 95
    $radioButtons = @()

    foreach ($lang in $languages) {
        $radio = New-Object System.Windows.Forms.RadioButton
        $radio.Text = "$($lang.Name)"
        $radio.Font = New-Object System.Drawing.Font("Segoe UI", 12)
        $radio.Location = New-Object System.Drawing.Point(100, $yPos)
        $radio.Size = New-Object System.Drawing.Size(220, 30)
        $radio.Tag = $lang.Code
        $form.Controls.Add($radio)
        $radioButtons += $radio
        $yPos += 35
    }

    # Sélectionner le premier par défaut
    if ($radioButtons.Count -gt 0) { $radioButtons[0].Checked = $true }

    # Espacement avant le bouton
    $btnY = $yPos + 15

    # Bouton confirmer (bilingue)
    $btnConfirm = New-Object System.Windows.Forms.Button
    $btnConfirm.Text = "Confirmer / Confirm"
    $btnConfirm.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnConfirm.Location = New-Object System.Drawing.Point(100, $btnY)
    $btnConfirm.Size = New-Object System.Drawing.Size(220, 40)
    $btnConfirm.BackColor = [System.Drawing.Color]::FromArgb(111, 66, 193)
    $btnConfirm.ForeColor = [System.Drawing.Color]::White
    $btnConfirm.FlatStyle = "Flat"
    $btnConfirm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnConfirm)
    $form.AcceptButton = $btnConfirm

    # Ajuster la hauteur du formulaire — marge suffisante sous le bouton
    $formHeight = $btnY + 40 + 55
    $form.ClientSize = New-Object System.Drawing.Size(400, $formHeight)

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $selected = $radioButtons | Where-Object { $_.Checked } | Select-Object -First 1
        $form.Dispose()
        if ($selected) { return $selected.Tag }
    }

    $form.Dispose()
    return $null
}

function Initialize-Language {
    <#
    .SYNOPSIS
        Initialise le système de langue.
        Vérifie settings.json, sinon affiche la popup de sélection.
        Charge le fichier JSON de la langue choisie dans $global:M365Strings.
    .PARAMETER RootPath
        Répertoire racine de l'application (contient Lang/).
    .PARAMETER UserDataPath
        Répertoire des données utilisateur (contient settings.json).
        Si non spécifié, utilise RootPath (retro-compatibilité).
    .OUTPUTS
        [bool] — $true si initialisé avec succès, $false si annulé.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $false)]
        [string]$UserDataPath
    )

    # Si pas de UserDataPath, fallback sur RootPath (retro-compatibilité)
    if ([string]::IsNullOrWhiteSpace($UserDataPath)) {
        $UserDataPath = $RootPath
    }

    # 1. Vérifier si une langue est déjà sauvegardée
    $langCode = Get-SavedLanguage -RootPath $UserDataPath

    # 2. Sinon, afficher la popup de sélection
    if ([string]::IsNullOrWhiteSpace($langCode)) {
        # Vérifier d'abord que le dossier Lang/ existe et contient des fichiers
        $langFolder = Join-Path -Path $RootPath -ChildPath "Lang"
        if (-not (Test-Path $langFolder)) {
            Write-Warning "Lang folder not found: $langFolder"
            return $false
        }

        $langCode = Show-LanguageSelector -RootPath $RootPath
        if ([string]::IsNullOrWhiteSpace($langCode)) {
            return $false
        }
        # Sauvegarder le choix dans UserDataPath
        Save-LanguageChoice -RootPath $UserDataPath -LangCode $langCode
    }

    # 3. Charger le fichier de langue depuis RootPath
    $langFile = Join-Path -Path $RootPath -ChildPath "Lang\$langCode.json"
    if (-not (Test-Path $langFile)) {
        # Fallback vers fr.json
        $langFile = Join-Path -Path $RootPath -ChildPath "Lang\fr.json"
        if (-not (Test-Path $langFile)) {
            return $false
        }
    }

    try {
        $raw = Get-Content -Path $langFile -Raw -Encoding UTF8
        # Nettoyer le BOM si present (caractere Unicode U+FEFF)
        if ($raw.Length -gt 0 -and [int]$raw[0] -eq 65279) {
            $raw = $raw.Substring(1)
        }
        $global:M365Strings = $raw | ConvertFrom-Json
        $global:M365CurrentLang = $langCode
        return $true
    }
    catch {
        Write-Warning "Failed to parse language file: $($_.Exception.Message)"
        return $false
    }
}

function Get-Text {
    <#
    .SYNOPSIS
        Récupère une chaîne traduite par son chemin (notation pointée).

    .PARAMETER Key
        Chemin vers la chaîne, ex: "main_menu.tile_onboarding"

    .PARAMETER Params
        Paramètres de remplacement pour les placeholders {0}, {1}, etc.

    .EXAMPLE
        Get-Text "onboarding.title" $Config.client_name
        # Retourne : "Onboarding - Nouvel employé - ClientA"

    .OUTPUTS
        [string] — La chaîne traduite, ou la clé elle-même si non trouvée.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Key,

        [Parameter(ValueFromRemainingArguments = $true)]
        [object[]]$Params
    )

    if ($null -eq $global:M365Strings) { return $Key }

    # Naviguer dans l'objet JSON par notation pointée
    $parts = $Key.Split(".")
    $current = $global:M365Strings

    foreach ($part in $parts) {
        if ($null -eq $current) { return $Key }
        if ($current.PSObject.Properties[$part]) {
            $current = $current.$part
        }
        else {
            return $Key
        }
    }

    # Résultat doit être une chaîne
    if ($current -isnot [string]) { return $Key }

    # Remplacement des placeholders {0}, {1}, etc.
    if ($Params -and $Params.Count -gt 0) {
        try {
            $current = $current -f $Params
        }
        catch {
            # En cas d'erreur de formatage, retourner tel quel
        }
    }

    return $current
}
