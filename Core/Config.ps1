<#
.FICHIER
    Core/Config.ps1

.ROLE
    Chargement et validation du fichier de configuration client JSON.
    Expose la variable globale $Config utilisée par tous les autres modules.

.DEPENDANCES
    Aucune (premier module chargé)

.AUTEUR
    [Équipe IT — GestionRH-AzureAD]
#>

function Load-ClientConfig {
    <#
    .SYNOPSIS
        Charge et valide un fichier de configuration client JSON.

    .PARAMETER ConfigPath
        Chemin complet vers le fichier .json du client.

    .OUTPUTS
        [PSCustomObject] — Objet de configuration parsé et validé.
        Lève une exception si la validation échoue.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )

    # Vérification de l'existence du fichier
    if (-not (Test-Path -Path $ConfigPath)) {
        throw "Le fichier de configuration '$ConfigPath' est introuvable."
    }

    # Lecture et parsing du JSON
    try {
        $jsonContent = Get-Content -Path $ConfigPath -Raw -Encoding UTF8
        $configObject = $jsonContent | ConvertFrom-Json
    }
    catch {
        throw "Erreur lors du parsing JSON du fichier '$ConfigPath' : $($_.Exception.Message)"
    }

    # Validation des champs obligatoires
    $champsObligatoires = @(
        'client_name',
        'tenant_id',
        'client_id',
        'auth_method',
        'smtp_domain',
        'license_groups',
        'membership_groups',
        'offboarding',
        'notifications',
        'password_policy'
    )

    foreach ($champ in $champsObligatoires) {
        if (-not ($configObject.PSObject.Properties.Name -contains $champ)) {
            throw "Champ obligatoire manquant dans la configuration : '$champ'"
        }
    }

    # Validation du format tenant_id (GUID)
    $guidRegex = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
    if ($configObject.tenant_id -notmatch $guidRegex) {
        throw "Le tenant_id n'est pas un GUID valide : '$($configObject.tenant_id)'"
    }
    if ($configObject.client_id -notmatch $guidRegex) {
        throw "Le client_id n'est pas un GUID valide : '$($configObject.client_id)'"
    }

    # Validation de auth_method
    $methodesAutorisees = @('interactive_browser', 'device_code', 'client_secret')
    if ($configObject.auth_method -notin $methodesAutorisees) {
        throw "Méthode d'authentification invalide : '$($configObject.auth_method)'. Valeurs acceptées : $($methodesAutorisees -join ', ')"
    }

    # Validation de la politique de mot de passe
    if ($configObject.password_policy.length -lt 8) {
        throw "La longueur minimale du mot de passe doit être d'au moins 8 caractères."
    }

    # Validation offboarding
    $offboardingChamps = @('disabled_ou_group', 'revoke_licenses', 'remove_all_groups', 'retention_days')
    foreach ($champ in $offboardingChamps) {
        if (-not ($configObject.offboarding.PSObject.Properties.Name -contains $champ)) {
            throw "Champ obligatoire manquant dans offboarding : '$champ'"
        }
    }

    # Validation smtp_domain commence par @
    if (-not $configObject.smtp_domain.StartsWith('@')) {
        throw "Le smtp_domain doit commencer par '@'. Valeur actuelle : '$($configObject.smtp_domain)'"
    }

    # CHOIX: On ajoute le chemin du fichier source dans l'objet config pour référence
    $configObject | Add-Member -NotePropertyName '_config_path' -NotePropertyValue $ConfigPath -Force

    return $configObject
}

function Get-ClientList {
    <#
    .SYNOPSIS
        Retourne la liste des fichiers de configuration client disponibles.

    .PARAMETER ClientsFolder
        Chemin vers le dossier Clients/.

    .OUTPUTS
        [Array] — Liste d'objets avec le nom du client et le chemin du fichier.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientsFolder
    )

    if (-not (Test-Path -Path $ClientsFolder)) {
        throw "Le dossier Clients '$ClientsFolder' est introuvable."
    }

    $fichiers = Get-ChildItem -Path $ClientsFolder -Filter "*.json" | Where-Object { $_.Name -ne '_Template.json' }

    $listeClients = @()
    foreach ($fichier in $fichiers) {
        try {
            $json = Get-Content -Path $fichier.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            $listeClients += [PSCustomObject]@{
                Name     = $json.client_name
                FileName = $fichier.Name
                FullPath = $fichier.FullName
            }
        }
        catch {
            # Fichier JSON invalide — on l'ignore mais on logue un avertissement
            Write-Warning "Impossible de lire le fichier client '$($fichier.Name)' : $($_.Exception.Message)"
        }
    }

    return $listeClients
}

# Point d'attention :
# - La variable $Config doit être assignée dans le scope appelant via dot-sourcing
#   Exemple dans Main.ps1 : $Config = Load-ClientConfig -ConfigPath $cheminChoisi
# - Le _Template.json est exclu de la liste des clients disponibles
