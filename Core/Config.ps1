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

function Invoke-ConfigMigration {
    <#
    .SYNOPSIS
        Migre un fichier de configuration client depuis l'ancien format vers le nouveau.
        Détecte license_groups / membership_groups et les remplace par license_group_prefix.
        Sauvegarde automatiquement le fichier si des changements sont effectués.

    .PARAMETER ConfigObject
        Objet de configuration parsé depuis le JSON.

    .PARAMETER ConfigPath
        Chemin du fichier JSON pour la sauvegarde automatique.

    .OUTPUTS
        [PSCustomObject] — Objet de configuration migré.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $ConfigObject,

        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )

    $migrated = $false

    # Migration license_groups → license_group_prefix
    if ($ConfigObject.PSObject.Properties["license_groups"] -and -not $ConfigObject.PSObject.Properties["license_group_prefix"]) {
        # Tenter d'extraire un préfixe commun des noms de groupes existants
        $prefix = "LIC_"
        $existingGroups = @($ConfigObject.license_groups)
        if ($existingGroups.Count -gt 0) {
            # Chercher un préfixe commun parmi les groupes (ex: LIC-, LIC_, License_)
            $firstGroup = $existingGroups[0]
            $separators = @('_', '-')
            foreach ($sep in $separators) {
                $parts = $firstGroup.Split($sep)
                if ($parts.Count -ge 2) {
                    $candidate = "$($parts[0])$sep"
                    $allMatch = -not ($existingGroups | Where-Object { -not $_.StartsWith($candidate) })
                    if ($allMatch) {
                        $prefix = $candidate
                        break
                    }
                }
            }
        }

        $ConfigObject | Add-Member -NotePropertyName 'license_group_prefix' -NotePropertyValue $prefix -Force
        $ConfigObject.PSObject.Properties.Remove('license_groups')
        $migrated = $true
        Write-Warning "Migration config : license_groups → license_group_prefix = '$prefix'"
    }

    # Suppression de membership_groups (plus utilisé — recherche dynamique)
    if ($ConfigObject.PSObject.Properties["membership_groups"]) {
        $ConfigObject.PSObject.Properties.Remove('membership_groups')
        $migrated = $true
        Write-Warning "Migration config : membership_groups supprimé (recherche dynamique activée)"
    }

    # Sauvegarde automatique si migration effectuée
    if ($migrated) {
        try {
            $jsonContent = $ConfigObject | ConvertTo-Json -Depth 10
            $jsonContent | Out-File -FilePath $ConfigPath -Encoding UTF8 -Force
            Write-Warning "Migration config : fichier sauvegardé automatiquement → $ConfigPath"
        }
        catch {
            Write-Warning "Migration config : impossible de sauvegarder le fichier migré : $($_.Exception.Message)"
        }
    }

    return $ConfigObject
}

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

    # Migration automatique des anciens formats de configuration
    $configObject = Invoke-ConfigMigration -ConfigObject $configObject -ConfigPath $ConfigPath

    # Validation des champs obligatoires
    $champsObligatoires = @(
        'client_name',
        'tenant_id',
        'client_id',
        'auth_method',
        'smtp_domain',
        'offboarding',
        'notifications',
        'password_policy'
    )

    foreach ($champ in $champsObligatoires) {
        if (-not $configObject.PSObject.Properties[$champ]) {
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
        if (-not $configObject.offboarding.PSObject.Properties[$champ]) {
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