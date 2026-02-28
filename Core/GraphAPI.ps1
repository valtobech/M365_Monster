<#
.FICHIER
    Core/GraphAPI.ps1

.ROLE
    Wrappers nommés sur les appels Microsoft Graph les plus utilisés.
    Chaque fonction encapsule un appel Graph, gère les erreurs et retourne
    un objet structuré {Success, Data, Error}.

.DEPENDANCES
    - Module Microsoft.Graph (PowerShell SDK)
    - Core/Functions.ps1 (Write-Log)
    - Connexion Graph active (via Core/Connect.ps1)

.AUTEUR
    [Équipe IT — GestionRH-AzureAD]
#>

function Get-AzUser {
    <#
    .SYNOPSIS
        Récupère un utilisateur Azure AD par son ID ou UPN.

    .PARAMETER UserId
        ID ou UserPrincipalName de l'utilisateur.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Data: object, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    try {
        Write-Log -Level "INFO" -Action "GET_USER" -UPN $UserId -Message "Récupération de l'utilisateur."
        $user = Get-MgUser -UserId $UserId -Property "id,displayName,userPrincipalName,accountEnabled,department,jobTitle,mobilePhone,mail,employeeType,assignedLicenses" -ErrorAction Stop
        return [PSCustomObject]@{ Success = $true; Data = $user; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "GET_USER" -UPN $UserId -Message "Erreur : $errMsg"
        return [PSCustomObject]@{ Success = $false; Data = $null; Error = $errMsg }
    }
}

function Search-AzUsers {
    <#
    .SYNOPSIS
        Recherche des utilisateurs Azure AD par prénom, nom, displayName ou adresse mail.

    .PARAMETER SearchTerm
        Terme de recherche — recherche partielle (contains) sur displayName, givenName, surname et mail.
        Exemples : "Martin", "ferrand", "kevin", "kevin.ferrand@domaine.com"

    .PARAMETER MaxResults
        Nombre maximum de résultats (défaut : 20).

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Data: array, Error: string}

    .NOTES
        Utilise $search (OData) avec ConsistencyLevel: eventual — permet la recherche partielle
        sur plusieurs champs simultanément, contrairement à startsWith qui exige le début exact.
        Syntaxe : "displayName:terme" OR "givenName:terme" OR "surname:terme" OR "mail:terme"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SearchTerm,

        [int]$MaxResults = 20
    )

    try {
        Write-Log -Level "INFO" -Action "SEARCH_USERS" -Message "Recherche d'utilisateurs : '$SearchTerm'"

        # $search avec ConsistencyLevel eventual — recherche partielle (contains) sur plusieurs champs
        # Couvre : "Martin" → trouve "Kevin Martin", "Martinez", "martin@..."
        # Couvre : "kevin" → trouve "Kevin Ferrand", "Kevin Martin"
        # Couvre : "ferrand" → trouve "Kevin Ferrand" même si displayName commence par "Kevin"
        $searchQuery = "`"displayName:$SearchTerm`" OR `"givenName:$SearchTerm`" OR `"surname:$SearchTerm`" OR `"mail:$SearchTerm`""
        $users = Get-MgUser `
            -Search $searchQuery `
            -Top $MaxResults `
            -Property "id,displayName,userPrincipalName,department,jobTitle,accountEnabled,mail,givenName,surname" `
            -ConsistencyLevel "eventual" `
            -CountVariable countVar `
            -ErrorAction Stop

        return [PSCustomObject]@{ Success = $true; Data = $users; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "SEARCH_USERS" -Message "Erreur recherche '$SearchTerm' : $errMsg"
        return [PSCustomObject]@{ Success = $false; Data = @(); Error = $errMsg }
    }
}

function New-AzUser {
    <#
    .SYNOPSIS
        Crée un nouvel utilisateur dans Azure AD.

    .PARAMETER UserParams
        Hashtable contenant les propriétés de l'utilisateur :
        - DisplayName, GivenName, Surname, UserPrincipalName, MailNickname
        - Password, Department, JobTitle, MobilePhone, EmployeeType
        - ForceChangePasswordNextSignIn

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Data: object, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$UserParams
    )

    $upn = $UserParams.UserPrincipalName

    try {
        Write-Log -Level "INFO" -Action "CREATE_USER" -UPN $upn -Message "Création de l'utilisateur."

        # Construction du body pour New-MgUser
        $passwordProfile = @{
            Password                      = $UserParams.Password
            ForceChangePasswordNextSignIn  = $UserParams.ForceChangePasswordNextSignIn
        }

        $params = @{
            DisplayName       = $UserParams.DisplayName
            GivenName         = $UserParams.GivenName
            Surname           = $UserParams.Surname
            UserPrincipalName = $UserParams.UserPrincipalName
            MailNickname      = $UserParams.MailNickname
            PasswordProfile   = $passwordProfile
            AccountEnabled    = $true
        }

        # UsageLocation est obligatoire pour l'assignation de licence
        if (-not [string]::IsNullOrWhiteSpace($UserParams.UsageLocation)) {
            $params.UsageLocation = $UserParams.UsageLocation
        }

        # Ajout des champs optionnels s'ils sont renseignés
        if (-not [string]::IsNullOrWhiteSpace($UserParams.Department)) {
            $params.Department = $UserParams.Department
        }
        if (-not [string]::IsNullOrWhiteSpace($UserParams.JobTitle)) {
            $params.JobTitle = $UserParams.JobTitle
        }
        if (-not [string]::IsNullOrWhiteSpace($UserParams.MobilePhone)) {
            $params.MobilePhone = $UserParams.MobilePhone
        }
        if (-not [string]::IsNullOrWhiteSpace($UserParams.EmployeeType)) {
            $params.EmployeeType = $UserParams.EmployeeType
        }
        if (-not [string]::IsNullOrWhiteSpace($UserParams.EmployeeId)) {
            $params.EmployeeId = $UserParams.EmployeeId
        }

        $newUser = New-MgUser -BodyParameter $params -ErrorAction Stop

        # Log de succès — ne PAS logger le mot de passe
        Write-Log -Level "SUCCESS" -Action "CREATE_USER" -UPN $upn -Message "Utilisateur créé avec succès. ID: $($newUser.Id)"

        return [PSCustomObject]@{ Success = $true; Data = $newUser; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "CREATE_USER" -UPN $upn -Message "Erreur création : $errMsg"
        return [PSCustomObject]@{ Success = $false; Data = $null; Error = $errMsg }
    }
}

function Set-AzUser {
    <#
    .SYNOPSIS
        Met à jour les attributs d'un utilisateur Azure AD.

    .PARAMETER UserId
        ID ou UPN de l'utilisateur à modifier.

    .PARAMETER Properties
        Hashtable des propriétés à mettre à jour.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [hashtable]$Properties
    )

    try {
        $propsStr = ($Properties.Keys -join ', ')
        Write-Log -Level "INFO" -Action "UPDATE_USER" -UPN $UserId -Message "Mise à jour des propriétés : $propsStr"

        Update-MgUser -UserId $UserId -BodyParameter $Properties -ErrorAction Stop

        Write-Log -Level "SUCCESS" -Action "UPDATE_USER" -UPN $UserId -Message "Propriétés mises à jour : $propsStr"
        return [PSCustomObject]@{ Success = $true; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "UPDATE_USER" -UPN $UserId -Message "Erreur mise à jour : $errMsg"
        return [PSCustomObject]@{ Success = $false; Error = $errMsg }
    }
}

function Disable-AzUser {
    <#
    .SYNOPSIS
        Désactive un compte utilisateur Azure AD.

    .PARAMETER UserId
        ID ou UPN de l'utilisateur.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    try {
        Write-Log -Level "INFO" -Action "DISABLE_USER" -UPN $UserId -Message "Désactivation du compte."

        Update-MgUser -UserId $UserId -BodyParameter @{ AccountEnabled = $false } -ErrorAction Stop

        Write-Log -Level "SUCCESS" -Action "DISABLE_USER" -UPN $UserId -Message "Compte désactivé."
        return [PSCustomObject]@{ Success = $true; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "DISABLE_USER" -UPN $UserId -Message "Erreur désactivation : $errMsg"
        return [PSCustomObject]@{ Success = $false; Error = $errMsg }
    }
}

function Enable-AzUser {
    <#
    .SYNOPSIS
        Réactive un compte utilisateur Azure AD.

    .PARAMETER UserId
        ID ou UPN de l'utilisateur.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    try {
        Write-Log -Level "INFO" -Action "ENABLE_USER" -UPN $UserId -Message "Réactivation du compte."

        Update-MgUser -UserId $UserId -BodyParameter @{ AccountEnabled = $true } -ErrorAction Stop

        Write-Log -Level "SUCCESS" -Action "ENABLE_USER" -UPN $UserId -Message "Compte réactivé."
        return [PSCustomObject]@{ Success = $true; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "ENABLE_USER" -UPN $UserId -Message "Erreur réactivation : $errMsg"
        return [PSCustomObject]@{ Success = $false; Error = $errMsg }
    }
}

function Add-AzUserToGroup {
    <#
    .SYNOPSIS
        Ajoute un utilisateur à un groupe Azure AD (par nom de groupe).

    .PARAMETER UserId
        ID de l'utilisateur.

    .PARAMETER GroupName
        Nom du groupe Azure AD.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$GroupName
    )

    try {
        Write-Log -Level "INFO" -Action "ADD_GROUP" -UPN $UserId -Message "Ajout au groupe '$GroupName'."

        # Recherche du groupe par nom
        $group = Get-MgGroup -Filter "displayName eq '$GroupName'" -ErrorAction Stop
        if ($null -eq $group) {
            throw "Groupe '$GroupName' introuvable dans Azure AD."
        }

        # CHOIX: Si plusieurs groupes portent le même nom, on prend le premier
        $groupId = $group.Id
        if ($groupId -is [array]) { $groupId = $groupId[0] }
        if ($group -is [array]) { $groupId = $group[0].Id }

        $bodyParam = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId"
        }

        New-MgGroupMemberByRef -GroupId $groupId -BodyParameter $bodyParam -ErrorAction Stop

        Write-Log -Level "SUCCESS" -Action "ADD_GROUP" -UPN $UserId -Message "Ajouté au groupe '$GroupName'."
        return [PSCustomObject]@{ Success = $true; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        # CHOIX: On ne considère pas "déjà membre" comme une erreur bloquante
        if ($errMsg -like "*already exist*" -or $errMsg -like "*One or more added object references already exist*") {
            Write-Log -Level "WARNING" -Action "ADD_GROUP" -UPN $UserId -Message "Déjà membre du groupe '$GroupName'."
            return [PSCustomObject]@{ Success = $true; Error = $null }
        }
        Write-Log -Level "ERROR" -Action "ADD_GROUP" -UPN $UserId -Message "Erreur ajout groupe '$GroupName' : $errMsg"
        return [PSCustomObject]@{ Success = $false; Error = $errMsg }
    }
}

function Remove-AzUserGroups {
    <#
    .SYNOPSIS
        Retire un utilisateur de tous ses groupes Azure AD.

    .PARAMETER UserId
        ID de l'utilisateur.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, RemovedCount: int, Errors: array, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    try {
        Write-Log -Level "INFO" -Action "REMOVE_GROUPS" -UPN $UserId -Message "Retrait de tous les groupes."

        $groupes = Get-MgUserMemberOf -UserId $UserId -ErrorAction Stop
        $erreurs = @()
        $compteur = 0

        foreach ($membre in $groupes) {
            # CHOIX: On ne retire que les groupes (pas les rôles d'annuaire)
            if ($membre.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group') {
                try {
                    Remove-MgGroupMemberByRef -GroupId $membre.Id -DirectoryObjectId $UserId -ErrorAction Stop
                    $compteur++
                }
                catch {
                    $erreurs += "Groupe $($membre.Id) : $($_.Exception.Message)"
                }
            }
        }

        $msg = "Retiré de $compteur groupe(s)."
        if ($erreurs.Count -gt 0) {
            $msg += " $($erreurs.Count) erreur(s)."
            Write-Log -Level "WARNING" -Action "REMOVE_GROUPS" -UPN $UserId -Message $msg
        }
        else {
            Write-Log -Level "SUCCESS" -Action "REMOVE_GROUPS" -UPN $UserId -Message $msg
        }

        return [PSCustomObject]@{
            Success      = ($erreurs.Count -eq 0)
            RemovedCount = $compteur
            Errors       = $erreurs
            Error        = if ($erreurs.Count -gt 0) { $erreurs -join "`n" } else { $null }
        }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "REMOVE_GROUPS" -UPN $UserId -Message "Erreur retrait groupes : $errMsg"
        return [PSCustomObject]@{ Success = $false; RemovedCount = 0; Errors = @($errMsg); Error = $errMsg }
    }
}

function Set-AzUserLicense {
    <#
    .SYNOPSIS
        Assigne une licence à un utilisateur Azure AD.

    .PARAMETER UserId
        ID de l'utilisateur.

    .PARAMETER SkuId
        SKU de la licence à assigner (format : TenantName:SKUPARTNAME).

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$SkuId
    )

    try {
        Write-Log -Level "INFO" -Action "ASSIGN_LICENSE" -UPN $UserId -Message "Attribution de la licence '$SkuId'."

        # Récupération du SKU ID réel depuis le nom
        $skuPartName = $SkuId.Split(':')[-1]
        $subscribedSkus = Get-MgSubscribedSku -ErrorAction Stop
        $sku = $subscribedSkus | Where-Object { $_.SkuPartNumber -eq $skuPartName }

        if ($null -eq $sku) {
            throw "SKU '$skuPartName' introuvable dans le tenant. Licences disponibles : $(($subscribedSkus | ForEach-Object { $_.SkuPartNumber }) -join ', ')"
        }

        $params = @{
            AddLicenses    = @(
                @{
                    SkuId = $sku.SkuId
                }
            )
            RemoveLicenses = @()
        }

        Set-MgUserLicense -UserId $UserId -BodyParameter $params -ErrorAction Stop

        Write-Log -Level "SUCCESS" -Action "ASSIGN_LICENSE" -UPN $UserId -Message "Licence '$SkuId' attribuée."
        return [PSCustomObject]@{ Success = $true; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "ASSIGN_LICENSE" -UPN $UserId -Message "Erreur licence : $errMsg"
        return [PSCustomObject]@{ Success = $false; Error = $errMsg }
    }
}

function Remove-AzUserLicenses {
    <#
    .SYNOPSIS
        Révoque toutes les licences d'un utilisateur Azure AD.

    .PARAMETER UserId
        ID de l'utilisateur.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, RemovedCount: int, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    try {
        Write-Log -Level "INFO" -Action "REVOKE_LICENSES" -UPN $UserId -Message "Révocation de toutes les licences."

        $user = Get-MgUser -UserId $UserId -Property "assignedLicenses" -ErrorAction Stop
        $licences = $user.AssignedLicenses

        if ($null -eq $licences -or $licences.Count -eq 0) {
            Write-Log -Level "INFO" -Action "REVOKE_LICENSES" -UPN $UserId -Message "Aucune licence à révoquer."
            return [PSCustomObject]@{ Success = $true; RemovedCount = 0; Error = $null }
        }

        $skuIds = $licences | ForEach-Object { $_.SkuId }

        $params = @{
            AddLicenses    = @()
            RemoveLicenses = $skuIds
        }

        Set-MgUserLicense -UserId $UserId -BodyParameter $params -ErrorAction Stop

        $count = $skuIds.Count
        Write-Log -Level "SUCCESS" -Action "REVOKE_LICENSES" -UPN $UserId -Message "$count licence(s) révoquée(s)."
        return [PSCustomObject]@{ Success = $true; RemovedCount = $count; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "REVOKE_LICENSES" -UPN $UserId -Message "Erreur révocation licences : $errMsg"
        return [PSCustomObject]@{ Success = $false; RemovedCount = 0; Error = $errMsg }
    }
}

function Revoke-AzUserSessions {
    <#
    .SYNOPSIS
        Révoque toutes les sessions actives d'un utilisateur Azure AD.

    .PARAMETER UserId
        ID ou UPN de l'utilisateur.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    try {
        Write-Log -Level "INFO" -Action "REVOKE_SESSIONS" -UPN $UserId -Message "Révocation des sessions actives."

        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$UserId/revokeSignInSessions" -ErrorAction Stop

        Write-Log -Level "SUCCESS" -Action "REVOKE_SESSIONS" -UPN $UserId -Message "Sessions révoquées."
        return [PSCustomObject]@{ Success = $true; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "REVOKE_SESSIONS" -UPN $UserId -Message "Erreur révocation sessions : $errMsg"
        return [PSCustomObject]@{ Success = $false; Error = $errMsg }
    }
}

function Set-AzUserManager {
    <#
    .SYNOPSIS
        Définit le manager d'un utilisateur Azure AD.

    .PARAMETER UserId
        ID de l'utilisateur.

    .PARAMETER ManagerId
        ID du manager.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$ManagerId
    )

    try {
        Write-Log -Level "INFO" -Action "SET_MANAGER" -UPN $UserId -Message "Attribution du manager : $ManagerId"

        $bodyParam = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/users/$ManagerId"
        }

        Set-MgUserManagerByRef -UserId $UserId -BodyParameter $bodyParam -ErrorAction Stop

        Write-Log -Level "SUCCESS" -Action "SET_MANAGER" -UPN $UserId -Message "Manager défini : $ManagerId"
        return [PSCustomObject]@{ Success = $true; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "SET_MANAGER" -UPN $UserId -Message "Erreur attribution manager : $errMsg"
        return [PSCustomObject]@{ Success = $false; Error = $errMsg }
    }
}

function Reset-AzUserPassword {
    <#
    .SYNOPSIS
        Réinitialise le mot de passe d'un utilisateur Azure AD.

    .PARAMETER UserId
        ID ou UPN de l'utilisateur.

    .PARAMETER NewPassword
        Nouveau mot de passe.

    .PARAMETER ForceChange
        Force le changement au prochain login.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [SecureString]$NewPassword,

        [bool]$ForceChange = $true
    )

    try {
        # Ne PAS logger le mot de passe — conversion SecureString -> plaintext uniquement pour l'appel Graph
        Write-Log -Level "INFO" -Action "RESET_PASSWORD" -UPN $UserId -Message "Réinitialisation du mot de passe."

        $plainPassword   = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($NewPassword)
        )
        $passwordProfile = @{
            Password                      = $plainPassword
            ForceChangePasswordNextSignIn  = $ForceChange
        }

        Update-MgUser -UserId $UserId -PasswordProfile $passwordProfile -ErrorAction Stop

        Write-Log -Level "SUCCESS" -Action "RESET_PASSWORD" -UPN $UserId -Message "Mot de passe réinitialisé."
        return [PSCustomObject]@{ Success = $true; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "RESET_PASSWORD" -UPN $UserId -Message "Erreur reset mot de passe : $errMsg"
        return [PSCustomObject]@{ Success = $false; Error = $errMsg }
    }
}

function Get-AzUserManager {
    <#
    .SYNOPSIS
        Récupère le manager d'un utilisateur Azure AD.

    .PARAMETER UserId
        ID ou UPN de l'utilisateur.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Data: object, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    try {
        $manager = Get-MgUserManager -UserId $UserId -ErrorAction Stop
        return [PSCustomObject]@{ Success = $true; Data = $manager; Error = $null }
    }
    catch {
        # Pas de manager n'est pas forcément une erreur
        if ($_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*Resource*Not Found*") {
            return [PSCustomObject]@{ Success = $true; Data = $null; Error = $null }
        }
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "GET_MANAGER" -UPN $UserId -Message "Erreur : $errMsg"
        return [PSCustomObject]@{ Success = $false; Data = $null; Error = $errMsg }
    }
}

# Point d'attention :
# - Chaque fonction retourne TOUJOURS un PSCustomObject avec au minimum Success et Error
# - Les mots de passe ne sont JAMAIS inclus dans les logs
# - Les erreurs "déjà membre" lors d'ajout à un groupe sont traitées comme des succès

function Test-AzUserExists {
    <#
    .SYNOPSIS
        Vérifie si un UserPrincipalName existe déjà dans Azure AD.

    .PARAMETER UPN
        UserPrincipalName à vérifier.

    .OUTPUTS
        [bool] — $true si l'UPN existe, $false sinon.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UPN
    )

    try {
        $null = Get-MgUser -UserId $UPN -Property "id" -ErrorAction Stop
        return $true
    }
    catch {
        # 404 = n'existe pas = c'est ce qu'on veut
        return $false
    }
}

function Get-AzDistinctValues {
    <#
    .SYNOPSIS
        Récupère les valeurs distinctes d'un attribut utilisateur dans le tenant.
        Utilisé pour alimenter dynamiquement les listes déroulantes
        (department, jobTitle, employeeType, usageLocation).

    .PARAMETER Property
        Nom de la propriété Graph à récupérer.

    .PARAMETER MaxUsers
        Nombre maximum d'utilisateurs à scanner (défaut : 999).

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Data: string[], Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("department", "jobTitle", "employeeType", "usageLocation", "officeLocation")]
        [string]$Property,

        [int]$MaxUsers = 999
    )

    try {
        Write-Log -Level "INFO" -Action "GET_DISTINCT" -Message "Récupération des valeurs distinctes pour '$Property'."

        # Certaines propriétés (employeeType, usageLocation) ne supportent pas $filter dans Graph.
        # On récupère les utilisateurs avec la propriété demandée, sans filtre serveur.
        # Le filtrage des valeurs non-vides se fait côté client.
        $propertiesFilter = @("department", "jobTitle")

        if ($Property -in $propertiesFilter) {
            # Ces propriétés supportent le $filter côté serveur
            $filter = "$Property ne null"
            $users = Get-MgUser -Filter $filter -Top $MaxUsers -Property $Property -ConsistencyLevel "eventual" -CountVariable countVar -ErrorAction Stop
        }
        else {
            # Fallback : récupérer tous les users avec la propriété, filtrage côté client
            $users = Get-MgUser -Top $MaxUsers -Property $Property -ErrorAction Stop
        }

        # Extraire les valeurs distinctes et trier
        $values = @()
        foreach ($user in $users) {
            $val = $user.$Property
            if (-not [string]::IsNullOrWhiteSpace($val) -and $val -notin $values) {
                $values += $val
            }
        }
        $values = $values | Sort-Object

        Write-Log -Level "INFO" -Action "GET_DISTINCT" -Message "'$Property' : $($values.Count) valeur(s) distincte(s) trouvée(s)."
        return [PSCustomObject]@{ Success = $true; Data = $values; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "GET_DISTINCT" -Message "Erreur pour '$Property' : $errMsg"
        return [PSCustomObject]@{ Success = $false; Data = @(); Error = $errMsg }
    }
}

function Search-AzGroups {
    <#
    .SYNOPSIS
        Recherche des groupes Entra ID par nom (recherche partielle).
        Utilisé par l'éditeur de profils d'accès pour résoudre les groupes
        sans saisie manuelle de GUID.

    .PARAMETER SearchTerm
        Terme de recherche (partiel) sur le displayName du groupe.

    .PARAMETER MaxResults
        Nombre maximum de résultats retournés (défaut : 20).

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Data: array, Error: string}
        Data contient des objets avec : Id, DisplayName, GroupTypes, MailEnabled, SecurityEnabled

    .EXAMPLE
        Search-AzGroups -SearchTerm "Finance"
        # Retourne tous les groupes dont le nom contient "Finance"

    .NOTES
        Utilise $search avec ConsistencyLevel: eventual pour la recherche partielle.
        Identique au pattern de Search-AzUsers dans GraphAPI.ps1.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SearchTerm,

        [int]$MaxResults = 20
    )

    try {
        Write-Log -Level "INFO" -Action "SEARCH_GROUPS" -Message "Recherche de groupes : '$SearchTerm'"

        # $search avec ConsistencyLevel eventual — recherche partielle (contains)
        $groups = Get-MgGroup `
            -Search "`"displayName:$SearchTerm`"" `
            -Top $MaxResults `
            -Property "id,displayName,groupTypes,mailEnabled,securityEnabled,description" `
            -ConsistencyLevel "eventual" `
            -CountVariable countVar `
            -ErrorAction Stop

        Write-Log -Level "INFO" -Action "SEARCH_GROUPS" -Message "$($groups.Count) groupe(s) trouvé(s) pour '$SearchTerm'."

        return [PSCustomObject]@{ Success = $true; Data = $groups; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "SEARCH_GROUPS" -Message "Erreur recherche groupes '$SearchTerm' : $errMsg"
        return [PSCustomObject]@{ Success = $false; Data = @(); Error = $errMsg }
    }
}

function Remove-AzUserFromGroup {
    <#
    .SYNOPSIS
        Retire un utilisateur d'un groupe Entra ID spécifique (par ID de groupe).
        Complément à Add-AzUserToGroup pour la gestion fine des profils d'accès.

    .PARAMETER UserId
        ID Entra de l'utilisateur.

    .PARAMETER GroupId
        ID du groupe Entra ID.

    .PARAMETER GroupName
        Nom du groupe (pour le logging uniquement).

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [string]$GroupName = ""
    )

    $label = if ($GroupName) { "'$GroupName' ($GroupId)" } else { $GroupId }

    try {
        Write-Log -Level "INFO" -Action "REMOVE_FROM_GROUP" -UPN $UserId -Message "Retrait du groupe $label."

        Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $UserId -ErrorAction Stop

        Write-Log -Level "SUCCESS" -Action "REMOVE_FROM_GROUP" -UPN $UserId -Message "Retiré du groupe $label."
        return [PSCustomObject]@{ Success = $true; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        # "Not a member" ou "does not exist" = déjà retiré
        if ($errMsg -like "*not found*" -or $errMsg -like "*does not exist*") {
            Write-Log -Level "WARNING" -Action "REMOVE_FROM_GROUP" -UPN $UserId -Message "Déjà absent du groupe $label — ignoré."
            return [PSCustomObject]@{ Success = $true; Error = $null }
        }
        Write-Log -Level "ERROR" -Action "REMOVE_FROM_GROUP" -UPN $UserId -Message "Erreur retrait groupe $label : $errMsg"
        return [PSCustomObject]@{ Success = $false; Error = $errMsg }
    }
}






