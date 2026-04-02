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
        if (-not [string]::IsNullOrWhiteSpace($UserParams.OfficeLocation)) {
            $params.OfficeLocation = $UserParams.OfficeLocation
        }
        if (-not [string]::IsNullOrWhiteSpace($UserParams.CompanyName)) {
            $params.CompanyName = $UserParams.CompanyName
        }
        if (-not [string]::IsNullOrWhiteSpace($UserParams.EmployeeHireDate)) {
            $params.EmployeeHireDate = $UserParams.EmployeeHireDate
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
        Remplace le jobTitle par le format :
        "DISABLE - A supprimer le JJ/MM/AAAA | Titre d'origine"
        (date = aujourd'hui + 3 mois)
        Ceci déclenche aussi l'exclusion des groupes dynamiques dont la règle
        contient (user.jobTitle -notContains "DISABLE").

    .PARAMETER UserId
        ID ou UPN de l'utilisateur.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Error: string, OriginalJobTitle: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    try {
        Write-Log -Level "INFO" -Action "DISABLE_USER" -UPN $UserId -Message "Désactivation du compte."

        # Récupérer le jobTitle actuel
        $user = Get-MgUser -UserId $UserId -Property "jobTitle" -ErrorAction Stop
        $originalJobTitle = $user.JobTitle

        # Paramètres de base : désactiver le compte
        $bodyParams = @{ AccountEnabled = $false }

        # Remplacement du jobTitle : "DISABLE - A supprimer le JJ/MM/AAAA | Titre d'origine"
        $deleteDate = (Get-Date).AddMonths(3).ToString("dd/MM/yyyy")
        $titlePart = if (-not [string]::IsNullOrWhiteSpace($originalJobTitle)) { $originalJobTitle } else { "N/A" }

        # Ne modifier que si pas déjà au format DISABLE
        if ($originalJobTitle -notlike "DISABLE -*") {
            $newJobTitle = "DISABLE - A supprimer le $deleteDate | $titlePart"
            $bodyParams["JobTitle"] = $newJobTitle
            Write-Log -Level "INFO" -Action "DISABLE_USER" -UPN $UserId -Message "JobTitle modifié : '$originalJobTitle' → '$newJobTitle'"
        }

        Update-MgUser -UserId $UserId -BodyParameter $bodyParams -ErrorAction Stop

        Write-Log -Level "SUCCESS" -Action "DISABLE_USER" -UPN $UserId -Message "Compte désactivé."
        return [PSCustomObject]@{
            Success          = $true
            Error            = $null
            OriginalJobTitle = $originalJobTitle
        }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "DISABLE_USER" -UPN $UserId -Message "Erreur désactivation : $errMsg"
        return [PSCustomObject]@{ Success = $false; Error = $errMsg; OriginalJobTitle = $null }
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
        if ($group -is [array]) { $group = $group[0] }
        $groupId = $group.Id

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
        Les groupes dynamiques (DynamicMembership) sont automatiquement ignorés
        car leurs membres sont gérés par règle — toute tentative de retrait
        provoquerait une erreur 400 de Graph API.

    .PARAMETER UserId
        ID de l'utilisateur.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, RemovedCount: int, SkippedDynamic: int, Errors: array, Error: string}
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
        $dynamicSkipped = 0

        foreach ($membre in $groupes) {
            # On ne retire que les groupes (pas les rôles d'annuaire)
            if ($membre.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group') {

                # Vérifier si c'est un groupe dynamique — on skip
                $groupTypes = $membre.AdditionalProperties.groupTypes
                if ($groupTypes -and ($groupTypes -contains "DynamicMembership")) {
                    $grpName = $membre.AdditionalProperties.displayName
                    Write-Log -Level "INFO" -Action "REMOVE_GROUPS" -UPN $UserId -Message "Groupe dynamique ignoré : '$grpName' ($($membre.Id))"
                    $dynamicSkipped++
                    continue
                }

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
        if ($dynamicSkipped -gt 0) {
            $msg += " $dynamicSkipped groupe(s) dynamique(s) ignoré(s)."
        }
        if ($erreurs.Count -gt 0) {
            $msg += " $($erreurs.Count) erreur(s)."
            Write-Log -Level "WARNING" -Action "REMOVE_GROUPS" -UPN $UserId -Message $msg
        }
        else {
            Write-Log -Level "SUCCESS" -Action "REMOVE_GROUPS" -UPN $UserId -Message $msg
        }

        return [PSCustomObject]@{
            Success        = ($erreurs.Count -eq 0)
            RemovedCount   = $compteur
            SkippedDynamic = $dynamicSkipped
            Errors         = $erreurs
            Error          = if ($erreurs.Count -gt 0) { $erreurs -join "`n" } else { $null }
        }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "REMOVE_GROUPS" -UPN $UserId -Message "Erreur retrait groupes : $errMsg"
        return [PSCustomObject]@{ Success = $false; RemovedCount = 0; SkippedDynamic = 0; Errors = @($errMsg); Error = $errMsg }
    }
}


function Remove-AzUserLicenses {
    <#
    .SYNOPSIS
        Révoque les licences assignées directement à un utilisateur Azure AD.
        Les licences héritées de groupes (group-based licensing) sont ignorées
        car elles ne peuvent pas être retirées directement de l'utilisateur.

    .PARAMETER UserId
        ID de l'utilisateur.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, RemovedCount: int, SkippedInherited: int, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    try {
        Write-Log -Level "INFO" -Action "REVOKE_LICENSES" -UPN $UserId -Message "Révocation de toutes les licences."

        $user = Get-MgUser -UserId $UserId -Property "assignedLicenses,licenseAssignmentStates" -ErrorAction Stop
        $licences = $user.AssignedLicenses

        if ($null -eq $licences -or $licences.Count -eq 0) {
            Write-Log -Level "INFO" -Action "REVOKE_LICENSES" -UPN $UserId -Message "Aucune licence à révoquer."
            return [PSCustomObject]@{ Success = $true; RemovedCount = 0; SkippedInherited = 0; Error = $null }
        }

        # Identifier les licences assignées directement vs héritées de groupes
        # licenseAssignmentStates contient l'origine de chaque licence :
        #   - assignedByGroup = $null ou vide → assignée directement
        #   - assignedByGroup = GUID → héritée d'un groupe
        $directSkuIds = @()
        $inheritedCount = 0

        $assignmentStates = $user.LicenseAssignmentStates
        if ($null -ne $assignmentStates -and $assignmentStates.Count -gt 0) {
            foreach ($state in $assignmentStates) {
                if ([string]::IsNullOrWhiteSpace($state.AssignedByGroup)) {
                    # Licence assignée directement
                    $directSkuIds += $state.SkuId
                }
                else {
                    $inheritedCount++
                    Write-Log -Level "INFO" -Action "REVOKE_LICENSES" -UPN $UserId -Message "Licence héritée ignorée : SkuId=$($state.SkuId), Groupe=$($state.AssignedByGroup)"
                }
            }
            # Dédupliquer (une licence peut apparaître en direct ET en groupe)
            $directSkuIds = $directSkuIds | Select-Object -Unique
        }
        else {
            # Fallback : si licenseAssignmentStates n'est pas disponible,
            # tenter de retirer toutes les licences et gérer l'erreur par SKU
            $directSkuIds = $licences | ForEach-Object { $_.SkuId }
            Write-Log -Level "WARNING" -Action "REVOKE_LICENSES" -UPN $UserId -Message "LicenseAssignmentStates indisponible — tentative de retrait de toutes les licences."
        }

        if ($directSkuIds.Count -eq 0) {
            $msg = "Aucune licence directe à révoquer. $inheritedCount licence(s) héritée(s) de groupes ignorée(s)."
            Write-Log -Level "INFO" -Action "REVOKE_LICENSES" -UPN $UserId -Message $msg
            return [PSCustomObject]@{ Success = $true; RemovedCount = 0; SkippedInherited = $inheritedCount; Error = $null }
        }

        $params = @{
            AddLicenses    = @()
            RemoveLicenses = $directSkuIds
        }

        Set-MgUserLicense -UserId $UserId -BodyParameter $params -ErrorAction Stop

        $msg = "$($directSkuIds.Count) licence(s) directe(s) révoquée(s)."
        if ($inheritedCount -gt 0) {
            $msg += " $inheritedCount licence(s) héritée(s) de groupes ignorée(s)."
        }
        Write-Log -Level "SUCCESS" -Action "REVOKE_LICENSES" -UPN $UserId -Message $msg

        return [PSCustomObject]@{
            Success          = $true
            RemovedCount     = $directSkuIds.Count
            SkippedInherited = $inheritedCount
            Error            = $null
        }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "REVOKE_LICENSES" -UPN $UserId -Message "Erreur révocation licences : $errMsg"
        return [PSCustomObject]@{ Success = $false; RemovedCount = 0; SkippedInherited = 0; Error = $errMsg }
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
        [ValidateSet("department", "jobTitle", "employeeType", "usageLocation", "officeLocation", "companyName")]
        [string]$Property,

        [int]$MaxUsers = 999
    )

    try {
        Write-Log -Level "INFO" -Action "GET_DISTINCT" -Message "Récupération des valeurs distinctes pour '$Property'."

        # Certaines propriétés (employeeType, usageLocation) ne supportent pas $filter dans Graph.
        # On récupère les utilisateurs avec la propriété demandée, sans filtre serveur.
        # Le filtrage des valeurs non-vides se fait côté client.
        if ($Property -in @("department", "jobTitle")) {
            # Ces propriétés supportent le $filter côté serveur
            $filter = "$Property ne null"
            $users = Get-MgUser -Filter $filter -Top $MaxUsers -Property $Property -ConsistencyLevel "eventual" -CountVariable countVar -ErrorAction Stop
        }
        else {
            # Fallback : récupérer tous les users avec la propriété, filtrage côté client
            $users = Get-MgUser -Top $MaxUsers -Property $Property -ErrorAction Stop
        }

        # Extraire les valeurs distinctes et trier
        $values = $users |
            ForEach-Object { $_.$Property } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique

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


# =============================================================================
# FONCTIONS EXCHANGE ONLINE — Offboarding (conversion boîte partagée, GAL, taille)
# =============================================================================

function Get-AzMailboxSize {
    <#
    .SYNOPSIS
        Récupère la taille de la boîte aux lettres d'un utilisateur via Exchange Online.

    .PARAMETER Identity
        UPN ou adresse email de l'utilisateur.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, SizeGB: double, SizeDisplay: string, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )

    try {
        Write-Log -Level "INFO" -Action "GET_MAILBOX_SIZE" -UPN $Identity -Message "Récupération de la taille de la boîte aux lettres."

        $stats = Get-EXOMailboxStatistics -Identity $Identity -ErrorAction Stop

        # Extraction de la taille en octets depuis TotalItemSize
        # Format typique : "1.234 GB (1,324,567,890 bytes)" ou variantes localisées
        $totalItemSize = $stats.TotalItemSize.ToString()
        $sizeBytes = 0

        if ($totalItemSize -match '\(([0-9,\.]+)\s+bytes\)') {
            $sizeBytes = [double]($Matches[1] -replace '[,\.]', '')
        }
        elseif ($totalItemSize -match '([0-9,\.]+)\s+bytes') {
            $sizeBytes = [double]($Matches[1] -replace '[,\.]', '')
        }
        else {
            # Fallback : valeur brute de la propriété Value si disponible
            try {
                $sizeBytes = $stats.TotalItemSize.Value.ToBytes()
            }
            catch {
                Write-Log -Level "WARNING" -Action "GET_MAILBOX_SIZE" -UPN $Identity -Message "Impossible de parser la taille : $totalItemSize"
            }
        }

        $sizeGB = [math]::Round($sizeBytes / 1GB, 2)

        Write-Log -Level "SUCCESS" -Action "GET_MAILBOX_SIZE" -UPN $Identity -Message "Taille : $sizeGB Go ($totalItemSize)"
        return [PSCustomObject]@{
            Success     = $true
            SizeGB      = $sizeGB
            SizeDisplay = $totalItemSize
            Error       = $null
        }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "GET_MAILBOX_SIZE" -UPN $Identity -Message "Erreur récupération taille BAL : $errMsg"
        return [PSCustomObject]@{ Success = $false; SizeGB = 0; SizeDisplay = ""; Error = $errMsg }
    }
}

function Convert-AzMailboxToShared {
    <#
    .SYNOPSIS
        Convertit une boîte aux lettres utilisateur en boîte partagée (Shared Mailbox).
        Requiert une session Exchange Online active.

    .PARAMETER Identity
        UPN ou adresse email de l'utilisateur.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Error: string}

    .NOTES
        - Les boîtes partagées ont une limite de 50 Go sans licence.
        - Au-delà de 50 Go, une licence Exchange Online Plan 1 ou Plan 2 est requise.
        - La conversion doit se faire AVANT la révocation des licences Exchange.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )

    try {
        Write-Log -Level "INFO" -Action "CONVERT_SHARED" -UPN $Identity -Message "Conversion de la BAL en boîte partagée."

        Set-Mailbox -Identity $Identity -Type Shared -ErrorAction Stop

        Write-Log -Level "SUCCESS" -Action "CONVERT_SHARED" -UPN $Identity -Message "BAL convertie en boîte partagée."
        return [PSCustomObject]@{ Success = $true; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "CONVERT_SHARED" -UPN $Identity -Message "Erreur conversion BAL partagée : $errMsg"
        return [PSCustomObject]@{ Success = $false; Error = $errMsg }
    }
}

function Hide-AzMailboxFromGAL {
    <#
    .SYNOPSIS
        Cache une boîte aux lettres du carnet d'adresses global (GAL).
        Tente d'abord via Exchange Online (Set-Mailbox), puis fallback Graph API.

    .PARAMETER Identity
        UPN ou adresse email de l'utilisateur.

    .PARAMETER UserId
        ID Entra de l'utilisateur (utilisé pour le fallback Graph).

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Method: string, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity,

        [string]$UserId = ""
    )

    try {
        Write-Log -Level "INFO" -Action "HIDE_FROM_GAL" -UPN $Identity -Message "Masquage de la BAL dans le carnet d'adresses global."

        # Tentative Exchange Online d'abord
        try {
            Set-Mailbox -Identity $Identity -HiddenFromAddressListsEnabled $true -ErrorAction Stop
            Write-Log -Level "SUCCESS" -Action "HIDE_FROM_GAL" -UPN $Identity -Message "BAL masquée du GAL via Exchange Online."
            return [PSCustomObject]@{ Success = $true; Method = "ExchangeOnline"; Error = $null }
        }
        catch {
            Write-Log -Level "WARNING" -Action "HIDE_FROM_GAL" -UPN $Identity -Message "Exchange Online échoué, tentative Graph API. Erreur EXO : $($_.Exception.Message)"
        }

        # Fallback Graph API — showInAddressList
        if (-not [string]::IsNullOrWhiteSpace($UserId)) {
            Update-MgUser -UserId $UserId -BodyParameter @{ ShowInAddressList = $false } -ErrorAction Stop
            Write-Log -Level "SUCCESS" -Action "HIDE_FROM_GAL" -UPN $Identity -Message "BAL masquée du GAL via Graph API."
            return [PSCustomObject]@{ Success = $true; Method = "GraphAPI"; Error = $null }
        }
        else {
            throw "Exchange Online échoué et aucun UserId fourni pour le fallback Graph API."
        }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "HIDE_FROM_GAL" -UPN $Identity -Message "Erreur masquage GAL : $errMsg"
        return [PSCustomObject]@{ Success = $false; Method = ""; Error = $errMsg }
    }
}

function Grant-AzMailboxFullAccess {
    <#
    .SYNOPSIS
        Ajoute les permissions FullAccess (Read and Manage) sur une boîte aux lettres
        pour un utilisateur délégué. Requiert Exchange Online.

    .PARAMETER MailboxIdentity
        UPN ou adresse email de la boîte aux lettres cible.

    .PARAMETER DelegateUPN
        UPN de l'utilisateur qui recevra l'accès FullAccess.

    .PARAMETER AutoMapping
        Si $true (défaut), la boîte apparaîtra automatiquement dans Outlook du délégué.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Error: string}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxIdentity,

        [Parameter(Mandatory = $true)]
        [string]$DelegateUPN,

        [bool]$AutoMapping = $true
    )

    try {
        Write-Log -Level "INFO" -Action "GRANT_FULLACCESS" -UPN $MailboxIdentity -Message "Ajout FullAccess pour '$DelegateUPN' (AutoMapping=$AutoMapping)."

        Add-MailboxPermission -Identity $MailboxIdentity -User $DelegateUPN `
            -AccessRights FullAccess -InheritanceType All `
            -AutoMapping $AutoMapping -ErrorAction Stop | Out-Null

        Write-Log -Level "SUCCESS" -Action "GRANT_FULLACCESS" -UPN $MailboxIdentity -Message "FullAccess accordé à '$DelegateUPN'."
        return [PSCustomObject]@{ Success = $true; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "GRANT_FULLACCESS" -UPN $MailboxIdentity -Message "Erreur ajout FullAccess pour '$DelegateUPN' : $errMsg"
        return [PSCustomObject]@{ Success = $false; Error = $errMsg }
    }
}

# Point d'attention :
# - Chaque fonction retourne TOUJOURS un PSCustomObject avec au minimum Success et Error
# - Les mots de passe ne sont JAMAIS inclus dans les logs
# - Les erreurs "déjà membre" lors d'ajout à un groupe sont traitées comme des succès