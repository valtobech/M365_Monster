<#
.FICHIER
    Core/PIMFunctions.ps1

.ROLE
    Fonctions metier pour la gestion PIM (Privileged Identity Management).
    Chargement des roles Entra ID, audit des groupes PIM, creation,
    assignation et mise a jour des roles via Graph API.

.DEPENDANCES
    - Core/Functions.ps1  (Write-Log, Show-ConfirmDialog, Show-ResultDialog)
    - Core/GraphAPI.ps1   (connexion Graph active via Invoke-MgGraphRequest)
    - Core/Lang.ps1       (Get-Text)
    - Variable globale    $Config

.AUTEUR
    [Equipe IT -- M365 Monster]
#>

# === Variables partagees du module PIM ===
$script:PimData          = [ordered]@{}
$script:PimGroupStatus   = @{}
$script:PimAuditResults  = @{}
$script:AllEntraRoles    = @{}

function Test-PimActiveAssignment {
    <#
    .SYNOPSIS
        Determine si un type de groupe PIM utilise des assignations Active (vs Eligible).
        Role_Fixe et Groupe* = Active, sinon Eligible.
    .PARAMETER GroupType
        Type du groupe : Role, Role_Fixe, Groupe, Groupe_Critical.
    .OUTPUTS
        [bool] -- $true si Active, $false si Eligible.
    #>
    param([string]$GroupType)
    return ($GroupType -like 'Groupe*' -or $GroupType -eq 'Role_Fixe')
}

function Initialize-PimConfig {
    <#
    .SYNOPSIS
        Verifie et initialise la section pim_role_groups dans la config client.
        Si la section n'existe pas, elle est creee automatiquement avec un
        template vide et sauvegardee dans le fichier JSON client.
    .OUTPUTS
        [bool] -- $true si la section existe ou a ete creee, $false si erreur.
    #>

    if ($Config.PSObject.Properties["pim_role_groups"]) {
        Write-Log -Level "INFO" -Action "PIM_INIT" -Message "Section pim_role_groups trouvee dans la config."
        return $true
    }

    # Section absente -- creation automatique
    Write-Log -Level "WARNING" -Action "PIM_INIT" -Message "Section pim_role_groups absente. Injection automatique..."

    try {
        # Ajouter la section vide a l'objet config en memoire
        $emptySection = [PSCustomObject]@{}
        $Config | Add-Member -NotePropertyName "pim_role_groups" -NotePropertyValue $emptySection -Force

        # Sauvegarder dans le fichier JSON client
        $configPath = $Config._config_path
        if ($configPath -and (Test-Path $configPath)) {
            # Relire le JSON brut, ajouter la section, reecrire
            $jsonContent = Get-Content -Path $configPath -Raw -Encoding UTF8
            $jsonObject  = $jsonContent | ConvertFrom-Json
            $jsonObject | Add-Member -NotePropertyName "pim_role_groups" -NotePropertyValue ([PSCustomObject]@{}) -Force
            $jsonObject | ConvertTo-Json -Depth 10 | Out-File -FilePath $configPath -Encoding UTF8 -Force
            Write-Log -Level "SUCCESS" -Action "PIM_INIT" -Message "Section pim_role_groups ajoutee dans '$configPath'."
        }
        else {
            Write-Log -Level "WARNING" -Action "PIM_INIT" -Message "Chemin config introuvable -- section ajoutee en memoire uniquement."
        }
        return $true
    }
    catch {
        Write-Log -Level "ERROR" -Action "PIM_INIT" -Message "Erreur injection pim_role_groups : $($_.Exception.Message)"
        return $false
    }
}

function Import-PimData {
    <#
    .SYNOPSIS
        Charge les donnees PIM depuis la config client dans $script:PimData.
    #>

    $script:PimData = [ordered]@{}
    $script:PimGroupStatus   = @{}
    $script:PimAuditResults  = @{}

    foreach ($key in $Config.pim_role_groups.PSObject.Properties.Name) {
        $src = $Config.pim_role_groups.$key
        $script:PimData[$key] = @{
            Description = if ($src.description) { $src.description } else { '' }
            Type        = if ($src.type) { $src.type } else { 'Role' }
            Roles       = @(if ($src.roles) { $src.roles } else { @() })
        }
    }

    foreach ($g in $script:PimData.Keys) {
        $script:PimGroupStatus[$g] = 'Pending'
    }

    Write-Log -Level "INFO" -Action "PIM_INIT" -Message "$($script:PimData.Count) groupe(s) PIM charges depuis la config."
}

function Load-EntraRoles {
    <#
    .SYNOPSIS
        Charge toutes les definitions de roles Entra ID (built-in + custom)
        via Graph API REST. Peuple $script:AllEntraRoles.
    .OUTPUTS
        [bool] -- $true si charge, $false si erreur.
    #>

    Write-Log -Level "INFO" -Action "PIM_ROLES" -Message "Chargement des definitions de roles Entra ID..."
    try {
        # Charger toutes les definitions de roles sans filtre OData
        # (evite les problemes d'echappement $select/$top dans les URIs)
        $uri = 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions'
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop

        $script:AllEntraRoles = @{}
        $allRoles = @($response.value)

        # Pagination si necessaire
        while ($response.'@odata.nextLink') {
            $response = Invoke-MgGraphRequest -Method GET -Uri $response.'@odata.nextLink' -ErrorAction Stop
            $allRoles += $response.value
        }

        foreach ($role in $allRoles) {
            # Pour les roles built-in, Id == TemplateId
            # Pour les roles custom, templateId peut etre different de id
            $roleId = if ($role.templateId) { $role.templateId } else { $role.id }
            $script:AllEntraRoles[$role.displayName] = @{
                Id         = $role.id
                TemplateId = $roleId
                IsBuiltIn  = [bool]$role.isBuiltIn
            }
        }

        Write-Log -Level "SUCCESS" -Action "PIM_ROLES" -Message "$($script:AllEntraRoles.Count) roles Entra ID charges."
        return $true
    }
    catch {
        $errDetail = $_.Exception.Message
        # Tenter d'extraire le detail de l'erreur Graph (body JSON)
        if ($_.ErrorDetails.Message) {
            $errDetail += " | Detail: $($_.ErrorDetails.Message)"
        }
        Write-Log -Level "ERROR" -Action "PIM_ROLES" -Message "Erreur chargement roles : $errDetail"
        return $false
    }
}

function Resolve-RoleName {
    <#
    .SYNOPSIS
        Resout un roleDefinitionId (GUID) vers un displayName.
        Cherche d'abord dans $AllEntraRoles (local), puis appel Graph direct.
        Enrichit $AllEntraRoles pour les prochains lookups.
    .PARAMETER RoleDefId
        GUID du roleDefinitionId a resoudre.
    .OUTPUTS
        [string] -- displayName du role, ou $null si non resolu.
    #>
    param([string]$RoleDefId)

    if ([string]::IsNullOrWhiteSpace($RoleDefId)) { return $null }

    # Lookup rapide dans le cache existant (par Id ou TemplateId)
    foreach ($entry in $script:AllEntraRoles.GetEnumerator()) {
        if ($entry.Value.Id -eq $RoleDefId -or $entry.Value.TemplateId -eq $RoleDefId) {
            return $entry.Key
        }
    }

    # Appel Graph direct pour resoudre le role inconnu
    try {
        $roleResp = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/$RoleDefId" `
            -ErrorAction Stop
        $name  = $roleResp.displayName
        $tplId = if ($roleResp.templateId) { $roleResp.templateId } else { $roleResp.id }

        # Enrichir le cache pour les prochains lookups
        # Stocker avec les 3 IDs possibles : id, templateId, et le RoleDefId original
        # (pour les custom roles, le schedule peut utiliser n'importe lequel)
        if (-not $script:AllEntraRoles.ContainsKey($name)) {
            $script:AllEntraRoles[$name] = @{
                Id         = $roleResp.id
                TemplateId = $tplId
                OriginalId = $RoleDefId
                IsBuiltIn  = [bool]$roleResp.isBuiltIn
            }
        }
        else {
            # Ajouter l'OriginalId si pas deja present
            if (-not $script:AllEntraRoles[$name].OriginalId) {
                $script:AllEntraRoles[$name].OriginalId = $RoleDefId
            }
        }
        Write-Log -Level "INFO" -Action "PIM_RESOLVE" -Message "Role resolu via Graph : $name (Id=$RoleDefId, Custom=$(-not $roleResp.isBuiltIn))"
        return $name
    }
    catch {
        Write-Log -Level "WARNING" -Action "PIM_RESOLVE" -Message "Impossible de resoudre le role $RoleDefId : $($_.Exception.Message)"
        return $null
    }
}

function Get-PimSchedules {
    <#
    .SYNOPSIS
        Charge les schedules PIM (eligibility + assignment) en une seule fois.
        Gere la pagination.
    .OUTPUTS
        [hashtable] -- @{ Eligible = @(...); Active = @(...) } ou $null si erreur.
    #>

    Write-Log -Level "INFO" -Action "PIM_SCHEDULES" -Message "Chargement des assignations PIM..."
    try {
        # Eligible schedules
        $eligUri = 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilitySchedules'
        $eligResponse = Invoke-MgGraphRequest -Method GET -Uri $eligUri -ErrorAction Stop
        $eligSchedules = @($eligResponse.value)
        while ($eligResponse.'@odata.nextLink') {
            $eligResponse = Invoke-MgGraphRequest -Method GET -Uri $eligResponse.'@odata.nextLink' -ErrorAction Stop
            $eligSchedules += $eligResponse.value
        }

        # Active schedules
        $actUri = 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentSchedules'
        $actResponse = Invoke-MgGraphRequest -Method GET -Uri $actUri -ErrorAction Stop
        $actSchedules = @($actResponse.value)
        while ($actResponse.'@odata.nextLink') {
            $actResponse = Invoke-MgGraphRequest -Method GET -Uri $actResponse.'@odata.nextLink' -ErrorAction Stop
            $actSchedules += $actResponse.value
        }

        return @{ Eligible = $eligSchedules; Active = $actSchedules }
    }
    catch {
        Write-Log -Level "ERROR" -Action "PIM_SCHEDULES" -Message "Erreur chargement : $($_.Exception.Message)"
        return $null
    }
}

function Invoke-PimAudit {
    <#
    .SYNOPSIS
        Audite tous les groupes PIM definis dans $script:PimData.
        Met a jour $script:PimGroupStatus et $script:PimAuditResults.
    .PARAMETER ProgressCallback
        Scriptblock appele avec (done, total) pour mettre a jour la barre de progression.
    .PARAMETER StepCallback
        Scriptblock appele avec (message) pour afficher l'etape en cours.
    #>
    param([scriptblock]$ProgressCallback, [scriptblock]$StepCallback)

    Write-Log -Level "INFO" -Action "PIM_AUDIT" -Message "=== Debut audit des groupes PIM ==="
    if ($StepCallback) { & $StepCallback (Get-Text "pim.step_loading_schedules") }

    $schedules = Get-PimSchedules
    if (-not $schedules) {
        Show-ResultDialog -Titre (Get-Text "pim.title") `
            -Message (Get-Text "pim.audit_load_error" "Impossible de charger les schedules PIM.") `
            -IsSuccess $false
        return
    }

    $eligSchedules = $schedules.Eligible
    $actSchedules  = $schedules.Active
    $done  = 0
    $total = $script:PimData.Count

    foreach ($gName in $script:PimData.Keys) {
        $def    = $script:PimData[$gName]
        $issues = [System.Collections.Generic.List[PSCustomObject]]::new()
        $isActive = Test-PimActiveAssignment -GroupType $def.Type

        if ($StepCallback) { & $StepCallback ("[$gName] " + (Get-Text "pim.step_checking_group")) }

        # Recherche du groupe dans Entra
        $grp = $null
        try {
            $grpUri = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$gName'&`$select=id,displayName,isAssignableToRole,description&`$top=1"
            $grpResponse = Invoke-MgGraphRequest -Method GET -Uri $grpUri -ErrorAction Stop
            $grp = $grpResponse.value | Select-Object -First 1
        }
        catch { $grp = $null }

        if (-not $grp) {
            $issues.Add([PSCustomObject]@{ Text = (Get-Text "pim.audit_group_missing"); Status = 'Error' })
            $script:PimGroupStatus[$gName]  = 'Missing'
            $script:PimAuditResults[$gName] = $issues
            $done++
            if ($ProgressCallback) { & $ProgressCallback $done $total }
            Write-Log -Level "ERROR" -Action "PIM_AUDIT" -Message "[$gName] MANQUANT"
            continue
        }

        # IsAssignableToRole
        if ($grp.isAssignableToRole) {
            $issues.Add([PSCustomObject]@{ Text = "IsAssignableToRole = True"; Status = 'OK' })
        }
        else {
            $issues.Add([PSCustomObject]@{ Text = "IsAssignableToRole = False (non corrigeable)"; Status = 'Error' })
        }

        # Description (information seulement, pas de drift)
        if ($grp.description -ne $def.Description) {
            $descText = (Get-Text "pim.audit_desc_info") -f $def.Description, $grp.description
            $issues.Add([PSCustomObject]@{ Text = $descText; Status = 'Info' })
        }

        # Verification de chaque role
        foreach ($roleName in @($def.Roles)) {
            $roleInfo = $script:AllEntraRoles[$roleName]

            # Si le role n'est pas trouve, tenter de resoudre un pattern Unknown (guid)
            if (-not $roleInfo -and $roleName -match '^Unknown\s*\(([0-9a-fA-F\-]{36})\)$') {
                $guid = $Matches[1]
                $resolvedName = Resolve-RoleName -RoleDefId $guid
                if ($resolvedName) {
                    # Renommer le role dans PimData (remplacer Unknown par le vrai nom)
                    $def.Roles = @($def.Roles | ForEach-Object {
                        if ($_ -eq $roleName) { $resolvedName } else { $_ }
                    })
                    $script:PimDirty = $true
                    Write-Log -Level "SUCCESS" -Action "PIM_AUDIT" -Message "[$gName] Role resolu : '$roleName' -> '$resolvedName'"
                    $issues.Add([PSCustomObject]@{
                        Text   = "[$resolvedName] $(Get-Text 'pim.audit_role_resolved') ($guid)"
                        Status = 'OK'
                    })
                    $roleName = $resolvedName
                    $roleInfo = $script:AllEntraRoles[$roleName]
                }
            }

            if (-not $roleInfo) {
                $issues.Add([PSCustomObject]@{
                    Text   = "[$roleName] $(Get-Text 'pim.audit_role_unknown')"
                    Status = 'Error'
                })
                continue
            }

            # Matcher par Id, TemplateId et OriginalId (les custom roles peuvent
            # utiliser n'importe lequel de ces IDs dans les schedules PIM)
            $rid = $roleInfo.Id
            $tid = $roleInfo.TemplateId
            $oid = $roleInfo.OriginalId  # peut etre $null pour les built-in
            $found = $null

            $targetSchedules = if ($isActive) { $actSchedules } else { $eligSchedules }
            $found = $targetSchedules | Where-Object {
                $_.principalId -eq $grp.id -and
                ($_.roleDefinitionId -eq $rid -or $_.roleDefinitionId -eq $tid -or ($oid -and $_.roleDefinitionId -eq $oid)) -and
                $_.status -eq 'Provisioned'
            } | Select-Object -First 1

            # Si non trouve dans la liste ciblee, chercher dans l'autre liste
            # (un custom role peut etre assigne en Eligible meme si le groupe est type Groupe)
            if ($null -eq $found) {
                $altSchedules = if ($isActive) { $eligSchedules } else { $actSchedules }
                $altFound = $altSchedules | Where-Object {
                    $_.principalId -eq $grp.id -and
                    ($_.roleDefinitionId -eq $rid -or $_.roleDefinitionId -eq $tid -or ($oid -and $_.roleDefinitionId -eq $oid)) -and
                    $_.status -eq 'Provisioned'
                } | Select-Object -First 1
                if ($altFound) {
                    $found = $altFound
                    $altLabel = if ($isActive) { 'Eligible' } else { 'Active' }
                    Write-Log -Level "WARNING" -Action "PIM_AUDIT" -Message "[$gName] [$roleName] Trouve en $altLabel au lieu de $(if ($isActive) {'Active'} else {'Eligible'})"
                }
            }

            if ($null -ne $found) {
                $expiry = 'permanent'
                if ($found.scheduleInfo.expiration.type -ne 'noExpiration') {
                    $endDt = $found.scheduleInfo.expiration.endDateTime
                    if ($endDt) { $expiry = (Get-Text "pim.audit_expires") + " " + ([datetime]$endDt).ToString('yyyy-MM-dd') }
                }
                $customTag = if (-not $roleInfo.IsBuiltIn) { ' [Custom]' } else { '' }
                $issues.Add([PSCustomObject]@{ Text = "[$roleName]$customTag OK -- $expiry"; Status = 'OK' })
            }
            else {
                $typeLabel = if ($isActive) { 'Active' } else { 'Eligible' }
                $customTag = if (-not $roleInfo.IsBuiltIn) { ' [Custom]' } else { '' }
                $issues.Add([PSCustomObject]@{ Text = "[$roleName]$customTag $(Get-Text 'pim.audit_role_missing') ($typeLabel)"; Status = 'Error' })
                # Log detaille pour diagnostic des custom roles
                Write-Log -Level "WARNING" -Action "PIM_AUDIT" -Message "[$gName] [$roleName] MANQUANT -- RoleId=$rid TemplateId=$tid OriginalId=$oid GroupId=$($grp.id) IsActive=$isActive"
            }
        }

        # Bilan (seules les erreurs declenchent un Drift, pas les Info/Warn)
        $errCount  = @($issues | Where-Object { $_.Status -eq 'Error' }).Count
        $script:PimGroupStatus[$gName] = if ($errCount -gt 0) { 'Drift' } else { 'OK' }
        $script:PimAuditResults[$gName] = $issues

        $logLevel = if ($errCount -gt 0) { "WARNING" } else { "SUCCESS" }
        Write-Log -Level $logLevel -Action "PIM_AUDIT" -Message "[$gName] $errCount erreur(s), $warnCount avertissement(s)"

        $done++
        if ($ProgressCallback) { & $ProgressCallback $done $total }
    }

    if ($StepCallback) { & $StepCallback (Get-Text "pim.step_complete") }
    Write-Log -Level "SUCCESS" -Action "PIM_AUDIT" -Message "=== Audit PIM termine ==="
}

function Wait-PimGroupReplication {
    <#
    .SYNOPSIS
        Attend la replication Entra d'un groupe (polling 5s, max 60s).
    .OUTPUTS
        [bool] -- $true si replique, $false si timeout.
    #>
    param([string]$GroupId, [string]$GroupName)

    Write-Log -Level "WARNING" -Action "PIM_CREATE" -Message "[$GroupName] Attente replication Entra (max 60s)..."
    for ($i = 1; $i -le 12; $i++) {
        Start-Sleep -Seconds 5
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $null = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId" -ErrorAction Stop
            Write-Log -Level "SUCCESS" -Action "PIM_CREATE" -Message "[$GroupName] Groupe replique apres $($i*5)s"
            return $true
        }
        catch {}
    }
    Write-Log -Level "ERROR" -Action "PIM_CREATE" -Message "[$GroupName] Timeout replication"
    return $false
}

function New-PimGroupEntra {
    <#
    .SYNOPSIS
        Cree un groupe PIM dans Entra ID (IsAssignableToRole = true, irreversible).
    .OUTPUTS
        [string] -- Id du groupe cree ou existant, $null si annule/erreur.
    #>
    param([string]$Name, [hashtable]$Def)

    # Verifier si existe deja
    $existing = $null
    try {
        $resp = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$Name'&`$select=id&`$top=1" -ErrorAction Stop
        $existing = $resp.value | Select-Object -First 1
    }
    catch {}

    if ($existing) {
        Write-Log -Level "WARNING" -Action "PIM_CREATE" -Message "[$Name] Groupe existant (Id: $($existing.id))"
        $script:PimGroupStatus[$Name] = 'Skipped'
        return $existing.id
    }

    # Confirmation utilisateur
    $confirmMsg = (Get-Text "pim.create_confirm") -f $Name, $Def.Description
    $confirm = Show-ConfirmDialog -Titre (Get-Text "pim.create_confirm_title") -Message $confirmMsg
    if (-not $confirm) {
        Write-Log -Level "WARNING" -Action "PIM_CREATE" -Message "[$Name] Annule par l'utilisateur"
        $script:PimGroupStatus[$Name] = 'Skipped'
        return $null
    }

    try {
        Write-Log -Level "INFO" -Action "PIM_CREATE" -Message "[$Name] Creation en cours..."
        $body = @{
            displayName        = $Name
            description        = $Def.Description
            mailEnabled        = $false
            mailNickname       = ($Name -replace '[^a-zA-Z0-9]','_')
            securityEnabled    = $true
            isAssignableToRole = $true
            groupTypes         = @()
        }
        $grp = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" `
            -Body ($body | ConvertTo-Json -Depth 5) -ContentType "application/json" -ErrorAction Stop

        Write-Log -Level "SUCCESS" -Action "PIM_CREATE" -Message "[$Name] Groupe cree (Id: $($grp.id))"
        $script:PimGroupStatus[$Name] = 'Success'
        return $grp.id
    }
    catch {
        Write-Log -Level "ERROR" -Action "PIM_CREATE" -Message "[$Name] Erreur creation : $($_.Exception.Message)"
        $script:PimGroupStatus[$Name] = 'Error'
        return $null
    }
}

function Add-PimRolesToGroup {
    <#
    .SYNOPSIS
        Assigne les roles a un groupe PIM via schedule requests.
        Gere retry SubjectNotFound et fallback duree max tenant.
    #>
    param(
        [string]$GroupId, [string]$GroupName,
        [string[]]$Roles, [string]$GroupType
    )

    if (-not (Wait-PimGroupReplication -GroupId $GroupId -GroupName $GroupName)) { return }

    Write-Log -Level "WARNING" -Action "PIM_ASSIGN" -Message "[$GroupName] Attente propagation PIM (15s)..."
    Start-Sleep -Seconds 15
    [System.Windows.Forms.Application]::DoEvents()

    $isActive    = Test-PimActiveAssignment -GroupType $GroupType
    $assignLabel = if ($isActive) { 'Active' } else { 'Eligible' }
    $cmdUri      = if ($isActive) {
        "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests"
    } else {
        "https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleRequests"
    }

    foreach ($roleName in ($Roles | Where-Object { $_ -notmatch '\[Custom\]' })) {
        $roleInfo = $script:AllEntraRoles[$roleName]
        if (-not $roleInfo) {
            Write-Log -Level "WARNING" -Action "PIM_ASSIGN" -Message "[$GroupName] Role introuvable : $roleName"
            continue
        }

        $roleDefId   = $roleInfo.Id
        $startDt     = (Get-Date).ToUniversalTime().ToString('o')
        $endEligible = (Get-Date).AddYears(1).AddDays(-1).ToUniversalTime().ToString('o')
        $endActive   = (Get-Date).AddMonths(6).AddDays(-1).ToUniversalTime().ToString('o')
        $endFallback = if ($isActive) { $endActive } else { $endEligible }

        $bodyNoExp = @{
            action = 'adminAssign'; justification = 'M365 Monster PIM Manager - assignation initiale'
            principalId = $GroupId; roleDefinitionId = $roleDefId; directoryScopeId = '/'
            scheduleInfo = @{ startDateTime = $startDt; expiration = @{ type = 'noExpiration' } }
        }
        $bodyFallback = @{
            action = 'adminAssign'; justification = 'M365 Monster PIM Manager - assignation (duree max tenant)'
            principalId = $GroupId; roleDefinitionId = $roleDefId; directoryScopeId = '/'
            scheduleInfo = @{ startDateTime = $startDt; expiration = @{ type = 'afterDateTime'; endDateTime = $endFallback } }
        }

        $attempt = 0; $ok = $false; $usedFallback = $false
        while (-not $ok -and $attempt -lt 4) {
            $attempt++
            $body = if ($usedFallback) { $bodyFallback } else { $bodyNoExp }
            try {
                $null = Invoke-MgGraphRequest -Method POST -Uri $cmdUri `
                    -Body ($body | ConvertTo-Json -Depth 10) -ContentType "application/json" -ErrorAction Stop
                $suffix = if ($usedFallback) { if ($isActive) { ' [6 mois]' } else { ' [1 an]' } } else { '' }
                $lvl = if ($usedFallback) { "WARNING" } else { "SUCCESS" }
                Write-Log -Level $lvl -Action "PIM_ASSIGN" -Message "[$GroupName] [$assignLabel] $roleName$suffix"
                $ok = $true
            }
            catch {
                $errMsg = $_.Exception.Message
                if ($errMsg -match 'SubjectNotFound' -and $attempt -lt 4) {
                    Write-Log -Level "WARNING" -Action "PIM_ASSIGN" -Message "[$GroupName] SubjectNotFound '$roleName' -- retry $attempt/3 dans 10s"
                    Start-Sleep -Seconds 10
                    [System.Windows.Forms.Application]::DoEvents()
                }
                elseif ($errMsg -match 'noExpiration|maximumGrantPeriod|PolicyViolated|ExpirationRule' -and -not $usedFallback) {
                    $label = if ($isActive) { '6 mois' } else { '1 an' }
                    Write-Log -Level "WARNING" -Action "PIM_ASSIGN" -Message "[$GroupName] Expiration refusee '$roleName' -- fallback $label"
                    $usedFallback = $true; $attempt = 0
                }
                else {
                    Write-Log -Level "ERROR" -Action "PIM_ASSIGN" -Message "[$GroupName] Erreur '$roleName' : $errMsg"
                    $ok = $true
                }
            }
        }
    }

    # Rappel pour les roles [Custom]
    foreach ($r in ($Roles | Where-Object { $_ -match '\[Custom\]' })) {
        Write-Log -Level "WARNING" -Action "PIM_ASSIGN" -Message "[$GroupName] Role custom -- assignation manuelle requise : $r"
    }
}

function Invoke-PimUpdate {
    <#
    .SYNOPSIS
        Met a jour les groupes PIM selectionnes :
        - Corrige displayName et description si drift
        - Ajoute les roles manquants
        - Retire les roles en trop (avec confirmation)
    .PARAMETER GroupNames
        Noms des groupes a mettre a jour.
    .PARAMETER ProgressCallback
        Scriptblock appele avec (done, total) pour la barre de progression.
    .PARAMETER StepCallback
        Scriptblock appele avec (message) pour afficher l'etape en cours.
    #>
    param(
        [string[]]$GroupNames,
        [scriptblock]$ProgressCallback,
        [scriptblock]$StepCallback
    )

    Write-Log -Level "INFO" -Action "PIM_UPDATE" -Message "=== Debut mise a jour -- $($GroupNames.Count) groupe(s) ==="
    if ($StepCallback) { & $StepCallback (Get-Text "pim.step_loading_schedules") }

    $schedules = Get-PimSchedules
    if (-not $schedules) {
        Write-Log -Level "ERROR" -Action "PIM_UPDATE" -Message "Impossible de charger les schedules PIM."
        return
    }

    $eligSchedules = $schedules.Eligible
    $actSchedules  = $schedules.Active
    $total = $GroupNames.Count
    $done  = 0

    foreach ($gName in $GroupNames) {
        $def      = $script:PimData[$gName]
        $isActive = Test-PimActiveAssignment -GroupType $def.Type

        if ($StepCallback) { & $StepCallback ("[$gName] " + (Get-Text "pim.step_checking_group")) }

        # Recherche du groupe avec description et displayName
        $grp = $null
        try {
            $resp = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$gName'&`$select=id,displayName,description&`$top=1" -ErrorAction Stop
            $grp = $resp.value | Select-Object -First 1
        }
        catch {}

        if (-not $grp) {
            Write-Log -Level "WARNING" -Action "PIM_UPDATE" -Message "[$gName] Groupe introuvable -- ignore"
            $done++; if ($ProgressCallback) { & $ProgressCallback $done $total }
            continue
        }

        # Description : pas de correction automatique (informatif seulement)
        if ($grp.description -ne $def.Description) {
            Write-Log -Level "INFO" -Action "PIM_UPDATE" -Message "[$gName] Description differente (non corrigee) -- Local='$($def.Description)' Entra='$($grp.description)'"
        }

        # === Calcul du delta des roles ===
        if ($StepCallback) { & $StepCallback ("[$gName] " + (Get-Text "pim.step_comparing_roles")) }

        # Construire un mapping bidirectionnel : tous les IDs possibles d'un role -> Id canonique
        # Pour les custom roles, le schedule peut utiliser Id, TemplateId ou OriginalId
        $expectedCanonical = @{}  # canonicalId -> roleName
        $allExpectedIds = @{}     # anyId -> canonicalId
        foreach ($rn in $def.Roles) {
            $ri = $script:AllEntraRoles[$rn]
            if (-not $ri) { continue }
            $canonical = $ri.Id
            $expectedCanonical[$canonical] = $rn
            $allExpectedIds[$ri.Id] = $canonical
            if ($ri.TemplateId) { $allExpectedIds[$ri.TemplateId] = $canonical }
            if ($ri.OriginalId) { $allExpectedIds[$ri.OriginalId] = $canonical }
        }

        $targetSchedules = if ($isActive) {
            $actSchedules | Where-Object { $_.principalId -eq $grp.id -and $_.status -eq 'Provisioned' }
        } else {
            $eligSchedules | Where-Object { $_.principalId -eq $grp.id -and $_.status -eq 'Provisioned' }
        }
        $currentIds = @($targetSchedules | ForEach-Object { $_.roleDefinitionId })

        # Determiner les roles manquants (attendus mais pas dans les schedules)
        $matchedCanonicals = @{}
        foreach ($cid in $currentIds) {
            $canonical = $allExpectedIds[$cid]
            if ($canonical) { $matchedCanonicals[$canonical] = $true }
        }
        $toAdd = @($expectedCanonical.Keys | Where-Object { -not $matchedCanonicals.ContainsKey($_) })

        # Determiner les roles en trop (dans les schedules mais pas attendus)
        $toRemove = @($currentIds | Where-Object { -not $allExpectedIds.ContainsKey($_) })

        Write-Log -Level "INFO" -Action "PIM_UPDATE" -Message "[$gName] A ajouter : $($toAdd.Count) -- A retirer : $($toRemove.Count)"

        # === Retrait (avec confirmation) ===
        if ($toRemove.Count -gt 0) {
            $removeNames = $toRemove | ForEach-Object {
                $roleId = $_
                $match = $script:AllEntraRoles.GetEnumerator() | Where-Object { $_.Value.Id -eq $roleId -or $_.Value.TemplateId -eq $roleId } | Select-Object -First 1
                if ($match) { $match.Key } else { $roleId }
            }
            $confirmMsg = (Get-Text "pim.update_remove_confirm") -f $gName, ($removeNames -join "`n")
            $confirmed = Show-ConfirmDialog -Titre (Get-Text "pim.update_remove_title") -Message $confirmMsg

            if ($confirmed) {
                $removeUri = if ($isActive) {
                    "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests"
                } else {
                    "https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleRequests"
                }
                foreach ($roleId in $toRemove) {
                    $roleName = ($script:AllEntraRoles.GetEnumerator() | Where-Object { $_.Value.Id -eq $roleId -or $_.Value.TemplateId -eq $roleId } | Select-Object -First 1).Key
                    if (-not $roleName) { $roleName = $roleId }
                    if ($StepCallback) { & $StepCallback ("[$gName] " + (Get-Text "pim.step_removing_role") + " $roleName") }
                    try {
                        $null = Invoke-MgGraphRequest -Method POST -Uri $removeUri `
                            -Body (@{ action='adminRemove'; justification='M365 Monster PIM Manager - retrait drift'
                                      principalId=$grp.id; roleDefinitionId=$roleId; directoryScopeId='/' } | ConvertTo-Json -Depth 5) `
                            -ContentType "application/json" -ErrorAction Stop
                        Write-Log -Level "SUCCESS" -Action "PIM_UPDATE" -Message "[$gName] Retire : $roleName"
                    }
                    catch {
                        Write-Log -Level "ERROR" -Action "PIM_UPDATE" -Message "[$gName] Erreur retrait '$roleName' : $($_.Exception.Message)"
                    }
                }
            }
            else {
                Write-Log -Level "WARNING" -Action "PIM_UPDATE" -Message "[$gName] Retrait annule par l'utilisateur"
            }
        }

        # === Ajout des roles manquants ===
        if ($toAdd.Count -gt 0) {
            if ($StepCallback) { & $StepCallback ("[$gName] " + (Get-Text "pim.step_waiting_pim")) }
            Write-Log -Level "WARNING" -Action "PIM_UPDATE" -Message "[$gName] Attente propagation PIM (15s)..."
            Start-Sleep -Seconds 15
            [System.Windows.Forms.Application]::DoEvents()

            $addUri = if ($isActive) {
                "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests"
            } else {
                "https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleRequests"
            }
            $assignLabel = if ($isActive) { 'Active' } else { 'Eligible' }

            foreach ($roleId in $toAdd) {
                $roleName = ($script:AllEntraRoles.GetEnumerator() | Where-Object { $_.Value.Id -eq $roleId -or $_.Value.TemplateId -eq $roleId } | Select-Object -First 1).Key
                if (-not $roleName) { $roleName = $roleId }

                if ($StepCallback) { & $StepCallback ("[$gName] [$assignLabel] " + (Get-Text "pim.step_adding_role") + " $roleName") }

                $startDt     = (Get-Date).ToUniversalTime().ToString('o')
                $endFallback = if ($isActive) { (Get-Date).AddMonths(6).AddDays(-1).ToUniversalTime().ToString('o') } `
                               else { (Get-Date).AddYears(1).AddDays(-1).ToUniversalTime().ToString('o') }

                $bodyNoExp = @{
                    action='adminAssign'; justification='M365 Monster PIM Manager - ajout manquant'
                    principalId=$grp.id; roleDefinitionId=$roleId; directoryScopeId='/'
                    scheduleInfo=@{ startDateTime=$startDt; expiration=@{ type='noExpiration' } }
                }
                $bodyFallback = @{
                    action='adminAssign'; justification='M365 Monster PIM Manager - ajout manquant (duree max tenant)'
                    principalId=$grp.id; roleDefinitionId=$roleId; directoryScopeId='/'
                    scheduleInfo=@{ startDateTime=$startDt; expiration=@{ type='afterDateTime'; endDateTime=$endFallback } }
                }

                $attempt = 0; $ok = $false; $usedFallback = $false
                while (-not $ok -and $attempt -lt 4) {
                    $attempt++
                    $body = if ($usedFallback) { $bodyFallback } else { $bodyNoExp }
                    try {
                        $null = Invoke-MgGraphRequest -Method POST -Uri $addUri `
                            -Body ($body | ConvertTo-Json -Depth 10) -ContentType "application/json" -ErrorAction Stop
                        $suffix = if ($usedFallback) { if ($isActive) { ' [6 mois]' } else { ' [1 an]' } } else { '' }
                        $lvl = if ($usedFallback) { "WARNING" } else { "SUCCESS" }
                        Write-Log -Level $lvl -Action "PIM_UPDATE" -Message "[$gName] [$assignLabel] Ajoute : $roleName$suffix"
                        $ok = $true
                    }
                    catch {
                        $errMsg = $_.Exception.Message
                        if ($errMsg -match 'SubjectNotFound' -and $attempt -lt 4) {
                            Start-Sleep -Seconds 10; [System.Windows.Forms.Application]::DoEvents()
                        }
                        elseif ($errMsg -match 'noExpiration|maximumGrantPeriod|PolicyViolated|ExpirationRule' -and -not $usedFallback) {
                            $usedFallback = $true; $attempt = 0
                        }
                        else {
                            Write-Log -Level "ERROR" -Action "PIM_UPDATE" -Message "[$gName] Erreur '$roleName' : $errMsg"
                            $ok = $true
                        }
                    }
                }
            }
        }

        $done++
        if ($ProgressCallback) { & $ProgressCallback $done $total }
        Write-Log -Level "SUCCESS" -Action "PIM_UPDATE" -Message "[$gName] Mise a jour terminee"
    }
    if ($StepCallback) { & $StepCallback (Get-Text "pim.step_complete") }
    Write-Log -Level "SUCCESS" -Action "PIM_UPDATE" -Message "=== Update PIM termine ==="
}

function Export-PimCsvReport {
    <#
    .SYNOPSIS
        Exporte un rapport CSV de l'etat des groupes PIM.
    #>

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter   = 'CSV|*.csv'
    $sfd.FileName = "PIM_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $rows = foreach ($gn in $script:PimData.Keys) {
        $d = $script:PimData[$gn]
        $auditItems = $script:PimAuditResults[$gn]
        foreach ($r in $d.Roles) {
            $auditNote = ''
            if ($auditItems) {
                $match = $auditItems | Where-Object { $_.Text -match [regex]::Escape($r) } | Select-Object -First 1
                if ($match) { $auditNote = $match.Text -replace "`n",' ' }
            }
            [PSCustomObject]@{
                Groupe      = $gn
                Type        = $d.Type
                Description = $d.Description
                Role        = $r
                Statut      = $script:PimGroupStatus[$gn]
                IsCustom    = ($r -match '\[Custom\]')
                AuditNote   = $auditNote
                ExportDate  = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            }
        }
    }
    $rows | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
    Write-Log -Level "SUCCESS" -Action "PIM_EXPORT" -Message "Rapport PIM exporte : $($sfd.FileName)"
    Show-ResultDialog -Titre (Get-Text "pim.title") `
        -Message ((Get-Text "pim.export_success") -f $sfd.FileName) `
        -IsSuccess $true
}

function Save-PimConfig {
    <#
    .SYNOPSIS
        Sauvegarde $script:PimData dans la section pim_role_groups du JSON client.
    .OUTPUTS
        [bool] -- $true si sauvegarde, $false si erreur.
    #>

    $configPath = $Config._config_path
    if (-not $configPath -or -not (Test-Path $configPath)) {
        Write-Log -Level "ERROR" -Action "PIM_SAVE" -Message "Chemin config introuvable : $configPath"
        return $false
    }

    try {
        # Relire le JSON complet
        $jsonContent = Get-Content -Path $configPath -Raw -Encoding UTF8
        $jsonObject  = $jsonContent | ConvertFrom-Json

        # Reconstruire la section pim_role_groups depuis $script:PimData
        $pimSection = [PSCustomObject]@{}
        foreach ($key in $script:PimData.Keys) {
            $def = $script:PimData[$key]
            $groupObj = [PSCustomObject]@{
                description = $def.Description
                type        = $def.Type
                roles       = @($def.Roles)
            }
            $pimSection | Add-Member -NotePropertyName $key -NotePropertyValue $groupObj -Force
        }

        # Remplacer la section dans l'objet JSON
        if ($jsonObject.PSObject.Properties["pim_role_groups"]) {
            $jsonObject.pim_role_groups = $pimSection
        }
        else {
            $jsonObject | Add-Member -NotePropertyName "pim_role_groups" -NotePropertyValue $pimSection -Force
        }

        # Ecrire le fichier
        $jsonObject | ConvertTo-Json -Depth 10 | Out-File -FilePath $configPath -Encoding UTF8 -Force

        # Mettre a jour $Config en memoire
        $Config.pim_role_groups = $pimSection

        Write-Log -Level "SUCCESS" -Action "PIM_SAVE" -Message "Configuration PIM sauvegardee ($($script:PimData.Count) groupes)."
        return $true
    }
    catch {
        Write-Log -Level "ERROR" -Action "PIM_SAVE" -Message "Erreur sauvegarde : $($_.Exception.Message)"
        return $false
    }
}

function Remove-PimGroupFromConfig {
    <#
    .SYNOPSIS
        Supprime un groupe PIM de $script:PimData (en memoire uniquement).
        Appeler Save-PimConfig ensuite pour persister.
    #>
    param([string]$GroupName)

    if ($script:PimData.Contains($GroupName)) {
        $script:PimData.Remove($GroupName)
        $script:PimGroupStatus.Remove($GroupName)
        $script:PimAuditResults.Remove($GroupName)
        Write-Log -Level "INFO" -Action "PIM_EDIT" -Message "Groupe '$GroupName' retire de la config (en memoire)."
        return $true
    }
    return $false
}

function Import-PimGroupsFromTenant {
    <#
    .SYNOPSIS
        Decouvre les groupes PIM existants dans le tenant (prefixe PIM_,
        IsAssignableToRole = true) et lit leurs roles PIM assignes
        (eligible + active) depuis les schedules Graph API.
    .PARAMETER StepCallback
        Scriptblock (message) pour afficher l'etape en cours.
    .OUTPUTS
        [ordered] hashtable { GroupName => @{ Description; Type; Roles; GroupId; IsNew } }
        ou $null si erreur / aucun resultat.
    #>
    param([scriptblock]$StepCallback)

    Write-Log -Level "INFO" -Action "PIM_IMPORT" -Message "=== Debut decouverte des groupes PIM du tenant ==="

    # --- 1. Charger les roles Entra si pas deja fait ---
    if ($script:AllEntraRoles.Count -eq 0) {
        if ($StepCallback) { & $StepCallback (Get-Text "pim.import_step_loading_roles") }
        if (-not (Load-EntraRoles)) {
            Write-Log -Level "ERROR" -Action "PIM_IMPORT" -Message "Impossible de charger les roles Entra."
            return $null
        }
    }

    # Index inverse : roleDefinitionId -> displayName (cache local rapide)
    # Couvre Id et TemplateId pour les built-in et custom
    $roleIdToName = @{}
    foreach ($entry in $script:AllEntraRoles.GetEnumerator()) {
        $roleIdToName[$entry.Value.Id]         = $entry.Key
        $roleIdToName[$entry.Value.TemplateId]  = $entry.Key
    }

    # --- 2. Chercher les groupes PIM_ dans Entra (IsAssignableToRole) ---
    if ($StepCallback) { & $StepCallback (Get-Text "pim.import_step_searching") }

    $discovered = [ordered]@{}
    try {
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=startsWith(displayName,'PIM_') and isAssignableToRole eq true&`$select=id,displayName,description&`$top=100"
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        $groups = @($resp.value)
        while ($resp.'@odata.nextLink') {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $resp.'@odata.nextLink' -ErrorAction Stop
            $groups += $resp.value
        }
    }
    catch {
        Write-Log -Level "ERROR" -Action "PIM_IMPORT" -Message "Erreur recherche groupes PIM_ : $($_.Exception.Message)"
        return $null
    }

    if ($groups.Count -eq 0) {
        Write-Log -Level "WARNING" -Action "PIM_IMPORT" -Message "Aucun groupe PIM_ trouve dans le tenant."
        return $null
    }
    Write-Log -Level "INFO" -Action "PIM_IMPORT" -Message "$($groups.Count) groupe(s) PIM_ trouves dans Entra."

    # --- 3. Charger les schedules PIM (eligible + active) ---
    if ($StepCallback) { & $StepCallback (Get-Text "pim.import_step_loading_schedules") }

    $schedules = Get-PimSchedules
    if (-not $schedules) {
        Write-Log -Level "ERROR" -Action "PIM_IMPORT" -Message "Impossible de charger les schedules PIM."
        return $null
    }

    $eligSchedules = $schedules.Eligible
    $actSchedules  = $schedules.Active

    # Index des groupes par Id pour lookup rapide
    $groupIdMap = @{}
    foreach ($g in $groups) { $groupIdMap[$g.id] = $g }

    # --- 4. Pour chaque groupe, resoudre les roles assignes ---
    foreach ($grp in ($groups | Sort-Object displayName)) {
        $gName = $grp.displayName
        if ($StepCallback) { & $StepCallback ("[$gName] " + (Get-Text "pim.import_step_resolving")) }

        # Roles eligibles pour ce groupe
        $eligRoles = @($eligSchedules | Where-Object {
            $_.principalId -eq $grp.id -and $_.status -eq 'Provisioned'
        })
        # Roles actifs pour ce groupe
        $actRoles = @($actSchedules | Where-Object {
            $_.principalId -eq $grp.id -and $_.status -eq 'Provisioned'
        })

        # Resoudre les noms de roles (avec fallback Graph pour les custom)
        $resolvedRoles = @()
        $hasEligible = $false
        $hasActive   = $false

        foreach ($sched in $eligRoles) {
            $rid = $sched.roleDefinitionId
            $rName = $roleIdToName[$rid]
            if (-not $rName) {
                $rName = Resolve-RoleName -RoleDefId $rid
                if ($rName) { $roleIdToName[$rid] = $rName }
            }
            if (-not $rName) { $rName = "Unknown ($rid)" }
            if ($rName -notin $resolvedRoles) { $resolvedRoles += $rName }
            $hasEligible = $true
        }
        foreach ($sched in $actRoles) {
            $rid = $sched.roleDefinitionId
            $rName = $roleIdToName[$rid]
            if (-not $rName) {
                $rName = Resolve-RoleName -RoleDefId $rid
                if ($rName) { $roleIdToName[$rid] = $rName }
            }
            if (-not $rName) { $rName = "Unknown ($rid)" }
            if ($rName -notin $resolvedRoles) { $resolvedRoles += $rName }
            $hasActive = $true
        }

        # Determiner le type selon le pattern de roles trouves
        $inferredType = if ($hasActive -and -not $hasEligible) { 'Groupe' }
                        elseif ($hasActive -and $hasEligible)  { 'Groupe' }
                        else { 'Role' }
        # Si le nom contient des indices de type
        if ($gName -match 'Fixe')     { $inferredType = 'Role_Fixe' }
        if ($gName -match 'Critical') { $inferredType = 'Groupe_Critical' }

        $isNew = -not $script:PimData.Contains($gName)

        $discovered[$gName] = @{
            Description = if ($grp.description) { $grp.description } else { '' }
            Type        = $inferredType
            Roles       = @($resolvedRoles | Sort-Object)
            GroupId     = $grp.id
            IsNew       = $isNew
            RoleCount   = $resolvedRoles.Count
        }

        Write-Log -Level "INFO" -Action "PIM_IMPORT" -Message "[$gName] $($resolvedRoles.Count) role(s) -- Type=$inferredType -- Nouveau=$isNew"
    }

    Write-Log -Level "SUCCESS" -Action "PIM_IMPORT" -Message "=== Decouverte terminee : $($discovered.Count) groupe(s) ==="
    return $discovered
}