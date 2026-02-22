<#
.FICHIER
    Core/Connect.ps1

.ROLE
    Gère l'authentification à Microsoft Graph via le module Microsoft.Graph SDK.
    Supporte trois modes d'authentification :
      - interactive_browser : popup navigateur avec login utilisateur (RECOMMANDÉ)
      - device_code : code à saisir sur https://microsoft.com/devicelogin
      - client_secret : Service Principal (secret dans le JSON — déconseillé)

.DEPENDANCES
    - Module Microsoft.Graph (PowerShell SDK)
    - Core/Config.ps1 (variable $Config chargée)
    - Core/Functions.ps1 (Write-Log)

.AUTEUR
    [Équipe IT — GestionRH-AzureAD]
#>

function Test-GraphModule {
    <#
    .SYNOPSIS
        Vérifie si le module Microsoft.Graph est installé.
        Propose l'installation si absent.

    .OUTPUTS
        [bool] — $true si le module est disponible, $false sinon.
    #>
    [CmdletBinding()]
    param()

    $module = Get-Module -ListAvailable -Name "Microsoft.Graph.Authentication" -ErrorAction SilentlyContinue

    if (-not $module) {
        Write-Log -Level "WARNING" -Action "PREREQUIS" -Message "Module Microsoft.Graph non trouvé."

        $reponse = [System.Windows.Forms.MessageBox]::Show(
            "Le module Microsoft.Graph n'est pas installé.`nVoulez-vous l'installer maintenant ?`n`n(Commande : Install-Module Microsoft.Graph -Scope CurrentUser)",
            "Module manquant",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($reponse -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Write-Log -Level "INFO" -Action "INSTALL" -Message "Installation du module Microsoft.Graph en cours..."
                Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
                Write-Log -Level "SUCCESS" -Action "INSTALL" -Message "Module Microsoft.Graph installé avec succès."
                return $true
            }
            catch {
                Write-Log -Level "ERROR" -Action "INSTALL" -Message "Échec de l'installation : $($_.Exception.Message)"
                [System.Windows.Forms.MessageBox]::Show(
                    "Échec de l'installation du module.`n$($_.Exception.Message)",
                    "Erreur",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return $false
            }
        }
        else {
            return $false
        }
    }

    return $true
}

function Connect-GraphAPI {
    <#
    .SYNOPSIS
        Établit la connexion à Microsoft Graph en utilisant la méthode
        définie dans la configuration client.

    .DESCRIPTION
        Trois méthodes supportées :

        1. interactive_browser (RECOMMANDÉ) :
           Ouvre une fenêtre de navigateur pour que le technicien se connecte
           avec son propre compte Microsoft (delegated access).
           Aucun secret stocké nulle part — le token est éphémère en mémoire.
           Compatible avec le MFA et le Conditional Access du tenant.

        2. device_code :
           Affiche un code dans la console, l'utilisateur va sur
           https://microsoft.com/devicelogin pour s'authentifier.
           Utile si le Device Code Flow n'est PAS bloqué par le tenant.

        3. client_secret :
           Auth silencieuse par Service Principal — nécessite un secret en
           clair dans le JSON. Déconseillé sauf cas automatisé isolé.

    .OUTPUTS
        [PSCustomObject] — {Success: bool, Error: string}
        En cas de succès, le contexte Graph est actif pour les appels suivants.
    #>
    [CmdletBinding()]
    param()

    # Vérification du module
    if (-not (Test-GraphModule)) {
        return [PSCustomObject]@{ Success = $false; Error = "Module Microsoft.Graph non disponible." }
    }

    # Scopes nécessaires pour les opérations RH (delegated permissions)
    $scopes = @(
        "User.ReadWrite.All",
        "Group.ReadWrite.All",
        "Directory.ReadWrite.All",
        "Mail.Send"
    )

    $tenantId = $Config.tenant_id
    $clientId = $Config.client_id
    $authMethod = $Config.auth_method

    Write-Log -Level "INFO" -Action "CONNEXION" -Message "Connexion Graph en cours — Méthode: $authMethod, Tenant: $tenantId"

    try {
        switch ($authMethod) {
            "interactive_browser" {
                # MÉTHODE RECOMMANDÉE — Delegated access via popup navigateur
                # Le SDK Microsoft.Graph ouvre automatiquement le navigateur par défaut.
                # L'utilisateur se connecte avec son compte, le MFA est géré par Entra ID.
                # Aucun secret n'est stocké — le token est éphémère en mémoire.
                # Le paramètre -ClientId pointe vers l'App Registration du projet,
                # qui doit avoir "Allow public client flows" activé et la redirect URI
                # http://localhost configurée.
                Connect-MgGraph -TenantId $tenantId -ClientId $clientId -Scopes $scopes -NoWelcome -ErrorAction Stop

                Write-Log -Level "SUCCESS" -Action "CONNEXION" -Message "Connexion Graph réussie (interactive_browser)."
            }

            "device_code" {
                # Auth par Device Code — l'utilisateur doit aller sur https://microsoft.com/devicelogin
                # Souvent bloqué par les politiques Conditional Access des tenants clients.
                Connect-MgGraph -TenantId $tenantId -ClientId $clientId -Scopes $scopes -UseDeviceCode -NoWelcome -ErrorAction Stop

                Write-Log -Level "SUCCESS" -Action "CONNEXION" -Message "Connexion Graph réussie (device_code)."
            }

            "client_secret" {
                # Authentification par Service Principal — secret dans le JSON
                # DÉCONSEILLÉ pour un usage interactif — réservé à l'automatisation
                if ([string]::IsNullOrWhiteSpace($Config.client_secret)) {
                    throw "client_secret requis pour l'authentification Service Principal mais non renseigné dans la configuration."
                }

                $secureSecret = ConvertTo-SecureString -String $Config.client_secret -AsPlainText -Force
                $credential = New-Object System.Management.Automation.PSCredential($clientId, $secureSecret)

                Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $credential -NoWelcome -ErrorAction Stop

                Write-Log -Level "SUCCESS" -Action "CONNEXION" -Message "Connexion Graph réussie (client_secret)."
            }

            default {
                throw "Méthode d'authentification inconnue : '$authMethod'. Valeurs acceptées : interactive_browser, device_code, client_secret"
            }
        }

        # Vérification post-connexion
        $context = Get-MgContext
        if ($null -eq $context) {
            throw "La connexion semble avoir réussi mais aucun contexte Graph n'est actif."
        }

        # Stockage du contexte en variable globale pour référence
        $script:GraphContext = $context

        return [PSCustomObject]@{ Success = $true; Error = $null }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log -Level "ERROR" -Action "CONNEXION" -Message "Échec de connexion Graph : $errMsg"
        return [PSCustomObject]@{ Success = $false; Error = $errMsg }
    }
}

function Disconnect-GraphAPI {
    <#
    .SYNOPSIS
        Déconnecte la session Microsoft Graph active.

    .OUTPUTS
        [void]
    #>
    [CmdletBinding()]
    param()

    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        $script:GraphContext = $null
        Write-Log -Level "INFO" -Action "DECONNEXION" -Message "Session Graph déconnectée."
    }
    catch {
        Write-Log -Level "WARNING" -Action "DECONNEXION" -Message "Erreur lors de la déconnexion : $($_.Exception.Message)"
    }
}

function Get-GraphConnectionStatus {
    <#
    .SYNOPSIS
        Retourne l'état de la connexion Graph.

    .OUTPUTS
        [bool] — $true si connecté, $false sinon.
    #>
    [CmdletBinding()]
    param()

    try {
        $context = Get-MgContext
        return ($null -ne $context)
    }
    catch {
        return $false
    }
}

# Point d'attention :
# - interactive_browser est la méthode RECOMMANDÉE : popup navigateur, zéro secret, MFA compatible
# - device_code est conservé mais souvent bloqué chez les clients (Conditional Access)
# - client_secret est conservé pour les cas d'automatisation isolés — secret en clair dans le JSON
# - Les scopes sont définis en dur car ils correspondent aux permissions requises par l'outil
# - -NoWelcome supprime le message "Welcome to Microsoft Graph!" dans la console
