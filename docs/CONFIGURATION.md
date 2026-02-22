# Guide de configuration — M365 Monster

> **Version :** 2.0 — Février 2026
> **Public cible :** Techniciens IT / MSP qui déploient l'outil chez un nouveau client

---

## Table des matières

1. [Vue d'ensemble](#1-vue-densemble)
2. [Prérequis Azure](#2-prérequis-azure)
3. [Créer l'App Registration dans Entra ID](#3-créer-lapp-registration-dans-entra-id)
4. [Méthodes d'authentification](#4-méthodes-dauthentification)
5. [Référence complète du fichier JSON](#5-référence-complète-du-fichier-json)
6. [Exemple complet commenté](#6-exemple-complet-commenté)
7. [Trouver les SKU de licences](#7-trouver-les-sku-de-licences)
8. [Dépannage](#8-dépannage)

---

## 1. Vue d'ensemble

Chaque client géré par l'outil a son propre fichier `.json` dans le dossier `Clients/`. Ce fichier contient **toute** la configuration spécifique au client : identifiants Azure, domaine mail, groupes, licences, départements, etc.

**Aucune valeur n'est codée en dur dans les scripts** — tout passe par ce JSON.

Pour ajouter un nouveau client :

1. **Via l'interface** (recommandé) : à l'écran de sélection du client, cliquer **⚙ Nouveau client / Paramétrage**
2. **Manuellement** : copier `Clients/_Template.json` → `Clients/MonClient.json` et remplir les valeurs

---

## 2. Prérequis Azure

Avant de configurer le JSON, vous avez besoin de :

- Un accès **Administrateur** (ou Administrateur d'applications) sur le tenant Entra ID du client
- Le **Tenant ID** du client (visible dans Entra ID → Vue d'ensemble)
- Installer le module PowerShell :

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
```

---

## 3. Créer l'App Registration dans Entra ID

L'outil a besoin d'une **App Registration** sur le tenant du client pour s'authentifier. Voici la procédure pas à pas.

### 3.1 Créer l'application

1. Aller dans **Microsoft Entra ID → App registrations → + New registration**
2. Remplir :
   - **Name** : `M365 Monster`
   - **Supported account types** : `Accounts in this organizational directory only`
   - **Redirect URI** :
     - Plateforme : **Mobile and desktop applications**
     - URI : `http://localhost`
3. Cliquer **Register**
4. **Noter le Application (client) ID** — c'est votre `client_id` pour le JSON

### 3.2 Ajouter le redirect URI WAM (obligatoire depuis SDK 2.34+)

Depuis décembre 2025, le SDK Microsoft.Graph utilise le WAM (Web Account Manager) par défaut. Un second redirect URI est obligatoire :

1. Dans l'App Registration → **Authentication** → **Add a platform** → **Mobile and desktop**
2. Ajouter : `ms-appx-web://Microsoft.AAD.BrokerPlugin/<votre-client-id>`
   - Remplacez `<votre-client-id>` par l'Application (client) ID noté à l'étape 3.1
3. **Sauvegarder**

> ⚠️ Sans ce redirect URI, l'erreur `AADSTS50011` apparaîtra lors de la connexion. M365 Monster affichera un message d'aide avec l'URI à ajouter et proposera de la copier dans le presse-papier.

### 3.3 Activer le flux public (obligatoire pour interactive_browser)

C'est l'étape la plus souvent oubliée :

1. Dans l'App Registration → **Authentication**
2. Tout en bas de la page, section **Advanced settings**
3. Mettre **"Allow public client flows"** sur **Yes**
4. **Sauvegarder**

> ⚠️ Sans cette option, la connexion interactive échouera avec une erreur AADSTS…

### 3.3 Configurer les permissions API

1. Aller dans **API permissions → + Add a permission → Microsoft Graph**
2. Choisir **Delegated permissions** et ajouter :

| Permission | Usage |
|---|---|
| `User.ReadWrite.All` | Créer, modifier, désactiver des comptes |
| `Group.ReadWrite.All` | Ajouter/retirer des utilisateurs des groupes |
| `Directory.ReadWrite.All` | Accéder aux informations du tenant |
| `Mail.Send` | Envoyer les notifications email |

3. Cliquer **Grant admin consent for [Tenant]** (bouton bleu en haut)
4. Vérifier que chaque permission affiche ✅ **Granted**

### 3.5 Récapitulatif — ce qu'il faut noter

| Information | Où la trouver | Champ JSON |
|---|---|---|
| Tenant ID | Entra ID → Overview | `tenant_id` |
| Application (client) ID | App Registration → Overview | `client_id` |
| Allow public client flows | App Registration → Authentication | Doit être **Yes** |
| Redirect URI `http://localhost` | App Registration → Authentication | Obligatoire |
| Redirect URI WAM `ms-appx-web://...` | App Registration → Authentication | Obligatoire (SDK 2.34+) |

**Vous n'avez PAS besoin de créer un secret.** La méthode `interactive_browser` n'en utilise pas.

---

## 4. Méthodes d'authentification

### 4.1 `interactive_browser` — ✅ RECOMMANDÉ

```json
"auth_method": "interactive_browser"
```

**Comment ça marche :**
- Au lancement de l'outil, une fenêtre de navigateur s'ouvre automatiquement
- Le technicien se connecte avec son propre compte Microsoft (celui qui a les droits admin)
- Le MFA du tenant est géré nativement par Entra ID
- Un token éphémère est créé en mémoire — rien n'est stocké sur disque
- Le token est automatiquement rafraîchi par le SDK tant que la session est ouverte

**Pourquoi c'est la bonne méthode :**
- **Zéro secret dans les fichiers** — rien à protéger, rien à faire fuiter
- **Compatible MFA et Conditional Access** — fonctionne même avec les politiques les plus strictes
- **Chaque action est tracée** au nom de l'utilisateur connecté (audit trail)
- **Fonctionne partout** — aucun flux bloqué par les politiques tenant (contrairement au device_code)

**Prérequis côté tenant :**
- App Registration avec "Allow public client flows" activé
- Redirect URI `http://localhost` configurée (plateforme Mobile and desktop)
- Redirect URI WAM `ms-appx-web://Microsoft.AAD.BrokerPlugin/<client-id>` (obligatoire SDK 2.34+)
- Admin consent accordé sur les permissions

**Prérequis côté poste technicien :**
- **PowerShell 7+** (recommandé pour la compatibilité WAM)
- Un navigateur par défaut configuré (Edge, Chrome, Firefox…)
- Le compte du technicien doit avoir un rôle admin suffisant (User Administrator minimum)

### 4.2 `device_code` — ⚠️ Souvent bloqué

```json
"auth_method": "device_code"
```

**Comment ça marche :**
- Un code s'affiche dans la console PowerShell
- Le technicien ouvre `https://microsoft.com/devicelogin` dans un navigateur et saisit le code
- L'authentification se fait dans le navigateur

**Pourquoi c'est souvent bloqué :**
- De nombreux tenants ont des **politiques Conditional Access** qui bloquent le Device Code Flow
- C'est considéré comme un vecteur d'attaque (phishing par Device Code)
- Microsoft recommande de le désactiver dans les environnements de production

**Quand l'utiliser :**
- Si le poste du technicien n'a pas de navigateur par défaut fonctionnel
- Si vous devez vous connecter depuis un serveur sans GUI

### 4.3 `client_secret` — ❌ Déconseillé

```json
"auth_method": "client_secret",
"client_secret": "votre-secret-ici"
```

**Pourquoi c'est déconseillé :**
- Le secret est **en clair dans le fichier JSON** sur le disque
- Toute personne ayant accès au fichier a accès au tenant
- Le secret expire (1 ou 2 ans max) — maintenance à prévoir
- Pas d'audit trail nominatif — les actions sont au nom de l'application, pas du technicien
- Nécessite des **Application permissions** (pas Delegated) — accès plus large

**Quand l'utiliser malgré tout :**
- Automatisation headless (script planifié, sans interaction humaine)
- Si utilisé, protéger le fichier JSON par des permissions NTFS restrictives

### 4.4 Tableau comparatif

| Critère | `interactive_browser` | `device_code` | `client_secret` |
|---|---|---|---|
| Secret sur disque | ❌ Aucun | ❌ Aucun | ⚠️ Oui, en clair |
| Compatible MFA | ✅ Natif | ✅ Natif | ❌ Non applicable |
| Conditional Access | ✅ Compatible | ⚠️ Souvent bloqué | ❌ Ignoré |
| Audit nominatif | ✅ Oui | ✅ Oui | ❌ Au nom de l'app |
| Interaction requise | Popup navigateur | Code + navigateur | Aucune |
| Cas d'usage | **Usage quotidien** | Fallback | Automatisation |

---

## 5. Référence complète du fichier JSON

### Identité du client

| Champ | Type | Obligatoire | Description |
|---|---|---|---|
| `client_name` | string | ✅ | Nom affiché dans l'interface (libre) |
| `tenant_id` | string (GUID) | ✅ | Tenant ID Entra ID du client |
| `client_id` | string (GUID) | ✅ | Application ID de l'App Registration |
| `auth_method` | string | ✅ | `interactive_browser`, `device_code` ou `client_secret` |
| `client_secret` | string | ❌ | Uniquement si `auth_method` = `client_secret` |

### Nommage et domaine

| Champ | Type | Obligatoire | Description |
|---|---|---|---|
| `smtp_domain` | string | ✅ | Domaine email — **doit commencer par `@`** (ex: `@acmecorp.com`) |
| `naming_convention` | string | ✅ | Patron de génération du UPN. Variables : `{prenom}`, `{nom}` |

**Exemples de conventions de nommage :**

| Convention | Résultat pour "Jean-Pierre Côté" |
|---|---|
| `{prenom}.{nom}` | `jean-pierre.cote@domaine.com` |
| `{prenom}{nom}` | `jean-pierrecote@domaine.com` |
| `{nom}.{prenom}` | `cote.jean-pierre@domaine.com` |

> Les accents sont automatiquement retirés (é→e, ç→c, etc.) et tout est mis en minuscules.

### Groupes par défaut

```json
"default_groups": [
  "Tous-Employes",
  "VPN-Access",
  "Teams-General"
]
```

Liste des **noms exacts** (displayName) des groupes Entra ID auxquels chaque nouvel employé est ajouté automatiquement à l'onboarding. Les groupes doivent déjà exister dans le tenant.

### Licences

```json
"license_map": {
  "cdi":         "AcmeCorp:ENTERPRISEPACK",
  "cdd":         "AcmeCorp:ENTERPRISEPACK",
  "stagiaire":   "AcmeCorp:DESKLESSPACK",
  "contractuel": "AcmeCorp:ENTERPRISEPACK"
}
```

**Clés** : en minuscules, doivent correspondre aux valeurs de `contract_types` (aussi en minuscules).

**Valeurs** : au format `TenantName:SKU_PART_NUMBER`. Voir [section 7](#7-trouver-les-sku-de-licences) pour trouver les bons SKU.

### Départements et types de contrat

```json
"departments": ["Informatique", "RH", "Finance", "Marketing"],
"contract_types": ["CDI", "CDD", "Stagiaire", "Contractuel"]
```

Ces listes alimentent les menus déroulants dans les formulaires. Adaptez-les à l'organigramme du client.

> **Important :** les valeurs de `contract_types` doivent matcher les clés de `license_map` (en minuscules). Exemple : `"CDI"` dans contract_types → clé `"cdi"` dans license_map.

### Configuration offboarding

```json
"offboarding": {
  "disabled_ou_group": "Comptes-Désactivés",
  "mailbox_forward_to": "",
  "revoke_licenses": true,
  "remove_all_groups": true,
  "retention_days": 30
}
```

| Champ | Description |
|---|---|
| `disabled_ou_group` | Nom du groupe où placer les comptes désactivés (doit exister) |
| `mailbox_forward_to` | Email de redirection par défaut (vide = pas de redirection par défaut) |
| `revoke_licenses` | Pré-cocher l'option de révocation des licences dans le formulaire |
| `remove_all_groups` | Pré-cocher l'option de retrait des groupes |
| `retention_days` | Indicatif — nombre de jours avant suppression définitive (non appliqué automatiquement) |

### Notifications

```json
"notifications": {
  "enabled": true,
  "recipients": ["it@acmecorp.com", "rh@acmecorp.com"]
}
```

Si `enabled` est `true`, un email est envoyé via Microsoft Graph aux destinataires listés après chaque onboarding et offboarding. L'email est envoyé **au nom de l'utilisateur connecté** (celui qui s'est authentifié) — ce compte doit avoir une boîte mail Exchange Online.

### Politique de mot de passe

```json
"password_policy": {
  "length": 16,
  "force_change_at_login": true,
  "include_special_chars": true
}
```

| Champ | Description |
|---|---|
| `length` | Longueur du mot de passe généré (minimum 8) |
| `force_change_at_login` | L'utilisateur devra changer son mot de passe à la première connexion |
| `include_special_chars` | Inclure des caractères spéciaux (`!@#$%&*-_=+`) |

---

## 6. Exemple complet commenté

Voici un fichier prêt à l'emploi pour un client fictif. Copiez-le, remplacez les valeurs en `⬅` :

```json
{
  "client_name": "Brasserie Montréal Inc.",
  "tenant_id": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "client_id": "98765432-abcd-1234-efgh-567890abcdef",

  "auth_method": "interactive_browser",

  "smtp_domain": "@brasseriemtl.com",
  "naming_convention": "{prenom}.{nom}",

  "default_groups": [
    "Tous-Employes",
    "VPN-Access"
  ],

  "license_map": {
    "cdi":         "BrasserieMTL:ENTERPRISEPACK",
    "cdd":         "BrasserieMTL:ENTERPRISEPACK",
    "stagiaire":   "BrasserieMTL:DESKLESSPACK",
    "contractuel": "BrasserieMTL:ENTERPRISEPACK"
  },

  "departments": [
    "Production",
    "Comptabilité",
    "Ventes",
    "Logistique",
    "Direction"
  ],

  "contract_types": [
    "CDI",
    "CDD",
    "Stagiaire",
    "Contractuel"
  ],

  "offboarding": {
    "disabled_ou_group": "Anciens-Employes",
    "mailbox_forward_to": "",
    "revoke_licenses": true,
    "remove_all_groups": true,
    "retention_days": 30
  },

  "notifications": {
    "enabled": true,
    "recipients": [
      "it@brasseriemtl.com",
      "rh@brasseriemtl.com"
    ]
  },

  "password_policy": {
    "length": 16,
    "force_change_at_login": true,
    "include_special_chars": true
  }
}
```

---

## 7. Trouver les SKU de licences

Les SKU de licences ne sont pas évidents. Voici comment les trouver pour un tenant.

### Méthode 1 — Via PowerShell (recommandé)

Après connexion avec l'outil ou manuellement :

```powershell
Connect-MgGraph -Scopes "Organization.Read.All"
Get-MgSubscribedSku | Select-Object SkuPartNumber, SkuId, ConsumedUnits | Format-Table
```

### Méthode 2 — Noms courants

| Nom commercial | SKU Part Number |
|---|---|
| Microsoft 365 E3 | `SPE_E3` |
| Microsoft 365 E5 | `SPE_E5` |
| Office 365 E1 | `STANDARDPACK` |
| Office 365 E3 | `ENTERPRISEPACK` |
| Office 365 E5 | `ENTERPRISEPREMIUM` |
| Microsoft 365 Business Basic | `O365_BUSINESS_ESSENTIALS` |
| Microsoft 365 Business Standard | `O365_BUSINESS_PREMIUM` |
| Microsoft 365 Business Premium | `SPB` |
| Microsoft 365 F1 | `M365_F1_COMM` |
| Microsoft 365 F3 | `SPE_F1` |
| Office 365 F3 (Firstline / Deskless) | `DESKLESSPACK` |
| Exchange Online Plan 1 | `EXCHANGESTANDARD` |
| Exchange Online Plan 2 | `EXCHANGEENTERPRISE` |

> La liste complète est maintenue par Microsoft : chercher "Microsoft 365 licensing service plan reference" sur learn.microsoft.com.

### Format dans le JSON

Le format est `NomTenant:SKU_PART_NUMBER`. Le `NomTenant` est souvent le nom de l'organisation tel qu'affiché dans Entra ID (sans espaces). En pratique, seule la partie après `:` est utilisée par le script pour matcher avec `Get-MgSubscribedSku`.

---

## 8. Dépannage

### "AADSTS50011: The redirect URI 'ms-appx-web://...' does not match"

**Cause** : Le redirect URI WAM n'est pas configuré dans l'App Registration. Depuis le SDK Microsoft.Graph 2.34+ (décembre 2025), le WAM est obligatoire.

**Fix** : Ajouter `ms-appx-web://Microsoft.AAD.BrokerPlugin/<votre-client-id>` dans App Registration → Authentication → Mobile and desktop. M365 Monster affiche l'URI exacte à ajouter et propose de la copier.

### "InteractiveBrowserCredential authentication failed"

**Cause** : "Allow public client flows" n'est pas activé sur l'App Registration.

**Fix** : App Registration → Authentication → Advanced settings → Allow public client flows → **Yes** → Save

### "AADSTS50126: Invalid username or password"

**Cause** : Le compte utilisé n'existe pas sur ce tenant ou le mot de passe est incorrect.

**Fix** : Vérifier que vous vous connectez avec un compte du bon tenant.

### "Insufficient privileges to complete the operation"

**Cause** : Le compte connecté n'a pas les droits admin nécessaires, OU les permissions API n'ont pas reçu le consentement admin.

**Fix** :
1. Vérifier que les 4 permissions Delegated sont accordées avec admin consent
2. Vérifier que le compte a au minimum le rôle **User Administrator**

### "Could not load type from assembly" (PS 5.1)

**Cause** : Conflit de version du module Microsoft.Graph ou utilisation de PowerShell 5.1.

**Fix** :
```powershell
# Recommandé : utiliser PowerShell 7
winget install Microsoft.PowerShell

# Ou réinstaller le module
Uninstall-Module Microsoft.Graph -AllVersions -Force
Install-Module Microsoft.Graph -Scope CurrentUser -Force
```

### Le raccourci Bureau lance PowerShell 5.1 au lieu de 7

**Cause** : L'installation a été faite avec un ancien Install.ps1 qui hardcodait `powershell.exe`.

**Fix** : Réinstaller avec le nouvel Install.ps1 qui détecte automatiquement `pwsh.exe`.

### Erreur d'accès en écriture dans Program Files

**Cause** : Ancienne version qui écrivait les logs dans le dossier d'installation.

**Fix** : Mettre à jour vers la dernière version. Les logs sont désormais dans `%APPDATA%\M365Monster\Logs`.

### La langue n'est pas proposée au démarrage

**Cause** : Un `settings.json` existe déjà avec un choix de langue.

**Fix** : Supprimer `%APPDATA%\M365Monster\settings.json` pour re-proposer le choix.

### "Groupe 'X' introuvable dans Azure AD"

**Cause** : Le nom du groupe dans `default_groups` ou `disabled_ou_group` ne correspond pas exactement au `displayName` dans Entra ID.

**Fix** : Vérifier l'orthographe exacte (case sensitive) du groupe dans Entra ID → Groups.

---

*Fin du guide de configuration*
