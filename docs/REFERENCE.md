# Référence Projet — M365 Monster

> **Version :** 3.0
> **Date :** 2026-03-24
> **Portée :** Gestion du cycle de vie employé dans Microsoft 365 / Entra ID, avec interface graphique WinForms, multi-client, multi-langue.

---

## 1. Contexte et objectifs

M365 Monster est un outil PowerShell avec interface graphique (GUI) permettant à un administrateur IT ou un technicien helpdesk de gérer le cycle de vie des employés dans Microsoft Entra ID (Azure AD), sans ligne de commande.

L'outil est **agnostique au client** : un même set de scripts sert n'importe quelle organisation via un fichier de configuration JSON par client. Il est conçu pour les **MSP** (Managed Service Providers) qui gèrent plusieurs tenants.

---

## 2. Stack technique

| Composant | Technologie |
|---|---|
| Langage | PowerShell 7+ |
| GUI | Windows Forms (WinForms) via `[System.Windows.Forms]` |
| Annuaire | Microsoft Entra ID via **Microsoft Graph API** |
| Authentification | MSAL interactive_browser via SDK Graph (WAM depuis 2.34+) |
| Configuration client | Fichiers `.json` par client dans `Clients/` |
| Internationalisation | Multi-langue (FR/EN) via `Lang/*.json` + `Core/Lang.ps1` |
| Journalisation | Fichier `.log` horodaté par session dans `%APPDATA%` |
| Auto-update | GitHub Releases via `Core/Update.ps1` + `version.json` |
| Installation | `Install.ps1` / `Uninstall.ps1` (détection auto PS7) |
| Dépendances modules | `Microsoft.Graph` (PowerShell SDK) |

---

## 3. Architecture des fichiers

```
📁 M365Monster/
│
├── Main.ps1                        # Point d'entrée — orchestre tout
├── Install.ps1                     # Installateur
├── Uninstall.ps1                   # Désinstallateur (self-relocate vers temp)
├── version.json                    # Version courante (pour auto-update)
├── update_config.example.json      # Modèle de config auto-update (versionné)
│
├── 📁 Core/
│   ├── Config.ps1                  # Chargement et validation du JSON client
│   ├── Connect.ps1                 # Authentification Microsoft Graph + Exchange Online
│   ├── Functions.ps1               # Utilitaires (logs, mdp, dialogs, profils d'accès)
│   ├── GraphAPI.ps1                # Wrappers sur les appels Graph + Search-AzGroups
│   ├── Lang.ps1                    # Système i18n (Get-Text, Initialize-Language)
│   ├── PIMFunctions.ps1            # Logique métier PIM (audit, assignation, création)
│   └── Update.ps1                  # Auto-update depuis GitHub Releases
│
├── 📁 Modules/
│   ├── GUI_Main.ps1                # Fenêtre principale (10 tuiles)
│   ├── GUI_Onboarding.ps1          # Formulaire d'arrivée employé (+ profils d'accès)
│   ├── GUI_Offboarding.ps1         # Formulaire de départ employé
│   ├── GUI_Modification.ps1        # Formulaire de modification (+ changement profils)
│   ├── GUI_AccessProfiles.ps1      # Gestionnaire de profils d'accès + réconciliation
│   ├── GUI_SharedMailboxAudit.ps1  # Audit des boîtes partagées
│   ├── GUI_NestedGroupAudit.ps1   # Audit des groupes mixtes (Users+Devices) + Intune
│   ├── GUI_PIMManager.ps1         # Gestionnaire PIM (dashboard, éditeur, audit)
│   └── GUI_Settings.ps1            # Interface de paramétrage client
│
├── 📁 Lang/
│   ├── fr.json                     # Chaînes en français
│   └── en.json                     # Chaînes en anglais
│
├── 📁 Scripts/                     # Scripts externes (Employee Types, Stale Devices)
│   ├── AzureAD_EmployeeTypeManageGUI.ps1
│   └── AzureAD_CleanStaleDeviceGUI.ps1
│
├── 📁 Clients/
│   └── _Template.json              # Template — copier pour nouveau client
│
└── 📁 Assets/
    └── M365Monster.ico             # Icône de l'application
```

> **Note :** `update_config.json` est exclu du versioning (`.gitignore`).
> Il est créé automatiquement par `Install.ps1` à partir de `update_config.example.json`.

### Données utilisateur (hors Program Files)

```
%APPDATA%\M365Monster/
├── settings.json                   # Langue choisie
├── .last_update_check              # Horodatage de la dernière vérification MAJ
└── Logs/
    └── session_YYYY-MM-DD_HH-mm.log
```

---

## 4. Flux d'exécution — `Main.ps1`

```
1. Détermination du répertoire racine ($RootPath)
2. Déblocage NTFS des fichiers (Unblock-File)
3. Chargement WinForms
4. Auto-update (Core/Update.ps1) → redémarrage si MAJ appliquée
5. Initialisation i18n (Core/Lang.ps1) → popup choix langue si premier lancement
6. Dot-sourcing Core/ (Config, Functions, GraphAPI, Connect)
7. Initialisation logs dans %APPDATA%\M365Monster\Logs
8. Chargement anticipé de GUI_Settings (pour le sélecteur de client)
9. Sélection du client (liste déroulante + bouton "Nouveau client")
10. Chargement de la configuration JSON
11. Connexion Microsoft Graph (interactive_browser)
12. Chargement des modules GUI restants
13. Affichage de la fenêtre principale (8 tuiles)
14. Déconnexion Graph à la fermeture
```

---

## 5. Système d'internationalisation (i18n)

### Architecture

- **`Core/Lang.ps1`** : moteur i18n, expose `Get-Text "section.clé"` et `Initialize-Language`
- **`Lang/fr.json`**, **`Lang/en.json`** : fichiers de chaînes traduites (notation pointée)
- **`settings.json`** (dans `%APPDATA%`) : stocke le choix de langue

### Règles

- **Aucune chaîne GUI n'est hardcodée** — tout passe par `Get-Text`
- Au premier lancement, une popup bilingue propose le choix
- Le choix est sauvegardé dans `settings.json` pour les lancements suivants
- Supprimer `settings.json` pour re-proposer le choix

### Ajouter une langue

1. Copier `Lang/fr.json` → `Lang/xx.json`
2. Modifier `_code` et `_language` dans le nouveau fichier
3. Traduire toutes les chaînes
4. La nouvelle langue apparaît automatiquement dans le sélecteur

---

## 6. Auto-update

### Fonctionnement

1. `Main.ps1` appelle `Invoke-AutoUpdate` à chaque démarrage
2. Lit `update_config.json` (repo, branche, token, intervalle)
3. Compare `version.json` local vs `version.json` distant sur GitHub (raw)
4. Si nouvelle version disponible : popup de proposition → téléchargement du `.zip` → extraction → remplacement des fichiers → redémarrage

### Configuration (`update_config.json`)

```json
{
  "github_repo": "valtobech/M365_Monster",
  "branch": "main",
  "github_token": "",
  "download_url": "",
  "check_interval_hours": 0
}
```

- `check_interval_hours: 0` = vérification à chaque lancement
- `check_interval_hours: 24` = vérification toutes les 24h
- `github_token` = uniquement pour repo privé
- `download_url` = laisser vide pour utiliser GitHub Releases automatiquement

### Éléments préservés lors des mises à jour

- `Clients/` (configurations client)
- `update_config.json`
- `settings.json` (dans AppData)
- `Logs/` (dans AppData)

### Publier une nouvelle version

Voir `docs/RELEASE_PROCESS.md`.

---

## 7. Installation et désinstallation

### Install.ps1

- Détecte `pwsh.exe` (PS7) en priorité pour les raccourcis
- Copie : `Main.ps1`, `version.json`, `Core/`, `Modules/`, `Scripts/`, `Assets/`, `Lang/`
- Crée `Clients/` avec `_Template.json` uniquement
- Crée `update_config.json` automatiquement depuis `update_config.example.json`
- Raccourcis Bureau + Menu Démarrer avec icône
- Auto-update activé par défaut, aucune intervention requise

#### Options

```powershell
.\Install.ps1 -InstallPath "D:\Outils\M365Monster"    # Chemin custom
.\Install.ps1 -SkipModules                              # Sans install modules PS
.\Install.ps1 -SkipShortcuts                            # Sans raccourcis
.\Install.ps1 -SkipUpdateConfig                         # Sans config auto-update
```

### Uninstall.ps1

- Se relance depuis `%TEMP%` pour pouvoir supprimer son propre répertoire
- Change le répertoire courant (`Set-Location`) pour libérer le verrou
- Propose la conservation des fichiers `Clients/`
- Supprime aussi `%APPDATA%\M365Monster` (logs, settings)

---

## 8. Conventions de code

| Convention | Détail |
|---|---|
| Commentaires | En **français** |
| Nommage fonctions | `Verbe-Nom` (PowerShell approved verbs) |
| Chaînes GUI | Via `Get-Text "section.clé"` (jamais hardcodées) |
| Variables partagées | `$global:` ou `$script:` selon le scope |
| Gestion d'erreur | `try/catch` + `Write-Log` sur chaque appel Graph |
| Sécurité | Jamais de mot de passe en clair dans les logs ou fichiers |
| Logs | Écrits dans `%APPDATA%\M365Monster\Logs` |
| PowerShell | PS7 requis, détection auto dans Install/Main |

---

---

## 9. Permissions API Microsoft Graph

> Toutes les permissions sont de type **Délégué** (`Delegated`) — connexion interactive uniquement.
> Admin consent requis sur chaque tenant client.

| Permission | Usage dans M365 Monster |
|---|---|
| `User.ReadWrite.All` | Créer, modifier (profil, téléphones, UPN), désactiver/réactiver des comptes |
| `Group.ReadWrite.All` | Ajouter/retirer des utilisateurs des groupes (licences, sécurité) ; créer des groupes (remediation nested) |
| `Directory.ReadWrite.All` | Lire les domaines vérifiés du tenant, accès annuaire étendu |
| `Mail.Send` | Envoyer les notifications email via `/me/sendMail` |
| `UserAuthenticationMethod.ReadWrite.All` | Lire et supprimer les méthodes MFA (module Modification — Reset MFA) |
| `AuditLog.Read.All` | Lire les journaux de connexion (module Modification — Dernières connexions ; Shared Mailbox — last sign-in) |
| `Device.Read.All` | Lire les devices Entra (module Nested Group Audit — classification des membres) |
| `DeviceManagementConfiguration.Read.All` | Lire les policies Intune : configuration, compliance, ADMX, Autopilot, updates (module Nested Group Audit) |
| `DeviceManagementApps.Read.All` | Lire les applications Intune et leurs assignations (module Nested Group Audit) |
| `DeviceManagementManagedDevices.Read.All` | Lire les devices managés Intune et les scripts de remédiation (module Nested Group Audit) |
| `RoleManagement.ReadWrite.Directory` | Gérer les rôles PIM : lire les roleDefinitions, créer/supprimer des roleEligibilityScheduleRequests et roleAssignmentScheduleRequests (module PIM Manager) |

### Notes importantes

- **Téléphones et alias email** : `Update-MgUser` est bloqué par Exchange Online sur `mobilePhone`, `businessPhones` et `proxyAddresses`. L'outil utilise `Invoke-MgGraphRequest PATCH` directement sur `/v1.0/users/{id}` pour contourner cette restriction.
- **Token en cache** : si `Forbidden (403)` apparaît après ajout d'un scope, fermer et relancer l'outil pour forcer un nouveau token.
- **proxyAddresses** : Exchange Online gère les alias de façon autonome. L'ajout/suppression via Graph fonctionne uniquement si la boîte Exchange Online est active et que le compte connecté a les droits suffisants.
- **Endpoints Intune (beta)** : le module Nested Group Audit utilise les endpoints `beta` de Microsoft Graph pour les policies Intune (`/beta/deviceManagement/...`). Ces endpoints peuvent évoluer sans préavis. Chaque catégorie est scannée dans un `try/catch` individuel pour garantir la résilience.
- **Graph Batch API** : le scan des groupes utilise `/$batch` (paquets de 20 requêtes parallèles) pour accélérer l'analyse des membres. Anti-throttling de 150ms entre chaque lot.
- **Renommage de groupes** : le module Nested Group Audit utilise `Update-MgGroup` pour renommer le groupe d'origine lors de la séparation Users/Devices. Le `mailNickname` est mis à jour simultanément (caractères non-alphanumériques supprimés).
- **Suppression de membres** : `Remove-MgGroupMemberByRef` est utilisé pour retirer les membres transférés du groupe source. L'opération est unitaire (un appel par membre) avec progression visuelle et compteur d'erreurs.
- **Profils d'accès** : les profils sont stockés dans `access_profiles` du JSON client. La réconciliation bidirectionnelle utilise `_pending_removals` pour détecter les groupes retirés d'un template. Diff intelligent : seuls les ajouts/retraits nécessaires sont exécutés.
- **PIM — Création de groupes** : `IsAssignableToRole = true` est irréversible sur un groupe Entra. Le module exige une double confirmation. Le polling de réplication Entra (5s × 12 = 60s max) attend que le groupe soit visible avant l'assignation.
- **PIM — Assignation de rôles** : utilise `roleEligibilityScheduleRequests` (eligible) ou `roleAssignmentScheduleRequests` (active) selon le type de groupe. Retry automatique sur `SubjectNotFound` (3 tentatives × 10s). Fallback `noExpiration` → `afterDateTime` si le tenant impose une durée maximale.
- **PIM — Chargement dynamique des rôles** : tous les rôles built-in et custom du tenant sont chargés via `/v1.0/roleManagement/directory/roleDefinitions`. Aucune map statique à maintenir.
- **Employee Type — Session partagée** : le script externe `AzureAD_EmployeeTypeManageGUI.ps1` détecte `Get-MgContext` au lancement. Si une session existe (lancé depuis M365 Monster), `$script:ExternalSession = $true` et `Disconnect-MgGraph` est ignoré à la fermeture.

---

## 10. Historique des versions

Voir [CHANGELOG.md](CHANGELOG.md) pour le détail complet de chaque version.

| Version | Date | Résumé |
|---|---|---|
| `0.1.8` | 2026-03-24 | Nouveau module PIM Manager + Employee Type session Graph partagée |
| `0.1.7` | 2026-02-28 | Nouveau module Profils d'accès (composables, réconciliation bidirectionnelle) |
| `0.1.6` | 2026-02-25 | Performance SharedMailboxAudit : architecture 2 passes $filter (6min → 10s) |
| `0.1.5` | 2026-02-25 | Audit Nested : renommage croisé groupe d'origine + suppression membres transférés |
| `0.1.4` | 2026-02-25 | Nouveau module Audit Groupes Nested (Users+Devices) avec scan Intune |
| `0.1.3` | 2026-02-23 | Alias email via Exchange Online (Set-Mailbox), connexion EXO, Shared Mailbox Audit |
| `0.1.2` | 2026-02-22 | Corrections module Modification : alias, téléphones, groupes, UX |
| `0.1.1` | 2026-02-22 | Corrections UX module Modification : scroll, combos, permissions |
| `0.1.0` | 2026-02-22 | Version bêta initiale |

---

*Fin du document de référence*