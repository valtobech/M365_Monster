# Référence Projet — M365 Monster

> **Version :** 2.5
> **Date :** 2026-02-28
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
| Dépendances modules | `Microsoft.Graph` (PowerShell SDK), `ExchangeOnlineManagement` |

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
│   ├── GraphAPI.ps1                # Wrappers sur les appels Graph (dont Search-AzGroups)
│   ├── Lang.ps1                    # Système i18n (Get-Text, Initialize-Language)
│   └── Update.ps1                  # Auto-update depuis GitHub Releases
│
├── 📁 Modules/
│   ├── GUI_Main.ps1                # Fenêtre principale (8 tuiles)
│   ├── GUI_Onboarding.ps1          # Formulaire d'arrivée employé (+ profils d'accès)
│   ├── GUI_Offboarding.ps1         # Formulaire de départ employé
│   ├── GUI_Modification.ps1        # Formulaire de modification (+ changement de profils)
│   ├── GUI_AccessProfiles.ps1      # Gestionnaire de profils d'accès + réconciliation
│   ├── GUI_SharedMailboxAudit.ps1  # Audit des boîtes partagées
│   ├── GUI_NestedGroupAudit.ps1   # Audit des groupes mixtes (Users+Devices) + Intune
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
11. Connexion Microsoft Graph (interactive_browser) + Exchange Online
12. Chargement des modules GUI restants (dont GUI_AccessProfiles)
13. Affichage de la fenêtre principale (8 tuiles)
14. Déconnexion Graph + EXO à la fermeture
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

## 6. Profils d'accès

### Concept

Les profils d'accès sont des packages composables de groupes Entra ID. Chaque profil regroupe un ensemble logique de groupes (ex : « Finance » = 3 groupes spécifiques). Le système permet de standardiser les attributions de groupes lors de l'onboarding et de détecter/corriger les écarts en production.

### Architecture

- **Stockage** : section `access_profiles` dans le JSON client. Chaque profil a une clé unique, un `display_name`, une `description`, un flag `is_baseline`, et un tableau `groups` (objets `{id, display_name}`).
- **Profil baseline** : marqué `is_baseline: true`, appliqué automatiquement à tous les employés (onboarding et réconciliation). Un seul baseline par client.
- **Profils additionnels** : sélectionnables individuellement lors de l'onboarding ou de la modification.

### Fonctions backend (`Core/Functions.ps1`)

| Fonction | Rôle |
|---|---|
| `Get-AccessProfiles` | Liste des profils du client, filtrage baseline optionnel |
| `Get-BaselineProfile` | Retourne le profil baseline ou `$null` |
| `Compare-AccessProfileGroups` | Diff intelligent entre anciens et nouveaux profils (ToAdd, ToRemove, ToKeep) |
| `Invoke-AccessProfileChange` | Applique les ajouts/retraits chirurgicaux sur un utilisateur |
| `Get-UserActiveProfiles` | Détecte les profils actifs par correspondance complète des groupes |
| `Get-ProfileReconciliation` | Scan batch des écarts template vs production |
| `Invoke-ProfileReconciliation` | Applique les corrections en lot avec progression |

### Algorithme de diff (`Compare-AccessProfileGroups`)

Le diff est conçu pour minimiser les interruptions de service :

1. Collecter les groupes des anciens profils → `$oldGroups`
2. Collecter les groupes des nouveaux profils → `$newGroups`
3. Baseline ajouté dans `$newGroups` toujours, et dans `$oldGroups` uniquement si `OldProfileKeys.Count > 0` (évite le bug onboarding)
4. Intersection = `$toKeep` (pas touché), nouveaux uniquement = `$toAdd`, anciens uniquement = `$toRemove`

### Algorithme de réconciliation (`Get-ProfileReconciliation`)

Complexité O(N) où N = nombre de groupes dans le profil :

1. Pour chaque groupe G du profil : `Get-MgGroupMember -All` → liste des membres
2. Construction d'une map `userId → {UPN, DisplayName, PresentGroupIds}`
3. Seuil : un utilisateur est candidat si `PresentGroupIds.Count >= max(1, N-1)` mais `< N`
4. Les groupes manquants sont identifiés par différence ensembliste

### Points d'intégration

- **Onboarding** (`GUI_Onboarding.ps1`) : section profils d'accès avec baseline en lecture seule, profils additionnels en CheckedListBox, prévisualisation, application à la création.
- **Modification** (`GUI_Modification.ps1`) : `Show-ChangeAccessProfile` avec détection des profils actuels et diff interactif.
- **Gestionnaire** (`GUI_AccessProfiles.ps1`) : CRUD complet + recherche Graph + réconciliation.
- **Paramètres** (`GUI_Settings.ps1`) : bouton d'accès au gestionnaire.

---

## 7. Auto-update

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

## 8. Installation et désinstallation

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

## 9. Conventions de code

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

## 10. Permissions API Microsoft Graph

> Toutes les permissions sont de type **Délégué** (`Delegated`) — connexion interactive uniquement.
> Admin consent requis sur chaque tenant client.

| Permission | Usage dans M365 Monster |
|---|---|
| `User.ReadWrite.All` | Créer, modifier (profil, téléphones, UPN), désactiver/réactiver des comptes |
| `Group.ReadWrite.All` | Ajouter/retirer des utilisateurs des groupes (licences, sécurité, profils d'accès) ; créer des groupes (remediation nested) |
| `Directory.ReadWrite.All` | Lire les domaines vérifiés du tenant, accès annuaire étendu |
| `Mail.Send` | Envoyer les notifications email via `/me/sendMail` |
| `UserAuthenticationMethod.ReadWrite.All` | Lire et supprimer les méthodes MFA (module Modification — Reset MFA) |
| `AuditLog.Read.All` | Lire les journaux de connexion (module Modification — Dernières connexions ; Shared Mailbox — last sign-in) |
| `Device.Read.All` | Lire les devices Entra (module Nested Group Audit — classification des membres) |
| `DeviceManagementConfiguration.Read.All` | Lire les policies Intune : configuration, compliance, ADMX, Autopilot, updates (module Nested Group Audit) |
| `DeviceManagementApps.Read.All` | Lire les applications Intune et leurs assignations (module Nested Group Audit) |
| `DeviceManagementManagedDevices.Read.All` | Lire les devices managés Intune et les scripts de remédiation (module Nested Group Audit) |

### Notes importantes

- **Profils d'accès** : utilisent `Group.ReadWrite.All` pour les ajouts/retraits de groupes et `Get-MgGroupMember` pour la réconciliation. Aucune permission supplémentaire requise.
- **Téléphones et alias email** : `Update-MgUser` est bloqué par Exchange Online sur `mobilePhone`, `businessPhones` et `proxyAddresses`. L'outil utilise `Invoke-MgGraphRequest PATCH` directement sur `/v1.0/users/{id}` pour contourner cette restriction.
- **Token en cache** : si `Forbidden (403)` apparaît après ajout d'un scope, fermer et relancer l'outil pour forcer un nouveau token.
- **proxyAddresses** : Exchange Online gère les alias de façon autonome. L'ajout/suppression via Graph fonctionne uniquement si la boîte Exchange Online est active et que le compte connecté a les droits suffisants.
- **Endpoints Intune (beta)** : le module Nested Group Audit utilise les endpoints `beta` de Microsoft Graph pour les policies Intune (`/beta/deviceManagement/...`). Ces endpoints peuvent évoluer sans préavis. Chaque catégorie est scannée dans un `try/catch` individuel pour garantir la résilience.
- **Graph Batch API** : le scan des groupes et la réconciliation utilisent des stratégies batch pour minimiser les appels API. Anti-throttling intégré entre chaque lot.
- **Renommage de groupes** : le module Nested Group Audit utilise `Update-MgGroup` pour renommer le groupe d'origine lors de la séparation Users/Devices. Le `mailNickname` est mis à jour simultanément (caractères non-alphanumériques supprimés).
- **Suppression de membres** : `Remove-MgGroupMemberByRef` est utilisé pour retirer les membres transférés du groupe source. L'opération est unitaire (un appel par membre) avec progression visuelle et compteur d'erreurs.

---

## 11. Historique des versions

Voir [CHANGELOG.md](CHANGELOG.md) pour le détail complet de chaque version.

| Version | Date | Résumé |
|---|---|---|
| `0.1.7` | 2026-02-28 | Profils d'accès : gestion, onboarding, modification, réconciliation |
| `0.1.5` | 2026-02-25 | Audit Nested : renommage croisé groupe d'origine + suppression membres transférés |
| `0.1.4` | 2026-02-25 | Nouveau module Audit Groupes Nested (Users+Devices) avec scan Intune |
| `0.1.3` | 2026-02-23 | Alias email via Exchange Online (Set-Mailbox), connexion EXO, Shared Mailbox Audit |
| `0.1.2` | 2026-02-22 | Corrections module Modification : alias, téléphones, groupes, UX |
| `0.1.1` | 2026-02-22 | Corrections UX module Modification : scroll, combos, permissions |
| `0.1.0` | 2026-02-22 | Version bêta initiale |

---

*Fin du document de référence*