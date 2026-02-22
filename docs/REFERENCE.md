# Référence Projet — M365 Monster

> **Version :** 2.0
> **Date :** 2026-02-22
> **Portée :** Gestion du cycle de vie employé dans Microsoft 365 / Entra ID, avec interface graphique WinForms, multi-client, multi-langue.

---

## 1. Contexte et objectifs

M365 Monster est un outil PowerShell avec interface graphique (GUI) permettant à un administrateur IT ou un technicien helpdesk de gérer le cycle de vie des employés dans Microsoft Entra ID (Azure AD), sans ligne de commande.

L'outil est **agnostique au client** : un même set de scripts sert n'importe quelle organisation via un fichier de configuration JSON par client. Il est conçu pour les **MSP** (Managed Service Providers) qui gèrent plusieurs tenants.

---

## 2. Stack technique

| Composant | Technologie |
|---|---|
| Langage | PowerShell 7+ (recommandé), compatible 5.1 |
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
├── update_config.json              # Configuration GitHub auto-update
│
├── 📁 Core/
│   ├── Config.ps1                  # Chargement et validation du JSON client
│   ├── Connect.ps1                 # Authentification Microsoft Graph
│   ├── Functions.ps1               # Utilitaires (logs, mdp, dialogs...)
│   ├── GraphAPI.ps1                # Wrappers sur les appels Graph
│   ├── Lang.ps1                    # Système i18n (Get-Text, Initialize-Language)
│   └── Update.ps1                  # Auto-update depuis GitHub Releases
│
├── 📁 Modules/
│   ├── GUI_Main.ps1                # Fenêtre principale (6 tuiles)
│   ├── GUI_Onboarding.ps1          # Formulaire d'arrivée employé
│   ├── GUI_Offboarding.ps1         # Formulaire de départ employé
│   ├── GUI_Modification.ps1        # Formulaire de modification
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

### Données utilisateur (hors Program Files)

```
%APPDATA%\M365Monster/
├── settings.json                   # Langue choisie
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
13. Affichage de la fenêtre principale (6 tuiles)
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

1. `Main.ps1` appelle `Invoke-AutoUpdate` au démarrage
2. Vérifie `update_config.json` (repo, branche, token, intervalle)
3. Compare `version.json` local vs distant (GitHub)
4. Si nouvelle version : propose le téléchargement, extrait le .zip, redémarre

### Éléments préservés lors des mises à jour

- `Clients/` (configurations client)
- `update_config.json`
- `settings.json` (dans AppData)
- `Logs/` (dans AppData)

---

## 7. Installation et désinstallation

### Install.ps1

- Détecte `pwsh.exe` (PS7) en priorité pour les raccourcis
- Copie : `Main.ps1`, `version.json`, `update_config.json`, `Core/`, `Modules/`, `Scripts/`, `Assets/`, `Lang/`
- Crée `Clients/` avec `_Template.json` uniquement
- Raccourcis Bureau + Menu Démarrer avec icône
- Configuration interactive de l'auto-update GitHub

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
| PowerShell | PS7 recommandé, détection auto dans Install/Main |

---

*Fin du document de référence*
