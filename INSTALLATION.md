# Installation & Utilisation — M365 Monster

## Prérequis

- **Windows 10/11** ou Windows Server 2019+
- **PowerShell 7+** — [Installer PowerShell 7](https://learn.microsoft.com/fr-fr/powershell/scripting/install/installing-powershell-on-windows)
- Module **Microsoft.Graph** (installé automatiquement par l'installeur)
- Module **ExchangeOnlineManagement** (installé automatiquement par l'installeur — requis pour la gestion des alias email)
- Droits **administrateur local** pour l'installation

> ⚠️ PowerShell 7 est requis. Le SDK Microsoft.Graph 2.34+ utilise le WAM (Web Account Manager) qui nécessite PowerShell 7 pour un fonctionnement optimal. L'installeur détecte et configure automatiquement PowerShell 7 si présent.

---

## 1. Préparer Azure (une fois par tenant client)

1. Portail Azure → **Entra ID** → **App registrations** → **New registration**
2. Nom : `M365 Monster`, Type : **Single tenant**
3. **Authentication** :
   - Cocher **"Allow public client flows"** → Save
   - **Redirect URIs** — plateforme **Mobile and desktop**, ajouter **les deux** URI :
     - `http://localhost`
     - `ms-appx-web://Microsoft.AAD.BrokerPlugin/<votre-client-id>`

   > ⚠️ **Le second redirect URI est obligatoire depuis le SDK Microsoft.Graph 2.34+.**
   > Le WAM (Web Account Manager) est activé par défaut et ne peut plus être désactivé.
   > Sans ce redirect URI, l'erreur `AADSTS50011` apparaîtra à la connexion.
   > Remplacez `<votre-client-id>` par l'Application (client) ID de votre App Registration.

4. **API permissions** → ajouter les permissions **déléguées** → **Grant admin consent** :
   - `User.ReadWrite.All`
   - `Group.ReadWrite.All`
   - `Directory.ReadWrite.All`
   - `Mail.Send` (si notifications)
   - `UserAuthenticationMethod.ReadWrite.All` (reset MFA)
   - `AuditLog.Read.All` (journaux de connexion)

5. **Permissions Exchange Online** — le compte connecté doit avoir le rôle **Exchange Administrator** ou **Recipient Management** sur le tenant pour pouvoir exécuter `Set-Mailbox` (gestion des alias email).

---

## 2. Installer M365 Monster

Télécharger la dernière release depuis GitHub, extraire le .zip, puis :

```powershell
# PowerShell en tant qu'administrateur
cd "C:\chemin\vers\dossier\extrait\M365Monster"
.\Install.ps1
```

L'installateur :
- Installe les modules `Microsoft.Graph` et `ExchangeOnlineManagement` si absents
- Copie les fichiers vers `C:\Program Files\M365Monster`
- Détecte PowerShell 7 (`pwsh.exe`) et configure les raccourcis en conséquence
- Crée un raccourci Bureau + Menu Démarrer (avec icône)
- Configure l'auto-update automatiquement (aucune intervention requise)

### Options

```powershell
.\Install.ps1 -InstallPath "D:\Outils\M365Monster"    # Chemin custom
.\Install.ps1 -SkipModules                              # Sans install modules PS
.\Install.ps1 -SkipShortcuts                            # Sans raccourcis
.\Install.ps1 -SkipUpdateConfig                         # Sans config auto-update
```

### Structure installée

```
C:\Program Files\M365Monster\
├── Main.ps1                    # Point d'entrée
├── version.json                # Version courante
├── update_config.json          # Config auto-update (créé par l'installeur)
├── Uninstall.ps1               # Désinstallateur
├── Core/                       # Modules fonctionnels
│   ├── Config.ps1
│   ├── Connect.ps1
│   ├── Functions.ps1
│   ├── GraphAPI.ps1
│   ├── Lang.ps1
│   └── Update.ps1
├── Modules/                    # Interfaces GUI
│   ├── GUI_Main.ps1
│   ├── GUI_Onboarding.ps1
│   ├── GUI_Offboarding.ps1
│   ├── GUI_Modification.ps1
│   └── GUI_Settings.ps1
├── Lang/                       # Fichiers de langue
│   ├── fr.json
│   └── en.json
├── Scripts/                    # Scripts externes
├── Assets/                     # Icônes et ressources
└── Clients/                    # Configurations client (préservé lors des MAJ)
    └── _Template.json
```

### Données utilisateur

Les logs, préférences et données de cache sont stockés dans `%APPDATA%\M365Monster\` (et non dans Program Files) pour éviter les problèmes de permissions :

```
%APPDATA%\M365Monster\
├── settings.json               # Langue choisie
├── .last_update_check          # Horodatage dernière vérification MAJ
└── Logs/
    └── session_YYYY-MM-DD_HH-mm.log
```

---

## 3. Configurer un client

### Option A — Via l'interface graphique (recommandé)

1. Lancer **M365 Monster** depuis le Bureau
2. À l'écran de sélection du client, cliquer **⚙ Nouveau client / Paramétrage**
3. Remplir le formulaire et sauvegarder
4. Le nouveau client apparaît dans la liste déroulante

### Option B — Édition manuelle du JSON

1. Ouvrir `C:\Program Files\M365Monster\Clients\`
2. Copier `_Template.json` → `NomDuClient.json`
3. Remplir au minimum :

```json
{
  "client_name": "Mon Client",
  "tenant_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "client_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "auth_method": "interactive_browser",
  "smtp_domain": "@client.com"
}
```

> Voir le [Guide de configuration](GUIDE_CONFIGURATION.md) pour la référence complète de tous les champs.

---

## 4. Utilisation

Double-cliquer **M365 Monster** sur le Bureau.

1. **Premier lancement** : choisir la langue (FR/EN)
2. Sélectionner le client dans la liste (ou en créer un nouveau)
3. S'authentifier via le navigateur (MFA supporté) — **deux authentifications successives** :
   - Microsoft Graph (Entra ID) — popup navigateur
   - Exchange Online — popup navigateur (même compte, peut s'enchaîner automatiquement)
4. Utiliser les 6 tuiles du menu principal :
   - **Onboarding** — Créer un nouveau compte employé
   - **Offboarding** — Gérer le départ d'un employé
   - **Modification** — Modifier les attributs d'un employé
   - **Types d'employé** — Gestion des Employee Types dans Entra ID
   - **Devices inactifs** — Nettoyage des devices inactifs
   - **Paramétrage** — Gérer les configurations client

---

## 5. Auto-update

L'auto-update est activé par défaut à l'installation. M365 Monster vérifie automatiquement les mises à jour à **chaque lancement**.

La configuration se trouve dans `C:\Program Files\M365Monster\update_config.json` :

```json
{
  "github_repo": "valtobech/M365_Monster",
  "branch": "main",
  "github_token": "",
  "download_url": "",
  "check_interval_hours": 0
}
```

- `check_interval_hours: 0` = vérification à chaque lancement (défaut)
- `check_interval_hours: 24` = vérification toutes les 24h
- `github_token` = uniquement nécessaire pour un repo privé
- `download_url` = laisser vide pour utiliser GitHub Releases automatiquement

Les éléments suivants sont **préservés** lors des mises à jour :
- `Clients/` (configurations client)
- `update_config.json`

---

## 6. Désinstallation

```powershell
& "C:\Program Files\M365Monster\Uninstall.ps1"
```

Ou via Menu Démarrer → M365 Monster → Désinstaller.

Le désinstallateur :
- Propose de conserver les fichiers de configuration client
- Supprime les raccourcis (Bureau + Menu Démarrer)
- Supprime le dossier d'installation
- Supprime les données utilisateur (`%APPDATA%\M365Monster`)

---

## 7. Dépannage

| Problème | Cause | Solution |
|---|---|---|
| `AADSTS50011` redirect URI mismatch | Redirect URI WAM manquant | Ajouter `ms-appx-web://Microsoft.AAD.BrokerPlugin/<client-id>` dans App Registration → Authentication |
| Erreur d'accès en écriture dans Program Files | Ancienne version (logs dans dossier d'install) | Mettre à jour vers la dernière version (logs dans AppData) |
| Le raccourci lance PowerShell 5.1 | PowerShell 7 non installé au moment de l'installation | Installer PS7 puis réinstaller M365 Monster |
| `InteractiveBrowserCredential failed` | "Allow public client flows" désactivé | App Registration → Authentication → Allow public client flows → **Yes** |
| Langue non proposée au démarrage | `settings.json` existe déjà | Supprimer `%APPDATA%\M365Monster\settings.json` |
| Mise à jour non détectée | Throttling actif | Supprimer `%APPDATA%\M365Monster\.last_update_check` |
| Alias email — "La session Exchange Online n'est pas active" | Connexion EXO échouée au démarrage | Fermer et relancer l'outil ; vérifier que le compte a le rôle Exchange Administrator |
| Alias email — erreur 400 ou "access denied" via Set-Mailbox | Permissions Exchange insuffisantes | Attribuer le rôle **Recipient Management** ou **Exchange Administrator** au compte dans EAC |
| Module ExchangeOnlineManagement absent | Non installé ou installation échouée | Exécuter manuellement : `Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force` |