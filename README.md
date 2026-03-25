# M365 Monster

> Outil PowerShell GUI pour la gestion du cycle de vie employé dans Microsoft 365 / Entra ID.
> Conçu pour les MSP et les équipes IT multi-clients.

![Version](https://img.shields.io/github/v/release/valtobech/M365_Monster)
![PowerShell](https://img.shields.io/badge/PowerShell-7%2B-blue)
![License](https://img.shields.io/github/license/valtobech/M365_Monster)

---

## Fonctionnalités

### Gestion du cycle de vie

- **Onboarding** — Création complète d'un compte employé : identité, poste, département, office location, company name, hire date, Employee Type, gestionnaire, licences dynamiques par préfixe, groupes d'appartenance, profils d'accès, mot de passe généré automatiquement.
- **Offboarding** — Désactivation du compte avec renommage automatique du jobTitle (`DISABLE - A supprimer le JJ/MM/AAAA | Titre`), révocation des sessions, retrait des groupes (groupes dynamiques ignorés automatiquement), masquage du GAL, conversion en boîte partagée avec vérification de la taille et sélecteur de licence Exchange, délégation FullAccess, révocation des licences (héritées de groupes ignorées), notification email.
- **Modification** — Modifier les attributs d'un employé existant : profil (département, titre, bureau, téléphones), UPN, alias email via Exchange Online, groupes, licences, reset MFA. Fiche utilisateur avec dernières connexions.

### Administration & Audit

- **Profils d'accès** — Définir des paquets de groupes par rôle (baseline + profils métier). Réconciliation bidirectionnelle : détection et correction automatique des écarts entre le template et les groupes réels des utilisateurs.
- **Gestionnaire PIM** — Création, audit et mise à jour des groupes de rôles PIM dans Entra ID. Import depuis le tenant, assignation automatique avec retry et polling de réplication, export CSV.
- **Audit Shared Mailbox** — Identifier les boîtes partagées avec un compte actif, des licences assignées inutilement ou un sign-in interactif détecté. Chargement optimisé en deux passes (Graph v1.0 + beta).
- **Audit Groupes Nested** — Identifier les groupes Entra contenant à la fois des utilisateurs et des devices. Scan des policies et applications Intune impactées (14 catégories par batch Graph API). Création de groupes séparés User/Device avec transfert automatique des membres.
- **Types d'employé** — Gestion en masse des Employee Types dans Entra ID via import CSV.
- **Devices inactifs** — Nettoyage des appareils inactifs dans Entra ID.

### Caractéristiques techniques

- **Multi-client (MSP)** — Un fichier JSON par tenant, sélection au lancement. Pas de valeurs hardcodées.
- **Multi-langue** — Français et anglais, sélection au premier lancement. 500+ clés i18n.
- **Interface graphique WinForms** — Zéro ligne de commande pour l'utilisateur final.
- **Auto-update** — Vérification automatique via GitHub Releases avec téléchargement et installation intégrés. Élévation UAC transparente pour les installations Program Files.
- **Performance** — Batch Graph API avec anti-throttling, deux passes d'enrichissement pour les opérations lourdes, pagination automatique.
- **Exchange Online** — Intégration native pour les opérations que Graph API ne supporte pas (alias email, conversion SharedMailbox, masquage GAL, délégation FullAccess).

---

## Captures d'écran

> *À venir — screenshots de l'interface principale, onboarding, offboarding, PIM Manager, audit.*

---

## Installation rapide

```powershell
# 1. Télécharger la dernière release depuis GitHub
# 2. Extraire le .zip
# 3. Exécuter en tant qu'administrateur :
.\Install.ps1
```

L'installeur configure automatiquement PowerShell 7 et installe les modules requis (`Microsoft.Graph`, `ExchangeOnlineManagement`).

👉 **[Guide d'installation complet](INSTALLATION.md)**

---

## Configuration Azure

Chaque tenant client nécessite une App Registration dans Entra ID avec les permissions déléguées appropriées.

👉 **[Guide de configuration Azure & JSON client](CONFIGURATION.md)**

---

## Prérequis

- Windows 10/11 ou Windows Server 2019+
- PowerShell 7+
- Modules `Microsoft.Graph` et `ExchangeOnlineManagement` (installés automatiquement)
- Une App Registration Entra ID par tenant avec permissions déléguées
- Rôle Exchange Administrator ou Recipient Management pour les opérations Exchange Online

---

## Permissions API requises

### Microsoft Graph (permissions déléguées)

| Permission | Module(s) |
|---|---|
| `User.ReadWrite.All` | Onboarding, Modification, Offboarding, Shared Mailbox Audit |
| `Group.ReadWrite.All` | Onboarding, Modification, Offboarding, Profils d'accès, Nested Group Audit |
| `Directory.ReadWrite.All` | Tous |
| `Mail.Send` | Notifications email |
| `UserAuthenticationMethod.ReadWrite.All` | Modification (reset MFA) |
| `AuditLog.Read.All` | Modification (sign-in logs), Shared Mailbox Audit |
| `Device.Read.All` | Nested Group Audit |
| `DeviceManagementConfiguration.Read.All` | Nested Group Audit (Intune) |
| `DeviceManagementApps.Read.All` | Nested Group Audit (Intune) |
| `DeviceManagementManagedDevices.Read.All` | Nested Group Audit (Intune) |
| `RoleManagement.ReadWrite.Directory` | Gestionnaire PIM |

### Exchange Online

Le compte connecté doit avoir le rôle **Exchange Administrator** ou **Recipient Management** pour les opérations `Set-Mailbox` (alias, conversion SharedMailbox, masquage GAL, délégation FullAccess).

---

## Architecture

```
M365Monster/
├── Main.ps1                    # Orchestrateur principal
├── Core/
│   ├── Config.ps1              # Chargement et migration des JSON clients
│   ├── Connect.ps1             # Authentification Graph + Exchange Online
│   ├── Functions.ps1           # Fonctions utilitaires partagées
│   ├── GraphAPI.ps1            # Wrappers Graph API + Exchange Online
│   ├── Lang.ps1                # Système i18n (Get-Text)
│   ├── PIMFunctions.ps1        # Fonctions PIM (rôles, schedules)
│   └── Update.ps1              # Auto-update via GitHub Releases
├── Modules/
│   ├── GUI_Main.ps1            # Menu principal (tuiles)
│   ├── GUI_Onboarding.ps1      # Formulaire d'onboarding
│   ├── GUI_Offboarding.ps1     # Formulaire d'offboarding
│   ├── GUI_Modification.ps1    # Modification d'attributs
│   ├── GUI_AccessProfiles.ps1  # Gestionnaire de profils d'accès
│   ├── GUI_PIMManager.ps1      # Gestionnaire PIM
│   ├── GUI_SharedMailboxAudit.ps1
│   ├── GUI_NestedGroupAudit.ps1
│   └── GUI_Settings.ps1        # Paramétrage client
├── Lang/
│   ├── fr.json                 # Traductions françaises
│   └── en.json                 # Traductions anglaises
├── Clients/
│   └── _Template.json          # Template de configuration client
├── version.json
├── Install.ps1
└── Uninstall.ps1
```

---

## Documentation

| Document | Contenu |
|---|---|
| **[INSTALLATION.md](INSTALLATION.md)** | Installation, prérequis, utilisation, dépannage |
| **[CONFIGURATION.md](CONFIGURATION.md)** | App Registration Azure, configuration JSON par client |
| **[REFERENCE.md](REFERENCE.md)** | Architecture technique, conventions, flux d'exécution |
| **[CHANGELOG.md](CHANGELOG.md)** | Historique complet des versions |
| **[RELEASE_PROCESS.md](RELEASE_PROCESS.md)** | Procédure de publication interne |

---

## Développé par

**VALTO Bech** 

## Licence

[MIT](LICENSE)
