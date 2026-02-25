# M365 Monster
> Outil PowerShell GUI pour la gestion du cycle de vie employé dans Microsoft 365 / Entra ID

![Version](https://img.shields.io/github/v/release/valtobech/M365_Monster)
![PowerShell](https://img.shields.io/badge/PowerShell-7%2B-blue)

## Fonctionnalités

- **Onboarding** — Créer un compte employé complet (profil, groupes, licences, mot de passe)
- **Offboarding** — Gérer le départ (désactivation, révocation licences, retrait groupes, redirection mail)
- **Modification** — Modifier les attributs d'un employé (profil, UPN, alias, téléphones, groupes, licences, MFA)
- **Types d'employé** — Gestion des Employee Types dans Entra ID
- **Devices inactifs** — Nettoyage des appareils inactifs dans Entra
- **Audit Shared Mailbox** — Identifier les BAL partagées avec un compte actif, des licences assignées ou un sign-in interactif
- **Audit Groupes Nested** — Identifier les groupes Entra contenant à la fois des utilisateurs et des devices, scanner les policies/applications Intune impactées, créer des groupes séparés

### Caractéristiques techniques

- Multi-client (MSP) via fichiers JSON — un fichier par tenant
- Multi-langue (FR / EN) avec sélection au premier lancement
- Interface graphique WinForms — zéro ligne de commande
- Auto-update via GitHub Releases
- Scan Intune par batch Graph API (14 catégories : apps, configuration, compliance, scripts, updates, Autopilot...)

## Installation rapide
```powershell
# 1. Télécharger la dernière release
# 2. Extraire le .zip
# 3. Exécuter en tant qu'administrateur :
.\Install.ps1
```

👉 [Guide d'installation complet](INSTALLATION.md)
👉 [Configuration Azure App Registration](CONFIGURATION.md)

## Prérequis
- Windows 10/11
- PowerShell 7+
- Modules `Microsoft.Graph` et `ExchangeOnlineManagement` (installés automatiquement)
- Une App Registration Entra ID avec les permissions Graph déléguées

## Permissions API requises

| Permission | Module(s) |
|---|---|
| `User.ReadWrite.All` | Onboarding, Modification, Offboarding, Shared Mailbox |
| `Group.ReadWrite.All` | Onboarding, Modification, Offboarding, Nested Group Audit |
| `Directory.ReadWrite.All` | Tous |
| `Mail.Send` | Notifications email |
| `AuditLog.Read.All` | Modification, Shared Mailbox Audit |
| `Device.Read.All` | Nested Group Audit |
| `DeviceManagementConfiguration.Read.All` | Nested Group Audit (Intune) |
| `DeviceManagementApps.Read.All` | Nested Group Audit (Intune) |
| `DeviceManagementManagedDevices.Read.All` | Nested Group Audit (Intune) |

## Documentation

- [INSTALLATION.md](INSTALLATION.md) — Installation, utilisation, dépannage
- [CONFIGURATION.md](CONFIGURATION.md) — Configuration JSON par client, App Registration, licences
- [REFERENCE.md](REFERENCE.md) — Architecture, conventions, flux d'exécution
- [CHANGELOG.md](CHANGELOG.md) — Historique des versions
- [RELEASE_PROCESS.md](RELEASE_PROCESS.md) — Procédure de publication

## Licence

[MIT](LICENSE)