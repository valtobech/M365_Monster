# M365 Monster

> Outil PowerShell GUI pour la gestion du cycle de vie employé dans Microsoft 365 / Entra ID

![Version](https://img.shields.io/github/v/release/valtobech/M365_Monster)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)

## Fonctionnalités
- Onboarding / Offboarding / Modification d'employés
- Multi-client (MSP) via JSON
- Multi-langue (FR / EN)
- Interface graphique WinForms — zéro ligne de commande

## Installation rapide
```powershell
# 1. Télécharger la dernière release
# 2. Extraire le .zip
# 3. Exécuter en tant qu'administrateur :
.\Install.ps1
```

👉 [Guide d'installation complet](INSTALLATION.md)
👉 [Configuration Azure App Registration](docs/CONFIGURATION.md)

## Prérequis
- Windows 10/11
- PowerShell 7+ 
- Module `Microsoft.Graph` (installé automatiquement)
- Une App Registration Entra ID avec les permissions Graph déléguées
