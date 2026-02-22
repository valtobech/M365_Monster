# Changelog — M365 Monster

Toutes les modifications notables sont documentées ici.
Format basé sur [Keep a Changelog](https://keepachangelog.com/fr/1.0.0/).
Versioning selon [Semantic Versioning](https://semver.org/lang/fr/).

---

## [0.1.2] — 2026-02-22

### Corrigé

**Module Modification (`GUI_Modification.ps1`)**

- **Alias email — `BadRequest (400)`** : remplacement de `Update-MgUser -BodyParameter` par `Invoke-MgGraphRequest PATCH` avec sérialisation explicite en `[string[]]` pour garantir un tableau JSON valide. Le cast `[string[]]` évite la dégradation en string unique sur les listes d'un seul élément.
- **Téléphone mobile / poste fixe — `Forbidden (403)`** : même correction PATCH direct. Le message d'erreur distingue désormais `Forbidden` (scope absent du token — instructions de reconnexion) de `Authorization_RequestDenied` (rôle Entra insuffisant).
- **Boutons Rafraîchir / Fermer invisibles** : les boutons étaient positionnés à Y=730 directement sur `$form`, dépassant la zone cliente réelle. Déplacés dans un `$pnlFooter` en `Dock::Bottom` de 44px — toujours visibles quelle que soit la taille de la fenêtre.
- **Licences — boutons Retirer/Assigner inactifs** : bug de closure PowerShell dans `BeginInvoke` : les variables `$btnRevoke`/`$btnAssign` n'étaient pas capturées dans le scriptblock. Corrigé avec des variables `$script:licBtn*` et calcul du delta dans l'événement `ItemCheck`.
- **Groupes — titre incorrect** : "Désassigner des groupes" renommé en "Modifier les groupes" (clé `action_groups_manage`).
- **Groupes — ajout impossible** : `Show-RemoveGroups` remplacé par `Show-ManageGroups` avec deux colonnes — groupes assignés (retrait) et recherche Graph live (assignation).
- **Scroll menu lateral** : `AutoScrollMinSize` calculé dynamiquement après la boucle de construction — le scroll s'active correctement quand le contenu dépasse la hauteur visible.

**Internationalisation**

- 11 nouvelles clés i18n ajoutées (FR + EN) : `action_groups_manage`, `groups_assigned_label`, `groups_btn_assign`, `groups_confirm_assign`, `groups_search_hint`, `groups_search_label`, `groups_search_placeholder`, `groups_success_assign`, `error_phone_forbidden`, `error_forbidden_reconnect`, `error_proxy_badrequest`.

### Ajouté

- **Fiche utilisateur dans la zone droite** : affichage automatique après sélection — nom, UPN, département, titre, pays, bureau, mobile, poste, statut.
- **Dernières connexions (5)** : `DataGridView` chargé via `auditLogs/signIns` à la sélection de l'utilisateur. Fallback explicite si le scope `AuditLog.Read.All` est absent.

---

## [0.1.1] — 2026-02-22

### Corrigé

**Module Modification (`GUI_Modification.ps1`) — session précédente**

- Fenêtre trop petite et menu latéral non scrollable : taille portée à 860×780, `FormBorderStyle = Sizable`, `AutoScroll = $true` sur le panel menu.
- Département / Titre / EmployeeType / Bureau : menus déroulants vides remplacés par `Show-ModifyComboField` avec chargement dynamique depuis Graph (valeurs existantes dans le tenant).
- Saisie libre dans les combos : `DropDownStyle = DropDown` + `AutoCompleteMode = SuggestAppend`.
- Recherche manager par alias (ex. `adupontel`) : filtre `Search-AzUsers` étendu à `mail`.
- Permissions téléphone : hint affiché avant clic Appliquer + `Format-GraphErrorMessage` pour contextualiser les erreurs.
- `Reset-AzUserPassword` : paramètre converti en `SecureString` (conformité PSScriptAnalyzer).

**Internationalisation**

- Système i18n complet : 140 clés FR + EN pour le module Modification.

---

## [0.1.0] — 2026-02-22

### Ajouté

- Version bêta initiale de M365 Monster.
- Modules : Onboarding, Offboarding, Modification, Settings.
- Architecture multi-client via fichiers `Clients/*.json`.
- Internationalisation FR/EN via `Lang/*.json` + `Core/Lang.ps1`.
- Auto-update via `Core/Update.ps1` + GitHub Releases.
- Installateur/désinstallateur (`Install.ps1` / `Uninstall.ps1`).
- Authentification Microsoft Graph : `interactive_browser`, `device_code`, `client_secret`.

---

*Voir [RELEASE_PROCESS.md](RELEASE_PROCESS.md) pour la procédure de publication.*