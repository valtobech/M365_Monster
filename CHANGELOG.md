# Changelog — M365 Monster

Toutes les modifications notables sont documentées ici.
Format basé sur [Keep a Changelog](https://keepachangelog.com/fr/1.0.0/).
Versioning selon [Semantic Versioning](https://semver.org/lang/fr/).

---

## [0.1.7] — 2026-02-28

### Ajouté

**Nouveau module — Profils d'accès (`GUI_AccessProfiles.ps1`)**

Système complet de gestion de profils d'accès composables : chaque profil regroupe un ensemble de groupes Entra ID qui forment un package cohérent (ex : « Finance » = 3 groupes). Les profils sont stockés dans le JSON client et s'appliquent lors de l'onboarding, de la modification et via la réconciliation.

- **Gestionnaire de profils** (Paramètres → Gestion des profils d'accès) : interface CRUD pour créer, éditer et supprimer des profils. Recherche de groupes Entra ID en temps réel via Graph API, profil baseline (appliqué automatiquement à tous les employés), sauvegarde persistante dans le JSON client.
- **Intégration Onboarding** : section « Profils d'accès » dans le formulaire de création. Le profil baseline est affiché en lecture seule (toujours appliqué). Les profils additionnels sont sélectionnables via CheckedListBox. Bouton « Prévisualiser les groupes » affiche la liste complète des groupes résultants avec leur profil source. Les groupes sont ajoutés à l'utilisateur lors de la création.
- **Intégration Modification** : changement de profils dans le module Modification via `Show-ChangeAccessProfile`. Affichage des profils actuels (détectés) et sélection des nouveaux profils. Diff intelligent : les groupes à l'intersection ne sont pas touchés, seuls les ajouts/retraits chirurgicaux sont exécutés.
- **Réconciliation** (bouton « Réconcilier » dans le gestionnaire de profils) : scan des utilisateurs associés à un profil et détection des groupes manquants. Algorithme batch O(N) — N appels Graph (un par groupe du profil), pas de requête par utilisateur. Seuil configurable pour minimiser les faux positifs. DataGridView avec lignes rouges (écarts) / vertes (corrigés). Application en lot avec barre de progression. Export CSV avec piste d'audit complète (UPN, groupes manquants, profil source).

**Fonctions backend — `Core/Functions.ps1`**

- `Get-AccessProfiles` : retourne la liste des profils du client courant avec filtrage baseline optionnel.
- `Get-BaselineProfile` : retourne le profil marqué `is_baseline = true`.
- `Compare-AccessProfileGroups` : calcule le diff intelligent entre deux ensembles de profils (ToAdd, ToRemove, ToKeep). Inclut automatiquement le baseline. Gère correctement le cas onboarding (OldProfileKeys vide).
- `Get-UserActiveProfiles` : détecte les profils actifs d'un utilisateur par correspondance complète de ses appartenances de groupes.
- `Invoke-AccessProfileChange` : applique un changement de profils (ajouts + retraits chirurgicaux). Gère les cas « already member » et « not found » silencieusement.
- `Get-ProfileReconciliation` : scan batch des écarts entre un template de profil et les utilisateurs en production.
- `Invoke-ProfileReconciliation` : applique les corrections avec progression et gestion d'erreur granulaire.

**Fonctions Graph API — `Core/GraphAPI.ps1`**

- `Search-AzGroups` : recherche de groupes Entra ID par nom (préfixe) pour l'éditeur de profils.

**Configuration client — `_Template.json`**

- Nouvelle section optionnelle `access_profiles` avec 7 profils par défaut : Base commune [B], Direction, Finance, Juridique, Maintenance, Ressources humaines, Technologies de l'information.
- Rétrocompatible : les clients sans `access_profiles` ne sont pas impactés.

**Internationalisation**

- 30+ nouvelles clés i18n (FR + EN) dans les sections `access_profiles` et `onboarding` : éditeur de profils, recherche de groupes, prévisualisation, réconciliation, messages de confirmation et d'erreur.

### Corrigé

- **Profils d'accès — bouton « Nouveau profil » tronqué** : largeur élargie de 95px à 110px pour afficher le texte FR complet.
- **Profils d'accès — bouton recherche affichant un carré** : l'emoji 🔍 (non supporté par WinForms) remplacé par un bouton texte « Rechercher » via `Get-Text`.
- **Profils d'accès — bouton « Réconcilier » masqué** : zone de résultats de recherche réduite (120→100px), repositionnement du hint et des boutons d'action pour éviter les chevauchements.
- **Profils d'accès — faux avertissement « non sauvegardé »** : le flag `$script:APDirty` n'était jamais remis à `$false` après sauvegarde. Ajout du reset après persist réussi.
- **Profils d'accès — message de fermeture hardcodé** : chaîne FR remplacée par `Get-Text "access_profiles.unsaved_warning"`.
- **Onboarding — baseline non appliqué** : le diff `Compare-AccessProfileGroups` plaçait le baseline dans `$oldGroups` même pour un onboarding (`OldProfileKeys = @()`), causant son classement en `$toKeep` au lieu de `$toAdd`. Corrigé par condition `$OldProfileKeys.Count -gt 0`.
- **Onboarding — profils additionnels non appliqués** : la garde `Get-Variable -Scope Local` créait un `$clbProfiles = $null` local qui masquait la variable du scope parent. Bloc supprimé.

### Mis à jour

- `Core/Functions.ps1` : +7 fonctions profils d'accès (Get-AccessProfiles, Get-BaselineProfile, Compare-AccessProfileGroups, Invoke-AccessProfileChange, Get-UserActiveProfiles, Get-ProfileReconciliation, Invoke-ProfileReconciliation).
- `Core/GraphAPI.ps1` : +1 fonction (Search-AzGroups).
- `Modules/GUI_AccessProfiles.ps1` : nouveau module ~820 lignes (éditeur + réconciliation).
- `Modules/GUI_Onboarding.ps1` : 3 insertions (section GUI profils, récapitulatif confirmation, appel Invoke-AccessProfileChange).
- `Modules/GUI_Modification.ps1` : intégration Show-ChangeAccessProfile.
- `Modules/GUI_Settings.ps1` : bouton d'accès au gestionnaire de profils.
- `Main.ps1` : dot-sourcing de `GUI_AccessProfiles.ps1`.
- `Clients/_Template.json` : section `access_profiles` avec 7 profils par défaut.
- `Lang/fr.json`, `Lang/en.json` : 30+ nouvelles clés i18n.
- `REFERENCE.md` : v2.5, section profils d'accès, architecture, historique.
- `INSTALLATION.md` : module GUI_AccessProfiles dans la structure, section profils dans l'utilisation.
- `CONFIGURATION.md` : v2.1, référence complète de la section `access_profiles`.

---

## [0.1.5] — 2026-02-25

### Ajouté

**Audit Nested Groups — Renommage croisé du groupe d'origine**

- Lors de la création d'un groupe `_Device`, le groupe d'origine peut être automatiquement renommé en `_User` (et inversement).
- Checkbox « Renommer le groupe d'origine en _Device / _User » ajoutée dans l'onglet Remédiation (cochée par défaut).
- Nouvelle fonction interne `Rename-EntraGroup` : appelle `Update-MgGroup` pour mettre à jour le `displayName` et le `mailNickname` du groupe source.

**Audit Nested Groups — Suppression des membres transférés du groupe d'origine**

- Les membres (users ou devices) transférés vers le nouveau groupe sont désormais retirés du groupe d'origine via `Remove-MgGroupMemberByRef`.
- Checkbox « Supprimer les membres transférés du groupe d'origine » ajoutée dans l'onglet Remédiation (cochée par défaut).
- Nouvelle fonction interne `Remove-MembersFromSourceGroup` : boucle de suppression avec progression visuelle, compteur d'erreurs et journalisation.

**Confirmation dynamique et messages de succès enrichis**

- Les messages de confirmation affichent désormais les actions sélectionnées (suppression, renommage) avec des indicateurs visuels (✔).
- Le message de succès récapitule toutes les opérations effectuées : création du groupe, nombre de membres ajoutés, nombre de membres retirés, renommage du groupe d'origine.

**Internationalisation**

- 15 nouvelles clés i18n (FR + EN) dans la section `nested_group_audit` : checkboxes, confirmations dynamiques, messages de succès, journalisation des opérations de renommage et suppression.

### Mis à jour

- `GUI_NestedGroupAudit.ps1` : onglet Remédiation agrandi (GroupBox 200px → 240px) pour accueillir les checkboxes, journal des actions repositionné.
- `REFERENCE.md` : version du document, note technique sur `Update-MgGroup` / `Remove-MgGroupMemberByRef`, tableau historique des versions.

---

## [0.1.4] — 2026-02-25

### Ajouté

**Nouveau module — Audit des groupes mixtes / Nested Groups (`GUI_NestedGroupAudit.ps1`)**

- **Onglet Scan** : analyse tous les groupes Entra ID et identifie ceux contenant à la fois des utilisateurs et des appareils (groupes « nested » / mixtes). Barre de progression, filtre texte, export CSV.
- **Onglet Membres** : affichage côte à côte des utilisateurs (DisplayName, UPN, poste) et des devices (nom, OS, DeviceId) pour le groupe sélectionné.
- **Onglet Impact Intune** : scanne 14 catégories de policies/applications Intune pour identifier celles qui référencent le groupe sélectionné. Catégories couvertes : Applications (Win32, Store, LOB...), App Configurations, Configuration Policies (Settings Catalog), Device Configurations (legacy), Compliance Policies, Group Policy Configurations (ADMX), Autopilot Profiles, Feature Updates, Quality Updates, Driver Updates, Remediation Scripts, Platform Scripts. Affiche le type d'assignation (Inclus/Exclu), l'intent (Required/Available), la plateforme. Export CSV.
- **Onglet Remédiation** : création de groupes de sécurité séparés `{Original}_User` / `{Original}_Device` avec transfert automatique des membres. Noms pré-remplis et éditables, description automatique, confirmation double avant exécution. Journal des actions en temps réel. Les membres sont ajoutés au nouveau groupe sans être retirés du groupe original.

**Optimisation du scan — Graph Batch API (`/$batch`)**

- Le scan des membres de groupes utilise le endpoint `/$batch` de Microsoft Graph pour envoyer jusqu'à 20 requêtes en parallèle par lot.
- Gain de performance estimé ~15-20x par rapport au scan séquentiel (ex: 500 groupes = ~25 appels batch au lieu de 500 appels individuels).
- Anti-throttling de 150ms entre chaque lot.

**Scopes Graph — 4 nouvelles permissions déléguées (permanentes)**

- `Device.Read.All` — lecture des devices Entra (classification des membres)
- `DeviceManagementConfiguration.Read.All` — lecture des policies Intune (config, compliance, ADMX, Autopilot, updates)
- `DeviceManagementApps.Read.All` — lecture des applications Intune et leurs assignations
- `DeviceManagementManagedDevices.Read.All` — lecture des devices managés et scripts de remédiation

**Intégration au menu principal**

- Nouvelle tuile « Groupes mixtes (Nested) » dans `GUI_Main.ps1` (ligne 4, colonne droite, couleur teal).
- Fenêtre principale agrandie de 840px à 948px pour accueillir la tuile supplémentaire.

**Internationalisation**

- 95+ nouvelles clés i18n (FR + EN) dans la section `nested_group_audit` : onglets, boutons, colonnes, catégories Intune, messages de confirmation, erreurs, journal d'actions.
- 2 nouvelles clés `main_menu` pour la tuile du menu principal.

### Mis à jour

- `Core/Connect.ps1` : 4 scopes Intune ajoutés au tableau `$scopes` (permanents, aucun impact si non utilisés).
- `Main.ps1` : dot-sourcing de `GUI_NestedGroupAudit.ps1`.
- `README.md` : fonctionnalités, tableau de permissions, documentation.
- `INSTALLATION.md` : permissions API, structure installée (modules), liste des tuiles (8).
- `CONFIGURATION.md` : nouvelles permissions dans le tableau de la section App Registration.
- `REFERENCE.md` : arborescence, flux d'exécution, permissions, notes techniques (batch, beta endpoints), historique.
- `RELEASE_PROCESS.md` : `CONFIGURATION.md` ajouté dans les fichiers du zip.

---

## [0.1.3] — 2026-02-22

### Corrigé

**Alias email — remplacement de Graph PATCH par Exchange Online (`Set-Mailbox`)**

- **Cause racine** : `proxyAddresses` est **read-only via Graph API** sur les boîtes Exchange Online (`400 Bad Request — Property 'proxyAddresses' is read-only`). Ni `Update-MgUser`, ni `Invoke-MgGraphRequest PATCH` ne fonctionnent sur ces boîtes.
- **Solution** : passage à `Set-Mailbox -EmailAddresses @{Add=...}` / `@{Remove=...}` via le module `ExchangeOnlineManagement`, qui est la méthode correcte et documentée par Microsoft.
- `Show-ManageAliases` (`GUI_Modification.ps1`) : lecture via `Get-Mailbox` + écriture via `Set-Mailbox`. Vérification de la connexion EXO à l'ouverture du formulaire, avec proposition de reconnexion si inactive.
- `Show-ModifyUPN` (`GUI_Modification.ps1`) : l'ajout de l'ancien UPN comme alias post-changement UPN utilise désormais `Set-Mailbox` à la place du PATCH Graph.

### Ajouté

**Module `Core/Connect.ps1` — nouvelles fonctions Exchange Online**

- `Test-ExchangeModule` : vérifie la présence du module `ExchangeOnlineManagement`, propose l'installation si absent (même pattern que `Test-GraphModule`).
- `Connect-ExchangeOnlineSession` : connexion interactive EXO, vérifie si déjà connecté via `Get-ConnectionInformation`.
- `Disconnect-ExchangeOnlineSession` : déconnexion propre (`Disconnect-ExchangeOnline -Confirm:$false`).
- `Get-ExchangeConnectionStatus` : retourne `[bool]` selon l'état de `Get-ConnectionInformation`.

**`Main.ps1` — connexion EXO au démarrage**

- Appel de `Connect-ExchangeOnlineSession` après la connexion Graph.
- Non bloquant : si EXO échoue, un avertissement est affiché mais l'outil reste fonctionnel (alias email désactivés uniquement).
- Appel de `Disconnect-ExchangeOnlineSession` à la fermeture.

**`Install.ps1` — nouveau prérequis**

- `ExchangeOnlineManagement` ajouté dans `$requiredModules` — installé automatiquement avec `Microsoft.Graph` lors de l'installation.

### Mis à jour

- `REFERENCE.md` : stack technique, flux d'exécution, section permissions (Graph + EXO), notes sur les alias.
- `INSTALLATION.md` : prérequis, description installateur, permissions Azure, étape d'authentification, tableau de dépannage — tout mis à jour pour refléter la double connexion Graph + Exchange Online.

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