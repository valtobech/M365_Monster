# Changelog — M365 Monster

Toutes les modifications notables sont documentées ici.
Format basé sur [Keep a Changelog](https://keepachangelog.com/fr/1.0.0/).
Versioning selon [Semantic Versioning](https://semver.org/lang/fr/).

---

## [0.1.8] — 2026-03-24

### Ajouté

**Nouveau module — Gestionnaire PIM (`GUI_PIMManager.ps1` + `PIMFunctions.ps1`)**

- **Dashboard interactif** : les groupes PIM sont affichés sous forme de cartes empilées dans un ScrollPanel avec bordure colorée à gauche et pastille de statut. Couleurs selon l'état : vert (OK), orange (Drift), rouge (Missing/Error), gris (Pending). Compteur résumé en haut du panneau (ex : « 3 OK | 2 Drift | 1 Missing »).
- **Éditeur de groupes PIM** : panneau droit complet avec nom éditable, description, type (ComboBox : `Role_Fixe` / `Groupe` / `Groupe_Critical` / `Role`), liste des rôles assignés avec bouton « Retirer », sélecteur de rôles Entra ID avec `AutoCompleteMode = SuggestAppend` + bouton « Ajouter rôle ».
- **Chargement dynamique des rôles** : tous les rôles Entra ID du tenant (built-in + custom) sont découverts via `Invoke-MgGraphRequest` sur `/v1.0/roleManagement/directory/roleDefinitions`. Plus aucune map statique à maintenir.
- **Audit PIM** : scan des schedules `roleEligibilityScheduleInstances` et `roleAssignmentScheduleInstances` pour chaque groupe configuré. Détection des rôles manquants, en surplus (drift) ou OK.
- **Mise à jour PIM** : création automatique des groupes Entra avec `IsAssignableToRole = true` (irréversible, confirmation obligatoire), assignation des rôles avec retry `SubjectNotFound` (3 tentatives × 10s), fallback automatique `noExpiration` → durée max tenant (6 mois active / 1 an eligible).
- **Polling de réplication** : attente de propagation Entra (5s × 12 = 60s max) + attente PIM (15s) avant assignation.
- **Import depuis le tenant** : bouton « Importer » qui recherche les groupes nommés `PIM_*` existants dans Entra ID, charge leurs schedules et les ajoute à la configuration locale.
- **Gestion du cycle de vie** : boutons « + Nouveau groupe » et « Supprimer » dans le panneau gauche. Bouton « Sauvegarder » (violet) persiste toutes les modifications dans le JSON client via `Save-PimConfig`. Avertissement de modifications non sauvegardées à la fermeture si `$PimDirty = $true`.
- **Export CSV** : rapport complet avec nom du groupe, type, rôles, statut, date d'export.
- Nouvelle tuile « PIM Manager » dans `GUI_Main.ps1` (ligne 5). Fenêtre principale agrandie (820×1060).
- Scope `RoleManagement.ReadWrite.Directory` ajouté dans `Connect.ps1`.
- Section `pim_role_groups` ajoutée dans `_Template.json` avec 3 exemples (Role_Fixe, Groupe, Role).
- 60+ nouvelles clés i18n (FR + EN) dans la section `pim.*` + clés tuile dans `main_menu`.

**Module Employee Type — Réutilisation de la session Graph**

- Le script `AzureAD_EmployeeTypeManageGUI.ps1` détecte désormais une session Graph existante via `Get-MgContext`. Lorsqu'il est lancé depuis M365 Monster, il réutilise la session partagée sans demander de ré-authentification.
- Flag `$script:ExternalSession` : si `$true`, le script ne déconnecte pas Graph à la fermeture (préserve la session de l'application principale).
- Le bouton « Connecter » passe automatiquement en « ✓ Connecté » et les contrôles sont activés au lancement si la session existe.

### Mis à jour

- `Main.ps1` : dot-sourcing de `Core/PIMFunctions.ps1` et `Modules/GUI_PIMManager.ps1`.
- `GUI_Main.ps1` : tuile PIM ajoutée en ligne 5, fenêtre agrandie, footer repositionné.
- `Connect.ps1` : scope `RoleManagement.ReadWrite.Directory` ajouté au tableau `$scopes`.
- `_Template.json` : section `pim_role_groups` avec exemples.
- `en.json` / `fr.json` : section `pim.*` complète + clés tuile PIM dans `main_menu`.
- `REFERENCE.md` : version 3.0, arborescence mise à jour (PIMFunctions, GUI_PIMManager), permissions PIM, notes techniques (roleEligibilityScheduleRequests, IsAssignableToRole, polling), historique des versions.
- `INSTALLATION.md` : permission `RoleManagement.ReadWrite.Directory` ajoutée, structure installée mise à jour (Core/PIMFunctions.ps1, Modules/GUI_PIMManager.ps1), tuile PIM dans la liste des 10 tuiles.
- `CONFIGURATION.md` : permission PIM dans le tableau, section `pim_role_groups` documentée dans la référence JSON.

---

## [0.1.7] — 2026-02-28

### Ajouté

**Nouveau module — Profils d'accès (`GUI_AccessProfiles.ps1`)**

- Système de profils d'accès composables : chaque profil regroupe un ensemble de groupes Entra ID formant un package cohérent (ex : « Finance » = 3 groupes). Les profils sont stockés dans le JSON client (`access_profiles`) et s'appliquent lors de l'onboarding, de la modification et via la réconciliation.
- **Gestionnaire de profils** dans GUI_Settings : créer, éditer, supprimer des profils. Recherche de groupes Entra via `Search-AzGroups`. Profil baseline applicable automatiquement à tous les nouveaux employés.
- **Intégration Onboarding** : le profil baseline s'applique automatiquement. Profils additionnels sélectionnables via CheckedListBox. Bouton « Prévisualiser » pour vérifier les groupes avant création.
- **Intégration Modification** : changer les profils d'un employé existant avec diff intelligent (seuls les ajouts/retraits nécessaires sont exécutés via Graph).
- **Réconciliation bidirectionnelle** : détection et correction des écarts entre les profils templates et les groupes réels des utilisateurs. Gestion des `_pending_removals` dans le JSON pour traquer les groupes retirés d'un profil.
- Section `access_profiles` ajoutée dans `_Template.json` avec profils exemples (Common baseline, Finance, Legal, RH, Direction, TI, Maintenance).
- 40+ clés i18n (FR + EN) pour la gestion des profils d'accès.

### Mis à jour

- `Core/Functions.ps1` : `Get-ProfileReconciliation` avec paramètre `-RemovedGroups` pour la réconciliation bidirectionnelle.
- `Core/GraphAPI.ps1` : `Search-AzGroups` pour la recherche de groupes Entra.
- `Main.ps1` : dot-sourcing de `GUI_AccessProfiles.ps1`.
- `GUI_Main.ps1` : tuile « Profils d'accès » ajoutée.
- Documentation complète mise à jour (REFERENCE, INSTALLATION, CONFIGURATION).

---

## [0.1.6] — 2026-02-25

### Corrigé

**Audit Shared Mailbox — Performance et throttling Graph API**

- **Chargement initial très lent (~6 min pour 63 BAL)** : l'enrichissement Graph via `$batch` avec `signInActivity` déclenchait un throttling massif (429 Too Many Requests). Chaque batch de 20 sous-requêtes prenait ~80 secondes à cause du coût interne de `signInActivity` côté Microsoft.
- **Erreurs 429 fréquentes au rechargement** : les appels Graph saturaient le quota API du tenant, provoquant des rejets en cascade.

### Ajouté

**Architecture 2 passes `$filter` (remplacement du `$batch`)**

- **Passe 1** (v1.0) : `$filter=id eq '...'` par lots de 15 pour `accountEnabled`, `assignedLicenses`, `mail`, etc. Temps : ~2s.
- **Passe 2** (beta) : `$filter=id eq '...'` par lots de 15 pour `signInActivity`. Temps : ~6s.
- L'approche `$filter` (requête de liste) ne compte qu'une seule requête HTTP par lot côté rate limiter Graph, vs 20 sous-requêtes comptabilisées individuellement avec `$batch`. Résultat : **6 minutes → 10 secondes**.
- Pré-chargement du cache SKU licences avant le traitement.
- Retry avec backoff progressif (3 tentatives) sur les 429 résiduels.
- Throttle inter-lots (300ms) pour lisser la charge.

### Mis à jour

- `Modules/GUI_SharedMailboxAudit.ps1` : réécriture complète de la stratégie d'enrichissement.

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
