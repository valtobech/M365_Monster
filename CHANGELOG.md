# Changelog — M365 Monster

Toutes les modifications notables sont documentées ici.
Format basé sur [Keep a Changelog](https://keepachangelog.com/fr/1.0.0/).
Versioning selon [Semantic Versioning](https://semver.org/lang/fr/).

---

## [0.1.10] — 2026-03-24

### Refonte — Module Offboarding (`GUI_Offboarding.ps1`)

**Internationalisation complète**

- Toutes les chaînes hardcodées remplacées par `Get-Text "offboarding.*"`. 70+ nouvelles clés i18n (FR + EN) dans la section `offboarding`.

**Groupes dynamiques ignorés**

- `Remove-AzUserGroups` détecte les groupes avec `groupTypes` contenant `DynamicMembership` et les skip automatiquement au lieu de provoquer une erreur 400 de Graph API.
- Les groupes dynamiques skippés sont loggés individuellement et rapportés dans le récapitulatif (ex : « Retiré de 3 groupe(s). 2 groupe(s) dynamique(s) ignoré(s) »).
- Nouveau champ `SkippedDynamic` dans l'objet retour de `Remove-AzUserGroups`.

**Renommage automatique du jobTitle**

- `Disable-AzUser` remplace le jobTitle par le format `DISABLE - A supprimer le JJ/MM/AAAA | Titre d'origine` (date = aujourd'hui + 3 mois).
- Déclenche l'exclusion automatique des groupes dynamiques dont la règle contient `(user.jobTitle -notContains "DISABLE")`.
- Le displayName n'est pas modifié.

**Licences héritées de groupes ignorées**

- `Remove-AzUserLicenses` lit `licenseAssignmentStates` pour distinguer les licences assignées directement (retirables) des licences héritées de groupes (`AssignedByGroup` non vide = skip).
- Fallback si `licenseAssignmentStates` n'est pas disponible : tentative de retrait de toutes les licences avec gestion d'erreur par SKU.
- Nouveau champ `SkippedInherited` dans l'objet retour.

**Conversion en boîte partagée (Shared Mailbox)**

- Nouvelle checkbox « Convertir en boîte partagée » dans le formulaire, décoché par défaut.
- Activation conditionnelle : la checkbox n'est activable qu'après sélection d'un utilisateur et vérification réussie de la taille de la BAL via Exchange Online.
- Vérification automatique de la taille BAL dès la sélection de l'utilisateur : label vert si ≤ 50 Go, warning orange si > 50 Go.
- Si BAL > 50 Go et conversion cochée : panneau de sélection de licence Exchange affiché (même pattern `CheckedListBox` que l'onboarding, filtrage par `license_group_prefix`). Confirmation supplémentaire si aucune licence sélectionnée.
- La conversion (`Set-Mailbox -Type Shared`) s'exécute **avant** la révocation des licences dans la séquence d'offboarding.
- Nouvelle fonction `Get-AzMailboxSize` : wrapper `Get-EXOMailboxStatistics` avec parsing de `TotalItemSize` en Go.
- Nouvelle fonction `Convert-AzMailboxToShared` : `Set-Mailbox -Identity $upn -Type Shared`.

**Masquage du carnet d'adresses (GAL)**

- Nouvelle checkbox « Masquer du carnet d'adresses (GAL) » cochée par défaut.
- Dual-approach : Exchange Online (`Set-Mailbox -HiddenFromAddressListsEnabled $true`) avec fallback Graph API (`showInAddressList = $false`).
- Nouvelle fonction `Hide-AzMailboxFromGAL` avec retour `Method` indiquant la méthode utilisée.

**Délégation FullAccess (Read & Manage)**

- Le champ « Rediriger mail vers » (jamais implémenté) est remplacé par « Déléguer accès » avec hint explicatif.
- Si rempli, `Add-MailboxPermission -AccessRights FullAccess` est exécuté sur la boîte de l'utilisateur offboardé pour le délégué spécifié.
- Nouvelle fonction `Grant-AzMailboxFullAccess` avec `AutoMapping = $true`.

**Séquençage de l'offboarding**

- Nouvel ordre d'exécution : Disable + JobTitle → Sessions → Groups (skip dynamic) → Hide GAL → Convert Shared → Add License Exchange → Delegate FullAccess → Remove Licenses → Disabled group → Notification.

### Modifié — Module Onboarding (`GUI_Onboarding.ps1`)

**Nouveaux champs**

- **Employee Hire Date** : `DateTimePicker` pré-rempli avec la date du jour. Valeur transmise à Graph en format ISO 8601 (`yyyy-MM-ddT00:00:00Z`).
- **Office Location** : `ComboBox` éditable alimenté dynamiquement par `Get-AzDistinctValues -Property officeLocation`. L'utilisateur peut sélectionner une valeur existante du tenant ou saisir une nouvelle valeur. Champ obligatoire.
- **Company Name** : même pattern editable combo avec `Get-AzDistinctValues -Property companyName`. Champ obligatoire.

**Modifié dans `New-AzUser` (GraphAPI.ps1)**

- Gestion des propriétés `OfficeLocation`, `CompanyName` et `EmployeeHireDate` dans les paramètres optionnels de création.

**Modifié dans `Get-AzDistinctValues` (GraphAPI.ps1)**

- `companyName` ajouté au `ValidateSet` (en plus de `officeLocation` déjà présent).

### Mis à jour — Fichiers de langue (`fr.json`, `en.json`)

- 70+ nouvelles clés dans la section `offboarding` : formulaire, checkboxes, étapes, résultats, erreurs, confirmations, messages BAL/licence.
- 6 nouvelles clés dans la section `onboarding` : `field_office_location`, `field_company_name`, `field_hire_date`, `validation_office_location`, `validation_company_name`, `confirm_office_location`, `confirm_company_name`, `confirm_hire_date`.

---

## [0.1.9] — 2026-03-24

### Refonte — Module Onboarding (`GUI_Onboarding.ps1`)

**Gestionnaire obligatoire**

- Le champ Gestionnaire (Manager) est désormais obligatoire (marqué `*`, validation bloquante).
- La validation empêche la soumission si aucun gestionnaire n'est sélectionné via la recherche dynamique.
- Le gestionnaire est systématiquement inclus dans le récapitulatif de confirmation et la notification email.

**Licences — Recherche dynamique par préfixe**

- Suppression des groupes de licence hardcodés (`license_groups` dans le JSON client).
- Nouveau champ `license_group_prefix` dans les paramètres client (ex: `LIC_`, `LIC-`, `License_`).
- Les groupes de licence sont chargés dynamiquement via `Search-AzGroups` avec filtrage strict `startsWith` (le `$search` Graph fait un "contains", le filtrage côté client ne garde que les noms commençant par le préfixe).
- Passage d'un ComboBox single-select à un **CheckedListBox multi-sélection** : il est désormais possible d'assigner plusieurs groupes de licence lors de l'onboarding.
- Bouton ⟳ pour rafraîchir la liste des licences. Label d'info indiquant le préfixe actif.

**Groupes d'appartenance — Recherche dynamique**

- Suppression des groupes d'appartenance hardcodés (`membership_groups` dans le JSON client).
- Nouveau pattern de recherche dynamique : champ de recherche + `Search-AzGroups` → résultats dans une ListBox, double-clic pour ajouter au CheckedListBox des groupes sélectionnés.
- **Layout côte à côte** : résultats de recherche à gauche (380px), groupes sélectionnés à droite (380px), compteur de sélection en temps réel.
- Section déplacée **après les Profils d'accès** pour un flux logique (identité → poste → licence → profils → groupes → mot de passe).

**Fenêtre redimensionnable**

- Le formulaire d'onboarding est désormais redimensionnable (`Sizable`) avec une taille par défaut de 1060×920 et un minimum de 900×700.
- Le panel scrollable et les boutons Créer/Annuler sont ancrés pour suivre le redimensionnement.

**Optimisation du code**

- Chargement des données dynamiques AzAD factorisé en boucle `foreach` (4 appels → 1 bloc).
- Nouvelle fonction `Add-SectionHeader` éliminant la duplication pour les labels de section.
- Validation des champs obligatoires factorisée via un tableau `$validations`.

### Modifié — Module Paramétrage (`GUI_Settings.ps1`)

- Sections `GROUPES DE LICENCE` et `GROUPES D'APPARTENANCE` supprimées (plus de listes hardcodées).
- Nouvelle section `LICENCE (RECHERCHE DYNAMIQUE)` avec un champ `Préfixe licence`.
- Rétrocompatibilité dans `Set-FormFromConfig` : détection automatique de l'ancien format `license_groups` et migration vers `license_group_prefix`.
- Bouton « Gestionnaire de profils d'accès » désormais **grisé** (couleur `DarkGray` + `Enabled = $false`) tant qu'aucun client n'est chargé en édition. Repasse en bleu à la sélection d'un client. Tooltip explicatif au survol.
- Préservation des `pim_role_groups` en plus des `access_profiles` lors de la sauvegarde.

### Ajouté — Migration automatique des JSON clients (`Config.ps1`)

- Nouvelle fonction `Invoke-ConfigMigration` appelée au chargement de chaque fichier client.
- Détecte `license_groups` → extrait intelligemment le préfixe commun (ex: `["LIC-M365-E3", "LIC-M365-E5"]` → `"LIC-"`), ajoute `license_group_prefix`, supprime `license_groups`.
- Détecte `membership_groups` → le supprime (recherche dynamique prend le relais).
- **Sauvegarde automatique** du fichier JSON migré sur disque. La migration ne se déclenche qu'une fois par fichier.
- `license_groups` et `membership_groups` retirés des champs obligatoires dans la validation.

### Mis à jour — Fichiers de langue (`fr.json`, `en.json`)

- 11 nouvelles clés i18n dans la section `onboarding` : `validation_manager`, `license_prefix_info`, `license_no_prefix`, `group_search_label`, `group_search_help`, `group_min_chars`, `group_search_title`, `group_no_result`, `group_selected_label`, `group_selected_count`.

### Mis à jour — Template JSON (`_Template.json`)

- `license_groups` et `membership_groups` remplacés par `"license_group_prefix": "LIC_"`.

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
