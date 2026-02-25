# Processus de Release — M365 Monster

> Document de référence interne pour publier une nouvelle version de M365 Monster.

---

## Checklist rapide

```
[ ] 1. Mettre à jour version.json
[ ] 2. Commiter et pousser les changements
[ ] 3. Recréer M365Monster.zip
[ ] 4. Publier la Release sur GitHub
[ ] 5. Vérifier la Release
```

---

## Étape 1 — Mettre à jour `version.json`

À la racine du projet, modifier `version.json` :

```json
{
  "version": "X.Y.Z",
  "release_date": "YYYY-MM-DD",
  "minimum_ps_version": "7.0",
  "release_notes": "Description courte des changements"
}
```

### Convention de versioning (SemVer)

| Type de changement | Exemple | Version |
|---|---|---|
| Correction de bug | Fix throttling update | `0.1.1` → `0.1.2` |
| Nouvelle fonctionnalité | Ajout traduction module | `0.1.2` → `0.2.0` |
| Changement majeur / breaking | Refonte architecture | `0.2.0` → `1.0.0` |

---

## Étape 2 — Commiter et pousser

```powershell
git add .
git commit -m "chore: version X.Y.Z — description courte"
git push
```

---

## Étape 3 — Recréer `M365Monster.zip`

Exécuter ce script depuis la racine du repo :

```powershell
$source    = "D:\GIT\M365_Monster"
$outputZip = "D:\GIT\M365Monster.zip"
$tempDir   = "$env:TEMP\M365Monster_release"

# Nettoyage
if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
New-Item -Path "$tempDir\M365Monster" -ItemType Directory -Force | Out-Null

# Dossiers à inclure
$inclure = @("Assets","Core","Lang","Modules","Scripts","Clients")
foreach ($item in $inclure) {
    Copy-Item -Path "$source\$item" -Destination "$tempDir\M365Monster\$item" -Recurse -Force
}

# Fichiers racine à inclure
$fichiers = @(
    "Main.ps1",
    "Install.ps1",
    "Uninstall.ps1",
    "version.json",
    "update_config.example.json",
    "LICENSE",
    "README.md",
    "INSTALLATION.md",
    "CONFIGURATION.md"
)
foreach ($f in $fichiers) {
    if (Test-Path "$source\$f") {
        Copy-Item -Path "$source\$f" -Destination "$tempDir\M365Monster\$f" -Force
    }
}

# Créer le zip
Compress-Archive -Path "$tempDir\M365Monster" -DestinationPath $outputZip -Force

Write-Host "ZIP créé : $outputZip" -ForegroundColor Green
Remove-Item $tempDir -Recurse -Force
```

### Vérifier le contenu du zip

```powershell
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::OpenRead("D:\GIT\M365Monster.zip").Entries |
    Select-Object -First 20 FullName
```

La structure attendue :
```
M365Monster/
M365Monster/Main.ps1
M365Monster/version.json
M365Monster/update_config.example.json
M365Monster/Core/Update.ps1
M365Monster/Clients/_Template.json
...
```

> ⚠️ `update_config.json` ne doit **pas** être dans le zip (exclu par `.gitignore`).
> C'est `update_config.example.json` qui sert de modèle — `Install.ps1` le copie automatiquement.

---

## Étape 4 — Publier la Release sur GitHub

### Créer une nouvelle Release

1. Aller sur https://github.com/valtobech/M365_Monster/releases/new
2. **Tag** → `vX.Y.Z` → "Create new tag on publish"
3. **Title** → `M365 Monster vX.Y.Z — Titre court`
4. **Description** → notes de version (voir modèle ci-dessous)
5. **Attach files** → glisser-déposer `M365Monster.zip`
6. Ne **pas** cocher "Set as pre-release" (bloque l'API `/releases/latest`)
7. Cliquer **"Publish release"**

### Modèle de description

```markdown
## Changements

- Description du changement 1
- Description du changement 2

## Installation

Voir [INSTALLATION.md](INSTALLATION.md)

## Mise à jour depuis une version existante

La mise à jour est automatique au prochain lancement de M365 Monster.
```

---

## Étape 5 — Vérifier la Release

```powershell
# Vérifier que l'API GitHub retourne bien la nouvelle version
$response = Invoke-RestMethod `
    -Uri "https://api.github.com/repos/valtobech/M365_Monster/releases/latest" `
    -Headers @{ "User-Agent" = "M365Monster-Check" }

Write-Host "Tag     : $($response.tag_name)"
Write-Host "Assets  :"
$response.assets | Select-Object name, size, browser_download_url

# Vérifier que version.json distant est à jour
$v = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/valtobech/M365_Monster/main/version.json"
Write-Host "version.json distant : $($v.version)"
```

Résultat attendu :
```
Tag     : vX.Y.Z
Assets  :
name              size    browser_download_url
M365Monster.zip   XXXXX   https://github.com/...

version.json distant : X.Y.Z
```

---

## Notes importantes

- Le mécanisme d'auto-update compare `version.json` **local** vs `version.json` **distant sur le repo** (pas le tag GitHub). Les deux doivent être cohérents.
- `Clients/` et `update_config.json` sont **toujours préservés** lors d'une mise à jour automatique.
- Si `check_interval_hours = 0`, la vérification a lieu à chaque lancement. Le fichier `%APPDATA%\M365Monster\.last_update_check` est ignoré.
- Pour forcer une vérification manuelle : supprimer `%APPDATA%\M365Monster\.last_update_check`.