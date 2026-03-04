<!-- markdownlint-disable -->
# acadon Outlook Signatur – Deployment

Dieses Verzeichnis enthält alle statischen Dateien, die für das Outlook Signatur Add-in benötigt werden.
Es kann direkt als GitHub Pages Repo verwendet werden.

## Struktur

```
deploy/
├── src/                    # Add-in Code (HTML, JS, CSS)
│   ├── taskpane/           # Taskpane UI
│   ├── commands/           # Event-Handler (OnNewMessageCompose)
│   └── lib/                # Bibliotheken (Template-Engine, Graph Client etc.)
├── templates/              # Signatur-Vorlagen (pro Sprache)
│   ├── DE/
│   ├── EN/
│   └── ...
├── addons/                 # Optionale Textbausteine (Banner, Zertifizierungen)
│   ├── addons.json         # Registry aller verfügbaren Bausteine
│   └── *.htm               # HTML-Snippets für Bausteine
├── assets/                 # Icons, Logo
│   ├── logo.png            # Firmenlogo für Signaturen
│   ├── icon-16.png         # Add-in Icons
│   └── ...
└── README.md
```

## GitHub Pages einrichten

### 1. Neues Repo erstellen
```bash
cd deploy/
git init
git add .
git commit -m "Initial commit - Outlook Signatur Add-in"
```

Auf GitHub ein **neues Public-Repo** erstellen (z.B. `outlook-signatur`), dann:
```bash
git remote add origin https://github.com/3Dcut/deployOutlookSignatureAddin.git
git branch -M main
git push -u origin main
```

### 2. GitHub Pages aktivieren
1. Im Repo auf **Settings** → **Pages** (linke Seitenleiste).
2. Unter **Source** → **Deploy from a branch** auswählen.
3. **Branch**: `main`, **Folder**: `/ (root)` → **Save**.
4. Nach ca. 1 Minute ist die Seite unter `https://3dcut.github.io/deployOutlookSignatureAddin/` erreichbar.

### 3. Manifest anpassen
In `outlook-addin/manifest.xml` alle `https://localhost:3001` durch die GitHub Pages URL ersetzen:
```
https://localhost:3001  →  https://EUER-ORG.github.io/outlook-signatur
```

Alle zu ändernden Stellen sind im Manifest mit Kommentaren markiert (`<!-- ADDIN_BASE_URL -->`).

### 4. Updates deployen
Bei Änderungen am `deploy/`-Ordner einfach den Inhalt ins Pages-Repo kopieren und pushen:
```bash
git add .
git commit -m "Update templates/addons/code"
git push
```
GitHub Pages aktualisiert sich dann automatisch (ca. 30-60 Sekunden).
