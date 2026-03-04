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

## Verwendung

1. Dieses Verzeichnis als eigenes Git-Repo initialisieren.
2. GitHub Pages aktivieren (Settings → Pages → Branch: main).
3. Die URL (z.B. `https://org.github.io/outlook-signatur/`) in der `manifest.xml` des Add-ins eintragen.
