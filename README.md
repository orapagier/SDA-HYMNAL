# SDA Hymnal

A lightweight Windows desktop app for quickly searching and projecting Seventh-day Adventist hymns and responses during church service, using PowerPoint slideshows behind the scenes.

Built with **Python** and **PySide6 (Qt)**, it drives Microsoft PowerPoint via COM automation to launch hymn slideshows in full-screen Presenter View, controlled from a small on-screen operator panel.

## Features

- **Instant search** — type a hymn number or title keyword to filter the list in real time.
- **One-click / one-key launch** — double-click or press `Enter` to project the selected hymn.
- **Presenter View automation** — automatically enables PowerPoint's Presenter View setting and launches the slideshow with the correct advance/show settings, so you don't have to configure it manually.
- **Floating control bar** — while a hymn is playing, a small always-on-top control window lets the operator return to the search screen without leaving the presentation.
- **Keyboard-driven workflow**, so an operator can run the whole service without touching the mouse:
  | Key | Action |
  |---|---|
  | `Enter` / `Return` | Open the selected hymn |
  | `Esc` | Close the current hymn and refresh the list |
  | `Shift` | Toggle focus between the search bar and the results list |
  | `↑` / `↓` | Move selection up/down the results list |
- **Add Hymns** — import your own `.pps` / `.ppsx` / `.ppt` / `.pptx` files into the library from a file picker; they're copied into the app's data folder and indexed immediately.
- **Refresh** — clears the search and force-quits any running PowerPoint instance, resetting the app for the next song.
- **Automatic trusted-location setup** — registers the app's own folder as a trusted PowerPoint location on first run so slideshows open without security prompts.

## Included hymn library

The app ships with a bundled library under `_internal/Data`:

| Folder | Contents | Files |
|---|---|---|
| `1 Responses` | Sabbath service responses (opening/closing, tithes, etc.) | 7 |
| `2 Hymns` | "Phil Edition" hymnal slideshows | 239 |
| `3 SDA HYMNAL` | Full SDA Hymnal slideshows | 586 |

Hymns you add via **Add Hymns** are stored separately under `Data/Added Hymns` next to the running executable and are picked up automatically the next time the list is scanned.

## Requirements

- Windows (the app relies on the Windows registry and COM automation — it will not run on macOS/Linux)
- Microsoft PowerPoint (any recent version) installed
- Python 3.12+ (only if running/building from source)

## Installing the prebuilt release

If you just want to run the app without setting up Python, download the latest build from the [Releases page](https://github.com/orapagier/SDA-HYMNAL/releases).

Each release provides two things:

- `SDA Hymnal.exe` — the application executable
- `_internal` — a folder (may be provided as a `.zip`) containing the Python/Qt runtime and the bundled hymn library

**The `.exe` will not run on its own.** Both must be downloaded and placed **side by side in the same folder**:

```
SDA Hymnal/
├── SDA Hymnal.exe
└── _internal/
    ├── Data/           # hymn slideshows
    ├── PySide6/        # Qt runtime
    └── ...             # other runtime dependencies
```

**Steps:**

1. Download `SDA Hymnal.exe` and `_internal` (or `_internal.zip`) from the same release.
2. If `_internal` was downloaded as a `.zip`, extract it so you end up with an `_internal` folder.
3. Move both `SDA Hymnal.exe` and the `_internal` folder into the same directory, as shown above.
4. Double-click `SDA Hymnal.exe` to launch.

On first launch the app registers its own folder as a trusted PowerPoint location — if the folder moves later, keep `SDA Hymnal.exe` and `_internal` together, or the app won't find its hymn library.

## Running from source

```bash
pip install PySide6 pywin32
python "sdahymnal v1.0.1.py"
```

## Building a standalone executable

The project includes a PyInstaller spec file (`SDA Hymnal.spec`) that bundles the app, its data folder, and the icon into a distributable folder:

```bash
pip install pyinstaller
pyinstaller "SDA Hymnal.spec"
```

The output is written to `dist/SDA Hymnal/` (not tracked in this repository — see `.gitignore`). Copy that folder wherever you want to run the app, or zip it for distribution.

## Project structure

```
.
├── sdahymnal v1.0.1.py     # Application entry point (PySide6 UI + PowerPoint automation)
├── SDA Hymnal.spec         # PyInstaller build spec
├── LICENSE                 # MIT License
└── _internal/
    └── Data/
        ├── 1 Responses/    # Response slideshows
        ├── 2 Hymns/        # Phil Edition hymn slideshows
        ├── 3 SDA HYMNAL/   # SDA Hymnal slideshows
        ├── favicon.ico     # App icon
        └── search_icon.png
```

## How it works

1. On launch, the app locates its own base folder (handling both the frozen `.exe` case and running from source) and adds it as a trusted PowerPoint location via the Windows registry, then forces PowerPoint's `UsePresenterView` option on.
2. It recursively scans that folder for `.pps`, `.ppsx`, `.ppt`, and `.pptx` files and builds a searchable, alphabetically-sorted index.
3. Selecting a hymn spins up a background thread that opens the file in PowerPoint via `win32com.client`, configures the slideshow to advance on click in Presenter View, and runs it — while the main window hides and a small floating control bar appears so the operator can return to search at any time.
4. **Refresh** / `Esc` quits the active PowerPoint instance (via `GetActiveObject`) and resets the UI for the next hymn.

## License

MIT — see [LICENSE](LICENSE).

## Author

Jelmar A. Orapa — orapajelmar@gmail.com
