# STS2 Card Viewer

A statistics viewer for Slay the Spire 2 runs. Parse your game saves and view card pick/win rates and relic statistics organized by character class.

## Features

- **Card Statistics**: See how often each card is offered and picked, with pick rates and win rates per class
- **Relic Statistics**: Track relic appearances and win rates with descriptions, per-class breakdown
- **Event Statistics**: Track which event choices you pick and win rates
- **Ancient Relic Tracking**: See which relics are offered at ancient events (Neow, Tezcatara, etc.)
- **Encounter Tracking**: View damage taken per encounter with act numbers and enemy names
- **Sortable Columns**: Click any column header to sort in all tabs
- **Command-Line Version**: Use `sts2_stats.py` for headless environments
- **Multi-Class Support**: Ironclad, Silent, Defect, Necrobinder, and Regent
- **Cross-Class Cards**: Cards picked into decks are tracked with their original class
- **Auto-Detection**: Automatically finds STS2 save files on Windows and Linux
- **Excel Export**: Generates detailed spreadsheets with all statistics
- **Standalone Executable**: No Python required to run

## Download

Get the latest executable for your platform from the [Releases](https://github.com/alvinpeng01/sts2_viewer/releases) page.

## Save File Locations

The app automatically detects your save files, but if needed:

**Windows:**
```
%APPDATA%\SlayTheSpire2\steam\<profile>\saves\history\
```

**Linux:**
```
~/.local/share/SlayTheSpire2/steam/<profile>/saves/history/
```

## Building from Source

### Requirements
- Python 3.8+
- tkinter
- openpyxl

### Run directly
```bash
pip install openpyxl
python sts2_card_viewer.py
```

### Build executable
```bash
pip install pyinstaller
pyinstaller --onefile --name STS2_CardViewer --console sts2_card_viewer.py
```

The executable will be in `dist/STS2_CardViewer`.

## Usage

1. Run the app - data is automatically generated on startup
2. Click "Help" to see detected save folder path
3. Browse tabs: Cards, Relics, Events, Ancient Relics, Encounters
4. Use search to filter items
5. Click column headers to sort
6. Click "Refresh" to regenerate data

### Command-Line

```bash
python sts2_stats.py -o output.xlsx
```

Use `-d` to specify a custom save directory.

## Data Privacy

Personal data (xlsx files) are excluded from the repository via `.gitignore`.

## License

MIT
