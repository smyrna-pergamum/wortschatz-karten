# Wortschatz Karten

**English** · [Deutsch](README_DE.md)

A Windows flashcard app for learning German vocabulary – built for **DSH / C1** learners and working with an Excel vocabulary table (1,436 words in 10 units are included).

[![Download](https://img.shields.io/github/v/release/smyrna-pergamum/wortschatz-karten?label=Download&style=for-the-badge)](https://github.com/smyrna-pergamum/wortschatz-karten/releases/latest)
![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6?style=for-the-badge)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

![Front of a card (light theme)](docs/screenshot_light.png)

| Back of a card (dark theme) | Statistics |
|---|---|
| ![Back of a card](docs/screenshot_dark.png) | ![Statistics](docs/screenshot_stats.png) |

## Features

- **Spaced repetition (Leitner system)** – 5 boxes; known cards come back after 1, 2, 4, 8 and 16 days, unknown cards return at the end of the round
- **Learning modes** – review (due + new), all cards, new only, last answered wrong, favorites, difficult words
- **Filter** by unit (`Lektion`) and section (`Abschnitt`)
- **Back of the card** – explanation, synonyms, antonyms, English, Turkish and example sentence, colour-coded
- **Automatic translation** of the word and example sentence into your mother language (Turkish, English, Arabic, Russian, Persian, Spanish, French, Italian, Polish, Ukrainian) – cached locally
- **Pronunciation** of German words (🔊) – online with a local audio cache, offline with an installed German Windows voice
- **Favorites and difficult words**, exportable to Excel
- **Statistics** with Leitner box distribution and results per unit
- **Light and dark theme**, card-flip animation, sharp on high-DPI displays
- **Adapts to the window size** – works on small laptop screens and with Windows display scaling (125 % / 150 %)
- **Keyboard shortcuts** for learning without the mouse
- All data stays **local** on your computer; progress is kept when you update the Excel file

## Installation

1. Download **`WortschatzKarten_Setup.exe`** from the [latest release](https://github.com/smyrna-pergamum/wortschatz-karten/releases/latest).
2. Run the setup and follow the steps.
   > Windows SmartScreen may show a warning because the installer is not code-signed. Click **More info → Run anyway**.
3. Start **Wortschatz Karten**. On the first start the app finds the included vocabulary list and asks whether to import it – click **Yes**.

## Keyboard Shortcuts

| Key | Action |
|---|---|
| Space | Flip card |
| → / Enter | Next card |
| ← | Previous card |
| 1 | Not known |
| 2 | Known |
| S | Pronunciation |
| F | Toggle favorite |
| D | Toggle difficult |
| M | Shuffle |
| Ctrl+O | Update Excel |
| Ctrl+E | Export favorites & difficult words |

## Using Your Own Excel File

Choose **Daten › Excel aktualisieren…** and select an `.xlsx` file. Learning progress is kept for every word that still exists in the new file.

The first rows of a sheet are searched for these column headers (German, English and Turkish variants are recognised):

| Column | Meaning | Required |
|---|---|---|
| `Unite` / `Lektion` | Unit | – (otherwise taken from the sheet name) |
| `Abschnitt` | Section / topic | – |
| `Deutsche Word` / `Wort` | German word | **yes** |
| `Erklärung` | German explanation | – |
| `Synonym` | Synonyms | – |
| `Antonym` | Antonyms | – |
| `English` | English meaning | – |
| `Turkish` / `Muttersprache` | Turkish meaning | – |
| `Beispiel` | German example sentence | – |

The included file [`DSH_Wortschatz_Table.xlsx`](DSH_Wortschatz_Table.xlsx) can be used as a template.

## Notes

- **Internet** is only needed for automatic translation and for the first pronunciation of a word; results are cached afterwards. If the translation service is temporarily unavailable, the app says so – the vocabulary itself always works offline.
- **Local data** (progress, settings, caches) is stored in `%APPDATA%\WortschatzKarten`. Use **Daten › Lernfortschritt zurücksetzen…** to start over.

## For Developers

Requires Python 3.10 or newer.

```bash
pip install -r requirements.txt
python wortschatz_karten.py
```

Build the EXE and the installer ([Inno Setup 6](https://jrsoftware.org/isinfo.php)):

```bash
pip install pyinstaller
pyinstaller --noconsole --onedir --name WortschatzKarten --icon "Wortschatz_karten.ico" wortschatz_karten.py
ISCC.exe installer.iss
```

The installer is written to `installer_output\WortschatzKarten_Setup.exe`.

## License

[MIT](LICENSE)
