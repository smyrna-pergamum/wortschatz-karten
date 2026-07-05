# WortschatzKarten

WortschatzKarten is a minimalist Windows desktop application for learning German vocabulary with flashcards.

The app is designed especially for DSH/C1 learners and works with an Excel vocabulary table. It supports filtering by unit and section, spaced repetition, pronunciation, statistics, favorites, difficult words, and Excel export.

## Features

- Import vocabulary from an Excel file
- Filter cards by `Unite` and `Abschnitt`
- Learn the German target word shown as the main flashcard word
- Show explanations, synonyms, antonyms, English, Turkish, examples, and translations
- German pronunciation button
- Spaced repetition learning system
- Mark words as known or unknown
- Mark words as favorite or difficult
- Export favorite and difficult words as Excel
- Statistics window
- Light and dark theme
- Local database for learning progress
- Minimalist GUI

## Excel Format

The Excel file should contain the following columns:

| Column | Meaning |
|---|---|
| Lektion | Unit |
| Abschnitt | Section / topic |
| Deutsche Word | Main German word |
| Erklärung | German explanation |
| Synonym | Synonym |
| Antonym | Antonym |
| English | English meaning |
| MutterSprache | Mother Language meaning |
| Beispiel | German example sentence |

## Installation for Users

1. Download `WortschatzKarten.zip` from the GitHub release.
2. Extract the ZIP file.
3. Open the extracted folder.
4. Run the setup file, for example:

```text
WortschatzKarten_Setup.exe
```

5. Install the program.
6. After installation, start WortschatzKarten.
7. On the first start, the program asks for an Excel file.
8. Select the Excel file that was installed together with the program in the installation target folder.

Usually the Excel file is located in:

```text
C:\Program Files\WortschatzKarten\DSH_Wortschatz_Table.xlsx
```

After importing the Excel file once, the program stores the vocabulary in its local database. The Excel file does not need to be selected again unless you want to update the vocabulary data.

## Updating the Excel Data

If you want to use a newer Excel file later, open the program and choose:

```text
Daten > Excel aktualisieren...
```

Then select the updated Excel file.

## Local Data Storage

The app stores learning progress, statistics, favorites, difficult words, and cached data locally under:

```text
%APPDATA%\WortschatzKarten
```

## Developer Start

If you want to run the Python version directly:

```bash
pip install -r requirements.txt
python wortschatz_karten.py
```

## Build Windows EXE

```bash
pyinstaller --noconsole --onedir --name WortschatzKarten --icon "Wortschatz_karten.ico" wortschatz_karten.py
```

## Build Installer

The installer is created with Inno Setup:

```bash
ISCC.exe installer.iss
```

## License

MIT License
