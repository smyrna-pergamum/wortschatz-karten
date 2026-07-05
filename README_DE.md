# WortschatzKarten

WortschatzKarten ist eine minimalistische Windows-Desktop-App zum Lernen deutscher Vokabeln mit Karteikarten.

Die App wurde besonders für DSH/C1-Lernende entwickelt und arbeitet mit einer Excel-Wortschatztabelle. Sie unterstützt Filter nach Einheit und Abschnitt, Spaced Repetition, Aussprache, Statistik, Favoriten, schwierige Wörter und Excel-Export.

## Funktionen

- Vokabeln aus einer Excel-Datei importieren
- Karten nach `Unite` und `Abschnitt` filtern
- Das deutsche Zielwort als Hauptwort der Karte anzeigen
- Erklärung, Synonyme, Antonyme, Englisch, Türkisch, Beispiele und Übersetzungen anzeigen
- Aussprache-Button für deutsche Wörter
- Spaced-Repetition-Lernsystem
- Wörter als gewusst oder nicht gewusst markieren
- Wörter als Favorit oder schwierig markieren
- Favoriten und schwierige Wörter als Excel exportieren
- Statistikfenster
- Light Theme und Dark Theme
- Lokale Datenbank für den Lernfortschritt
- Apple-inspirierte minimalistische Oberfläche

## Excel-Format

Die Excel-Datei sollte folgende Spalten enthalten:

| Spalte | Bedeutung |
|---|---|
| Unite | Einheit |
| Abschnitt | Abschnitt / Thema |
| Deutsche Word | Deutsches Zielwort |
| Erklärung | Erklärung auf Deutsch |
| Synonym | Synonym |
| Antonym | Antonym |
| English | Englische Bedeutung |
| Turkish | Türkische Bedeutung |
| Beispiel | Deutscher Beispielsatz |

## Installation für Benutzer

1. Lade `WortschatzKarten.zip` aus dem GitHub-Release herunter.
2. Entpacke die ZIP-Datei.
3. Öffne den entpackten Ordner.
4. Starte die Setup-Datei, zum Beispiel:

```text
WortschatzKarten_Setup.exe
```

5. Installiere das Programm.
6. Starte danach WortschatzKarten.
7. Beim ersten Start fragt das Programm nach einer Excel-Datei.
8. Wähle die Excel-Datei aus, die zusammen mit dem Programm in den Installationsordner kopiert wurde.

Normalerweise befindet sich die Excel-Datei hier:

```text
C:\Program Files\WortschatzKarten\DSH_Wortschatz_Table.xlsx
```

Nachdem die Excel-Datei einmal importiert wurde, speichert das Programm die Vokabeln in einer lokalen Datenbank. Die Excel-Datei muss danach nicht erneut ausgewählt werden, außer wenn die Vokabeldaten aktualisiert werden sollen.

## Excel-Daten aktualisieren

Wenn später eine neue Excel-Datei verwendet werden soll, öffne im Programm:

```text
Daten > Excel aktualisieren...
```

Danach kann die neue Excel-Datei ausgewählt werden.

## Lokale Datenspeicherung

Die App speichert Lernfortschritt, Statistik, Favoriten, schwierige Wörter und Cache-Daten lokal unter:

```text
%APPDATA%\WortschatzKarten
```

## Start für Entwickler

Wenn die Python-Version direkt gestartet werden soll:

```bash
pip install -r requirements.txt
python wortschatz_karten.py
```

## Windows-EXE erstellen

```bash
pyinstaller --noconsole --onedir --name WortschatzKarten --icon "Wortschatz_karten.ico" wortschatz_karten.py
```

## Installer erstellen

Der Installer wird mit Inno Setup erstellt:

```bash
ISCC.exe installer.iss
```

## Lizenz

MIT License
