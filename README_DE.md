# Wortschatz Karten

[English](README.md) · **Deutsch**

Eine Karteikarten-App für Windows zum Lernen deutscher Vokabeln – entwickelt für **DSH-/C1-Lernende**, arbeitet mit einer Excel-Wortschatztabelle (1.436 Wörter in 10 Lektionen sind enthalten).

[![Download](https://img.shields.io/github/v/release/smyrna-pergamum/wortschatz-karten?label=Download&style=for-the-badge)](https://github.com/smyrna-pergamum/wortschatz-karten/releases/latest)
![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6?style=for-the-badge)
![Lizenz](https://img.shields.io/badge/Lizenz-MIT-green?style=for-the-badge)

![Vorderseite einer Karte (helles Design)](docs/screenshot_light.png)

| Rückseite (dunkles Design) | Statistik |
|---|---|
| ![Rückseite einer Karte](docs/screenshot_dark.png) | ![Statistik](docs/screenshot_stats.png) |

## Funktionen

- **Spaced Repetition (Leitner-System)** – 5 Boxen; gewusste Karten kommen nach 1, 2, 4, 8 und 16 Tagen wieder, nicht gewusste am Ende der Runde
- **Lernmodi** – Wiederholung (fällig + neu), alle Karten, nur neue, zuletzt nicht gewusst, Favoriten, schwierige Wörter
- **Filter** nach Lektion und Abschnitt
- **Rückseite** – Erklärung, Synonyme, Antonyme, Englisch, Türkisch und Beispielsatz, farblich gegliedert
- **Automatische Übersetzung** von Wort und Beispielsatz in die Muttersprache (Türkisch, Englisch, Arabisch, Russisch, Persisch, Spanisch, Französisch, Italienisch, Polnisch, Ukrainisch) – lokal zwischengespeichert
- **Aussprache** deutscher Wörter (🔊) – online mit lokalem Audio-Cache, offline mit einer installierten deutschen Windows-Stimme
- **Favoriten und schwierige Wörter**, exportierbar als Excel-Datei
- **Statistik** mit Leitner-Box-Verteilung und Auswertung je Lektion
- **Helles und dunkles Design**, Umdreh-Animation, scharf auf High-DPI-Bildschirmen
- **Passt sich der Fenstergröße an** – funktioniert auf kleinen Laptop-Bildschirmen und mit Windows-Skalierung (125 % / 150 %)
- **Tastenkürzel** für das Lernen ohne Maus
- Alle Daten bleiben **lokal** auf dem Computer; der Lernfortschritt bleibt beim Aktualisieren der Excel-Datei erhalten

## Installation

1. **`WortschatzKarten_Setup.exe`** aus dem [neuesten Release](https://github.com/smyrna-pergamum/wortschatz-karten/releases/latest) herunterladen.
2. Setup starten und den Schritten folgen.
   > Windows SmartScreen kann eine Warnung anzeigen, weil der Installer nicht signiert ist. Auf **Weitere Informationen → Trotzdem ausführen** klicken.
3. **Wortschatz Karten** starten. Beim ersten Start findet die App die mitgelieferte Wortliste und fragt, ob sie importiert werden soll – mit **Ja** bestätigen.

## Tastenkürzel

| Taste | Aktion |
|---|---|
| Leertaste | Karte umdrehen |
| → / Enter | Nächste Karte |
| ← | Vorherige Karte |
| 1 | Nicht gewusst |
| 2 | Gewusst |
| S | Aussprache |
| F | Favorit an/aus |
| D | Schwieriges Wort an/aus |
| M | Mischen |
| Strg+O | Excel aktualisieren |
| Strg+E | Favoriten & schwierige Wörter exportieren |

## Eigene Excel-Datei verwenden

Über **Daten › Excel aktualisieren…** eine `.xlsx`-Datei auswählen. Der Lernfortschritt bleibt für alle Wörter erhalten, die in der neuen Datei weiterhin vorkommen.

In den ersten Zeilen einer Tabelle wird nach diesen Spaltenüberschriften gesucht (deutsche, englische und türkische Varianten werden erkannt):

| Spalte | Bedeutung | Pflicht |
|---|---|---|
| `Unite` / `Lektion` | Einheit | – (sonst aus dem Tabellennamen) |
| `Abschnitt` | Abschnitt / Thema | – |
| `Deutsche Word` / `Wort` | Deutsches Wort | **ja** |
| `Erklärung` | Erklärung auf Deutsch | – |
| `Synonym` | Synonyme | – |
| `Antonym` | Antonyme | – |
| `English` | Englische Bedeutung | – |
| `Turkish` / `Muttersprache` | Türkische Bedeutung | – |
| `Beispiel` | Deutscher Beispielsatz | – |

Die mitgelieferte Datei [`DSH_Wortschatz_Table.xlsx`](DSH_Wortschatz_Table.xlsx) kann als Vorlage dienen.

## Hinweise

- **Internet** wird nur für die automatische Übersetzung und die erste Aussprache eines Wortes benötigt; danach wird zwischengespeichert. Ist der Übersetzungsdienst vorübergehend nicht erreichbar, zeigt die App das an – der Wortschatz selbst funktioniert immer offline.
- **Lokale Daten** (Lernfortschritt, Einstellungen, Caches) liegen unter `%APPDATA%\WortschatzKarten`. Mit **Daten › Lernfortschritt zurücksetzen…** kann neu begonnen werden.

## Für Entwickler

Benötigt Python 3.10 oder neuer.

```bash
pip install -r requirements.txt
python wortschatz_karten.py
```

EXE und Installer erstellen ([Inno Setup 6](https://jrsoftware.org/isinfo.php)):

```bash
pip install pyinstaller
pyinstaller --noconsole --onedir --name WortschatzKarten --icon "Wortschatz_karten.ico" wortschatz_karten.py
ISCC.exe installer.iss
```

Der Installer liegt danach unter `installer_output\WortschatzKarten_Setup.exe`.

## Lizenz

[MIT](LICENSE)
