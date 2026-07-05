# Wortschatz Karten

Minimalistische Lernkarten-App für Windows mit Excel-Import, lokaler Speicherung, Filtern nach Einheit/Abschnitt, Lernstatistik, Light/Dark Theme und automatischer Muttersprache-Übersetzung.

## Start unter Windows

1. ZIP-Datei entpacken.
2. `start_windows.bat` doppelklicken.
3. Beim ersten Start die Wortschatz-Excel-Datei auswählen.
4. Danach fragt das Programm nicht mehr nach der Excel-Datei. Neue Daten können über `Daten > Excel aktualisieren...` geladen werden.

## Neue Übersetzungsfunktion

Oben im Programm kann eine `Muttersprache` gewählt werden. Nach dem Klick auf `Umdrehen` zeigt die Kartenrückseite zusätzlich:

- automatische Übersetzung des deutschen Hauptworts in die gewählte Muttersprache
- automatische Übersetzung des deutschen Beispielsatzes in die gewählte Muttersprache

Die Übersetzungen werden lokal in der SQLite-Datenbank zwischengespeichert. Bereits übersetzte Begriffe werden daher nicht erneut online angefragt.

## Hinweise

- Die automatische Übersetzung nutzt das kostenlose Python-Paket `deep-translator` mit `GoogleTranslator`.
- Dafür ist beim ersten Übersetzen eine Internetverbindung erforderlich.
- Automatische Übersetzungen können je nach Kontext ungenau sein. Die Excel-Spalten `English` und `Turkish` bleiben weiterhin die primären, kuratierten Bedeutungen.
- Die Datenbank liegt unter `%APPDATA%\WortschatzKarten\wortschatz.db`.

## Abhängigkeiten

```bat
pip install -r requirements.txt
```
