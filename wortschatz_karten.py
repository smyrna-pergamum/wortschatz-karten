# -*- coding: utf-8 -*-
"""
Wortschatz Karten - Windows GUI Lernkartenprogramm

Start:
    python wortschatz_karten.py

Abhaengigkeiten:
    pip install -r requirements.txt   (openpyxl, deep-translator, gTTS)

Die App fragt beim ersten Start nach einer Excel-Datei, speichert die Daten danach
lokal in SQLite und fragt beim naechsten Start nicht erneut. Aktualisierung ueber
"Daten > Excel aktualisieren" - der Lernfortschritt bleibt dabei erhalten.

Tastenkuerzel:
    Leertaste  Karte umdrehen        Pfeil rechts / Enter  naechste Karte
    1          Nicht gewusst         Pfeil links           vorherige Karte
    2          Gewusst               S                     Aussprache
    F          Favorit               D                     Schwieriges Wort
    M          Mischen
"""

from __future__ import annotations

import os
import re
import math
import sys
import queue
import random
import sqlite3
import hashlib
import threading
import traceback
import subprocess
from dataclasses import dataclass, fields
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font
except ImportError as exc:
    raise SystemExit("Bitte zuerst installieren: pip install openpyxl") from exc

try:
    from deep_translator import GoogleTranslator
except ImportError:
    GoogleTranslator = None

try:
    from deep_translator import MyMemoryTranslator
except ImportError:
    MyMemoryTranslator = None

try:
    from gtts import gTTS
except ImportError:
    gTTS = None

try:
    from PIL import Image, ImageDraw, ImageTk  # fuer geglaettete Rundungen
except ImportError:
    Image = ImageDraw = ImageTk = None

APP_NAME = "Wortschatz Karten"
APP_VERSION = "1.1.0"
APP_DIR = Path(os.getenv("APPDATA", str(Path.home()))) / "WortschatzKarten"
DB_PATH = APP_DIR / "wortschatz.db"
AUDIO_DIR = APP_DIR / "audio"
BUNDLED_EXCEL_NAME = "DSH_Wortschatz_Table.xlsx"

# Spaltenreihenfolge fuer den Export - identisch mit dem Importformat.
STANDARD_COLUMNS = [
    "Unite",
    "Abschnitt",
    "Deutsche Word",
    "Erklärung",
    "Synonym",
    "Antonym",
    "English",
    "Turkish",
    "Beispiel",
]

MUTTERSPRACHEN = {
    "Türkisch": "tr",
    "Englisch": "en",
    "Arabisch": "ar",
    "Russisch": "ru",
    "Persisch": "fa",
    "Spanisch": "es",
    "Französisch": "fr",
    "Italienisch": "it",
    "Polnisch": "pl",
    "Ukrainisch": "uk",
    "Deutsch": "de",
}

# Sprachcodes fuer den Ersatz-Uebersetzer (MyMemory), falls Google nicht erreichbar ist.
MYMEMORY_CODES = {"tr": "tr-TR", "en": "en-GB", "ar": "ar-SA", "ru": "ru-RU", "fa": "fa-IR", "es": "es-ES",
                  "fr": "fr-FR", "it": "it-IT", "pl": "pl-PL", "uk": "uk-UA"}

# Sprachen, deren Uebersetzung bereits in der Excel-Tabelle steht (kein Online-Aufruf noetig).
EXCEL_LANGUAGE_FIELDS = {"tr": "turkish", "en": "english"}

# Robust mapping for the Excel variants in the uploaded DSH files.
COLUMN_ALIASES = {
    "unit": ["unite", "unit", "ünite", "lektion", "einheit", "teil"],
    "abschnitt": ["abschnitt", "bölüm", "bolum", "bereich", "kapitel", "thema"],
    "german_word": ["deutsche word", "deutsches wort", "almanca kelime", "kelime", "wort", "wort / ausdruck",
                    "kelime / ausdruck", "kelime / ifade", "word / ausdruck", "deutsch"],
    "erklaerung": ["erklärung", "erklaerung", "almanca açıklama", "almanca aciklama", "deutsche erklärung", "definition"],
    "synonym": ["synonym", "synonym(e)", "synonyme"],
    "antonym": ["antonym", "antonym(e)", "antonyme"],
    "english": ["english", "englisch", "en", "english / tuerkce", "english / türkçe"],
    "turkish": ["turkish", "türkçe", "tuerkce", "tr", "muttersprache"],
    "beispiel": ["beispiel", "beispielsatz", "almanca örnek cümle", "almanca ornek cumle", "deutscher beispielsatz"],
}

# Leitner-System: Wiederholungsabstand in Tagen je Box (Box 0 = neu, Box 5 = sicher).
LEITNER_INTERVALS = [0, 1, 2, 4, 8, 16]
MAX_BOX = len(LEITNER_INTERVALS) - 1

LERNMODI = {
    "Wiederholung (fällig + neu)": "due",
    "Alle Karten": "all",
    "Nur neue Karten": "new",
    "Zuletzt nicht gewusst": "unknown",
    "Favoriten": "favorite",
    "Schwierige Wörter": "difficult",
}

THEMES = {
    "light": {
        "bg": "#F2F2F7",
        "panel": "#FFFFFF",
        "field": "#F1F2F6",
        "soft": "#F5F7FB",
        "text": "#1D1D1F",
        "muted": "#6E6E73",
        "line": "#DCDCE1",
        "track": "#E5E5EA",
        "accent": "#007AFF",
        "accent_hover": "#0066D6",
        "success": "#28A745",
        "success_hover": "#218838",
        "danger": "#E5383B",
        "danger_hover": "#C62F32",
        "button": "#EEF0F4",
        "button_hover": "#E1E4EA",
        "button_text": "#1D1D1F",
        "on_accent": "#FFFFFF",
        "star": "#FF9F0A",
        "flag": "#E5383B",
        "purple": "#AF52DE",
        "chip": "#E8F1FF",
    },
    "dark": {
        "bg": "#121214",
        "panel": "#1C1C1F",
        "field": "#29292D",
        "soft": "#232831",
        "text": "#F5F5F7",
        "muted": "#A1A1A6",
        "line": "#34343A",
        "track": "#34343A",
        "accent": "#0A84FF",
        "accent_hover": "#3D9BFF",
        "success": "#2FB350",
        "success_hover": "#3CC85F",
        "danger": "#FF453A",
        "danger_hover": "#FF6A61",
        "button": "#2C2C30",
        "button_hover": "#3A3A40",
        "button_text": "#F5F5F7",
        "on_accent": "#FFFFFF",
        "star": "#FFB340",
        "flag": "#FF6A61",
        "purple": "#BF5AF2",
        "chip": "#1C2E47",
    },
}


def normalize_header(value: object) -> str:
    text = "" if value is None else str(value)
    # "İ".lower() ergibt "i" + kombinierender Punkt, deshalb vor lower() ersetzen.
    text = text.replace("İ", "i").strip().lower()
    text = text.replace("ı", "i")
    text = text.replace("ä", "a").replace("ö", "o").replace("ü", "u").replace("ß", "ss")
    text = text.replace("ç", "c").replace("ş", "s").replace("ğ", "g")
    text = text.replace("/", " / ")
    text = re.sub(r"\s+", " ", text)
    return text


def text_or_empty(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def clean_german_word(text: str) -> str:
    """Entfernt Grammatikangaben fuer Uebersetzung und Aussprache.

    "der Dialekt, -e" -> "der Dialekt", "das Lithium (nur Sg.)" -> "das Lithium",
    "der Verfechter, - / die Verfechterin, -nen" -> "der Verfechter".
    """
    text = re.sub(r"\([^)]*\)", " ", text or "")
    text = text.split(" / ")[0]
    text = re.sub(r",\s*[-–¨\"].*$", "", text)
    text = re.sub(r"\s+", " ", text).strip(" ,;")
    return text


def natural_sort(text: str) -> Tuple[int, str]:
    nums = re.findall(r"\d+", text or "")
    return (int(nums[0]) if nums else 9999, text or "")


@dataclass
class WordCard:
    id: int
    unit: str
    abschnitt: str
    german_word: str
    erklaerung: str
    synonym: str
    antonym: str
    english: str
    turkish: str
    beispiel: str
    known_count: int = 0
    unknown_count: int = 0
    box: int = 0
    due_date: Optional[str] = None
    favorite: int = 0
    difficult: int = 0
    last_result: Optional[str] = None

    @property
    def practiced(self) -> bool:
        return (self.known_count + self.unknown_count) > 0


WORD_CARD_FIELDS = [f.name for f in fields(WordCard)]


class Store:
    def __init__(self, path: Path = DB_PATH) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        self.path = path
        self.conn = sqlite3.connect(str(path))
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.init_db()

    def close(self) -> None:
        self.conn.close()

    def init_db(self) -> None:
        cur = self.conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS words (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                unit TEXT NOT NULL,
                abschnitt TEXT NOT NULL,
                german_word TEXT NOT NULL,
                erklaerung TEXT,
                synonym TEXT,
                antonym TEXT,
                english TEXT,
                turkish TEXT,
                beispiel TEXT,
                source_sheet TEXT,
                source_row INTEGER
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS progress (
                word_id INTEGER PRIMARY KEY,
                known_count INTEGER DEFAULT 0,
                unknown_count INTEGER DEFAULT 0,
                last_result TEXT,
                last_seen TEXT,
                FOREIGN KEY(word_id) REFERENCES words(id) ON DELETE CASCADE
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS translations (
                source_text TEXT NOT NULL,
                source_lang TEXT NOT NULL,
                target_lang TEXT NOT NULL,
                translation TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY (source_text, source_lang, target_lang)
            )
        """)
        # Migration fuer Datenbanken aus Version 1.0.
        existing = {r["name"] for r in cur.execute("PRAGMA table_info(progress)").fetchall()}
        for column, decl in [
            ("box", "INTEGER DEFAULT 0"),
            ("due_date", "TEXT"),
            ("favorite", "INTEGER DEFAULT 0"),
            ("difficult", "INTEGER DEFAULT 0"),
        ]:
            if column not in existing:
                cur.execute(f"ALTER TABLE progress ADD COLUMN {column} {decl}")
        self.conn.commit()

    def get_setting(self, key: str, default: str = "") -> str:
        row = self.conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
        return row["value"] if row else default

    def set_setting(self, key: str, value: str) -> None:
        self.conn.execute("INSERT OR REPLACE INTO settings(key, value) VALUES(?, ?)", (key, value))
        self.conn.commit()

    def has_words(self) -> bool:
        row = self.conn.execute("SELECT COUNT(*) AS c FROM words").fetchone()
        return bool(row and row["c"] > 0)

    def import_excel(self, excel_path: str) -> Tuple[int, int]:
        """Importiert die Excel-Datei. Gibt (Anzahl Karten, Anzahl uebernommener Lernstaende) zurueck."""
        rows = read_excel_words(excel_path)
        if not rows:
            raise ValueError("Keine passenden Wortschatzzeilen gefunden.\n\n"
                             "Benötigt wird mindestens eine Spalte \"Deutsche Word\" bzw. \"Wort\".")

        # Lernstand vor dem Neuimport sichern, damit er beim Aktualisieren erhalten bleibt.
        old_by_key: Dict[Tuple[str, str], dict] = {}
        old_by_word: Dict[str, List[dict]] = {}
        for r in self.conn.execute(
            "SELECT w.unit AS unit, w.german_word AS german_word, p.* FROM progress p JOIN words w ON w.id = p.word_id"
        ).fetchall():
            data = dict(r)
            old_by_key[(data["unit"], data["german_word"])] = data
            old_by_word.setdefault(data["german_word"], []).append(data)

        new_keys = {(r["unit"], r["german_word"]) for r in rows}
        used: set = set()
        restored = 0
        with self.conn:  # eine Transaktion: bei Fehlern bleibt der alte Stand erhalten
            cur = self.conn.cursor()
            cur.execute("DELETE FROM progress")
            cur.execute("DELETE FROM words")
            for r in rows:
                cur.execute(
                    """
                    INSERT INTO words(unit, abschnitt, german_word, erklaerung, synonym, antonym, english, turkish, beispiel, source_sheet, source_row)
                    VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        r["unit"], r["abschnitt"], r["german_word"], r["erklaerung"],
                        r["synonym"], r["antonym"], r["english"], r["turkish"], r["beispiel"],
                        r.get("source_sheet", ""), r.get("source_row", 0),
                    ),
                )
                prev = old_by_key.get((r["unit"], r["german_word"]))
                if prev is None:
                    # Einheit umbenannt? Nur zuordnen, wenn der alte Eintrag eindeutig ist und
                    # nicht ohnehin zu einer Zeile mit unveraenderter Einheit gehoert.
                    candidates = [p for p in old_by_word.get(r["german_word"], [])
                                  if (p["unit"], r["german_word"]) not in new_keys]
                    prev = candidates[0] if len(candidates) == 1 else None
                if prev and id(prev) not in used:
                    used.add(id(prev))
                    cur.execute(
                        """
                        INSERT INTO progress(word_id, known_count, unknown_count, last_result, last_seen, box, due_date, favorite, difficult)
                        VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (cur.lastrowid, prev["known_count"], prev["unknown_count"], prev["last_result"], prev["last_seen"],
                         prev["box"], prev["due_date"], prev["favorite"], prev["difficult"]),
                    )
                    restored += 1
            cur.execute("INSERT OR REPLACE INTO settings(key, value) VALUES('source_path', ?)", (excel_path,))
            cur.execute("INSERT OR REPLACE INTO settings(key, value) VALUES('last_import', ?)",
                        (datetime.now().isoformat(timespec="seconds"),))
        return len(rows), restored

    def reset_progress(self) -> None:
        with self.conn:
            self.conn.execute("DELETE FROM progress")

    def units(self) -> List[str]:
        rows = self.conn.execute("SELECT DISTINCT unit FROM words").fetchall()
        return sorted((r["unit"] for r in rows), key=natural_sort)

    def abschnitte(self, unit: Optional[str] = None) -> List[str]:
        if unit and unit != "Alle":
            rows = self.conn.execute("SELECT DISTINCT abschnitt FROM words WHERE unit = ? ORDER BY abschnitt", (unit,)).fetchall()
        else:
            rows = self.conn.execute("SELECT DISTINCT abschnitt FROM words ORDER BY abschnitt").fetchall()
        return [r["abschnitt"] for r in rows if r["abschnitt"]]

    def query_cards(self, unit: str = "Alle", abschnitt: str = "Alle", mode: str = "all") -> List[WordCard]:
        sql = """
            SELECT w.*,
                   COALESCE(p.known_count,0) AS known_count, COALESCE(p.unknown_count,0) AS unknown_count,
                   COALESCE(p.box,0) AS box, p.due_date AS due_date, p.last_result AS last_result,
                   COALESCE(p.favorite,0) AS favorite, COALESCE(p.difficult,0) AS difficult
            FROM words w
            LEFT JOIN progress p ON p.word_id = w.id
        """
        params: List[str] = []
        conditions = []
        if unit != "Alle":
            conditions.append("w.unit = ?")
            params.append(unit)
        if abschnitt != "Alle":
            conditions.append("w.abschnitt = ?")
            params.append(abschnitt)
        practiced = "COALESCE(p.known_count,0) + COALESCE(p.unknown_count,0) > 0"
        if mode == "due":
            conditions.append(f"(NOT ({practiced}) OR p.due_date IS NULL OR p.due_date <= ?)")
            params.append(date.today().isoformat())
        elif mode == "new":
            conditions.append(f"NOT ({practiced})")
        elif mode == "unknown":
            conditions.append("p.last_result = 'unknown'")
        elif mode == "favorite":
            conditions.append("p.favorite = 1")
        elif mode == "difficult":
            conditions.append("p.difficult = 1")
        if conditions:
            sql += " WHERE " + " AND ".join(conditions)
        sql += " ORDER BY w.id"
        rows = self.conn.execute(sql, params).fetchall()
        return [WordCard(**{k: row[k] for k in WORD_CARD_FIELDS}) for row in rows]

    def mark(self, word_id: int, result: str) -> Tuple[int, str]:
        """Speichert das Ergebnis und berechnet die naechste Wiederholung (Leitner-System)."""
        row = self.conn.execute("SELECT COALESCE(box,0) AS box FROM progress WHERE word_id = ?", (word_id,)).fetchone()
        box = row["box"] if row else 0
        today = date.today()
        if result == "known":
            box = min(box + 1, MAX_BOX)
            due = today + timedelta(days=LEITNER_INTERVALS[box])
        else:
            box = 1
            due = today
        field = "known_count" if result == "known" else "unknown_count"
        now = datetime.now().isoformat(timespec="seconds")
        self.conn.execute(
            f"""
            INSERT INTO progress(word_id, {field}, last_result, last_seen, box, due_date)
            VALUES(?, 1, ?, ?, ?, ?)
            ON CONFLICT(word_id) DO UPDATE SET
                {field} = COALESCE({field},0) + 1,
                last_result = excluded.last_result,
                last_seen = excluded.last_seen,
                box = excluded.box,
                due_date = excluded.due_date
            """,
            (word_id, result, now, box, due.isoformat()),
        )
        self.conn.commit()
        return box, due.isoformat()

    def set_flag(self, word_id: int, flag: str, value: bool) -> None:
        if flag not in {"favorite", "difficult"}:
            raise ValueError(flag)
        self.conn.execute(
            f"""
            INSERT INTO progress(word_id, {flag}) VALUES(?, ?)
            ON CONFLICT(word_id) DO UPDATE SET {flag} = excluded.{flag}
            """,
            (word_id, 1 if value else 0),
        )
        self.conn.commit()

    def flagged_words(self, flag: str) -> List[sqlite3.Row]:
        if flag not in {"favorite", "difficult"}:
            raise ValueError(flag)
        rows = self.conn.execute(
            f"SELECT w.* FROM words w JOIN progress p ON p.word_id = w.id WHERE p.{flag} = 1 ORDER BY w.id"
        ).fetchall()
        return rows

    def get_cached_translation(self, source_text: str, source_lang: str, target_lang: str) -> str:
        row = self.conn.execute(
            """
            SELECT translation FROM translations
            WHERE source_text = ? AND source_lang = ? AND target_lang = ?
            """,
            (source_text, source_lang, target_lang),
        ).fetchone()
        return row["translation"] if row else ""

    def save_cached_translation(self, source_text: str, source_lang: str, target_lang: str, translation: str) -> None:
        if not source_text or not translation:
            return
        self.conn.execute(
            """
            INSERT OR REPLACE INTO translations(source_text, source_lang, target_lang, translation, updated_at)
            VALUES(?, ?, ?, ?, ?)
            """,
            (source_text, source_lang, target_lang, translation, datetime.now().isoformat(timespec="seconds")),
        )
        self.conn.commit()

    def summary(self) -> Dict[str, int]:
        row = self.conn.execute(
            """
            SELECT COUNT(w.id) AS total,
                   SUM(CASE WHEN COALESCE(p.known_count,0)+COALESCE(p.unknown_count,0)=0 THEN 1 ELSE 0 END) AS new,
                   SUM(CASE WHEN COALESCE(p.known_count,0)+COALESCE(p.unknown_count,0)>0
                             AND (p.due_date IS NULL OR p.due_date <= ?) THEN 1 ELSE 0 END) AS due,
                   SUM(CASE WHEN COALESCE(p.box,0) >= ? THEN 1 ELSE 0 END) AS mastered,
                   SUM(COALESCE(p.favorite,0)) AS favorites,
                   SUM(COALESCE(p.difficult,0)) AS difficult
            FROM words w LEFT JOIN progress p ON p.word_id = w.id
            """,
            (date.today().isoformat(), MAX_BOX),
        ).fetchone()
        return {k: (row[k] or 0) for k in row.keys()}

    def stats(self) -> Dict[str, object]:
        overall = self.conn.execute(
            """
            SELECT COUNT(w.id) AS total,
                   SUM(COALESCE(p.known_count,0)) AS known,
                   SUM(COALESCE(p.unknown_count,0)) AS unknown,
                   SUM(CASE WHEN COALESCE(p.known_count,0)+COALESCE(p.unknown_count,0)>0 THEN 1 ELSE 0 END) AS practiced
            FROM words w LEFT JOIN progress p ON p.word_id = w.id
            """
        ).fetchone()
        by_unit = self.conn.execute(
            """
            SELECT w.unit AS unit, COUNT(w.id) AS total,
                   SUM(COALESCE(p.known_count,0)) AS known,
                   SUM(COALESCE(p.unknown_count,0)) AS unknown,
                   SUM(CASE WHEN COALESCE(p.known_count,0)+COALESCE(p.unknown_count,0)>0 THEN 1 ELSE 0 END) AS practiced,
                   SUM(CASE WHEN COALESCE(p.box,0) >= ? THEN 1 ELSE 0 END) AS mastered
            FROM words w LEFT JOIN progress p ON p.word_id = w.id
            GROUP BY w.unit
            """,
            (MAX_BOX,),
        ).fetchall()
        boxes = self.conn.execute(
            """
            SELECT COALESCE(p.box,0) AS box, COUNT(*) AS c
            FROM words w LEFT JOIN progress p ON p.word_id = w.id
            GROUP BY COALESCE(p.box,0)
            """
        ).fetchall()
        return {
            "overall": {k: (overall[k] or 0) for k in overall.keys()},
            # SQL "ORDER BY unit" sortiert "10. Lektion" vor "2. Lektion" - daher natuerlich sortieren.
            "by_unit": sorted((dict(r) for r in by_unit), key=lambda r: natural_sort(r["unit"])),
            "boxes": {r["box"]: r["c"] for r in boxes},
        }


def read_excel_words(excel_path: str) -> List[Dict[str, object]]:
    wb = load_workbook(excel_path, data_only=True, read_only=True)
    try:
        return _read_workbook_words(wb)
    finally:
        # Ohne close() bleibt die Datei im read-only-Modus gesperrt (Excel kann sie dann nicht speichern).
        wb.close()


def _read_workbook_words(wb) -> List[Dict[str, object]]:
    all_rows: List[Dict[str, object]] = []

    preferred_all_sheet = None
    for ws in wb.worksheets:
        if normalize_header(ws.title) in {"tum kelimeler", "alle worter", "all words"}:
            preferred_all_sheet = ws.title
            break

    for ws in wb.worksheets:
        if normalize_header(ws.title) in {"ozet", "summary", "zusammenfassung"}:
            continue
        if preferred_all_sheet and ws.title != preferred_all_sheet:
            continue

        header_row_idx, mapping = find_header(ws)
        if not header_row_idx or not mapping:
            continue

        unit_from_sheet = infer_unit_from_sheet(ws.title)
        for row_idx, row in enumerate(ws.iter_rows(min_row=header_row_idx + 1, values_only=True), start=header_row_idx + 1):
            if not row or not any(cell is not None for cell in row):
                continue
            row_values = {}
            for key, col in mapping.items():
                value = row[col - 1] if col - 1 < len(row) else None
                row_values[key] = text_or_empty(value)
            german_word = row_values.get("german_word", "")
            if not german_word or normalize_header(german_word) in {"kelime", "wort", "deutsche word", "almanca kelime"}:
                continue
            unit = row_values.get("unit", "") or unit_from_sheet
            abschnitt = row_values.get("abschnitt", "") or "Ohne Abschnitt"

            # Some combined EN/TR columns may exist in older source files.
            english = row_values.get("english", "")
            turkish = row_values.get("turkish", "")
            if english and not turkish and " / " in english:
                parts = [p.strip() for p in english.split(" / ", 1)]
                english, turkish = parts[0], parts[1]

            all_rows.append({
                "unit": unit or "Unbekannte Einheit",
                "abschnitt": abschnitt,
                "german_word": german_word,
                "erklaerung": row_values.get("erklaerung", ""),
                "synonym": row_values.get("synonym", ""),
                "antonym": row_values.get("antonym", ""),
                "english": english,
                "turkish": turkish,
                "beispiel": row_values.get("beispiel", ""),
                "source_sheet": ws.title,
                "source_row": row_idx,
            })
    return all_rows


def find_header(ws) -> Tuple[Optional[int], Dict[str, int]]:
    best_row = None
    best_mapping: Dict[str, int] = {}
    # iter_rows statt ws.cell(): im read-only-Modus wird sonst fuer jede Zelle neu geparst.
    aliases = {field: {normalize_header(a) for a in names} for field, names in COLUMN_ALIASES.items()}
    for r, raw in enumerate(ws.iter_rows(min_row=1, max_row=15, max_col=30, values_only=True), start=1):
        values = [normalize_header(v) for v in raw]
        mapping: Dict[str, int] = {}
        used_columns = set()
        for field, names in aliases.items():
            for idx, val in enumerate(values, start=1):
                if val in names and idx not in used_columns:
                    mapping[field] = idx
                    used_columns.add(idx)
                    break
        if "german_word" in mapping and len(mapping) > len(best_mapping):
            best_row, best_mapping = r, mapping
    return best_row, best_mapping


def infer_unit_from_sheet(sheet_name: str) -> str:
    match = re.search(r"(\d+)", sheet_name)
    if match:
        return f"Unit {match.group(1)}"
    return sheet_name.strip() or "Unbekannte Einheit"


def bundled_excel_path() -> Optional[Path]:
    """Die vom Installer mitgelieferte Wortliste neben der EXE bzw. dem Skript."""
    base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
    path = base / BUNDLED_EXCEL_NAME
    return path if path.exists() else None


def card_icon_text(card: Optional[WordCard]) -> str:
    if not card:
        return "★"
    text = " ".join([card.unit, card.abschnitt, card.german_word]).lower()
    if any(x in text for x in ["geld", "konto", "währung", "waehrung"]):
        return "€"
    if any(x in text for x in ["medizin", "schmerz", "therapie", "chirurgie"]):
        return "+"
    if any(x in text for x in ["kriminal", "gericht", "polizei", "straftat"]):
        return "§"
    if any(x in text for x in ["tier", "haustier", "wildtier", "zoologie"]):
        return "🐾"
    if any(x in text for x in ["chemie", "pestizid", "molekül", "element"]):
        return "⚗"
    if any(x in text for x in ["nachhalt", "ressource", "planet", "umwelt"]):
        return "♻"
    if any(x in text for x in ["sprache", "dialekt", "linguistik"]):
        return "Aa"
    word = clean_german_word(card.german_word)
    word = re.sub(r"^(der|die|das|sich)\s+", "", word, flags=re.IGNORECASE)
    return (word[:2] or "★").upper()


def color_for_text(text: str) -> str:
    palette = ["#007AFF", "#5856D6", "#AF52DE", "#FF2D55", "#FF9500", "#34C759", "#30B0C7"]
    digest = hashlib.md5(text.encode("utf-8")).hexdigest()
    return palette[int(digest[:2], 16) % len(palette)]


def mix(color_a: str, color_b: str, t: float) -> str:
    """Mischt zwei Hex-Farben (t=0 -> a, t=1 -> b)."""
    a = [int(color_a.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(color_b.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)]
    return "#" + "".join(f"{int(round(x + (y - x) * t)):02X}" for x, y in zip(a, b))


def lighten(hex_color: str, amount: float = 0.28) -> str:
    hex_color = hex_color.lstrip("#")
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    r = min(255, int(r + (255 - r) * amount))
    g = min(255, int(g + (255 - g) * amount))
    b = min(255, int(b + (255 - b) * amount))
    return f"#{r:02X}{g:02X}{b:02X}"


# ---------------------------------------------------------------------------
# Aussprache
# ---------------------------------------------------------------------------

_SAPI_SCRIPT = r"""
Add-Type -AssemblyName System.Speech
$s = New-Object System.Speech.Synthesis.SpeechSynthesizer
$v = $s.GetInstalledVoices() | Where-Object { $_.Enabled -and $_.VoiceInfo.Culture.Name -like 'de*' } | Select-Object -First 1
if (-not $v) { exit 3 }
$s.SelectVoice($v.VoiceInfo.Name)
$s.Speak($env:WK_SPEAK_TEXT)
"""


class Speaker:
    """Deutsche Aussprache: Audio-Cache -> gTTS (online) -> installierte deutsche Windows-Stimme."""

    def __init__(self, audio_dir: Path = AUDIO_DIR) -> None:
        self.audio_dir = audio_dir
        self._lock = threading.Lock()

    def speak(self, text: str) -> Optional[str]:
        """Spricht den Text (blockierend, im Hintergrund-Thread aufrufen). Gibt eine Fehlermeldung oder None zurueck."""
        text = clean_german_word(text)
        if not text:
            return None
        if not self._lock.acquire(blocking=False):
            return None  # es wird bereits gesprochen
        try:
            mp3 = self._cached_mp3(text)
            if mp3 and self._play_mp3(mp3):
                return None
            if self._speak_sapi(text):
                return None
            if gTTS is None:
                return ("Aussprache nicht verfügbar: Bitte \"pip install gTTS\" ausführen "
                        "oder eine deutsche Sprachausgabe in Windows installieren.")
            return ("Aussprache nicht verfügbar: Keine Internetverbindung und keine deutsche Windows-Stimme.\n"
                    "Windows-Einstellungen > Zeit und Sprache > Sprache > Deutsch hinzufügen (mit Sprachausgabe).")
        finally:
            self._lock.release()

    def _cached_mp3(self, text: str) -> Optional[Path]:
        path = self.audio_dir / (hashlib.sha1(text.encode("utf-8")).hexdigest() + ".mp3")
        if path.exists() and path.stat().st_size > 0:
            return path
        if gTTS is None:
            return None
        try:
            self.audio_dir.mkdir(parents=True, exist_ok=True)
            tmp = path.with_suffix(".part")
            gTTS(text=text, lang="de").save(str(tmp))
            os.replace(tmp, path)
            return path
        except Exception:
            return None

    @staticmethod
    def _play_mp3(path: Path) -> bool:
        if sys.platform != "win32":
            return False
        import ctypes
        mci = ctypes.windll.winmm.mciSendStringW
        alias = f"wk{threading.get_ident()}"
        if mci(f'open "{path}" type mpegvideo alias {alias}', None, 0, None) != 0:
            return False
        try:
            return mci(f"play {alias} wait", None, 0, None) == 0
        finally:
            mci(f"close {alias}", None, 0, None)

    @staticmethod
    def _speak_sapi(text: str) -> bool:
        if sys.platform != "win32":
            return False
        env = dict(os.environ, WK_SPEAK_TEXT=text)  # Text per Umgebungsvariable: keine Skript-Injektion
        try:
            result = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", _SAPI_SCRIPT],
                env=env, capture_output=True, timeout=30,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            return result.returncode == 0
        except Exception:
            return False


# ---------------------------------------------------------------------------
# UI-Bausteine
# ---------------------------------------------------------------------------

class ShapeRenderer:
    """Zeichnet Rundungen auf ein Tk-Canvas.

    Tk glaettet Canvas-Formen unter Windows nicht. Mit Pillow werden die Ecken deshalb als
    ueberabgetastete Bilder mit Alphakanal gezeichnet; ohne Pillow gibt es einfache Polygone.
    """

    SUPERSAMPLE = 4

    def __init__(self, root: tk.Misc) -> None:
        self.root = root
        self._cache: Dict[tuple, object] = {}

    def _corners(self, r: int, fill: str, outline: Optional[str]):
        key = ("corner", r, fill, outline)
        if key not in self._cache:
            ss = self.SUPERSAMPLE
            big = Image.new("RGBA", (r * ss, r * ss), (0, 0, 0, 0))
            ImageDraw.Draw(big).ellipse((0, 0, 2 * r * ss - 1, 2 * r * ss - 1), fill=fill,
                                        outline=outline, width=ss if outline else 0)
            tl = big.resize((r, r), Image.Resampling.LANCZOS)
            T = Image.Transpose
            self._cache[key] = [ImageTk.PhotoImage(img, master=self.root) for img in (
                tl, tl.transpose(T.FLIP_LEFT_RIGHT), tl.transpose(T.FLIP_TOP_BOTTOM), tl.transpose(T.ROTATE_180))]
        return self._cache[key]

    def _disc(self, d: int, fill: str):
        key = ("disc", d, fill)
        if key not in self._cache:
            ss = self.SUPERSAMPLE
            big = Image.new("RGBA", (d * ss, d * ss), (0, 0, 0, 0))
            ImageDraw.Draw(big).ellipse((0, 0, d * ss - 1, d * ss - 1), fill=fill)
            self._cache[key] = ImageTk.PhotoImage(big.resize((d, d), Image.Resampling.LANCZOS), master=self.root)
        return self._cache[key]

    def round_rect(self, canvas: tk.Canvas, x1: float, y1: float, x2: float, y2: float, r: float, fill: str,
                   outline: Optional[str] = None, tags: str = "shape") -> None:
        x1, y1, x2, y2 = (int(round(v)) for v in (x1, y1, x2, y2))
        w, h = x2 - x1, y2 - y1
        if w <= 0 or h <= 0:
            return
        r = int(max(0, min(r, w // 2, h // 2)))
        if Image is None or r < 2:
            pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
                   x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
            canvas.create_polygon(pts, smooth=True, fill=fill, outline=outline or "", tags=tags)
            return
        canvas.create_rectangle(x1 + r, y1, x2 - r, y2, fill=fill, outline="", tags=tags)
        canvas.create_rectangle(x1, y1 + r, x2, y2 - r, fill=fill, outline="", tags=tags)
        if outline:
            canvas.create_line(x1 + r, y1, x2 - r, y1, fill=outline, tags=tags)
            canvas.create_line(x1 + r, y2 - 1, x2 - r, y2 - 1, fill=outline, tags=tags)
            canvas.create_line(x1, y1 + r, x1, y2 - r, fill=outline, tags=tags)
            canvas.create_line(x2 - 1, y1 + r, x2 - 1, y2 - r, fill=outline, tags=tags)
        tl, tr, bl, br = self._corners(r, fill, outline)
        canvas.create_image(x1, y1, image=tl, anchor="nw", tags=tags)
        canvas.create_image(x2 - r, y1, image=tr, anchor="nw", tags=tags)
        canvas.create_image(x1, y2 - r, image=bl, anchor="nw", tags=tags)
        canvas.create_image(x2 - r, y2 - r, image=br, anchor="nw", tags=tags)

    def disc(self, canvas: tk.Canvas, cx: float, cy: float, d: int, fill: str, tags: str = "shape") -> None:
        if Image is None:
            canvas.create_oval(cx - d / 2, cy - d / 2, cx + d / 2, cy + d / 2, fill=fill, outline="", tags=tags)
        else:
            canvas.create_image(cx, cy, image=self._disc(int(d), fill), anchor="center", tags=tags)


class RoundButton(tk.Canvas):
    """Abgerundeter Knopf (Pillenform) mit Hover-Effekt."""

    VARIANTS = {
        # Variante: (Hintergrund, Hover, Text) - None = Farbe der Umgebung
        "default": ("button", "button_hover", "button_text"),
        "accent": ("accent", "accent_hover", "on_accent"),
        "success": ("success", "success_hover", "on_accent"),
        "danger": ("danger", "danger_hover", "on_accent"),
        "ghost": (None, "button_hover", "muted"),
        "chip": ("chip", "chip", "accent"),
        "toast": ("text", "text", "panel"),
    }

    def __init__(self, master, shapes: ShapeRenderer, text: str, command: Optional[Callable], variant: str = "default",
                 font=None, padx: int = 16, pady: int = 8, outer: str = "panel", radius: Optional[int] = None) -> None:
        super().__init__(master, highlightthickness=0, bd=0, takefocus=0)
        self.shapes = shapes
        self.text = text
        self.command = command
        self.variant = variant
        self.btn_font = font or tkfont.nametofont("TkDefaultFont")
        self.padx, self.pady = padx, pady
        self.outer = outer
        self.radius = radius
        self.colors: Optional[Dict[str, str]] = None
        self.fg_token: Optional[str] = None
        self.enabled = True
        self.hover = False
        self.pressed = False
        self._update_size()
        self.bind("<Enter>", lambda e: self._set_hover(True))
        self.bind("<Leave>", lambda e: self._set_hover(False))
        self.bind("<ButtonPress-1>", self._press)
        self.bind("<ButtonRelease-1>", self._release)
        self.bind("<Configure>", lambda e: self.redraw())

    def _update_size(self) -> None:
        width = self.btn_font.measure(self.text) + 2 * self.padx if self.text else 1
        height = self.btn_font.metrics("linespace") + 2 * self.pady
        tk.Canvas.configure(self, width=width, height=height)

    def configure(self, cnf=None, **kw):
        if isinstance(cnf, dict):
            kw.update(cnf)
        text_changed = "text" in kw
        if text_changed:
            self.text = kw.pop("text")
            self._update_size()
        result = tk.Canvas.configure(self, **kw) if kw else None
        if text_changed:
            self.redraw()
        return result

    config = configure

    def restyle(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.redraw()

    def set_enabled(self, enabled: bool) -> None:
        self.enabled = enabled
        if not enabled:
            self.hover = self.pressed = False
        self.redraw()

    def set_fg(self, token: Optional[str]) -> None:
        self.fg_token = token
        self.redraw()

    def redraw(self) -> None:
        if not self.colors:
            return
        c = self.colors
        self.delete("all")
        outer = c[self.outer]
        clickable = self.enabled and self.command is not None
        tk.Canvas.configure(self, bg=outer, cursor="hand2" if clickable else "arrow")
        if not self.text:
            return
        w = max(self.winfo_width(), 1) if self.winfo_ismapped() else int(float(self.cget("width")))
        h = max(self.winfo_height(), 1) if self.winfo_ismapped() else int(float(self.cget("height")))
        bg_key, hover_key, fg_key = self.VARIANTS[self.variant]
        fill = c[bg_key] if bg_key else outer
        fg = c[self.fg_token or fg_key]
        if not self.enabled:
            fill = outer if bg_key is None else c["button"]
            fg = c["muted"]
        elif self.hover or self.pressed:
            fill = c[hover_key]
        if fill != outer:
            self.shapes.round_rect(self, 0, 0, w, h, self.radius if self.radius is not None else h // 2, fill)
        self.create_text(w / 2, h / 2, text=self.text, fill=fg, font=self.btn_font)

    def _set_hover(self, inside: bool) -> None:
        if self.enabled and self.command is not None:
            self.hover = inside
            self.redraw()

    def _press(self, event) -> None:
        if self.enabled and self.command is not None:
            self.pressed = True
            self.redraw()

    def _release(self, event) -> None:
        was_pressed = self.pressed
        self.pressed = False
        self.redraw()
        inside = 0 <= event.x < self.winfo_width() and 0 <= event.y < self.winfo_height()
        if was_pressed and inside and self.enabled and self.command is not None:
            self.command()


class RoundedPanel(tk.Canvas):
    """Flaeche mit abgerundeten Ecken. Inhalte kommen in panel.inner (ein normaler Frame)."""

    def __init__(self, master, shapes: ShapeRenderer, radius: int, fill: str = "panel", outline: Optional[str] = "line",
                 outer: str = "bg", padding: int = 0, fit_height: bool = False, **kwargs) -> None:
        # Ein Canvas fordert sonst ca. 380x270 Pixel an und verdraengt bei wenig Platz andere Elemente.
        kwargs.setdefault("width", 1)
        kwargs.setdefault("height", 1)
        super().__init__(master, highlightthickness=0, bd=0, **kwargs)
        self.shapes = shapes
        self.radius = radius
        self.fill, self.outline, self.outer = fill, outline, outer
        self.padding = padding
        self.fit_height = fit_height
        self.colors: Optional[Dict[str, str]] = None
        self.squeeze = 1.0  # horizontale Stauchung fuer die Umdreh-Animation
        self.animating = False
        self.inner = tk.Frame(self, bd=0, highlightthickness=0)
        self._window = self.create_window(padding, padding, window=self.inner, anchor="nw")
        self.bind("<Configure>", self._on_configure)
        if fit_height:
            # Solange das Canvas 1 px hoch ist, wird der Inhalt nicht angezeigt und meldet kein <Configure>;
            # deshalb die Hoehe zusaetzlich nach dem Aufbau und bei jeder Groessenaenderung anpassen.
            self.inner.bind("<Configure>", self._fit_height, add="+")
            self.after_idle(self._fit_height)

    def _on_configure(self, event) -> None:
        options = {"width": max(1, event.width - 2 * self.padding)}
        if not self.fit_height:
            options["height"] = max(1, event.height - 2 * self.padding)
        self.itemconfigure(self._window, **options)
        if self.fit_height:
            self._fit_height()
        self.redraw()

    def _fit_height(self, event=None) -> None:
        height = self.inner.winfo_reqheight() + 2 * self.padding
        if int(float(self.cget("height"))) != height:
            tk.Canvas.configure(self, height=height)

    def restyle(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        tk.Canvas.configure(self, bg=colors[self.outer])
        self.redraw()

    def set_padding(self, padding: int) -> None:
        self.padding = padding
        self.coords(self._window, padding, padding)
        options = {"width": max(1, self.winfo_width() - 2 * padding)}
        if not self.fit_height:
            options["height"] = max(1, self.winfo_height() - 2 * padding)
        self.itemconfigure(self._window, **options)
        if self.fit_height:
            self._fit_height()

    def redraw(self) -> None:
        if not self.colors:
            return
        self.delete("shape")
        w, h = self.winfo_width(), self.winfo_height()
        if w <= 1 or h <= 1:
            return
        sw = w * self.squeeze
        x1 = (w - sw) / 2
        self.shapes.round_rect(self, x1, 0, x1 + sw, h, self.radius, self.colors[self.fill],
                               self.colors[self.outline] if self.outline else None)
        self.tag_lower("shape")

    def flip(self, on_middle: Callable[[], None], steps: int = 7, delay: int = 12) -> None:
        """Karten-Umdreh-Animation: Flaeche schrumpft auf 0 und waechst mit dem neuen Inhalt wieder auf."""
        if self.animating:
            return
        self.animating = True
        self.itemconfigure(self._window, state="hidden")
        frames = ([math.cos(i / steps * math.pi / 2) for i in range(1, steps + 1)] +
                  [math.sin(i / steps * math.pi / 2) for i in range(1, steps + 1)])

        def step(i: int) -> None:
            if i == steps:
                on_middle()
            if i >= len(frames):
                self.squeeze = 1.0
                self.redraw()
                self.itemconfigure(self._window, state="normal")
                self.animating = False
                return
            self.squeeze = max(0.02, frames[i])
            self.redraw()
            self.after(delay, step, i + 1)

        step(0)


def pick_font(root: tk.Misc, candidates: List[str]) -> str:
    available = set(tkfont.families(root))
    for name in candidates:
        if name in available:
            return name
    return "TkDefaultFont"


class WortschatzApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.store = Store()
        self.speaker = Speaker()
        self.theme_name = self.store.get_setting("theme", "light") or "light"
        if self.theme_name not in THEMES:
            self.theme_name = "light"
        self.colors = THEMES[self.theme_name]
        self.cards: List[WordCard] = []
        self.current: Optional[WordCard] = None
        self.is_back = False
        self.round_finished = False
        self.deck_index = 0
        self.translation_job_id = 0
        self.session = {"known": 0, "unknown": 0}
        self._themed: List[Tuple[tk.Misc, Dict[str, str]]] = []
        self._buttons: List[RoundButton] = []
        self._panels: List[RoundedPanel] = []
        self.shapes = ShapeRenderer(self)
        self._ui_queue: "queue.Queue[Callable[[], None]]" = queue.Queue()
        self._speaking = False
        self._fonts: Dict[tuple, tkfont.Font] = {}

        self.scale = self.winfo_fpixels("1i") / 96.0
        self.ui_font = pick_font(self, ["Segoe UI", "Helvetica", "Arial"])
        self.display_font = pick_font(self, ["Segoe UI Variable Display", "Segoe UI Variable", "Segoe UI Semibold", "Segoe UI"])

        saved_mutter = self.store.get_setting("muttersprache", "Türkisch") or "Türkisch"
        if saved_mutter not in MUTTERSPRACHEN:
            saved_mutter = "Türkisch"
        saved_mode = self.store.get_setting("lernmodus", "Wiederholung (fällig + neu)")
        if saved_mode not in LERNMODI:
            saved_mode = "Wiederholung (fällig + neu)"
        self.muttersprache_var = tk.StringVar(value=saved_mutter)
        self.mode_var = tk.StringVar(value=saved_mode)
        self.unit_var = tk.StringVar(value=self.store.get_setting("filter_unit", "Alle") or "Alle")
        self.abschnitt_var = tk.StringVar(value=self.store.get_setting("filter_abschnitt", "Alle") or "Alle")

        self._side_visible = True
        self._compact_actions = False
        self._illustration_visible = True
        self._compact_height = False
        self._resize_job: Optional[str] = None

        self.title(APP_NAME)
        self._set_window_icon()
        self._apply_initial_geometry()
        self.configure(bg=self.colors["bg"])
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.create_menu()
        self.create_widgets()
        self.bind_shortcuts()
        self.bind("<Configure>", self._on_root_configure)
        self.apply_theme()
        self.after(100, self._poll_ui_queue)
        self.after(150, self.ensure_data_loaded)

    # ----- Helfer -------------------------------------------------------

    def work_area(self) -> Tuple[int, int, int, int]:
        """Nutzbarer Bildschirmbereich ohne Taskleiste: (links, oben, Breite, Hoehe) in Pixeln."""
        if sys.platform == "win32":
            try:
                import ctypes
                from ctypes import wintypes
                rect = wintypes.RECT()
                if ctypes.windll.user32.SystemParametersInfoW(0x30, 0, ctypes.byref(rect), 0):  # SPI_GETWORKAREA
                    return rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top
            except Exception:
                pass
        return 0, 0, self.winfo_screenwidth(), self.winfo_screenheight() - 60

    def fit_size(self, width: int, height: int) -> Tuple[int, int]:
        """Begrenzt eine Fenstergroesse auf den Bildschirm (inkl. Titelleiste und Rahmen)."""
        _, _, area_w, area_h = self.work_area()
        return min(width, area_w - self.px(16)), min(height, area_h - self.px(40))

    def _apply_initial_geometry(self) -> None:
        left, top, area_w, area_h = self.work_area()
        min_w, min_h = self.fit_size(self.px(760), self.px(520))
        self.minsize(min_w, min_h)
        width, height = self.fit_size(self.px(1180), self.px(800))
        x = left + (area_w - width) // 2
        y = top + max(0, (area_h - height) // 2 - self.px(16))
        saved = re.fullmatch(r"(\d+)x(\d+)\+(-?\d+)\+(-?\d+)", self.store.get_setting("geometry", "") or "")
        if saved:
            sw, sh, sx, sy = map(int, saved.groups())
            sw, sh = self.fit_size(max(sw, min_w), max(sh, min_h))
            vroot_w, vroot_h = self.winfo_vrootwidth(), self.winfo_vrootheight()
            # Nur uebernehmen, wenn das Fenster sichtbar bleibt (z. B. zweiter Monitor inzwischen abgesteckt).
            if -vroot_w < sx < vroot_w - self.px(120) and -self.px(10) <= sy < vroot_h - self.px(120):
                width, height, x, y = sw, sh, sx, sy
        self.geometry(f"{width}x{height}+{x}+{y}")
        small_screen = area_w < self.px(1100) or area_h < self.px(760)
        if self.store.get_setting("zoomed", "0") == "1" or small_screen:
            try:
                self.state("zoomed")
            except tk.TclError:
                pass

    def px(self, value: float) -> int:
        """Pixelwerte an die Bildschirmskalierung (DPI) anpassen."""
        return int(round(value * self.scale))

    def font(self, size: int, weight: str = "normal", display: bool = False, slant: str = "roman") -> tkfont.Font:
        # Font-Objekte zwischenspeichern: Tk loescht benannte Fonts, sobald Python sie freigibt.
        key = (size, weight, display, slant)
        if key not in self._fonts:
            family = self.display_font if display else self.ui_font
            self._fonts[key] = tkfont.Font(self, family=family, size=size, weight=weight, slant=slant)
        return self._fonts[key]

    def themed(self, widget, **tokens: str):
        self._themed.append((widget, tokens))
        return widget

    def make_frame(self, parent, bg: str = "panel", border: bool = False, **kwargs) -> tk.Frame:
        frame = tk.Frame(parent, **kwargs)
        if border:
            frame.configure(highlightthickness=1)
            return self.themed(frame, bg=bg, highlightbackground="line", highlightcolor="line")
        return self.themed(frame, bg=bg)

    def make_label(self, parent, text: str = "", font=None, fg: str = "text", bg: str = "panel", **kwargs) -> tk.Label:
        label = tk.Label(parent, text=text, font=font, **kwargs)
        return self.themed(label, bg=bg, fg=fg)

    def make_button(self, parent, text: str, command: Optional[Callable], variant: str = "default",
                    outer: str = "panel", **kwargs) -> RoundButton:
        kwargs.setdefault("font", self.font(10, "bold"))
        kwargs.setdefault("padx", self.px(18))
        kwargs.setdefault("pady", self.px(9))
        button = RoundButton(parent, self.shapes, text, command, variant=variant, outer=outer, **kwargs)
        self._buttons.append(button)
        return button

    def make_panel(self, parent, fill: str = "panel", outline: Optional[str] = "line", outer: str = "bg",
                   radius: int = 18, padding: int = 8, fit_height: bool = False, **kwargs) -> RoundedPanel:
        panel = RoundedPanel(parent, self.shapes, self.px(radius), fill, outline, outer, self.px(padding), fit_height, **kwargs)
        self.themed(panel.inner, bg=fill)
        self._panels.append(panel)
        return panel

    def run_in_background(self, work: Callable[[], object], done: Callable[[object], None]) -> None:
        """Fuehrt work() im Hintergrund aus; done() laeuft danach im Tk-Hauptthread.

        Tkinter ist nicht thread-sicher - Hintergrund-Threads duerfen keine Tk-Methoden
        (auch nicht after()) aufrufen. Deshalb gehen Ergebnisse ueber eine Queue.
        """
        def worker() -> None:
            try:
                result = work()
            except Exception as exc:  # Fehler an den Hauptthread weitergeben statt den Thread still sterben zu lassen
                traceback.print_exc()
                result = exc
            self._ui_queue.put(lambda: done(result))
        threading.Thread(target=worker, daemon=True).start()

    def _poll_ui_queue(self) -> None:
        try:
            while True:
                callback = self._ui_queue.get_nowait()
                try:
                    callback()
                except Exception:
                    self.report_callback_exception(*sys.exc_info())
        except queue.Empty:
            pass
        finally:
            self.after(80, self._poll_ui_queue)

    def _set_window_icon(self) -> None:
        base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
        icon = base / "Wortschatz_karten.ico"
        if icon.exists():
            try:
                self.iconbitmap(default=str(icon))
            except tk.TclError:
                pass

    # ----- Aufbau -------------------------------------------------------

    def create_menu(self) -> None:
        menubar = tk.Menu(self)
        data_menu = tk.Menu(menubar, tearoff=0)
        data_menu.add_command(label="Excel aktualisieren...", accelerator="Strg+O", command=self.update_excel_dialog)
        data_menu.add_command(label="Favoriten & schwierige Wörter exportieren...", accelerator="Strg+E",
                              command=self.export_marked_words)
        data_menu.add_separator()
        data_menu.add_command(label="Lernfortschritt zurücksetzen...", command=self.reset_progress)
        data_menu.add_separator()
        data_menu.add_command(label="Beenden", command=self.on_close)
        menubar.add_cascade(label="Daten", menu=data_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Helles Design", command=lambda: self.set_theme("light"))
        view_menu.add_command(label="Dunkles Design", command=lambda: self.set_theme("dark"))
        view_menu.add_separator()
        view_menu.add_command(label="Statistik", command=self.show_stats)
        menubar.add_cascade(label="Ansicht", menu=view_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Tastenkürzel", command=self.show_shortcuts)
        help_menu.add_command(label="Über", command=self.show_about)
        menubar.add_cascade(label="Hilfe", menu=help_menu)
        self.config(menu=menubar)

    def create_widgets(self) -> None:
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        px = self.px

        self.main = self.make_frame(self, bg="bg")
        self.main.pack(fill="both", expand=True, padx=px(28), pady=(px(20), px(24)))

        # Kopfzeile
        header = self.make_frame(self.main, bg="bg")
        header.pack(fill="x")
        self.logo = self.themed(tk.Canvas(header, width=px(46), height=px(46), highlightthickness=0, bd=0), bg="bg")
        self.logo.pack(side="left")
        title_box = self.make_frame(header, bg="bg")
        title_box.pack(side="left", padx=(px(14), 0))
        self.make_label(title_box, "Wortschatz Karten", self.font(20, "bold", display=True), bg="bg").pack(anchor="w")
        self.subtitle_label = self.make_label(title_box, "Lernkarten für den DSH-Wortschatz", self.font(10), fg="muted", bg="bg")
        self.subtitle_label.pack(anchor="w")
        self.theme_button = self.make_button(header, "", self.toggle_theme, outer="bg")
        self.theme_button.pack(side="right")
        self.stats_button = self.make_button(header, "Statistik", self.show_stats, outer="bg")
        self.stats_button.pack(side="right", padx=(0, px(8)))

        # Filterleiste
        filterbar = self.make_panel(self.main, padding=16, fit_height=True)
        filterbar.pack(fill="x", pady=(px(18), px(16)))
        self.filterbar = filterbar
        inner = filterbar.inner
        # Kleine Mindestbreiten + Gewichte: die Felder wachsen mit dem Fenster und schrumpfen bei wenig Platz.
        combos = [
            ("EINHEIT", self.unit_var, 9, 2, self.on_unit_changed),
            ("ABSCHNITT", self.abschnitt_var, 12, 4, self.on_filter_changed),
            ("LERNMODUS", self.mode_var, 12, 3, self.on_filter_changed),
            ("MUTTERSPRACHE", self.muttersprache_var, 9, 2, self.on_muttersprache_changed),
        ]
        self.combos: List[ttk.Combobox] = []
        for col, (caption, var, width, weight, handler) in enumerate(combos):
            inner.grid_columnconfigure(col, weight=weight)
            pad = (px(4) if col == 0 else px(14), 0)
            self.make_label(inner, caption, self.font(8, "bold"), fg="muted").grid(row=0, column=col, sticky="w", padx=pad)
            combo = ttk.Combobox(inner, textvariable=var, state="readonly", width=width, font=self.font(10))
            combo.grid(row=1, column=col, sticky="we", padx=pad, pady=(px(5), 0))
            combo.bind("<<ComboboxSelected>>", handler)
            self.combos.append(combo)
        self.unit_combo, self.abschnitt_combo, self.mode_combo, self.muttersprache_combo = self.combos
        self.mode_combo.configure(values=list(LERNMODI.keys()))
        self.muttersprache_combo.configure(values=list(MUTTERSPRACHEN.keys()))
        self.shuffle_button = self.make_button(inner, "⤨  Mischen", self.shuffle_deck)
        self.shuffle_button.grid(row=1, column=4, sticky="e", padx=(px(14), px(4)), pady=(px(5), 0))

        # Inhalt: Karte + Seitenleiste
        content = self.make_frame(self.main, bg="bg")
        content.pack(fill="both", expand=True)

        self.side_panel = self.make_panel(content, width=px(256))
        self.side_panel.pack(fill="y", side="right", padx=(px(16), 0))
        self.build_side_panel(self.side_panel.inner)
        self.side_panel.inner.bind("<Configure>", lambda e: self.after_idle(self._fit_side), add="+")

        self.card_panel = self.make_panel(content)
        self.card_panel.pack(fill="both", expand=True, side="left")
        self.card_frame = self.card_panel.inner
        self.build_card(self.card_frame)

    def build_card(self, card: tk.Frame) -> None:
        px = self.px
        top = self.make_frame(card)
        top.pack(fill="x", padx=px(18), pady=(px(12), 0))
        self.card_meta = self.make_button(top, "", None, variant="chip", font=self.font(9, "bold"), padx=px(12), pady=px(5))
        self.card_meta.pack(side="left")
        icon = dict(variant="ghost", padx=px(9), pady=px(3))
        self.fav_button = self.make_button(top, "☆", lambda: self.toggle_flag("favorite"), font=self.font(14), **icon)
        self.fav_button.pack(side="right")
        self.difficult_button = self.make_button(top, "⚑", lambda: self.toggle_flag("difficult"), font=self.font(14), **icon)
        self.difficult_button.pack(side="right", padx=(0, px(2)))
        self.speak_button = self.make_button(top, "🔊", self.speak_current, font=self.font(13), **icon)
        self.speak_button.pack(side="right", padx=(0, px(2)))
        self.counter_label = self.make_label(top, "", self.font(10, "bold"), fg="muted")
        self.counter_label.pack(side="right", padx=(0, px(12)))

        self.progress_canvas = self.themed(tk.Canvas(card, height=px(6), highlightthickness=0, bd=0), bg="panel")
        self.progress_canvas.pack(fill="x", padx=px(20), pady=(px(12), 0))
        self.progress_canvas.bind("<Configure>", lambda e: self.draw_progress())

        # Die "Kartenflaeche" - sie wird beim Umdrehen animiert.
        self.face = self.make_panel(card, fill="soft", outline=None, outer="panel", radius=16, padding=12)
        self.face.pack(fill="both", expand=True, padx=px(20), pady=(px(14), px(10)))
        self.card_body = self.face.inner
        self.card_body.bind("<Configure>", self._on_body_resize)
        self.spacer_top = self.make_frame(self.card_body, bg="soft")
        self.illustration = self.themed(tk.Canvas(self.card_body, width=px(132), height=px(132), highlightthickness=0, bd=0), bg="soft")
        self.word_font_front = self.font(34, "bold", display=True)
        self.word_font_back = self.font(24, "bold", display=True)
        self.word_label = self.make_label(self.card_body, "", self.word_font_front, bg="soft", justify="center", wraplength=px(640))
        self.hint_label = self.make_label(self.card_body, "Leertaste drücken, um die Karte umzudrehen", self.font(10),
                                          fg="muted", bg="soft")
        self.spacer_bottom = self.make_frame(self.card_body, bg="soft")

        details_wrap = self.make_frame(self.card_body, bg="soft")
        self.details_wrap = details_wrap
        self.details_text = tk.Text(details_wrap, wrap="word", bd=0, highlightthickness=0, font=self.font(11), height=3,
                                    padx=px(12), pady=px(4), cursor="arrow")
        # Schlanke Scrollleiste ohne Pfeile
        self.style.layout("Slim.Vertical.TScrollbar", [("Vertical.Scrollbar.trough", {"sticky": "ns", "children": [
            ("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
        self.details_scroll = ttk.Scrollbar(details_wrap, orient="vertical", command=self.details_text.yview,
                                            style="Slim.Vertical.TScrollbar")
        self.details_text.configure(yscrollcommand=self._on_details_scroll, state="disabled")
        self.details_text.pack(side="left", fill="both", expand=True)
        indent = self.font(8, "bold").measure("●  ")
        self.details_text.tag_configure("label", font=self.font(8, "bold"), spacing1=px(12), spacing3=px(3))
        self.details_text.tag_configure("value", font=self.font(11), spacing3=px(2), lmargin1=indent, lmargin2=indent)
        self.details_text.tag_configure("highlight", font=self.font(12, "bold"), spacing3=px(2), lmargin1=indent, lmargin2=indent)
        self.details_text.tag_configure("example", font=self.font(11, slant="italic"), spacing3=px(2), lmargin1=indent, lmargin2=indent)
        self.details_text.tag_configure("error", font=self.font(10, slant="italic"), spacing3=px(2), lmargin1=indent, lmargin2=indent)
        self.details_text.tag_configure("info", font=self.font(11), justify="center", spacing1=px(8))
        self.details_text.tag_configure("first", spacing1=0)

        self.toast = self.make_button(card, "", None, variant="toast", outer="soft", font=self.font(9, "bold"),
                                      padx=px(16), pady=px(8))
        self._toast_job: Optional[str] = None

        # Vor der Kartenflaeche packen: bei wenig Hoehe schrumpft dann die Flaeche, nie die Knoepfe.
        actions = self.make_frame(card)
        actions.pack(fill="x", side="bottom", padx=px(20), pady=(px(4), px(14)), before=self.face)
        self.actions_frame = actions
        actions.bind("<Configure>", self._fit_actions)
        big = dict(font=self.font(10, "bold"), padx=px(18), pady=px(12))
        self.prev_button = self.make_button(actions, "←", self.prev_card, **dict(big, padx=px(16)))
        self.prev_button.pack(side="left")
        self.flip_button = self.make_button(actions, "Umdrehen", self.flip_card, variant="accent", **big)
        self.flip_button.pack(side="left", padx=(px(8), 0))
        self.next_button = self.make_button(actions, "Weiter  →", self.next_card, **big)
        self.next_button.pack(side="left", padx=(px(8), 0))
        self.known_button = self.make_button(actions, "✓  Gewusst", lambda: self.mark_current("known"), variant="success", **big)
        self.known_button.pack(side="right")
        self.unknown_button = self.make_button(actions, "✕  Nicht gewusst", lambda: self.mark_current("unknown"), variant="danger", **big)
        self.unknown_button.pack(side="right", padx=(0, px(8)))

    def build_side_panel(self, side: tk.Frame) -> None:
        px = self.px
        pad = dict(padx=px(14))
        self.side_sections: List[tk.Frame] = []

        def section(title: str, top: int = 14) -> tk.Frame:
            frame = self.make_frame(side)
            frame.pack(fill="x")
            self.make_label(frame, title, self.font(8, "bold"), fg="muted").pack(anchor="w", pady=(px(top), px(4)), **pad)
            self.side_sections.append(frame)
            return frame

        def stat_row(parent: tk.Frame, name: str) -> tk.Label:
            row = self.make_frame(parent)
            row.pack(fill="x", pady=px(1), **pad)
            self.make_label(row, name, self.font(10), fg="muted").pack(side="left")
            value = self.make_label(row, "0", self.font(10, "bold"))
            value.pack(side="right")
            return value

        sec = section("DIESE SITZUNG", top=10)
        self.stat_deck = stat_row(sec, "Karten im Stapel")
        self.stat_known = stat_row(sec, "Gewusst")
        self.stat_unknown = stat_row(sec, "Nicht gewusst")

        sec = section("ÜBERBLICK")
        self.stat_due = stat_row(sec, "Heute fällig")
        self.stat_new = stat_row(sec, "Neue Karten")
        self.stat_mastered = stat_row(sec, "Sicher gelernt (Box 5)")
        self.stat_fav = stat_row(sec, "Favoriten")
        self.stat_difficult = stat_row(sec, "Schwierige Wörter")

        sec = section("AKTUELLE KARTE")
        self.box_canvas = self.themed(tk.Canvas(sec, height=px(18), highlightthickness=0, bd=0), bg="panel")
        self.box_canvas.pack(fill="x", **pad)
        self.box_label = self.make_label(sec, "", self.font(9), fg="muted", justify="left", wraplength=px(210))
        self.box_label.pack(anchor="w", pady=(px(6), 0), **pad)

        sec = section("TASTENKÜRZEL")
        shortcuts = [("Leertaste", "Umdrehen"), ("← / →", "Zurück / Weiter"), ("1 / 2", "Nicht gewusst / Gewusst"),
                     ("S", "Aussprache"), ("F / D", "Favorit / Schwierig")]
        for key, action in shortcuts:
            row = self.make_frame(sec)
            row.pack(fill="x", pady=px(1), **pad)
            self.make_label(row, key, self.font(9, "bold")).pack(side="left")
            self.make_label(row, action, self.font(9), fg="muted").pack(side="right")

    # ----- Anpassung an die Fenstergroesse ------------------------------

    def _on_root_configure(self, event) -> None:
        if event.widget is not self:
            return
        if self._resize_job:
            self.after_cancel(self._resize_job)
        self._resize_job = self.after(40, self._apply_responsive)

    def _apply_responsive(self) -> None:
        """Schmale Fenster: Seitenleiste ausblenden (Statistik bleibt ueber den Knopf erreichbar)."""
        self._resize_job = None
        show_side = self.winfo_width() >= self.px(1040)
        if show_side != self._side_visible:
            self._side_visible = show_side
            if show_side:
                self.side_panel.pack(fill="y", side="right", padx=(self.px(16), 0), before=self.card_panel)
            else:
                self.side_panel.pack_forget()
        compact = self.winfo_height() < self.px(700)
        if compact != self._compact_height:
            self._set_compact_height(compact)
        self.after_idle(self._fit_side)
        self.after_idle(self._fit_actions)

    def _set_compact_height(self, compact: bool) -> None:
        """Niedrige Fenster: Abstaende verkleinern, Untertitel ausblenden, Logo und Schrift verkleinern."""
        px = self.px
        self._compact_height = compact
        self.main.pack_configure(pady=(px(10), px(10)) if compact else (px(20), px(24)))
        if compact:
            self.subtitle_label.pack_forget()
        else:
            self.subtitle_label.pack(anchor="w")
        size = px(34) if compact else px(46)
        self.logo.configure(width=size, height=size)
        self.draw_logo()
        self.filterbar.pack_configure(pady=(px(10), px(10)) if compact else (px(18), px(16)))
        self.filterbar.set_padding(px(10) if compact else px(16))
        self.progress_canvas.pack_configure(pady=(px(8), 0) if compact else (px(12), 0))
        self.face.pack_configure(pady=(px(8), px(6)) if compact else (px(14), px(10)))
        self.after_idle(self.render_card)

    def _fit_side(self) -> None:
        """Abschnitte der Seitenleiste von unten her ausblenden, wenn die Hoehe nicht reicht."""
        if not self._side_visible:
            return
        available = self.side_panel.inner.winfo_height() - self.px(8)
        if available <= 1:
            return
        visible = list(self.side_sections)
        for optional in (self.side_sections[3], self.side_sections[2]):  # erst Tastenkuerzel, dann aktuelle Karte
            if sum(sec.winfo_reqheight() for sec in visible) <= available:
                break
            visible.remove(optional)
        if [sec for sec in self.side_sections if sec.winfo_manager()] != visible:
            for sec in self.side_sections:
                sec.pack_forget()
            for sec in visible:
                sec.pack(fill="x")

    def _fit_actions(self, event=None) -> None:
        """Bei wenig Breite zeigen die Knoepfe nur Symbole statt Text."""
        available = self.actions_frame.winfo_width()
        if available <= 1:
            return
        f = self.font(10, "bold")
        pad = 2 * self.px(18)
        needed = (f.measure("←") + 2 * self.px(16) + sum(f.measure(t) + pad for t in (
            "Vorderseite", "Neue Runde  →", "✕  Nicht gewusst", "✓  Gewusst")) + 3 * self.px(8) + self.px(24))
        compact = available < needed
        if compact != self._compact_actions:
            self._compact_actions = compact
            self._update_action_labels()

    def _update_action_labels(self) -> None:
        compact = self._compact_actions
        new_round = self.round_finished or not self.cards
        self.next_button.configure(text="→" if compact else ("Neue Runde  →" if new_round else "Weiter  →"))
        self.unknown_button.configure(text="✕" if compact else "✕  Nicht gewusst")
        self.known_button.configure(text="✓" if compact else "✓  Gewusst")

    def _face_has_room_for_illustration(self) -> bool:
        height = self.card_body.winfo_height()
        if height <= 1:
            return True
        needed = (self.px(132) + self.px(16) + self.word_label.winfo_reqheight() + self.px(10)
                  + self.hint_label.winfo_reqheight() + self.px(16))
        return height >= needed

    def bind_shortcuts(self) -> None:
        self.bind_all("<KeyPress>", self._on_key)
        self.bind_all("<Control-o>", lambda e: self.update_excel_dialog())
        self.bind_all("<Control-e>", lambda e: self.export_marked_words())

    def _on_key(self, event) -> None:
        if event.widget.winfo_toplevel() is not self or (event.state & 0x4):  # andere Fenster / Strg
            return
        if self.face.animating:
            return "break"
        key = event.keysym.lower()
        actions = {
            "space": self.flip_card,
            "right": self.next_card,
            "return": self.next_card,
            "left": self.prev_card,
            "1": lambda: self.mark_current("unknown"),
            "kp_1": lambda: self.mark_current("unknown"),
            "2": lambda: self.mark_current("known"),
            "kp_2": lambda: self.mark_current("known"),
            "s": self.speak_current,
            "f": lambda: self.toggle_flag("favorite"),
            "d": lambda: self.toggle_flag("difficult"),
            "m": self.shuffle_deck,
        }
        action = actions.get(key)
        if action:
            action()
            return "break"

    # ----- Design -------------------------------------------------------

    def apply_theme(self) -> None:
        c = self.colors
        self.configure(bg=c["bg"])
        for widget, tokens in self._themed:
            try:
                widget.configure(**{option: c[token] for option, token in tokens.items()})
            except tk.TclError:
                pass
        for panel in self._panels:
            panel.restyle(c)
        for button in self._buttons:
            button.restyle(c)
        self.theme_button.configure(text="☾  Dunkel" if self.theme_name == "light" else "☀  Hell")
        self.draw_logo()

        # Eingabefelder ohne Rahmen, nur mit Flaechenfarbe; beim Fokus ein blauer Rand.
        self.style.configure("TCombobox", fieldbackground=c["field"], background=c["field"], foreground=c["text"],
                             arrowcolor=c["muted"], bordercolor=c["field"], lightcolor=c["field"], darkcolor=c["field"],
                             selectbackground=c["field"], selectforeground=c["text"], padding=(self.px(8), self.px(6)),
                             arrowsize=self.px(13), relief="flat")
        self.style.map("TCombobox",
                       fieldbackground=[("readonly", c["field"])],
                       background=[("readonly", c["field"]), ("active", c["field"])],
                       foreground=[("readonly", c["text"])],
                       selectbackground=[("readonly", c["field"])],
                       selectforeground=[("readonly", c["text"])],
                       bordercolor=[("focus", c["accent"]), ("hover", c["line"])],
                       lightcolor=[("focus", c["field"])], darkcolor=[("focus", c["field"])],
                       arrowcolor=[("hover", c["accent"])])
        self.style.configure("Slim.Vertical.TScrollbar", background=c["line"], troughcolor=c["soft"], bordercolor=c["soft"],
                             lightcolor=c["line"], darkcolor=c["line"], arrowsize=self.px(7), gripcount=0, relief="flat")
        self.style.map("Slim.Vertical.TScrollbar", background=[("active", c["muted"])],
                       lightcolor=[("active", c["muted"])], darkcolor=[("active", c["muted"])])
        # Aufklappliste der Comboboxen (ist ein normales Tk-Listbox-Widget).
        self.option_add("*TCombobox*Listbox.background", c["panel"])
        self.option_add("*TCombobox*Listbox.foreground", c["text"])
        self.option_add("*TCombobox*Listbox.selectBackground", c["accent"])
        self.option_add("*TCombobox*Listbox.selectForeground", "#FFFFFF")
        for combo in self.combos:
            try:
                popdown = self.tk.call("ttk::combobox::PopdownWindow", combo)
                self.tk.call(f"{popdown}.f.l", "configure", "-background", c["panel"], "-foreground", c["text"],
                             "-selectbackground", c["accent"], "-selectforeground", "#FFFFFF",
                             "-font", self.font(10), "-borderwidth", 0)
            except tk.TclError:
                pass

        self.details_text.configure(bg=c["soft"], fg=c["text"], insertbackground=c["text"],
                                    selectbackground=c["accent"], selectforeground="#FFFFFF")
        self.details_text.tag_configure("label", foreground=c["muted"])
        self.details_text.tag_configure("highlight", foreground=c["accent"])
        self.details_text.tag_configure("info", foreground=c["muted"])
        self.details_text.tag_configure("error", foreground=c["muted"])
        for token in ("accent", "success", "danger", "purple", "star", "muted"):
            self.details_text.tag_configure(f"dot_{token}", foreground=c[token])
        self.update_flag_buttons()
        self.draw_illustration(self.current)
        self.draw_progress()
        self.draw_box(self.current)

    def draw_logo(self) -> None:
        canvas = self.logo
        canvas.delete("all")
        size = int(float(canvas.cget("width")))
        self.shapes.round_rect(canvas, 0, 0, size, size, size * 0.28, self.colors["accent"])
        canvas.create_text(size / 2, size / 2, text="W", fill="#FFFFFF",
                           font=self.font(14 if size < self.px(40) else 18, "bold", display=True))

    def set_theme(self, theme: str) -> None:
        self.theme_name = theme
        self.colors = THEMES[theme]
        self.store.set_setting("theme", theme)
        self.apply_theme()

    def toggle_theme(self) -> None:
        self.set_theme("dark" if self.theme_name == "light" else "light")

    # ----- Daten --------------------------------------------------------

    def ensure_data_loaded(self) -> None:
        if not self.store.has_words():
            path = ""
            bundled = bundled_excel_path()
            if bundled and messagebox.askyesno(
                "Wortliste importieren",
                f"Die mitgelieferte Wortliste wurde gefunden:\n{bundled}\n\nJetzt importieren?",
            ):
                path = str(bundled)
            if not path:
                messagebox.showinfo("Excel-Datei laden", "Bitte wähle deine Wortschatz-Excel-Datei aus.")
                path = filedialog.askopenfilename(
                    title="Wortschatz-Excel auswählen",
                    initialdir=str(bundled.parent) if bundled else str(Path.home()),
                    filetypes=[("Excel-Dateien", "*.xlsx *.xlsm"), ("Alle Dateien", "*.*")],
                )
            if not path:
                self.set_status("Keine Daten geladen. Bitte über Daten > Excel aktualisieren eine Datei auswählen.")
                self.refresh_filters()
                self.load_deck()
                return
            if not self._import(path):
                self.refresh_filters()
                self.load_deck()
                return
        self.refresh_filters()
        self.load_deck()

    def _import(self, path: str) -> bool:
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            count, restored = self.store.import_excel(path)
        except Exception as exc:
            messagebox.showerror("Importfehler", f"Die Datei konnte nicht importiert werden:\n\n{exc}")
            return False
        finally:
            self.configure(cursor="")
        message = f"{count} Wortschatzkarten wurden importiert."
        if restored:
            message += f"\nDer Lernfortschritt von {restored} Karten wurde übernommen."
        messagebox.showinfo("Import abgeschlossen", message)
        return True

    def update_excel_dialog(self) -> None:
        default_path = self.store.get_setting("source_path", "")
        initialdir = str(Path(default_path).parent) if default_path else str(Path.home())
        path = filedialog.askopenfilename(
            title="Wortschatz-Excel aktualisieren",
            initialdir=initialdir,
            filetypes=[("Excel-Dateien", "*.xlsx *.xlsm"), ("Alle Dateien", "*.*")],
        )
        if not path or not self._import(path):
            return
        self.refresh_filters()
        self.load_deck()

    def export_marked_words(self) -> None:
        favorites = self.store.flagged_words("favorite")
        difficult = self.store.flagged_words("difficult")
        if not favorites and not difficult:
            messagebox.showinfo("Export", "Es sind noch keine Favoriten oder schwierigen Wörter markiert.\n\n"
                                          "Markiere Karten mit ☆ (F) oder ⚑ (D).")
            return
        path = filedialog.asksaveasfilename(
            title="Wörter exportieren",
            defaultextension=".xlsx",
            initialfile=f"Wortschatz_Export_{date.today().isoformat()}.xlsx",
            filetypes=[("Excel-Dateien", "*.xlsx")],
        )
        if not path:
            return
        db_fields = ["unit", "abschnitt", "german_word", "erklaerung", "synonym", "antonym", "english", "turkish", "beispiel"]
        wb = Workbook()
        ws = wb.active
        for index, (title, rows) in enumerate([("Favoriten", favorites), ("Schwierige Wörter", difficult)]):
            if index:
                ws = wb.create_sheet()
            ws.title = title
            ws.append(STANDARD_COLUMNS)
            for cell in ws[1]:
                cell.font = Font(bold=True)
            for row in rows:
                ws.append([row[f] or "" for f in db_fields])
            for col, width in zip("ABCDEFGHI", [14, 26, 28, 40, 26, 22, 24, 24, 50]):
                ws.column_dimensions[col].width = width
            ws.freeze_panes = "A2"
        try:
            wb.save(path)
        except PermissionError:
            messagebox.showerror("Export", "Die Datei konnte nicht gespeichert werden.\nIst sie vielleicht noch in Excel geöffnet?")
            return
        messagebox.showinfo("Export", f"{len(favorites)} Favoriten und {len(difficult)} schwierige Wörter exportiert:\n{path}")

    def reset_progress(self) -> None:
        if not messagebox.askyesno("Lernfortschritt zurücksetzen",
                                   "Alle Ergebnisse, Wiederholungstermine, Favoriten und Markierungen löschen?\n\n"
                                   "Dies kann nicht rückgängig gemacht werden.", icon="warning"):
            return
        self.store.reset_progress()
        self.session = {"known": 0, "unknown": 0}
        self.load_deck()
        self.set_status("Lernfortschritt wurde zurückgesetzt.")

    # ----- Filter & Stapel ----------------------------------------------

    def refresh_filters(self) -> None:
        units = ["Alle"] + self.store.units()
        self.unit_combo.configure(values=units)
        if self.unit_var.get() not in units:
            self.unit_var.set("Alle")
        self.refresh_abschnitte()

    def refresh_abschnitte(self) -> None:
        abschnitte = ["Alle"] + self.store.abschnitte(self.unit_var.get())
        self.abschnitt_combo.configure(values=abschnitte)
        if self.abschnitt_var.get() not in abschnitte:
            self.abschnitt_var.set("Alle")

    def on_unit_changed(self, event=None) -> None:
        self.refresh_abschnitte()
        self.on_filter_changed()

    def on_filter_changed(self, event=None) -> None:
        self.store.set_setting("filter_unit", self.unit_var.get())
        self.store.set_setting("filter_abschnitt", self.abschnitt_var.get())
        self.store.set_setting("lernmodus", self.mode_var.get())
        self.focus_set()  # Tastenkuerzel wieder aktivieren
        self.load_deck()

    def load_deck(self) -> None:
        mode = LERNMODI.get(self.mode_var.get(), "all")
        cards = self.store.query_cards(self.unit_var.get(), self.abschnitt_var.get(), mode)
        if mode == "due":
            # Faellige Wiederholungen zuerst, danach neue Karten.
            review = [c for c in cards if c.practiced]
            new = [c for c in cards if not c.practiced]
            random.shuffle(review)
            random.shuffle(new)
            cards = review + new
        else:
            random.shuffle(cards)
        self.cards = cards
        self.deck_index = 0
        self.round_finished = False
        self.is_back = False
        self.current = self.cards[0] if self.cards else None
        self.render_card()
        self.update_side_stats()

    def shuffle_deck(self) -> None:
        if not self.cards:
            return
        random.shuffle(self.cards)
        self.deck_index = 0
        self.round_finished = False
        self.is_back = False
        self.current = self.cards[0]
        self.render_card()
        self.set_status("Stapel wurde gemischt.")

    def next_card(self) -> None:
        if self.face.animating:
            return
        if self.round_finished or not self.cards:
            self.load_deck()
            return
        if self.deck_index + 1 >= len(self.cards):
            self.finish_round()
            return
        self.deck_index += 1
        self.current = self.cards[self.deck_index]
        self.is_back = False
        self.render_card()

    def prev_card(self) -> None:
        if not self.cards or self.face.animating:
            return
        if self.round_finished:
            self.round_finished = False
        elif self.deck_index > 0:
            self.deck_index -= 1
        else:
            return
        self.current = self.cards[self.deck_index]
        self.is_back = False
        self.render_card()

    def finish_round(self) -> None:
        self.round_finished = True
        self.current = None
        self.is_back = False
        self.render_card()

    # ----- Darstellung --------------------------------------------------

    def _on_details_scroll(self, first: str, last: str) -> None:
        # Scrollleiste nur zeigen, wenn der Inhalt nicht ganz hineinpasst.
        self.details_scroll.set(first, last)
        if float(first) <= 0.0 and float(last) >= 1.0:
            self.details_scroll.pack_forget()
        elif not self.details_scroll.winfo_ismapped():
            self.details_scroll.pack(side="right", fill="y", before=self.details_text)

    def _on_body_resize(self, event) -> None:
        self.word_label.configure(wraplength=max(160, event.width - self.px(40)))
        if not self.is_back and self._face_has_room_for_illustration() != self._illustration_visible:
            self.after_idle(self.render_card)

    def _layout_body(self, widgets: List[Tuple[tk.Widget, dict]]) -> None:
        for widget in (self.spacer_top, self.illustration, self.word_label, self.hint_label, self.details_wrap, self.spacer_bottom):
            widget.pack_forget()
        for widget, options in widgets:
            widget.pack(**options)

    def render_card(self) -> None:
        px = self.px
        card = self.current
        self.draw_progress()
        self.draw_box(card)
        self.update_flag_buttons()
        has_card = card is not None
        for button in (self.known_button, self.unknown_button, self.flip_button, self.speak_button,
                       self.fav_button, self.difficult_button):
            button.set_enabled(has_card)
        self.prev_button.set_enabled(bool(self.cards) and (self.deck_index > 0 or self.round_finished))
        self._update_action_labels()

        if not has_card:
            self.card_meta.configure(text="")
            self.counter_label.configure(text="")
            if self.round_finished:
                self.word_label.configure(text="Runde geschafft!", font=self.word_font_back)
                known, unknown = self.session["known"], self.session["unknown"]
                total = known + unknown
                rate = f"{round(known / total * 100)} %" if total else "–"
                n = len(self.cards)
                self.set_details([("info", f"Du hast den Stapel mit {n} {'Karte' if n == 1 else 'Karten'} durchgearbeitet.\n\n"
                                           f"Diese Sitzung: {known} gewusst · {unknown} nicht gewusst · Quote {rate}\n\n"
                                           "Mit „Neue Runde“ (Enter) geht es weiter.")])
            elif not self.store.has_words():
                self.word_label.configure(text="Noch keine Wörter", font=self.word_font_back)
                self.set_details([("info", "Importiere eine Excel-Datei über Daten › Excel aktualisieren (Strg+O).")])
            else:
                mode = LERNMODI.get(self.mode_var.get(), "all")
                self.word_label.configure(text="Alles erledigt" if mode == "due" else "Keine Karten gefunden",
                                          font=self.word_font_back)
                hint = ("Für diese Auswahl ist heute keine Wiederholung fällig. Wähle „Alle Karten“, um trotzdem zu üben."
                        if mode == "due" else "Bitte ändere Einheit, Abschnitt oder Lernmodus.")
                self.set_details([("info", hint)])
            self.draw_illustration(None)
            self._illustration_visible = self._face_has_room_for_illustration()
            self._layout_body([
                (self.spacer_top, dict(fill="both", expand=True)),
                *([(self.illustration, dict(pady=(0, px(10))))] if self._illustration_visible else []),
                (self.word_label, dict(pady=(0, px(10)))),
                (self.details_wrap, dict(fill="x")),
                (self.spacer_bottom, dict(fill="both", expand=True)),
            ])
            return

        self.card_meta.configure(text=f"{card.unit}  ·  {card.abschnitt}")
        self.counter_label.configure(text=f"{self.deck_index + 1} / {len(self.cards)}")
        self.word_label.configure(text=card.german_word)
        if self.is_back:
            self.word_label.configure(font=self.word_font_back)
            self.flip_button.configure(text="Vorderseite")
            self.set_details(self.back_rows(card))
            self._layout_body([
                (self.word_label, dict(pady=(px(4), px(12)))),
                (self.details_wrap, dict(fill="both", expand=True)),
            ])
        else:
            self.word_label.configure(font=self.word_font_back if self._compact_height else self.word_font_front)
            self.flip_button.configure(text="Umdrehen")
            self.draw_illustration(card)
            self._illustration_visible = self._face_has_room_for_illustration()
            body_h = self.card_body.winfo_height()
            show_hint = body_h <= 1 or body_h >= (self.word_label.winfo_reqheight() + self.hint_label.winfo_reqheight()
                                                  + self.px(14))
            self._layout_body([
                (self.spacer_top, dict(fill="both", expand=True)),
                *([(self.illustration, dict(pady=(0, px(16))))] if self._illustration_visible else []),
                (self.word_label, dict(pady=(0, px(10)))),
                *([(self.hint_label, dict())] if show_hint else []),
                (self.spacer_bottom, dict(fill="both", expand=True)),
            ])

    def set_details(self, rows: List[Tuple[str, str]]) -> None:
        """rows: Liste von (Tag, Text). "label:<farbe>" erzeugt eine Ueberschrift mit farbigem Punkt."""
        self.details_text.configure(state="normal")
        self.details_text.delete("1.0", "end")
        for index, (tag, text) in enumerate(rows):
            first = ("first",) if index == 0 else ()
            if tag.startswith("label:"):
                token = tag.split(":", 1)[1]
                self.details_text.insert("end", "●  ", ("label", f"dot_{token}") + first)
                self.details_text.insert("end", text + "\n", ("label",) + first)
            else:
                self.details_text.insert("end", text + "\n", (tag,) + first)
        self.details_text.configure(state="disabled")
        self.details_text.yview_moveto(0)

    def back_rows(self, card: WordCard, word_translation: str = "", example_translation: str = "",
                  loading: bool = False, word_error: bool = False, example_error: bool = False) -> List[Tuple[str, str]]:
        lang_name = self.muttersprache_var.get()
        lang_code = MUTTERSPRACHEN.get(lang_name, "tr")
        word_translation_ready = bool(word_translation)
        example_translation_ready = bool(example_translation)
        if loading:
            word_translation = word_translation or "Übersetzung wird geladen …"
            if card.beispiel:
                example_translation = example_translation or "Übersetzung wird geladen …"
        rows: List[Tuple[str, str]] = []

        def add(label: str, value: str, color: str, tag: str = "value", muted: bool = False) -> None:
            if value and value.strip() not in {"-", "–", "—", "/"}:  # Platzhalter aus der Tabelle ausblenden
                rows.append((f"label:{color}", label.upper()))
                rows.append(("error" if muted else tag, value))

        add("Erklärung", card.erklaerung, "accent")
        add("Synonym", card.synonym, "success")
        add("Antonym", card.antonym, "danger")
        add("English", card.english, "purple", "highlight" if lang_code == "en" else "value")
        add("Türkisch", card.turkish, "purple", "highlight" if lang_code == "tr" else "value")
        if lang_code not in EXCEL_LANGUAGE_FIELDS and lang_code != "de":
            add(f"Muttersprache ({lang_name})", word_translation, "purple", "highlight",
                word_error or (loading and not word_translation_ready))
        add("Beispiel", card.beispiel, "star", "example")
        if lang_code != "de":
            add(f"Beispiel auf {lang_name}", example_translation, "star", "example",
                example_error or (loading and not example_translation_ready))
        progress = f"Gewusst {card.known_count}× · Nicht gewusst {card.unknown_count}× · Box {card.box}/{MAX_BOX}"
        add("Fortschritt", progress, "muted")
        return rows

    def draw_illustration(self, card: Optional[WordCard]) -> None:
        canvas = self.illustration
        canvas.delete("all")
        size = self.px(132)
        if card:
            fill = color_for_text(card.unit + card.abschnitt + card.german_word)
            label = card_icon_text(card)
        else:
            fill = self.colors["success"] if self.round_finished else self.colors["accent"]
            label = "✓" if self.round_finished else "★"
        ring = mix(fill, self.colors["soft"], 0.72)
        self.shapes.disc(canvas, size / 2, size / 2, size - 2 * self.px(6), ring)
        self.shapes.disc(canvas, size / 2, size / 2, size - 2 * self.px(20), fill)
        canvas.create_text(size / 2, size / 2, text=label, fill="#FFFFFF", font=self.font(28, "bold"))

    def draw_progress(self) -> None:
        canvas = self.progress_canvas
        canvas.delete("all")
        width = canvas.winfo_width()
        height = self.px(6)
        if width <= 1:
            return
        self.shapes.round_rect(canvas, 0, 0, width, height, height // 2, self.colors["track"])
        if self.cards:
            done = len(self.cards) if self.round_finished else self.deck_index + 1
            x = max(height, width * done / len(self.cards))
            self.shapes.round_rect(canvas, 0, 0, x, height, height // 2, self.colors["accent"])

    def draw_box(self, card: Optional[WordCard]) -> None:
        canvas = self.box_canvas
        canvas.delete("all")
        d = self.px(14)
        gap = self.px(8)
        box = card.box if card else 0
        for i in range(1, MAX_BOX + 1):
            fill = self.colors["accent"] if i <= box else self.colors["track"]
            self.shapes.disc(canvas, (i - 1) * (d + gap) + d / 2 + 1, d / 2 + 2, d, fill)
        if not card:
            self.box_label.configure(text="")
        elif not card.practiced:
            self.box_label.configure(text="Neue Karte – noch nicht geübt.")
        else:
            due = card.due_date
            if not due or due <= date.today().isoformat():
                when = "heute"
            else:
                days = (date.fromisoformat(due) - date.today()).days
                when = "morgen" if days == 1 else f"in {days} Tagen ({date.fromisoformat(due).strftime('%d.%m.')})"
            self.box_label.configure(text=f"Leitner-Box {card.box} von {MAX_BOX}\nNächste Wiederholung: {when}")

    def update_flag_buttons(self) -> None:
        card = self.current
        fav = bool(card and card.favorite)
        diff = bool(card and card.difficult)
        self.fav_button.configure(text="★" if fav else "☆")
        self.fav_button.set_fg("star" if fav else None)
        self.difficult_button.set_fg("flag" if diff else None)

    def update_side_stats(self) -> None:
        summary = self.store.summary()
        self.stat_deck.configure(text=str(len(self.cards)))
        self.stat_known.configure(text=str(self.session["known"]))
        self.stat_unknown.configure(text=str(self.session["unknown"]))
        self.stat_due.configure(text=str(summary["due"]))
        self.stat_new.configure(text=str(summary["new"]))
        self.stat_mastered.configure(text=str(summary["mastered"]))
        self.stat_fav.configure(text=str(summary["favorites"]))
        self.stat_difficult.configure(text=str(summary["difficult"]))

    def set_status(self, text: str, duration_ms: int = 2600) -> None:
        """Kurze Meldung als Einblendung ueber den Aktionsknoepfen."""
        if self._toast_job:
            self.after_cancel(self._toast_job)
        self.toast.configure(text=text)
        self.toast.place(relx=0.5, rely=1.0, y=-self.px(86), anchor="s")
        self.toast.lift()
        self._toast_job = self.after(duration_ms, self._hide_toast)

    def _hide_toast(self) -> None:
        self._toast_job = None
        self.toast.place_forget()

    # ----- Aktionen -----------------------------------------------------

    def flip_card(self) -> None:
        if not self.current or self.face.animating:
            return

        def turn() -> None:
            self.is_back = not self.is_back
            self.render_card()
            if self.is_back and self.current:
                self.load_muttersprache_translations(self.current)

        self.face.flip(turn)

    def mark_current(self, result: str) -> None:
        card = self.current
        if not card or self.face.animating:
            return
        card.box, card.due_date = self.store.mark(card.id, result)
        card.last_result = result
        if result == "known":
            card.known_count += 1
            self.session["known"] += 1
        else:
            card.unknown_count += 1
            self.session["unknown"] += 1
            # Nicht gewusste Karten kommen am Ende der Runde noch einmal.
            if card not in self.cards[self.deck_index + 1:]:
                self.cards.append(card)
        self.update_side_stats()
        self.next_card()

    def toggle_flag(self, flag: str) -> None:
        card = self.current
        if not card:
            return
        value = not bool(getattr(card, flag))
        self.store.set_flag(card.id, flag, value)
        setattr(card, flag, int(value))
        self.update_flag_buttons()
        self.update_side_stats()
        names = {"favorite": "Favoriten", "difficult": "schwierigen Wörtern"}
        self.set_status(f"„{card.german_word}“ {'zu' if value else 'aus'} {names[flag]} "
                        f"{'hinzugefügt' if value else 'entfernt'}.")

    def speak_current(self) -> None:
        if not self.current or self._speaking:
            return
        self._speaking = True
        self.speak_button.set_enabled(False)

        def done(error: object) -> None:
            self._speaking = False
            self.speak_button.set_enabled(self.current is not None)
            if error:
                messagebox.showwarning("Aussprache", str(error))

        text = self.current.german_word
        self.run_in_background(lambda: self.speaker.speak(text), done)

    def on_muttersprache_changed(self, event=None) -> None:
        self.store.set_setting("muttersprache", self.muttersprache_var.get())
        self.focus_set()
        if self.is_back and self.current:
            self.render_card()
            self.load_muttersprache_translations(self.current)

    @staticmethod
    def translate_text_online(source_text: str, target_lang: str, allow_fallback: bool = False) -> Tuple[str, bool]:
        """Gibt (Text, erfolgreich) zurueck. Laeuft im Hintergrund-Thread - keine Tk- oder DB-Zugriffe!

        allow_fallback: bei Saetzen MyMemory als Ersatz nutzen. Fuer Einzelwoerter liefert
        MyMemory oft unpassende Treffer aus seinem Uebersetzungsspeicher - dort lieber nichts anzeigen.
        """
        source_text = (source_text or "").strip()
        if not source_text:
            return "", True
        if GoogleTranslator is None:
            return "Nicht verfügbar – bitte „pip install deep-translator“ ausführen.", False
        error: Optional[Exception] = None
        try:
            result = GoogleTranslator(source="de", target=target_lang).translate(source_text)
            if result:
                return result, True
        except Exception as exc:
            error = exc
        if allow_fallback and MyMemoryTranslator is not None and target_lang in MYMEMORY_CODES:
            try:
                result = MyMemoryTranslator(source="de-DE", target=MYMEMORY_CODES[target_lang]).translate(source_text)
                if result:
                    return result, True
            except Exception:
                pass
        name = type(error).__name__ if error else ""
        if "TooManyRequests" in name:
            return "Vorübergehend nicht verfügbar – Anfragelimit des Übersetzungsdienstes erreicht. Bitte später erneut versuchen.", False
        if "Connection" in name or "Timeout" in name:
            return "Nicht verfügbar – keine Internetverbindung?", False
        return "Übersetzung fehlgeschlagen. Bitte später erneut versuchen.", False

    def load_muttersprache_translations(self, card: WordCard) -> None:
        self.translation_job_id += 1
        job_id = self.translation_job_id
        lang_code = MUTTERSPRACHEN.get(self.muttersprache_var.get(), "tr")
        if lang_code == "de":
            return

        # Steht die Muttersprache schon in der Excel-Tabelle, wird nur das Beispiel uebersetzt.
        word_text = "" if lang_code in EXCEL_LANGUAGE_FIELDS else clean_german_word(card.german_word)
        example_text = card.beispiel.strip()
        cached_word = self.store.get_cached_translation(word_text, "de", lang_code) if word_text else ""
        cached_example = self.store.get_cached_translation(example_text, "de", lang_code) if example_text else ""

        if (cached_word or not word_text) and (cached_example or not example_text):
            self.set_details(self.back_rows(card, cached_word, cached_example))
            return

        self.set_details(self.back_rows(card, cached_word, cached_example, loading=True))

        def work():
            word = (cached_word, True) if cached_word or not word_text else self.translate_text_online(word_text, lang_code)
            example = ((cached_example, True) if cached_example or not example_text
                       else self.translate_text_online(example_text, lang_code, allow_fallback=True))
            return word, example

        def done(result) -> None:
            (word, word_ok), (example, example_ok) = result
            if word_ok and word and word_text:
                self.store.save_cached_translation(word_text, "de", lang_code, word)
            if example_ok and example and example_text:
                self.store.save_cached_translation(example_text, "de", lang_code, example)
            if job_id != self.translation_job_id or self.current is not card or not self.is_back:
                return
            self.set_details(self.back_rows(card, word, example, word_error=not word_ok, example_error=not example_ok))

        self.run_in_background(work, done)

    # ----- Fenster ------------------------------------------------------

    def show_stats(self) -> None:
        data = self.store.stats()
        summary = self.store.summary()
        c = self.colors
        px = self.px
        win = tk.Toplevel(self)
        win.title("Statistik")
        width, height = self.fit_size(px(820), px(680))
        win.geometry(f"{width}x{height}")
        win.minsize(*self.fit_size(px(560), px(420)))
        win.configure(bg=c["bg"])
        win.transient(self)

        tk.Label(win, text="Lernstatistik", font=self.font(22, "bold", display=True), bg=c["bg"], fg=c["text"]).pack(
            anchor="w", padx=px(24), pady=(px(20), px(12)))

        overall = data["overall"]
        known, unknown = overall["known"], overall["unknown"]
        attempts = known + unknown
        rate = f"{round(known / attempts * 100)} %" if attempts else "–"
        tiles = tk.Frame(win, bg=c["bg"])
        tiles.pack(fill="x", padx=px(24))
        for i, (caption, value) in enumerate([
            ("Karten", overall["total"]), ("Geübt", overall["practiced"]), ("Heute fällig", summary["due"]),
            ("Sicher gelernt", summary["mastered"]), ("Trefferquote", rate),
        ]):
            panel = RoundedPanel(tiles, self.shapes, px(14), padding=px(12), fit_height=True)
            panel.restyle(c)
            panel.inner.configure(bg=c["panel"])
            panel.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else px(10), 0))
            tiles.grid_columnconfigure(i, weight=1, uniform="tile")
            tk.Label(panel.inner, text=caption.upper(), font=self.font(8, "bold"), bg=c["panel"], fg=c["muted"]).pack(
                anchor="w", padx=px(4))
            tk.Label(panel.inner, text=str(value), font=self.font(20, "bold", display=True), bg=c["panel"], fg=c["text"]).pack(
                anchor="w", padx=px(4))

        # Verteilung auf die Leitner-Boxen
        boxes = data["boxes"]
        total = max(1, overall["total"])
        tk.Label(win, text="LEITNER-BOXEN", font=self.font(8, "bold"), bg=c["bg"], fg=c["muted"]).pack(
            anchor="w", padx=px(24), pady=(px(18), px(6)))
        bar = tk.Canvas(win, height=px(14), bg=c["bg"], highlightthickness=0)
        bar.pack(fill="x", padx=px(24))
        box_colors = [c["track"], "#FF6B6B", "#FFA94D", "#FFD43B", "#69DB7C", "#2FB350"]

        def draw_bar(event=None) -> None:
            bar.delete("all")
            width = bar.winfo_width()
            x = 0.0
            for b in range(MAX_BOX + 1):
                count = boxes.get(b, 0)
                if not count:
                    continue
                w = max(2.0, width * count / total)
                bar.create_rectangle(x, 0, x + w, px(14), fill=box_colors[b], outline="")
                x += w
        bar.bind("<Configure>", draw_bar)
        legend = tk.Frame(win, bg=c["bg"])
        legend.pack(anchor="w", padx=px(24), pady=(px(6), 0))
        for b in range(MAX_BOX + 1):
            tk.Label(legend, text="●", font=self.font(9), bg=c["bg"], fg=box_colors[b] if b else c["muted"]).pack(side="left")
            tk.Label(legend, text=f"{'Neu' if b == 0 else f'Box {b}'} {boxes.get(b, 0)}", font=self.font(9),
                     bg=c["bg"], fg=c["muted"]).pack(side="left", padx=(px(2), px(14)))

        tree_style = "Stats.Treeview"
        self.style.configure(tree_style, background=c["panel"], fieldbackground=c["panel"], foreground=c["text"],
                             rowheight=px(28), bordercolor=c["line"], borderwidth=0, font=self.font(10))
        self.style.configure(f"{tree_style}.Heading", background=c["panel"], foreground=c["muted"], relief="flat",
                             font=self.font(9, "bold"), bordercolor=c["panel"], lightcolor=c["panel"], darkcolor=c["panel"])
        self.style.map(tree_style, background=[("selected", c["accent"])], foreground=[("selected", "#FFFFFF")])
        self.style.layout(tree_style, [("Treeview.treearea", {"sticky": "nswe"})])
        self.style.map(f"{tree_style}.Heading", background=[("active", c["button_hover"])])

        cols = ("Einheit", "Karten", "Geübt", "Gewusst", "Nicht gewusst", "Box 5", "Quote")
        table = RoundedPanel(win, self.shapes, px(14), padding=px(10))
        table.restyle(c)
        table.inner.configure(bg=c["panel"])
        table.pack(fill="both", expand=True, padx=px(24), pady=(px(16), px(24)))
        tree = ttk.Treeview(table.inner, columns=cols, show="headings", style=tree_style)
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=px(90), anchor="center")
        tree.column("Einheit", width=px(170), anchor="w")
        tree.tag_configure("odd", background=c["soft"])
        for i, row in enumerate(data["by_unit"]):
            k = row["known"] or 0
            u = row["unknown"] or 0
            q = f"{round(k / (k + u) * 100)} %" if k + u else "–"
            tree.insert("", "end", values=(row["unit"], row["total"], row["practiced"], k, u, row["mastered"], q),
                        tags=("odd",) if i % 2 else ())
        tree.pack(fill="both", expand=True)

    def show_shortcuts(self) -> None:
        messagebox.showinfo(
            "Tastenkürzel",
            "Leertaste\tKarte umdrehen\n"
            "→ / Enter\tNächste Karte\n"
            "←\t\tVorherige Karte\n"
            "1\t\tNicht gewusst\n"
            "2\t\tGewusst\n"
            "S\t\tAussprache\n"
            "F\t\tFavorit an/aus\n"
            "D\t\tSchwieriges Wort an/aus\n"
            "M\t\tMischen\n"
            "Strg+O\t\tExcel aktualisieren\n"
            "Strg+E\t\tExport",
        )

    def show_about(self) -> None:
        last_import = self.store.get_setting("last_import", "Noch nicht importiert")
        messagebox.showinfo(
            "Über Wortschatz Karten",
            f"{APP_NAME} {APP_VERSION}\n\n"
            "Lernkarten-App für DSH-Wortschatz mit Leitner-Wiederholungssystem.\n"
            f"Die Daten werden lokal gespeichert:\n{APP_DIR}\n\n"
            f"Letzter Import: {last_import}",
        )

    def on_close(self) -> None:
        try:
            state = self.state()
            self.store.set_setting("zoomed", "1" if state == "zoomed" else "0")
            if state == "normal":
                self.store.set_setting("geometry", self.geometry())
            self.store.close()
        finally:
            self.destroy()


def enable_dpi_awareness() -> None:
    """Scharfe Darstellung auf Bildschirmen mit Skalierung > 100 % (sonst wird Windows die App unscharf hochskalieren)."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def main() -> None:
    enable_dpi_awareness()
    app = WortschatzApp()
    app.mainloop()


if __name__ == "__main__":
    main()
