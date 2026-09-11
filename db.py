"""
db.py – SQLite-Datenschicht für das Werkstatt-To-Do-Programm.
Speichert alles in einer lokalen Datei todo.db im selben Ordner.
"""
import sqlite3
import os
import shutil
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "todo.db")
ANHANG_DIR = os.path.join(BASE_DIR, "anhaenge")

STATUS_OPTIONS = ["offen", "erledigt", "teilerledigt", "warte_auf_teile", "warte_auf_kunde"]
STATUS_LABELS = {
    "offen": "Offen",
    "erledigt": "Erledigt",
    "teilerledigt": "Teilerledigt",
    "warte_auf_teile": "Warte auf Teile",
    "warte_auf_kunde": "Warte auf Kunde",
}
STATUS_COLORS = {
    "offen": "#e0e0e0",
    "erledigt": "#8bc34a",
    "teilerledigt": "#ffca28",
    "warte_auf_teile": "#ff8a65",
    "warte_auf_kunde": "#4fc3f7",
}


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            typ TEXT NOT NULL CHECK(typ IN ('aufgabe', 'rueckruf')),
            titel TEXT NOT NULL,
            notiz TEXT DEFAULT '',
            status TEXT NOT NULL DEFAULT 'offen',
            wiedervorlage TEXT,
            erstellt_am TEXT NOT NULL,
            erledigt_am TEXT,
            absender TEXT,
            betreff TEXT,
            email_datum TEXT,
            email_text TEXT,
            email_dateiname TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS anhaenge (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id INTEGER NOT NULL,
            dateiname TEXT NOT NULL,
            pfad TEXT NOT NULL,
            hinzugefuegt_am TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS notiz_eintraege (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id INTEGER NOT NULL,
            text TEXT NOT NULL,
            erstellt_am TEXT NOT NULL
        )
        """
    )
    conn.commit()
    conn.close()
    os.makedirs(ANHANG_DIR, exist_ok=True)


def add_task(typ, titel, notiz="", absender=None, betreff=None,
             email_datum=None, email_text=None, email_dateiname=None):
    conn = get_connection()
    conn.execute(
        """
        INSERT INTO tasks (typ, titel, notiz, status, erstellt_am,
                            absender, betreff, email_datum, email_text, email_dateiname)
        VALUES (?, ?, ?, 'offen', ?, ?, ?, ?, ?, ?)
        """,
        (typ, titel, notiz, datetime.now().isoformat(timespec="seconds"),
         absender, betreff, email_datum, email_text, email_dateiname),
    )
    conn.commit()
    conn.close()


def get_tasks(typ):
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM tasks WHERE typ = ? ORDER BY id DESC", (typ,)
    ).fetchall()
    conn.close()
    return rows


def update_status(task_id, status):
    conn = get_connection()
    erledigt_am = datetime.now().isoformat(timespec="seconds") if status == "erledigt" else None
    conn.execute(
        "UPDATE tasks SET status = ?, erledigt_am = ? WHERE id = ?",
        (status, erledigt_am, task_id),
    )
    conn.commit()
    conn.close()


def update_wiedervorlage(task_id, datum):
    """datum: 'YYYY-MM-DD' oder None zum Löschen."""
    conn = get_connection()
    conn.execute("UPDATE tasks SET wiedervorlage = ? WHERE id = ?", (datum, task_id))
    conn.commit()
    conn.close()


def update_notiz(task_id, notiz):
    conn = get_connection()
    conn.execute("UPDATE tasks SET notiz = ? WHERE id = ?", (notiz, task_id))
    conn.commit()
    conn.close()


def update_titel(task_id, neuer_titel):
    conn = get_connection()
    conn.execute("UPDATE tasks SET titel = ? WHERE id = ?", (neuer_titel, task_id))
    conn.commit()
    conn.close()


def delete_task(task_id):
    # zugehoerige Anhaenge (Dateien + Ordner) mit aufraeumen
    for anhang in get_attachments(task_id):
        try:
            if os.path.isfile(anhang["pfad"]):
                os.remove(anhang["pfad"])
        except OSError:
            pass
    ordner = os.path.join(ANHANG_DIR, str(task_id))
    if os.path.isdir(ordner):
        shutil.rmtree(ordner, ignore_errors=True)

    conn = get_connection()
    conn.execute("DELETE FROM anhaenge WHERE task_id = ?", (task_id,))
    conn.execute("DELETE FROM notiz_eintraege WHERE task_id = ?", (task_id,))
    conn.execute("DELETE FROM tasks WHERE id = ?", (task_id,))
    conn.commit()
    conn.close()


# ---------- Anhaenge (PDF etc.) ----------

def add_attachment(task_id, quellpfad):
    """Kopiert die Datei in einen lokalen Anhaenge-Ordner und legt einen DB-Eintrag an."""
    ordner = os.path.join(ANHANG_DIR, str(task_id))
    os.makedirs(ordner, exist_ok=True)

    dateiname = os.path.basename(quellpfad)
    base, ext = os.path.splitext(dateiname)
    ziel = os.path.join(ordner, dateiname)
    counter = 1
    while os.path.exists(ziel):
        dateiname = f"{base} ({counter}){ext}"
        ziel = os.path.join(ordner, dateiname)
        counter += 1

    shutil.copy2(quellpfad, ziel)

    conn = get_connection()
    conn.execute(
        "INSERT INTO anhaenge (task_id, dateiname, pfad, hinzugefuegt_am) VALUES (?, ?, ?, ?)",
        (task_id, dateiname, ziel, datetime.now().isoformat(timespec="seconds")),
    )
    conn.commit()
    conn.close()


def get_attachments(task_id):
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM anhaenge WHERE task_id = ? ORDER BY id", (task_id,)
    ).fetchall()
    conn.close()
    return rows


def delete_attachment(attachment_id):
    conn = get_connection()
    row = conn.execute("SELECT pfad FROM anhaenge WHERE id = ?", (attachment_id,)).fetchone()
    conn.execute("DELETE FROM anhaenge WHERE id = ?", (attachment_id,))
    conn.commit()
    conn.close()
    if row:
        try:
            if os.path.isfile(row["pfad"]):
                os.remove(row["pfad"])
        except OSError:
            pass


# ---------- Notizverlauf (datierte Eintraege, z.B. Telefonate) ----------

def add_notiz_eintrag(task_id, text):
    conn = get_connection()
    conn.execute(
        "INSERT INTO notiz_eintraege (task_id, text, erstellt_am) VALUES (?, ?, ?)",
        (task_id, text, datetime.now().isoformat(timespec="seconds")),
    )
    conn.commit()
    conn.close()


def get_notiz_eintraege(task_id):
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM notiz_eintraege WHERE task_id = ? ORDER BY id DESC", (task_id,)
    ).fetchall()
    conn.close()
    return rows


def delete_notiz_eintrag(eintrag_id):
    conn = get_connection()
    conn.execute("DELETE FROM notiz_eintraege WHERE id = ?", (eintrag_id,))
    conn.commit()
    conn.close()
