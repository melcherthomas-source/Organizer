"""
Werkstatt To-Do – Aufgaben & Rückrufe
=====================================
Lokales Programm, keine Internetverbindung noetig, keine laufenden Kosten.

Start:  python app.py

Benoetigt:
    pip install tkinterdnd2 extract-msg
"""
import os
import re
import sys
import subprocess
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox, filedialog
from datetime import datetime, timedelta

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

OUTLOOK_COM_AVAILABLE = False
if sys.platform.startswith("win"):
    try:
        import win32com.client  # noqa: F401
        OUTLOOK_COM_AVAILABLE = True
    except ImportError:
        OUTLOOK_COM_AVAILABLE = False

import db
from email_parser import parse_email_file, text_bereinigen


def outlook_auswahl_lesen():
    """
    Liest die aktuell in Outlook markierte(n) E-Mail(s) direkt per COM aus –
    ohne Datei-Umweg. Funktioniert nur unter Windows mit installiertem
    Outlook (klassische Desktop-Version) und dem Paket 'pywin32'.
    Gibt eine Liste von dicts zurueck: {betreff, absender, datum, text}
    """
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    try:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception:
            raise RuntimeError(
                "Outlook konnte nicht angesprochen werden. Bitte Outlook öffnen und erneut versuchen."
            )

        explorer = outlook.ActiveExplorer()
        if explorer is None:
            raise RuntimeError("Kein offenes Outlook-Fenster gefunden.")

        auswahl = explorer.Selection
        if auswahl is None or auswahl.Count == 0:
            raise RuntimeError("Bitte zuerst eine E-Mail in Outlook markieren.")

        ergebnisse = []
        for i in range(1, auswahl.Count + 1):
            item = auswahl.Item(i)
            # Nur echte Mail-Elemente verarbeiten (MailItem = Klasse 43)
            if getattr(item, "Class", None) != 43:
                continue

            betreff = (item.Subject or "").strip() or "(kein Betreff)"
            absender = (getattr(item, "SenderName", "") or "").strip()
            try:
                absender_mail = item.SenderEmailAddress
                if absender_mail:
                    absender = f"{absender} <{absender_mail}>" if absender else absender_mail
            except Exception:
                pass
            try:
                datum = item.ReceivedTime.strftime("%d.%m.%Y %H:%M")
            except Exception:
                datum = ""
            text = text_bereinigen(getattr(item, "Body", "") or "")

            ergebnisse.append({
                "betreff": betreff,
                "absender": absender or "(unbekannter Absender)",
                "datum": datum,
                "text": text,
            })

        if not ergebnisse:
            raise RuntimeError("Die Auswahl in Outlook enthält keine E-Mail.")
        return ergebnisse
    finally:
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# Farbschema & Typografie – Dark Mode. Hier laesst sich das ganze Design
# zentral anpassen.
# ---------------------------------------------------------------------------
COLOR_BG = "#15161C"
COLOR_HEADER_BG = "#1C1E27"
COLOR_CARD_BG = "#22242E"
COLOR_CARD_SHADOW = "#0A0B0F"
COLOR_TEXT_PRIMARY = "#F3F4F6"
COLOR_TEXT_SECONDARY = "#9CA3AF"
COLOR_TEXT_MUTED = "#6B7280"
COLOR_BORDER = "#33353F"

COLOR_SURFACE = "#20222B"
COLOR_SURFACE_HOVER = "#2B2D38"
COLOR_SCROLLBAR_HOVER = "#3A3D4A"

COLOR_ACCENT = "#34D399"
COLOR_ACCENT_HOVER = "#10B981"
COLOR_ACCENT_LIGHT = "#132C24"
COLOR_ACCENT_LIGHT_HOVER = "#17392F"

COLOR_DANGER = "#F87171"
COLOR_DANGER_HOVER = "#EF4444"
COLOR_DANGER_LIGHT = "#3A1F22"

STATUS_COLORS = {
    "offen": "#94A3B8",
    "erledigt": "#22C55E",
    "teilerledigt": "#F59E0B",
    "warte_auf_teile": "#8B5CF6",
    "warte_auf_kunde": "#38BDF8",
}
STATUS_INACTIVE_BG = "#262832"
STATUS_INACTIVE_FG = "#9CA3AF"
STATUS_INACTIVE_HOVER = "#32343F"

# Rückrufe zeigen bewusst nicht "Warte auf Kunde" (das ist ein Aufgaben-spezifischer Status)
RUECKRUF_STATUS_OPTIONS = ["offen", "erledigt", "teilerledigt", "warte_auf_teile"]

FILTER_OPTIONEN = {
    "aufgabe": [("Alle", "alle"), ("Offen", "offen"), ("Teilerledigt", "teilerledigt"),
                ("Warte auf Teile", "warte_auf_teile"), ("Warte auf Kunde", "warte_auf_kunde")],
    "rueckruf": [("Alle", "alle"), ("Offen", "offen"), ("Teilerledigt", "teilerledigt"),
                 ("Warte auf Teile", "warte_auf_teile")],
}

F_TITLE = ("Segoe UI", 19, "bold")
F_COLUMN_TITLE = ("Segoe UI", 13, "bold")
F_CARD_TITLE = ("Segoe UI", 11, "bold")
F_NORMAL = ("Segoe UI", 10)
F_SMALL = ("Segoe UI", 9)
F_PILL = ("Segoe UI", 9, "bold")


def heute_iso():
    return datetime.now().strftime("%Y-%m-%d")


def iso_zu_anzeige(iso_datum):
    if not iso_datum:
        return ""
    try:
        return datetime.strptime(iso_datum, "%Y-%m-%d").strftime("%d.%m.%Y")
    except ValueError:
        return iso_datum


def iso_zu_anzeige_datetime(iso_str):
    try:
        return datetime.fromisoformat(iso_str).strftime("%d.%m.%Y %H:%M")
    except ValueError:
        return iso_str


def datei_oeffnen(pfad):
    """Oeffnet eine Datei mit dem Standardprogramm des Betriebssystems."""
    if sys.platform.startswith("win"):
        os.startfile(pfad)  # noqa: nur unter Windows verfuegbar
    elif sys.platform == "darwin":
        subprocess.run(["open", pfad], check=False)
    else:
        subprocess.run(["xdg-open", pfad], check=False)


_TELEFON_LABEL_RE = re.compile(
    r"(?:tel(?:efon)?|mobil|handy|rückruf(?:nummer)?|erreichbar unter)\s*[:\-]?\s*"
    r"([+0][\d\s\-/()]{6,20}\d)",
    re.IGNORECASE,
)
_TELEFON_GENERISCH_RE = re.compile(
    r"(\+49[\s\-/]?\d[\d\s\-/]{7,17}\d|0\d{2,5}[\s\-/]?\d{3,}[\d\s\-/]{0,10}\d)"
)


def telefonnummer_ermitteln(*texte):
    """Sucht in den uebergebenen Texten nach einer Telefonnummer."""
    for text in texte:
        if not text:
            continue
        m = _TELEFON_LABEL_RE.search(text)
        if m:
            return m.group(1).strip()
    for text in texte:
        if not text:
            continue
        m = _TELEFON_GENERISCH_RE.search(text)
        if m:
            return m.group(1).strip()
    return None


def name_aus_absender(absender, fallback):
    if not absender:
        return fallback
    absender = absender.strip()
    m = re.match(r'^"?([^"<]*?)"?\s*<[^>]+>$', absender)
    if m and m.group(1).strip():
        return m.group(1).strip()
    inner = re.search(r"<([^>]+)>", absender)
    if inner:
        return inner.group(1)
    return absender or fallback


def grund_ermitteln(betreff, text):
    if betreff and betreff != "(kein Betreff)":
        return betreff
    if text:
        erste_zeile = text.strip().split("\n")[0].strip()
        if erste_zeile:
            if len(erste_zeile) > 100:
                erste_zeile = erste_zeile[:100] + "…"
            return erste_zeile
    return None


def ist_faellig(iso_datum, status):
    if not iso_datum or status == "erledigt":
        return False
    try:
        d = datetime.strptime(iso_datum, "%Y-%m-%d").date()
        return d <= datetime.now().date()
    except ValueError:
        return False


# ---------------------------------------------------------------------------
# Wiederverwendbare, modern gestaltete Widgets
# ---------------------------------------------------------------------------
class Pill(tk.Canvas):
    """Abgerundeter Button / Badge. command=None ergibt ein reines Badge."""

    def __init__(self, parent, text, bg, fg="#FFFFFF", command=None,
                 font=None, padx=14, pady=7, radius=14, container_bg=COLOR_SURFACE,
                 hover_bg=None, hover_fg=None, outline=""):
        self.font = tkfont.Font(font=font) if font else tkfont.Font(family="Segoe UI", size=9, weight="bold")
        text_w = self.font.measure(text)
        text_h = self.font.metrics("linespace")
        w = text_w + padx * 2
        h = text_h + pady * 2
        super().__init__(parent, width=w, height=h, bg=container_bg,
                          highlightthickness=0, bd=0,
                          cursor="hand2" if command else "arrow")
        self._pw, self._ph = w, h
        self._text = text
        self._bg = bg
        self._hover_bg = hover_bg or bg
        self._fg = fg
        self._hover_fg = hover_fg or fg
        self._radius = radius
        self._outline = outline
        self._draw(self._bg, self._fg)
        if command:
            self.bind("<Button-1>", lambda e: command())
            self.bind("<Enter>", lambda e: self._draw(self._hover_bg, self._hover_fg))
            self.bind("<Leave>", lambda e: self._draw(self._bg, self._fg))

    def _draw(self, bg, fg):
        self.delete("all")
        r = min(self._radius, self._ph / 2)
        self._rr(1, 1, self._pw - 1, self._ph - 1, r, fill=bg, outline=self._outline)
        self.create_text(self._pw / 2, self._ph / 2, text=self._text, fill=fg, font=self.font)

    def _rr(self, x1, y1, x2, y2, r, **kwargs):
        pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
               x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
        return self.create_polygon(pts, smooth=True, **kwargs)


class RoundedCard(tk.Frame):
    """Container mit abgerundeten Ecken + dezentem Schatten. Inhalt kommt in .body."""

    def __init__(self, parent, column_bg, card_bg=COLOR_CARD_BG, radius=14,
                 shadow=COLOR_CARD_SHADOW, accent=None, accent_width=5, pad=14):
        super().__init__(parent, bg=column_bg)
        self._radius = radius
        self._shadow = shadow
        self._card_bg = card_bg
        self._accent = accent
        self._accent_w = accent_width
        self._pad = pad
        self._width = 0

        self.canvas = tk.Canvas(self, bg=column_bg, highlightthickness=0, bd=0)
        self.canvas.pack(fill="both", expand=True)

        self.body = tk.Frame(self.canvas, bg=card_bg)
        self._left_inset = pad + (accent_width + 10 if accent else 0)
        self._win = self.canvas.create_window(self._left_inset, pad, window=self.body, anchor="nw")

        self.body.bind("<Configure>", lambda e: self._redraw())
        self.canvas.bind("<Configure>", self._on_resize)

    def _on_resize(self, event):
        self._width = event.width
        content_w = max(10, self._width - self._left_inset - self._pad)
        self.canvas.itemconfig(self._win, width=content_w)
        self._redraw()

    def _redraw(self):
        if self._width <= 0:
            return
        self.body.update_idletasks()
        h = self.body.winfo_reqheight()
        total_h = h + self._pad * 2
        self.canvas.config(height=total_h)
        self.canvas.delete("bg")
        w = self._width
        self._rr(self.canvas, 3, 4, w - 1, total_h, self._radius, fill=self._shadow, outline="", tag="bg")
        self._rr(self.canvas, 0, 0, w - 4, total_h - 4, self._radius, fill=self._card_bg, outline="", tag="bg")
        if self._accent:
            bx1 = 10
            bx2 = bx1 + self._accent_w
            self._rr(self.canvas, bx1, 12, bx2, total_h - 4 - 12, self._accent_w / 2,
                      fill=self._accent, outline="", tag="bg")
        self.canvas.tag_lower("bg")

    @staticmethod
    def _rr(canvas, x1, y1, x2, y2, r, tag="bg", **kwargs):
        r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
        pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
               x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
        return canvas.create_polygon(pts, smooth=True, tags=tag, **kwargs)


def draw_dropzone(canvas, w, h):
    canvas.delete("all")
    color = COLOR_BORDER
    bg = COLOR_SURFACE
    r = 12
    pts = [r, 2, w - r, 2, w - 2, r, w - 2, h - r, w - r, h - 2, r, h - 2, 2, h - r, 2, r]
    canvas.create_polygon(pts, smooth=True, fill=bg, outline="")
    canvas.create_rectangle(2, 2, w - 2, h - 2, outline=color, width=1.5, dash=(5, 3))


class StyledEntry(tk.Entry):
    """Entry mit farbigem Fokus-Rahmen statt Standard-Windows-Optik."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, font=F_NORMAL, relief="flat",
                          highlightthickness=1.5, highlightbackground=COLOR_BORDER,
                          highlightcolor=COLOR_ACCENT, bg=COLOR_SURFACE, fg=COLOR_TEXT_PRIMARY,
                          insertbackground=COLOR_TEXT_PRIMARY, **kwargs)


class StyledText(tk.Text):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, font=F_NORMAL, relief="flat",
                          highlightthickness=1.5, highlightbackground=COLOR_BORDER,
                          highlightcolor=COLOR_ACCENT, bg=COLOR_SURFACE, fg=COLOR_TEXT_PRIMARY,
                          insertbackground=COLOR_TEXT_PRIMARY, wrap="word", **kwargs)


# ---------------------------------------------------------------------------
# Dialoge
# ---------------------------------------------------------------------------
class BaseDialog(tk.Toplevel):
    def __init__(self, parent, title, width=420):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=COLOR_HEADER_BG)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.width = width

    def header(self, text):
        tk.Label(self, text=text, font=F_COLUMN_TITLE, bg=COLOR_HEADER_BG,
                 fg=COLOR_TEXT_PRIMARY).pack(anchor="w", padx=18, pady=(18, 4))

    def button_row(self, buttons):
        """buttons: Liste von (text, bg, fg, command)."""
        row = tk.Frame(self, bg=COLOR_HEADER_BG)
        row.pack(fill="x", padx=18, pady=(6, 18))
        for text, bg, fg, cmd in buttons:
            p = Pill(row, text, bg=bg, fg=fg, command=cmd, container_bg=COLOR_HEADER_BG,
                      hover_bg=COLOR_ACCENT_HOVER if bg == COLOR_ACCENT else None)
            p.pack(side="right", padx=(6, 0))


class NeuerEintragDialog(BaseDialog):
    def __init__(self, parent, typ, on_save):
        super().__init__(parent, "Neue Aufgabe" if typ == "aufgabe" else "Neuer Rückruf")
        self.on_save = on_save
        self.typ = typ

        self.header("Titel / Kunde")
        self.titel_entry = StyledEntry(self, width=38)
        self.titel_entry.pack(padx=18, pady=(0, 12), fill="x")
        self.titel_entry.focus_set()

        self.notiz_text = None
        if typ == "aufgabe":
            tk.Label(self, text="Notiz (optional)", font=F_NORMAL, bg=COLOR_HEADER_BG,
                     fg=COLOR_TEXT_SECONDARY).pack(anchor="w", padx=18)
            self.notiz_text = StyledText(self, width=38, height=4)
            self.notiz_text.pack(padx=18, pady=(4, 6), fill="x")

        self.button_row([
            ("Abbrechen", COLOR_SURFACE, COLOR_TEXT_SECONDARY, self.destroy),
            ("Speichern", COLOR_ACCENT, "#FFFFFF", self._speichern),
        ])
        self.bind("<Return>", lambda e: self._speichern())

    def _speichern(self):
        titel = self.titel_entry.get().strip()
        if not titel:
            self.titel_entry.focus_set()
            return
        notiz = self.notiz_text.get("1.0", "end").strip() if self.notiz_text else ""
        self.on_save(titel, notiz)
        self.destroy()


class TitelDialog(BaseDialog):
    def __init__(self, parent, dialogtitel, aktueller_wert, on_save):
        super().__init__(parent, dialogtitel)
        self.on_save = on_save
        self.header(dialogtitel)
        self.entry = StyledEntry(self, width=38)
        self.entry.pack(padx=18, pady=(0, 12), fill="x")
        self.entry.insert(0, aktueller_wert)
        self.entry.focus_set()
        self.entry.select_range(0, tk.END)
        self.button_row([
            ("Abbrechen", COLOR_SURFACE, COLOR_TEXT_SECONDARY, self.destroy),
            ("Speichern", COLOR_ACCENT, "#FFFFFF", self._speichern),
        ])
        self.bind("<Return>", lambda e: self._speichern())

    def _speichern(self):
        neuer_wert = self.entry.get().strip()
        if neuer_wert:
            self.on_save(neuer_wert)
        self.destroy()


class NotizDialog(BaseDialog):
    def __init__(self, parent, aktuelle_notiz, on_save):
        super().__init__(parent, "Notiz bearbeiten")
        self.on_save = on_save
        self.header("Notiz")
        self.text = StyledText(self, width=38, height=6)
        self.text.pack(padx=18, pady=(0, 6), fill="x")
        self.text.insert("1.0", aktuelle_notiz)
        self.text.focus_set()
        self.button_row([
            ("Abbrechen", COLOR_SURFACE, COLOR_TEXT_SECONDARY, self.destroy),
            ("Speichern", COLOR_ACCENT, "#FFFFFF", self._speichern),
        ])

    def _speichern(self):
        self.on_save(self.text.get("1.0", "end").strip())
        self.destroy()


class VerlaufDialog(BaseDialog):
    def __init__(self, parent, on_save):
        super().__init__(parent, "Notiz hinzufügen")
        self.on_save = on_save
        self.header("z.B. \"Telefonat mit Kunde, Termin für Freitag vereinbart\"")
        self.text = StyledText(self, width=38, height=5)
        self.text.pack(padx=18, pady=(0, 6), fill="x")
        self.text.focus_set()
        self.button_row([
            ("Abbrechen", COLOR_SURFACE, COLOR_TEXT_SECONDARY, self.destroy),
            ("Speichern", COLOR_ACCENT, "#FFFFFF", self._speichern),
        ])
        self.bind("<Control-Return>", lambda e: self._speichern())

    def _speichern(self):
        text = self.text.get("1.0", "end").strip()
        if text:
            self.on_save(text)
        self.destroy()


class WiedervorlageDialog(BaseDialog):
    def __init__(self, parent, aktuelles_datum, on_save):
        super().__init__(parent, "Wiedervorlage setzen")
        self.on_save = on_save

        self.header("Datum (TT.MM.JJJJ)")
        self.entry = StyledEntry(self, width=16)
        self.entry.pack(padx=18, pady=(0, 12), anchor="w")
        if aktuelles_datum:
            self.entry.insert(0, iso_zu_anzeige(aktuelles_datum))

        quick = tk.Frame(self, bg=COLOR_HEADER_BG)
        quick.pack(padx=18, pady=(0, 12), anchor="w")
        for label, tage in [("Heute", 0), ("Morgen", 1), ("+3 Tage", 3), ("Nächste Woche", 7)]:
            p = Pill(quick, label, bg=COLOR_ACCENT_LIGHT, fg=COLOR_ACCENT,
                      hover_bg=COLOR_ACCENT_LIGHT_HOVER, container_bg=COLOR_HEADER_BG,
                      command=lambda t=tage: self._quick(t))
            p.pack(side="left", padx=(0, 6))

        self.button_row([
            ("Entfernen", COLOR_SURFACE, COLOR_DANGER, self._entfernen),
            ("Speichern", COLOR_ACCENT, "#FFFFFF", self._speichern),
        ])

    def _quick(self, tage):
        d = datetime.now().date() + timedelta(days=tage)
        self.entry.delete(0, tk.END)
        self.entry.insert(0, d.strftime("%d.%m.%Y"))

    def _speichern(self):
        text = self.entry.get().strip()
        if not text:
            self._entfernen()
            return
        try:
            iso = datetime.strptime(text, "%d.%m.%Y").strftime("%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Ungültiges Datum", "Bitte im Format TT.MM.JJJJ eingeben, z.B. 15.09.2026")
            return
        self.on_save(iso)
        self.destroy()

    def _entfernen(self):
        self.on_save(None)
        self.destroy()


class EmailAnsicht(tk.Toplevel):
    def __init__(self, parent, row):
        super().__init__(parent)
        self.title(row["betreff"] or "E-Mail")
        self.geometry("580x520")
        self.configure(bg=COLOR_HEADER_BG)

        kopf = tk.Frame(self, bg=COLOR_HEADER_BG)
        kopf.pack(fill="x", padx=20, pady=(18, 6))
        tk.Label(kopf, text=row["betreff"] or "(kein Betreff)", font=F_COLUMN_TITLE,
                 bg=COLOR_HEADER_BG, fg=COLOR_TEXT_PRIMARY, wraplength=520,
                 justify="left", anchor="w").pack(anchor="w")
        tk.Label(kopf, text=f"Von: {row['absender'] or '-'}", font=F_NORMAL,
                 bg=COLOR_HEADER_BG, fg=COLOR_TEXT_SECONDARY).pack(anchor="w", pady=(6, 0))
        tk.Label(kopf, text=f"Datum: {row['email_datum'] or '-'}", font=F_NORMAL,
                 bg=COLOR_HEADER_BG, fg=COLOR_TEXT_SECONDARY).pack(anchor="w")
        if row["email_dateiname"]:
            tk.Label(kopf, text=f"Datei: {row['email_dateiname']}", font=F_SMALL,
                     bg=COLOR_HEADER_BG, fg=COLOR_TEXT_MUTED).pack(anchor="w", pady=(2, 0))

        text_widget = StyledText(self, height=18)
        text_widget.insert("1.0", row["email_text"] or "(kein Text)")
        text_widget.config(state="disabled")
        text_widget.pack(fill="both", expand=True, padx=20, pady=(6, 20))


class ScrollableFrame(tk.Frame):
    """Ein vertikal scrollbarer Container fuer die Karten einer Spalte."""

    def __init__(self, parent, bg=COLOR_BG, **kwargs):
        super().__init__(parent, bg=bg, **kwargs)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        scrollbar = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview,
                                  bg=bg, troughcolor=bg, activebackground=COLOR_SCROLLBAR_HOVER,
                                  bd=0, highlightthickness=0, width=10)
        self.inner = tk.Frame(self.canvas, bg=bg)

        self.inner.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas_window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width),
        )
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self.canvas.bind_all("<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units"), add="+")
        self.canvas.bind_all("<Button-5>", lambda e: self.canvas.yview_scroll(1, "units"), add="+")

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


# ---------------------------------------------------------------------------
# Hauptanwendung
# ---------------------------------------------------------------------------
class TodoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Werkstatt To-Do – Aufgaben & Rückrufe")
        self.root.geometry("1240x800")
        self.root.configure(bg=COLOR_BG)

        self.show_erledigt = {"aufgabe": False, "rueckruf": False}
        self.filter_status = {"aufgabe": "alle", "rueckruf": "alle"}
        self.toggle_pills = {}
        self.filter_pills = {"aufgabe": {}, "rueckruf": {}}
        self._count_holder = {}
        self.expanded_rueckrufe = set()
        self.expanded_aufgaben = set()
        self.search_query = ""

        self._build_header()
        self._build_layout()
        self.refresh_all()

    # ---------- Layout ----------
    def _build_header(self):
        header = tk.Frame(self.root, bg=COLOR_HEADER_BG)
        header.pack(fill="x")
        inner = tk.Frame(header, bg=COLOR_HEADER_BG)
        inner.pack(fill="x", padx=24, pady=16)
        tk.Label(inner, text="Werkstatt To-Do", font=F_TITLE, bg=COLOR_HEADER_BG,
                 fg=COLOR_TEXT_PRIMARY).pack(side="left")
        tk.Label(inner, text="Aufgaben & Rückrufe", font=F_NORMAL, bg=COLOR_HEADER_BG,
                 fg=COLOR_TEXT_MUTED).pack(side="left", padx=(12, 0), pady=(6, 0))

        such_frame = tk.Frame(inner, bg=COLOR_HEADER_BG)
        such_frame.pack(side="right")
        tk.Label(such_frame, text="🔍", font=F_NORMAL, bg=COLOR_HEADER_BG,
                 fg=COLOR_TEXT_MUTED).pack(side="left", padx=(0, 8))
        self.search_entry = StyledEntry(such_frame, width=30)
        self.search_entry.pack(side="left", ipady=3)
        self.search_entry.bind("<KeyRelease>", self._on_search_change)

        tk.Frame(self.root, bg=COLOR_ACCENT, height=3).pack(fill="x")

    def _on_search_change(self, event=None):
        self.search_query = self.search_entry.get().strip().lower()
        self.refresh_all()

    def _build_layout(self):
        columns = tk.Frame(self.root, bg=COLOR_BG)
        columns.pack(fill="both", expand=True, padx=16, pady=16)
        columns.columnconfigure(0, weight=1)
        columns.columnconfigure(1, weight=1)
        columns.rowconfigure(0, weight=1)

        self.aufgaben_col = self._build_column(columns, "Aufgaben", "aufgabe", ist_rueckruf=False)
        self.aufgaben_col.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        self.rueckrufe_col = self._build_column(columns, "Rückrufe", "rueckruf", ist_rueckruf=True)
        self.rueckrufe_col.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

    def _build_column(self, parent, titel, typ, ist_rueckruf):
        outer = tk.Frame(parent, bg=COLOR_BG)

        kopf = tk.Frame(outer, bg=COLOR_BG)
        kopf.pack(fill="x", pady=(0, 10))
        title_row = tk.Frame(kopf, bg=COLOR_BG)
        title_row.pack(fill="x")
        tk.Label(title_row, text=titel, font=F_COLUMN_TITLE, bg=COLOR_BG,
                 fg=COLOR_TEXT_PRIMARY).pack(side="left")
        count_badge = Pill(title_row, "0", bg=COLOR_ACCENT_LIGHT, fg=COLOR_ACCENT,
                            container_bg=COLOR_BG, font=F_SMALL, padx=9, pady=3, radius=10)
        count_badge.pack(side="left", padx=(8, 0))
        self._count_holder[typ] = count_badge

        toggle = Pill(title_row, "Erledigte: Aus", bg=COLOR_SURFACE, fg=COLOR_TEXT_SECONDARY,
                       hover_bg=STATUS_INACTIVE_HOVER, outline=COLOR_BORDER,
                       container_bg=COLOR_BG, font=F_SMALL, padx=10, pady=5,
                       command=lambda t=typ: self._toggle_erledigt(t))
        toggle.pack(side="right")
        self.toggle_pills[typ] = toggle

        aktionen = tk.Frame(outer, bg=COLOR_BG)
        aktionen.pack(fill="x", pady=(0, 8))
        Pill(aktionen, f"+ Neue{'r Rückruf' if ist_rueckruf else ' Aufgabe'}",
             bg=COLOR_ACCENT, fg="#FFFFFF", hover_bg=COLOR_ACCENT_HOVER,
             container_bg=COLOR_BG, font=F_PILL,
             command=lambda t=typ: self._neuer_eintrag(t)).pack(side="left")

        filter_row = tk.Frame(outer, bg=COLOR_BG)
        filter_row.pack(fill="x", pady=(0, 8))
        tk.Label(filter_row, text="Anzeigen:", font=F_SMALL, bg=COLOR_BG,
                 fg=COLOR_TEXT_MUTED).pack(side="left", padx=(0, 6))
        for label, key in FILTER_OPTIONEN[typ]:
            aktiv = self.filter_status[typ] == key
            pill = Pill(filter_row, label,
                        bg=COLOR_ACCENT if aktiv else COLOR_SURFACE,
                        fg="#FFFFFF" if aktiv else COLOR_TEXT_SECONDARY,
                        hover_bg=COLOR_ACCENT_HOVER if aktiv else STATUS_INACTIVE_HOVER,
                        outline="" if aktiv else COLOR_BORDER,
                        container_bg=COLOR_BG, font=F_SMALL, padx=9, pady=4,
                        command=lambda t=typ, k=key: self._set_filter(t, k))
            pill.pack(side="left", padx=(0, 5))
            self.filter_pills[typ][key] = pill

        if OUTLOOK_COM_AVAILABLE:
            Pill(aktionen, "📥 Markierte Outlook-Mail übernehmen", bg=COLOR_ACCENT, fg="#FFFFFF",
                 hover_bg=COLOR_ACCENT_HOVER, container_bg=COLOR_BG, font=F_PILL,
                 command=lambda t=typ: self._outlook_uebernehmen(t)).pack(side="left", padx=(8, 0))

        Pill(aktionen, "Datei auswählen", bg=COLOR_ACCENT_LIGHT, fg=COLOR_ACCENT,
             hover_bg=COLOR_ACCENT_LIGHT_HOVER, container_bg=COLOR_BG, font=F_PILL,
             command=lambda t=typ: self._datei_dialog(t)).pack(side="left", padx=(8, 0))

        drop_wrap = tk.Frame(outer, bg=COLOR_BG, height=54)
        drop_wrap.pack(fill="x", pady=(0, 10))
        drop_wrap.pack_propagate(False)
        drop_canvas = tk.Canvas(drop_wrap, bg=COLOR_BG, highlightthickness=0, bd=0)
        drop_canvas.pack(fill="both", expand=True)
        hinweis = ("Bereits gespeicherte .msg/.eml hier ablegen"
                    if DND_AVAILABLE else
                    "Drag & Drop nicht verfügbar – 'Datei auswählen' nutzen")

        def redraw_drop(event=None, canvas=drop_canvas, text=hinweis):
            w = canvas.winfo_width()
            h = canvas.winfo_height()
            if w > 1 and h > 1:
                draw_dropzone(canvas, w, h)
                canvas.create_text(w / 2, h / 2, text=text, font=F_SMALL, fill=COLOR_TEXT_MUTED)

        drop_canvas.bind("<Configure>", redraw_drop)

        if DND_AVAILABLE:
            drop_canvas.drop_target_register(DND_FILES)
            drop_canvas.dnd_bind("<<Drop>>", lambda e, t=typ: self._on_drop(e, t))

        scroll = ScrollableFrame(outer, bg=COLOR_BG)
        scroll.pack(fill="both", expand=True)

        if typ == "aufgabe":
            self.aufgaben_scroll = scroll
        else:
            self.rueckrufe_scroll = scroll

        return outer

    # ---------- Aktionen ----------
    def _toggle_erledigt(self, typ):
        self.show_erledigt[typ] = not self.show_erledigt[typ]
        pill = self.toggle_pills[typ]
        an = self.show_erledigt[typ]
        pill._text = f"Erledigte: {'An' if an else 'Aus'}"
        pill._bg = COLOR_ACCENT_LIGHT if an else COLOR_SURFACE
        pill._fg = COLOR_ACCENT if an else COLOR_TEXT_SECONDARY
        pill._hover_bg = COLOR_ACCENT_LIGHT_HOVER if an else STATUS_INACTIVE_HOVER
        pill._draw(pill._bg, pill._fg)
        self.refresh_all()

    def _set_filter(self, typ, key):
        self.filter_status[typ] = key
        for k, pill in self.filter_pills[typ].items():
            aktiv = (k == key)
            pill._bg = COLOR_ACCENT if aktiv else COLOR_SURFACE
            pill._fg = "#FFFFFF" if aktiv else COLOR_TEXT_SECONDARY
            pill._hover_bg = COLOR_ACCENT_HOVER if aktiv else STATUS_INACTIVE_HOVER
            pill._draw(pill._bg, pill._fg)
        self.refresh_all()

    def _neuer_eintrag(self, typ):
        def speichern(titel, notiz):
            db.add_task(typ, titel, notiz)
            self.refresh_all()
        NeuerEintragDialog(self.root, typ, speichern)

    def _task_aus_email_erstellen(self, typ, daten, dateiname=None):
        if typ == "rueckruf":
            titel = name_aus_absender(daten["absender"], daten["betreff"] or "(kein Betreff)")
        else:
            titel = daten["betreff"] or "(kein Betreff)"
        notiz = daten["text"] if typ == "aufgabe" else ""
        db.add_task(
            typ,
            titel=titel,
            notiz=notiz,
            absender=daten["absender"],
            betreff=daten["betreff"],
            email_datum=daten["datum"],
            email_text=daten["text"],
            email_dateiname=dateiname,
        )

    def _outlook_uebernehmen(self, typ="rueckruf"):
        try:
            mails = outlook_auswahl_lesen()
        except RuntimeError as e:
            messagebox.showwarning("Outlook", str(e))
            return
        except Exception as e:
            messagebox.showerror(
                "Outlook",
                f"Unerwarteter Fehler beim Zugriff auf Outlook:\n{e}",
            )
            return

        for daten in mails:
            self._task_aus_email_erstellen(typ, daten)
        self.refresh_all()

    def _datei_dialog(self, typ="rueckruf"):
        pfade = filedialog.askopenfilenames(
            title="Outlook-Mail auswählen",
            filetypes=[("E-Mail Dateien", "*.msg *.eml"), ("Alle Dateien", "*.*")],
        )
        for pfad in pfade:
            self._email_importieren(pfad, typ)

    def _on_drop(self, event, typ="rueckruf"):
        pfade = self.root.tk.splitlist(event.data)
        for pfad in pfade:
            self._email_importieren(pfad, typ)

    def _email_importieren(self, pfad, typ="rueckruf"):
        pfad = pfad.strip("{}")
        if not os.path.isfile(pfad):
            messagebox.showerror("Fehler", f"Datei nicht gefunden:\n{pfad}")
            return
        try:
            daten = parse_email_file(pfad)
        except ValueError as e:
            messagebox.showerror("Import fehlgeschlagen", str(e))
            return
        except Exception as e:
            messagebox.showerror("Import fehlgeschlagen", f"Unerwarteter Fehler:\n{e}")
            return

        self._task_aus_email_erstellen(typ, daten, dateiname=os.path.basename(pfad))
        self.refresh_all()

    def _status_setzen(self, task_id, status):
        db.update_status(task_id, status)
        self.refresh_all()

    def _wiedervorlage_dialog(self, task_id, aktuelles_datum):
        def speichern(iso_datum):
            db.update_wiedervorlage(task_id, iso_datum)
            self.refresh_all()
        WiedervorlageDialog(self.root, aktuelles_datum, speichern)

    def _notiz_bearbeiten(self, task_id, aktuelle_notiz):
        def speichern(neue_notiz):
            db.update_notiz(task_id, neue_notiz)
            self.refresh_all()
        NotizDialog(self.root, aktuelle_notiz, speichern)

    def _titel_bearbeiten(self, task_id, aktueller_titel, dialogtitel):
        def speichern(neuer_titel):
            db.update_titel(task_id, neuer_titel)
            self.refresh_all()
        TitelDialog(self.root, dialogtitel, aktueller_titel, speichern)

    def _loeschen(self, task_id, titel):
        if messagebox.askyesno("Löschen bestätigen", f"'{titel}' wirklich löschen?"):
            db.delete_task(task_id)
            self.refresh_all()

    def _email_anzeigen(self, row):
        EmailAnsicht(self.root, row)

    def _anhang_hinzufuegen(self, task_id):
        pfade = filedialog.askopenfilenames(
            title="PDF anhängen",
            filetypes=[("PDF Dateien", "*.pdf"), ("Alle Dateien", "*.*")],
        )
        for pfad in pfade:
            try:
                db.add_attachment(task_id, pfad)
            except Exception as e:
                messagebox.showerror("Fehler", f"Konnte Datei nicht anhängen:\n{e}")
        if pfade:
            self.refresh_all()

    def _anhang_oeffnen(self, pfad):
        if not os.path.isfile(pfad):
            messagebox.showerror("Fehler", f"Datei nicht gefunden:\n{pfad}")
            return
        try:
            datei_oeffnen(pfad)
        except Exception as e:
            messagebox.showerror("Fehler", f"Konnte Datei nicht öffnen:\n{e}")

    def _anhang_loeschen(self, attachment_id):
        if messagebox.askyesno("Anhang löschen", "Diesen Anhang wirklich löschen?"):
            db.delete_attachment(attachment_id)
            self.refresh_all()

    # ---------- Rendern ----------
    def refresh_all(self):
        self._render_column(self.aufgaben_scroll, "aufgabe")
        self._render_column(self.rueckrufe_scroll, "rueckruf")

    def _render_column(self, scroll_frame, typ):
        for widget in scroll_frame.inner.winfo_children():
            widget.destroy()

        alle = db.get_tasks(typ)
        offene_anzahl = len([r for r in alle if r["status"] != "erledigt"])
        badge = self._count_holder[typ]
        badge._text = str(offene_anzahl)
        badge._draw(COLOR_ACCENT_LIGHT, COLOR_ACCENT)

        rows = alle if self.show_erledigt[typ] else [r for r in alle if r["status"] != "erledigt"]

        filter_key = self.filter_status[typ]
        if filter_key != "alle":
            rows = [r for r in rows if r["status"] == filter_key]

        if self.search_query:
            rows = [r for r in rows if self._treffer_suche(r, typ)]

        rows = sorted(
            rows,
            key=lambda r: (not ist_faellig(r["wiedervorlage"], r["status"]),
                            r["wiedervorlage"] or "9999-99-99"),
        )

        if not rows:
            if self.search_query:
                leer = f"Keine Treffer für „{self.search_entry.get().strip()}“."
            else:
                leer = "Keine offenen Aufgaben." if typ == "aufgabe" else "Keine offenen Rückrufe."
            tk.Label(scroll_frame.inner, text=leer, bg=COLOR_BG,
                     fg=COLOR_TEXT_MUTED, font=F_NORMAL, wraplength=380, justify="left"
                     ).pack(pady=30)
            return

        for row in rows:
            self._render_card(scroll_frame.inner, row, typ)

    def _treffer_suche(self, row, typ):
        teile = [row["titel"] or "", row["notiz"] or "",
                 row["absender"] or "", row["betreff"] or "", row["email_text"] or ""]
        for eintrag in db.get_notiz_eintraege(row["id"]):
            teile.append(eintrag["text"])
        heuhaufen = " ".join(teile).lower()
        return self.search_query in heuhaufen

    def _toggle_expand(self, task_id, typ="rueckruf"):
        zielset = self.expanded_aufgaben if typ == "aufgabe" else self.expanded_rueckrufe
        if task_id in zielset:
            zielset.discard(task_id)
        else:
            zielset.add(task_id)
        self.refresh_all()

    def _klickbar_umschalten(self, widget, task_id, typ):
        """Macht ein Label/Frame anklickbar, um die Karte auf-/zuzuklappen."""
        widget.configure(cursor="hand2")
        widget.bind("<Button-1>", lambda e: self._toggle_expand(task_id, typ))

    def _render_card(self, parent, row, typ):
        if typ == "rueckruf":
            self._render_rueckruf_card(parent, row)
        else:
            self._render_aufgabe_card(parent, row)

    def _render_aufgabe_card(self, parent, row):
        faellig = ist_faellig(row["wiedervorlage"], row["status"])
        accent = COLOR_DANGER if faellig else STATUS_COLORS[row["status"]]
        expanded = row["id"] in self.expanded_aufgaben

        card = RoundedCard(parent, column_bg=COLOR_BG, accent=accent, accent_width=5, pad=14)
        card.pack(fill="x", pady=(0, 12))
        body = card.body

        kopf = tk.Frame(body, bg=COLOR_CARD_BG)
        kopf.pack(fill="x")
        titel_label = tk.Label(kopf, text=row["titel"], font=F_CARD_TITLE, bg=COLOR_CARD_BG,
                 fg=COLOR_TEXT_PRIMARY, wraplength=300, justify="left", anchor="w")
        titel_label.pack(side="left", fill="x", expand=True)
        Pill(kopf, "✕", bg=COLOR_CARD_BG, fg=COLOR_TEXT_MUTED, hover_bg=COLOR_DANGER_LIGHT,
             hover_fg=COLOR_DANGER, container_bg=COLOR_CARD_BG, font=F_SMALL,
             padx=6, pady=3, radius=10,
             command=lambda: self._loeschen(row["id"], row["titel"])).pack(side="right")
        Pill(kopf, "▾ Einklappen" if expanded else "▸ Details", bg=COLOR_CARD_BG, fg=COLOR_ACCENT,
             hover_bg=COLOR_ACCENT_LIGHT, container_bg=COLOR_CARD_BG, font=F_SMALL,
             padx=8, pady=3, radius=10,
             command=lambda t=row["id"]: self._toggle_expand(t, "aufgabe")).pack(side="right", padx=(0, 4))
        Pill(kopf, "✎", bg=COLOR_CARD_BG, fg=COLOR_TEXT_MUTED, hover_bg=COLOR_ACCENT_LIGHT,
             hover_fg=COLOR_ACCENT, container_bg=COLOR_CARD_BG, font=F_SMALL,
             padx=6, pady=3, radius=10,
             command=lambda t=row["id"], v=row["titel"]: self._titel_bearbeiten(t, v, "Titel bearbeiten")
             ).pack(side="right", padx=(0, 4))
        for w in (kopf, titel_label, body):
            self._klickbar_umschalten(w, row["id"], "aufgabe")

        if not expanded:
            if row["notiz"]:
                vorschau = row["notiz"].replace("\n", " ").strip()
                if len(vorschau) > 90:
                    vorschau = vorschau[:90] + "…"
                notiz_label = tk.Label(body, text=vorschau, font=F_NORMAL, bg=COLOR_CARD_BG, fg=COLOR_TEXT_SECONDARY,
                         anchor="w", wraplength=400, justify="left")
                notiz_label.pack(fill="x", pady=(4, 0))
                self._klickbar_umschalten(notiz_label, row["id"], "aufgabe")
            wv_text = "Wiedervorlage: " + (iso_zu_anzeige(row["wiedervorlage"]) if row["wiedervorlage"] else "–")
            farbe = COLOR_DANGER if faellig else COLOR_TEXT_MUTED
            wv_label = tk.Label(body, text=wv_text, font=F_SMALL, bg=COLOR_CARD_BG, fg=farbe,
                     anchor="w")
            wv_label.pack(fill="x", pady=(2, 0))
            self._klickbar_umschalten(wv_label, row["id"], "aufgabe")
            return

        if row["absender"]:
            info = row["absender"]
            if row["email_datum"]:
                info += f"   ·   {row['email_datum']}"
            info_label = tk.Label(body, text=info, font=F_SMALL, bg=COLOR_CARD_BG, fg=COLOR_TEXT_MUTED,
                     anchor="w", wraplength=400, justify="left")
            info_label.pack(fill="x", pady=(4, 0))
            self._klickbar_umschalten(info_label, row["id"], "aufgabe")

        if row["notiz"]:
            notiz_voll_label = tk.Label(body, text=row["notiz"], font=F_NORMAL, bg=COLOR_CARD_BG, fg=COLOR_TEXT_SECONDARY,
                     anchor="w", wraplength=400, justify="left")
            notiz_voll_label.pack(fill="x", pady=(4, 0))
            self._klickbar_umschalten(notiz_voll_label, row["id"], "aufgabe")
        Pill(body, "Notiz bearbeiten", bg=COLOR_ACCENT_LIGHT, fg=COLOR_ACCENT,
             hover_bg=COLOR_ACCENT_LIGHT_HOVER, container_bg=COLOR_CARD_BG, font=F_SMALL,
             command=lambda r=row: self._notiz_bearbeiten(r["id"], r["notiz"] or "")
             ).pack(anchor="w", pady=(8, 0))

        self._render_verlauf_block(body, row["id"])
        self._render_anhaenge_block(body, row["id"], enable_dnd=True)
        self._render_status_block(body, row, faellig)

    def _render_rueckruf_card(self, parent, row):
        faellig = ist_faellig(row["wiedervorlage"], row["status"])
        accent = COLOR_DANGER if faellig else STATUS_COLORS[row["status"]]
        expanded = row["id"] in self.expanded_rueckrufe

        card = RoundedCard(parent, column_bg=COLOR_BG, accent=accent, accent_width=5, pad=14)
        card.pack(fill="x", pady=(0, 12))
        body = card.body

        name = row["titel"]
        nummer = telefonnummer_ermitteln(row["email_text"], row["titel"]) or "–"
        grund = grund_ermitteln(row["betreff"], row["email_text"]) or "–"

        kopf = tk.Frame(body, bg=COLOR_CARD_BG)
        kopf.pack(fill="x")
        name_label = tk.Label(kopf, text=name, font=F_CARD_TITLE, bg=COLOR_CARD_BG,
                 fg=COLOR_TEXT_PRIMARY, wraplength=300, justify="left", anchor="w")
        name_label.pack(side="left", fill="x", expand=True)
        Pill(kopf, "✕", bg=COLOR_CARD_BG, fg=COLOR_TEXT_MUTED, hover_bg=COLOR_DANGER_LIGHT,
             hover_fg=COLOR_DANGER, container_bg=COLOR_CARD_BG, font=F_SMALL,
             padx=6, pady=3, radius=10,
             command=lambda: self._loeschen(row["id"], row["titel"])).pack(side="right")
        Pill(kopf, "▾ Einklappen" if expanded else "▸ Details", bg=COLOR_CARD_BG, fg=COLOR_ACCENT,
             hover_bg=COLOR_ACCENT_LIGHT, container_bg=COLOR_CARD_BG, font=F_SMALL,
             padx=8, pady=3, radius=10,
             command=lambda t=row["id"]: self._toggle_expand(t)).pack(side="right", padx=(0, 4))
        Pill(kopf, "✎", bg=COLOR_CARD_BG, fg=COLOR_TEXT_MUTED, hover_bg=COLOR_ACCENT_LIGHT,
             hover_fg=COLOR_ACCENT, container_bg=COLOR_CARD_BG, font=F_SMALL,
             padx=6, pady=3, radius=10,
             command=lambda t=row["id"], v=row["titel"]: self._titel_bearbeiten(t, v, "Namen bearbeiten")
             ).pack(side="right", padx=(0, 4))
        for w in (kopf, name_label, body):
            self._klickbar_umschalten(w, row["id"], "rueckruf")

        nummer_label = tk.Label(body, text=f"📞 {nummer}", font=F_NORMAL, bg=COLOR_CARD_BG, fg=COLOR_TEXT_SECONDARY,
                 anchor="w")
        nummer_label.pack(fill="x", pady=(4, 0))
        grund_label = tk.Label(body, text=f"Grund: {grund}", font=F_NORMAL, bg=COLOR_CARD_BG, fg=COLOR_TEXT_SECONDARY,
                 anchor="w", wraplength=400, justify="left")
        grund_label.pack(fill="x", pady=(2, 0))
        self._klickbar_umschalten(nummer_label, row["id"], "rueckruf")
        self._klickbar_umschalten(grund_label, row["id"], "rueckruf")

        if not expanded:
            return

        info = row["absender"] or "-"
        if row["email_datum"]:
            info += f"   ·   {row['email_datum']}"
        info_label = tk.Label(body, text=info, font=F_SMALL, bg=COLOR_CARD_BG, fg=COLOR_TEXT_MUTED,
                 anchor="w", wraplength=400, justify="left")
        info_label.pack(fill="x", pady=(10, 0))
        self._klickbar_umschalten(info_label, row["id"], "rueckruf")
        if row["email_text"]:
            text_label = tk.Label(body, text=row["email_text"].strip(), font=F_NORMAL, bg=COLOR_CARD_BG,
                     fg=COLOR_TEXT_SECONDARY, anchor="w", wraplength=400, justify="left")
            text_label.pack(fill="x", pady=(6, 0))
            self._klickbar_umschalten(text_label, row["id"], "rueckruf")

        Pill(body, "📵 Kunde nicht erreicht", bg=COLOR_SURFACE, fg="#F59E0B", outline=COLOR_BORDER,
             hover_bg=STATUS_INACTIVE_HOVER, container_bg=COLOR_CARD_BG, font=F_SMALL,
             padx=10, pady=5,
             command=lambda t=row["id"]: self._kunde_nicht_erreicht(t)).pack(anchor="w", pady=(8, 0))

        self._render_verlauf_block(body, row["id"])
        self._render_anhaenge_block(body, row["id"])
        self._render_status_block(body, row, faellig, RUECKRUF_STATUS_OPTIONS)

    def _render_verlauf_block(self, body, task_id):
        eintraege = db.get_notiz_eintraege(task_id)
        container = tk.Frame(body, bg=COLOR_CARD_BG)
        container.pack(fill="x", pady=(10, 0))

        if eintraege:
            tk.Label(container, text="Verlauf", font=F_SMALL, bg=COLOR_CARD_BG,
                     fg=COLOR_TEXT_MUTED).pack(anchor="w", pady=(0, 4))
            for eintrag in eintraege:
                zeile = tk.Frame(container, bg=COLOR_CARD_BG)
                zeile.pack(fill="x", pady=(0, 6))
                kopf = tk.Frame(zeile, bg=COLOR_CARD_BG)
                kopf.pack(fill="x")
                tk.Label(kopf, text=iso_zu_anzeige_datetime(eintrag["erstellt_am"]), font=F_SMALL,
                         bg=COLOR_CARD_BG, fg=COLOR_TEXT_MUTED).pack(side="left")
                Pill(kopf, "✕", bg=COLOR_CARD_BG, fg=COLOR_TEXT_MUTED, hover_bg=COLOR_DANGER_LIGHT,
                     hover_fg=COLOR_DANGER, container_bg=COLOR_CARD_BG, font=F_SMALL,
                     padx=5, pady=2, radius=8,
                     command=lambda e=eintrag["id"]: self._verlauf_eintrag_loeschen(e)).pack(side="right")
                tk.Label(zeile, text=eintrag["text"], font=F_NORMAL, bg=COLOR_CARD_BG,
                         fg=COLOR_TEXT_SECONDARY, anchor="w", wraplength=400, justify="left"
                         ).pack(fill="x", pady=(1, 0))

        Pill(container, "+ Notiz hinzufügen", bg=COLOR_ACCENT_LIGHT, fg=COLOR_ACCENT,
             hover_bg=COLOR_ACCENT_LIGHT_HOVER, container_bg=COLOR_CARD_BG, font=F_SMALL,
             command=lambda t=task_id: self._verlauf_eintrag_hinzufuegen(t)).pack(anchor="w", pady=(2, 0))

    def _verlauf_eintrag_hinzufuegen(self, task_id):
        def speichern(text):
            db.add_notiz_eintrag(task_id, text)
            self.refresh_all()
        VerlaufDialog(self.root, speichern)

    def _verlauf_eintrag_loeschen(self, eintrag_id):
        if messagebox.askyesno("Eintrag löschen", "Diesen Notiz-Eintrag wirklich löschen?"):
            db.delete_notiz_eintrag(eintrag_id)
            self.refresh_all()

    def _kunde_nicht_erreicht(self, task_id):
        db.add_notiz_eintrag(task_id, "☎ Kunde nicht erreicht")
        self.refresh_all()

    def _render_anhaenge_block(self, body, task_id, enable_dnd=False):
        anhaenge = db.get_attachments(task_id)
        anh_container = tk.Frame(body, bg=COLOR_CARD_BG)
        anh_container.pack(fill="x", pady=(10, 0))

        if enable_dnd and DND_AVAILABLE:
            drop_wrap = tk.Frame(anh_container, bg=COLOR_CARD_BG, height=32)
            drop_wrap.pack(fill="x", pady=(0, 6))
            drop_wrap.pack_propagate(False)
            drop_canvas = tk.Canvas(drop_wrap, bg=COLOR_CARD_BG, highlightthickness=0, bd=0)
            drop_canvas.pack(fill="both", expand=True)

            def redraw_pdf_drop(event=None, canvas=drop_canvas):
                w = canvas.winfo_width()
                h = canvas.winfo_height()
                if w > 1 and h > 1:
                    draw_dropzone(canvas, w, h)
                    canvas.create_text(w / 2, h / 2, text="📎 PDF hierher ziehen",
                                        font=F_SMALL, fill=COLOR_TEXT_MUTED)

            drop_canvas.bind("<Configure>", redraw_pdf_drop)
            drop_canvas.drop_target_register(DND_FILES)
            drop_canvas.dnd_bind("<<Drop>>", lambda e, t=task_id: self._pdf_drop(e, t))

        for anhang in anhaenge:
            zeile = tk.Frame(anh_container, bg=COLOR_CARD_BG)
            zeile.pack(fill="x", pady=(0, 3))
            Pill(zeile, f"📄 {anhang['dateiname']}", bg=STATUS_INACTIVE_BG, fg=COLOR_TEXT_PRIMARY,
                 hover_bg=STATUS_INACTIVE_HOVER, container_bg=COLOR_CARD_BG, font=F_SMALL, padx=9, pady=4,
                 command=lambda p=anhang["pfad"]: self._anhang_oeffnen(p)).pack(side="left")
            Pill(zeile, "✕", bg=COLOR_CARD_BG, fg=COLOR_TEXT_MUTED, hover_bg=COLOR_DANGER_LIGHT,
                 hover_fg=COLOR_DANGER, container_bg=COLOR_CARD_BG, font=F_SMALL,
                 padx=6, pady=3, radius=10,
                 command=lambda a=anhang["id"]: self._anhang_loeschen(a)).pack(side="left", padx=(4, 0))
        Pill(anh_container, "+ PDF anhängen", bg=COLOR_SURFACE, fg=COLOR_TEXT_SECONDARY,
             outline=COLOR_BORDER, hover_bg=STATUS_INACTIVE_HOVER, container_bg=COLOR_CARD_BG,
             font=F_SMALL, padx=9, pady=4,
             command=lambda t=task_id: self._anhang_hinzufuegen(t)).pack(anchor="w", pady=(2, 0))

        if enable_dnd and DND_AVAILABLE:
            # Auch der restliche Kartenkoerper nimmt Drops entgegen (best effort,
            # die dashed Zone oben ist das zuverlaessige Ziel).
            body.drop_target_register(DND_FILES)
            body.dnd_bind("<<Drop>>", lambda e, t=task_id: self._pdf_drop(e, t))

    def _pdf_drop(self, event, task_id):
        pfade = self.root.tk.splitlist(event.data)
        ignoriert = []
        hinzugefuegt = 0
        for pfad in pfade:
            pfad = pfad.strip("{}")
            if not os.path.isfile(pfad):
                continue
            if not pfad.lower().endswith(".pdf"):
                ignoriert.append(os.path.basename(pfad))
                continue
            try:
                db.add_attachment(task_id, pfad)
                hinzugefuegt += 1
            except Exception as e:
                messagebox.showerror("Fehler", f"Konnte Datei nicht anhängen:\n{e}")
        if ignoriert:
            messagebox.showwarning(
                "Nur PDF-Dateien",
                "Diese Datei(en) wurden ignoriert, da nur PDF unterstützt wird:\n" + "\n".join(ignoriert),
            )
        if hinzugefuegt:
            self.refresh_all()

    def _render_status_block(self, body, row, faellig, status_liste=None):
        if status_liste is None:
            status_liste = db.STATUS_OPTIONS
        zeile = None
        for i, status in enumerate(status_liste):
            if i % 4 == 0:
                zeile = tk.Frame(body, bg=COLOR_CARD_BG)
                zeile.pack(fill="x", pady=(12 if i == 0 else 4, 0))
            aktiv = row["status"] == status
            bg = STATUS_COLORS[status] if aktiv else STATUS_INACTIVE_BG
            fg = "#FFFFFF" if aktiv else STATUS_INACTIVE_FG
            hover = STATUS_COLORS[status] if aktiv else STATUS_INACTIVE_HOVER
            Pill(zeile, db.STATUS_LABELS[status], bg=bg, fg=fg, hover_bg=hover,
                 container_bg=COLOR_CARD_BG, font=F_PILL, padx=10, pady=5,
                 command=lambda t=row["id"], s=status: self._status_setzen(t, s)
                 ).pack(side="left", padx=(0, 6))

        wv_frame = tk.Frame(body, bg=COLOR_CARD_BG)
        wv_frame.pack(fill="x", pady=(10, 0))
        wv_text = "Wiedervorlage: " + (iso_zu_anzeige(row["wiedervorlage"]) if row["wiedervorlage"] else "–")
        farbe = COLOR_DANGER if faellig else COLOR_TEXT_SECONDARY
        tk.Label(wv_frame, text=wv_text, font=F_SMALL, bg=COLOR_CARD_BG, fg=farbe).pack(side="left")
        Pill(wv_frame, "Setzen", bg=COLOR_SURFACE, fg=COLOR_ACCENT, outline=COLOR_BORDER,
             hover_bg=COLOR_ACCENT_LIGHT, container_bg=COLOR_CARD_BG, font=F_SMALL,
             padx=10, pady=4,
             command=lambda r=row: self._wiedervorlage_dialog(r["id"], r["wiedervorlage"])
             ).pack(side="right")


def main():
    db.init_db()
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = TodoApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
