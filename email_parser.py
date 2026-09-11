"""
email_parser.py – Liest .msg (Outlook) und .eml Dateien aus und liefert
Betreff, Absender, Datum und Text zurück.
"""
import os
import re
import email
from email import policy
from email.utils import parsedate_to_datetime


def text_bereinigen(text):
    """Vereinheitlicht Zeilenumbrüche, entfernt Leerzeichen am Zeilenende
    und reduziert mehrfache Leerzeilen auf maximal eine – Outlook/HTML-Mails
    erzeugen beim Umwandeln in Text oft sehr viele davon."""
    if not text:
        return text
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    zeilen = [zeile.rstrip() for zeile in text.split("\n")]
    text = "\n".join(zeilen)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def parse_email_file(path):
    """
    Gibt ein dict zurueck: {betreff, absender, datum, text}
    datum ist ein String im Format DD.MM.YYYY HH:MM oder "" wenn unbekannt.
    Wirft eine ValueError mit verstaendlicher Meldung bei Problemen.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".msg":
        return _parse_msg(path)
    elif ext == ".eml":
        return _parse_eml(path)
    else:
        raise ValueError(
            f"Nicht unterstuetztes Dateiformat '{ext}'. "
            f"Bitte eine .msg (aus Outlook) oder .eml Datei verwenden."
        )


def _parse_msg(path):
    try:
        import extract_msg
    except ImportError:
        raise ValueError(
            "Das Paket 'extract-msg' fehlt. Bitte installieren mit:\n"
            "pip install extract-msg"
        )
    try:
        msg = extract_msg.Message(path)
    except Exception as e:
        raise ValueError(f"Konnte .msg Datei nicht lesen: {e}")

    betreff = msg.subject or "(kein Betreff)"
    absender = msg.sender or "(unbekannter Absender)"
    text = msg.body or ""
    datum = ""
    try:
        if msg.date:
            # msg.date ist meist bereits ein datetime oder ein String
            dt = msg.date
            if isinstance(dt, str):
                datum = dt
            else:
                datum = dt.strftime("%d.%m.%Y %H:%M")
    except Exception:
        datum = ""
    msg.close()
    return {"betreff": betreff, "absender": absender, "datum": datum, "text": text_bereinigen(text)}


def _parse_eml(path):
    try:
        with open(path, "rb") as f:
            msg = email.message_from_binary_file(f, policy=policy.default)
    except Exception as e:
        raise ValueError(f"Konnte .eml Datei nicht lesen: {e}")

    betreff = msg.get("subject", "(kein Betreff)")
    absender = msg.get("from", "(unbekannter Absender)")
    datum_raw = msg.get("date", "")
    datum = ""
    if datum_raw:
        try:
            dt = parsedate_to_datetime(datum_raw)
            datum = dt.strftime("%d.%m.%Y %H:%M")
        except Exception:
            datum = datum_raw

    text = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                try:
                    text = part.get_content()
                except Exception:
                    pass
                break
        if not text:
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    try:
                        text = part.get_content()
                    except Exception:
                        pass
                    break
    else:
        try:
            text = msg.get_content()
        except Exception:
            text = ""

    return {"betreff": betreff, "absender": absender, "datum": datum, "text": text_bereinigen(text)}
