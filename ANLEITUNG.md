# Werkstatt To-Do – Aufgaben & Rückrufe

Lokales Programm, läuft ohne Internet und ohne laufende Kosten (kein API-Key nötig).

## Installation (einmalig)

Voraussetzung: Python 3 ist installiert (auf dem Mac Mini i.d.R. schon vorhanden,
unter Windows von python.org, Haken bei "Add to PATH" setzen).

Terminal / Eingabeaufforderung im Programmordner öffnen und:

```
pip install tkinterdnd2 extract-msg
```

## Start

```
python app.py
```

(unter manchen Windows-Installationen heißt der Befehl `python3 app.py` oder `py app.py`)

Es öffnet sich ein Fenster mit zwei Spalten: **Aufgaben** und **Rückrufe**.
Alle Daten werden in der Datei `todo.db` im selben Ordner gespeichert –
beim nächsten Start sind sie automatisch wieder da. Diese Datei einfach
mitsichern/mitnehmen, wenn du auf einen anderen Rechner wechselst.

## Bedienung

**Neue Aufgabe / neuer Rückruf:** Button oben in der jeweiligen Spalte.

**Status setzen:** Auf einer Karte auf einen der vier Buttons klicken –
Offen / Erledigt / Teilerledigt / Warte auf Teile. Der aktive Status ist
farbig hervorgehoben.

**Wiedervorlage:** Über den "Setzen"-Button ein Datum eintragen (oder eine
der Schnellwahl-Optionen Heute/Morgen/+3 Tage/Nächste Woche nutzen). Ist
das Datum erreicht oder überschritten, wird die Karte orange hervorgehoben
und automatisch nach oben sortiert – solange der Status nicht "Erledigt" ist.

**Erledigte Einträge:** Werden standardmäßig ausgeblendet. Häkchen
"Erledigte anzeigen" oben rechts in der Spalte zeigt sie wieder an.

**E-Mail zu einem Rückruf hinterlegen – direkt aus Outlook (empfohlen):**

Echtes Drag & Drop einer Mail direkt aus der Outlook-Liste funktioniert aus
technischen Gründen nicht zuverlässig (Windows zeigt dabei das
durchgestrichene Verbotssymbol, weil Outlook beim Ziehen kein normales
Dateiformat anbietet). Stattdessen gibt es eine direkte Anbindung ohne
Zwischenspeichern:

1. Einmalig installieren: `pip install pywin32`
2. Outlook öffnen und die gewünschte Mail in der Liste **markieren**
   (anklicken reicht, nicht öffnen)
3. Im Programm auf **"📥 Markierte Outlook-Mail übernehmen"** klicken
   (Button erscheint nur unter Windows, wenn pywin32 installiert ist)
4. Betreff, Absender, Datum und Text werden automatisch übernommen –
   mehrere markierte Mails auf einmal funktionieren auch.

Voraussetzung: die klassische Outlook-Desktop-App muss geöffnet sein
(kein Web-Outlook, kein "neues Outlook" ohne klassische Ansicht).

**Alternative über Datei:**

Falls die Mail bereits als `.msg` oder `.eml` gespeichert ist (z.B. per
"Nachricht speichern unter" in Outlook, oder aus Thunderbird), entweder in
die Ablagefläche "Bereits gespeicherte .msg/.eml hier ablegen" ziehen oder
über "Datei auswählen" öffnen.

Über "E-Mail vollständig anzeigen" lässt sich der komplette Inhalt einer
hinterlegten Mail jederzeit wieder öffnen.

## PDF-Anhänge

Über "+ PDF anhängen" auf jeder Karte lassen sich beliebige PDFs (z.B.
Rechnungen, Kostenvoranschläge) hinterlegen. Sie werden in einen Ordner
`anhaenge/<Nummer>/` neben `todo.db` kopiert. Ein Klick auf den Dateinamen
öffnet die PDF mit dem Standardprogramm, "✕" daneben entfernt den Anhang
wieder.

## Anzeigen-Filter

Über die Leiste "Anzeigen: Alle / Offen / Teilerledigt / Warte auf Teile"
lässt sich jede Spalte auf einen Status eingrenzen.
