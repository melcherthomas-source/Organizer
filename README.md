# Werkstatt Todo

Dunkles Windows-Programm für Aufgaben und Rückrufe.

## Funktionen
- Aufgaben und Rückrufe in zwei Spalten
- Schnelleingang für neue Vorgänge
- Parkplatz für lose Notizen / Dinge ohne aktive Aufgabe
- Kunden-/Fahrzeugakte mit Suche über Name, Kennzeichen, Telefon, E-Mail und FIN
- Kennzeichensuche unabhängig von Bindestrichen und Leerzeichen
- Name, Telefonnummer, E-Mail
- Fahrzeug: Kennzeichen, Modell, FIN
- Wiedervorlage: Morgen, In 3 Tagen oder benutzerdefiniertes Datum
- Notizen mit Zeitstempel
- PDF/Dateianhänge
- Outlook-Mailimport (.msg/.eml/.oft) im Electron-Programm
- automatische Übernahme von E-Mail und Telefonnummer
- automatische Zuordnung von Fahrzeugdaten über Kennzeichen
- Suche über alle Felder
- Erledigte Einträge landen im Archiv
- Archiv ist separat aufrufbar, durchsuchbar und öffnet die vollständigen Einträge
- Archivierte Einträge können wiederhergestellt werden

## Daten
Die laufenden Aufgaben werden wie bisher lokal im Electron-Programm gespeichert. Die vorhandene `todo.db` wird von dieser Oberfläche nicht als Aufgabenquelle verwendet und wird deshalb beim Paketieren nicht benötigt.

## Windows-EXE
Das Projekt kann über GitHub Actions als Windows-EXE gebaut werden.

Workflow: **Actions → Windows EXE bauen → Run workflow**.
