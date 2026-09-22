# Werkstatt Todo

Dunkles Windows-Programm für Aufgaben und Rückrufe.

## Funktionen
- Aufgaben und Rückrufe in zwei Spalten
- Einheitliche kompakte Karten
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

## GitHub EXE
Das Projekt wird über GitHub Actions als Windows-EXE gebaut.

Workflow: **Actions → Windows EXE bauen → Run workflow**.

## Neue Funktionen (September 2026)
- Ausklappbare und direkt bearbeitbare Verkaufsliste auf Basis der mitgelieferten Verkaufsliste.
- Automatische JSON-Backups im Windows-Electron-Benutzerordner, zusätzlich alle 5 Minuten.
- Backup-Wiederherstellung und Öffnen des Backup-Ordners über die Oberfläche.
- Testbutton für Windows-Benachrichtigungen.
- Wiedervorlagen werden alle 15 Sekunden geprüft; fällige Einträge lösen eine Windows-Benachrichtigung aus.
