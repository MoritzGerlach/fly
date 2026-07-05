# Flight Duel — Setup

## Dateien

- `flight_duel_dashboard.html` — diese Datei als `index.html` in dein GitHub-Pages-Repository hochladen.
- `FlightDuel_Code.gs` — einmalig in Google Apps Script einfügen und als Web App deployen.
- `flight_log_template.xlsx` — optionale Excel-Vorlage/Backup-Struktur mit denselben Spalten.

## GitHub Pages

1. Neues Repository erstellen, z. B. `flight-duel`.
2. `flight_duel_dashboard.html` in `index.html` umbenennen.
3. Datei in das Repository hochladen.
4. In GitHub: Settings → Pages → Deploy from branch → `main` / root.
5. Die veröffentlichte GitHub-Pages-URL öffnen.

## Google Sheet Backend

1. Google Drive öffnen → New → Google Apps Script.
2. Inhalt von `FlightDuel_Code.gs` in `Code.gs` kopieren.
3. Deploy → New deployment → Type: Web app.
4. Execute as: `Me`.
5. Who has access: `Anyone`.
6. Deploy klicken und die Web-App-URL kopieren.

## Dashboard verbinden

1. Dashboard öffnen.
2. Setup öffnen.
3. Web-App-URL einfügen.
4. Einen Setup-PIN vergeben, z. B. 4–8 Zeichen.
5. AirLabs API-Key einfügen.
6. Backend einrichten klicken.

Danach können Moritz und Farid dieselbe GitHub-Seite öffnen und Flüge speichern. Die Daten landen im Google Sheet.

## Nutzung

- Person wählen: Moritz oder Farid.
- Flugnummer und Datum eingeben.
- `Live-Daten holen & speichern` klicken.
- Falls ein Flug nicht gefunden wird: `Manuell eintragen` nutzen.
- CSV Export erstellt eine lokale Sicherung.

## Hinweis

Die GitHub-Seite selbst ist statisch. Das gemeinsame Speichern funktioniert über Google Apps Script + Google Sheet. Ohne Backend speichert die Seite nur lokal im Browser.
