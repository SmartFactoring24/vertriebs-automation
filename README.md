# Sales Automation MVP

Dieses Verzeichnis enthaelt das Grundgeruest fuer eine lokale Windows-Automatisierung mit:

- `KI` als Hauptdatenquelle
- lokaler Zustandsspeicherung
- `CSV`- und optionalem `XLSX`-Export
- Benachrichtigung ueber `WhatsApp Web`

Der Fokus dieser ersten Version liegt auf einer robusten Struktur, nicht auf fertigen Selektoren fuer eure echte `KI`-Oberflaeche oder WhatsApp-Gruppe.

## Aktueller Stand

Der Desktop-Loginpfad fuer den DVAG-Client ist auf dem Entwicklungsgeraet bereits erfolgreich automatisiert bis zum geoeffneten `KI`-Hauptfenster. Die aktuelle Uebergabe ist dokumentiert in [docs/handoff-2026-03-31.md](./docs/handoff-2026-03-31.md).

## Zielbild

1. Der Bot startet lokal auf einem Arbeitsrechner.
2. Er prueft, ob `KI` und `WhatsApp Web` im erwarteten Zustand sind.
3. Er liest aktuelle Datensaetze aus `KI`.
4. Er erkennt Aenderungen gegenueber dem letzten bekannten Stand.
5. Er exportiert neue Daten in `CSV` oder `XLSX`.
6. Er sendet bei relevanten Aenderungen eine Nachricht an die definierte WhatsApp-Gruppe.

## Projektstruktur

```text
vertriebs-automation/
  docs/
    architecture.md
  data/
    exports/
    logs/
    screenshots/
    state/
  src/
    main.ts
    config.ts
    scheduler.ts
    state/
    connectors/
    detectors/
    exporters/
    notifiers/
    utils/
  package.json
  tsconfig.json
  .env.example
```

## Schnellstart

1. `.env.example` nach `.env` kopieren und Werte anpassen.
2. `.env` mit `KI_APP_PATH`, Login-Daten und WhatsApp-Gruppenname fuellen.
3. Optional Fenster-Titelhinweise fuer `KI_LOGIN_WINDOW_TITLE_HINT` und `KI_MAIN_WINDOW_TITLE_HINT` setzen, sobald ihr die echten Titel kennt.
4. Die native Login- und Fensterlogik in `src/connectors/ki.ts` und `src/connectors/ki-desktop.ts` sowie die WhatsApp-Navigation in `src/notifiers/whatsapp-web.ts` an eure echte Oberflaeche anpassen.
5. Abhaengigkeiten installieren:

```powershell
npm install
npx playwright install chromium
```

6. Zum lokalen Start:

```powershell
npm run dev
```

Fuer den reinen Desktop-Login-Test bis zum geoeffneten `KI`-Fenster:

```powershell
npm run dev -- --login-ki
```

Alternativ gibt es einen einfachen Windows-Launcher:

- [launcher/StartKiAutomation.exe](./launcher/StartKiAutomation.exe)

Der Launcher startet intern denselben Flow wie `npm run dev -- --login-ki`.

Er verwendet bevorzugt eine global installierte `Node`/`npm`-Version und faellt nur dann auf die portable Repo-Version unter `tools/` zurueck.

## Was bereits vorbereitet ist

- saubere Modultrennung
- typisierte Datensaetze
- persistenter State
- Delta-Erkennung
- `CSV`-Export
- Platz fuer `XLSX`-Export
- WhatsApp-Nachrichtenformatierung
- Logging und Screenshot-Pfade

## Was ihr noch konkretisieren muesst

- welche `KI`-Ansicht die Live-Daten enthaelt
- welche Felder zwingend benoetigt werden
- wie ein "relevantes Ereignis" fachlich definiert ist
- wie eure WhatsApp-Gruppe eindeutig validiert werden soll
- ob `CSV` reicht oder `XLSX` Pflicht ist

## API oder WhatsApp Web

Fuer euer aktuelles Ziel `interne Gruppe informieren` ist diese Struktur bewusst auf `WhatsApp Web` ausgelegt.

Wenn ihr spaeter auf die offizielle `WhatsApp Business Platform` wechselt, muesst ihr vor allem das Notifier-Modul austauschen. Die Module fuer `KI`, State, Export und Delta-Erkennung koennen bestehen bleiben.
