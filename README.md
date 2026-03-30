# Sales Automation MVP

Dieses Verzeichnis enthaelt das Grundgeruest fuer eine lokale Windows-Automatisierung mit:

- `KI` als Hauptdatenquelle
- lokaler Zustandsspeicherung
- `CSV`- und optionalem `XLSX`-Export
- Benachrichtigung ueber `WhatsApp Web`

Der Fokus dieser ersten Version liegt auf einer robusten Struktur, nicht auf fertigen Selektoren fuer eure echte `KI`-Oberflaeche oder WhatsApp-Gruppe.

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
2. Browser-Profile fuer `KI` und `WhatsApp Web` vorbereiten.
3. Selektoren und Navigationslogik in `src/connectors/ki.ts` und `src/notifiers/whatsapp-web.ts` auf eure echte Oberflaeche anpassen.
4. Abhaengigkeiten installieren:

```powershell
npm install
npx playwright install chromium
```

5. Zum lokalen Start:

```powershell
npm run dev
```

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

