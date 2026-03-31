# Sales Automation MVP

Dieses Verzeichnis enthält das Grundgerüst für eine lokale Windows-Automatisierung mit:

- `KI` als Hauptdatenquelle
- lokaler Zustandsspeicherung
- `CSV`- und optionalem `XLSX`-Export
- Benachrichtigung über `WhatsApp Web`

Der Fokus dieser ersten Version liegt auf einer robusten Struktur, nicht auf fertigen Selektoren für eure echte `KI`-Oberfläche oder WhatsApp-Gruppe.

## Aktueller Stand

Der Desktop-Loginpfad für den DVAG-Client ist auf dem Entwicklungsgerät bereits erfolgreich automatisiert bis zum geöffneten `KI`-Hauptfenster. Die aktuelle Übergabe ist dokumentiert in [docs/handoff-2026-03-31.md](./docs/handoff-2026-03-31.md).

## Zielbild

1. Der Bot startet lokal auf einem Arbeitsrechner.
2. Er prüft, ob `KI` und `WhatsApp Web` im erwarteten Zustand sind.
3. Er liest aktuelle Datensätze aus `KI`.
4. Er erkennt Änderungen gegenüber dem letzten bekannten Stand.
5. Er exportiert neue Daten in `CSV` oder `XLSX`.
6. Er sendet bei relevanten Änderungen eine Nachricht an die definierte WhatsApp-Gruppe.

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
2. `.env` mit `KI_APP_PATH`, Login-Daten und WhatsApp-Gruppenname füllen.
3. Optional Fenster-Titelhinweise für `KI_LOGIN_WINDOW_TITLE_HINT` und `KI_MAIN_WINDOW_TITLE_HINT` setzen, sobald ihr die echten Titel kennt.
4. Die native Login- und Fensterlogik in `src/connectors/ki.ts` und `src/connectors/ki-desktop.ts` sowie die WhatsApp-Navigation in `src/notifiers/whatsapp-web.ts` an eure echte Oberfläche anpassen.
5. Abhängigkeiten installieren:

```powershell
npm install
npx playwright install chromium
```

6. Zum lokalen Start:

```powershell
npm run dev
```

Für den reinen Desktop-Login-Test bis zum geöffneten `KI`-Fenster:

```powershell
npm run dev -- --login-ki
```

Für einen gezielten Test des Crash- und Recovery-Ablaufs:

```powershell
npm run dev -- --test-recovery
```

Dieser Testmodus simuliert absichtlich einen ersten Fehler, den Neustart-Dialog und danach einen zweiten Fehler mit abschließendem Support-Hinweis. So lässt sich das Troubleshooting gezielt prüfen, ohne die `.env` absichtlich verfälschen zu müssen.

Alternativ gibt es einen einfachen Windows-Launcher:

- [launcher/StartKiAutomation.exe](./launcher/StartKiAutomation.exe)

Der Launcher startet intern denselben Flow wie `npm run dev -- --login-ki`.

Er verwendet bevorzugt eine global installierte `Node`/`npm`-Version und fällt nur dann auf die portable Repo-Version unter `tools/` zurück.

Zusätzlich gibt es einen vereinfachten Troubleshooting-Einstieg direkt im Launcher:

- Während des Startfensters kann innerhalb von 4 Sekunden `F4` gedrückt werden.
- Dadurch öffnet sich ein kleines Debug- und Troubleshooting-Menü in PowerShell.
- Mit `1` wird der Recovery- und Crash-Test gestartet.
- Mit `2` wird eine KI-Statusdiagnose ausgegeben.
- Mit `3` werden alle zugehörigen KI-Prozesse sauber beendet.
- Mit `0` wird der normale Start fortgesetzt.

Wichtige Hinweise für Nutzer:

- Für den laufenden Automationsbetrieb muss das PowerShell-Fenster geöffnet bleiben.
- Die DVAG-2FA muss während des Anmeldevorgangs manuell freigegeben werden.
- Das Freigabegerät sollte deshalb beim Start des Bots unmittelbar bereitliegen.
- Crash- und Diagnoseprotokolle werden lokal unter `./data/logs` gespeichert.

## Was bereits vorbereitet ist

- saubere Modultrennung
- typisierte Datensätze
- persistenter State
- Delta-Erkennung
- `CSV`-Export
- Platz für `XLSX`-Export
- WhatsApp-Nachrichtenformatierung
- Logging und Screenshot-Pfade

## Was ihr noch konkretisieren müsst

- welche `KI`-Ansicht die Live-Daten enthält
- welche Felder zwingend benötigt werden
- wie ein "relevantes Ereignis" fachlich definiert ist
- wie eure WhatsApp-Gruppe eindeutig validiert werden soll
- ob `CSV` reicht oder `XLSX` Pflicht ist

## API oder WhatsApp Web

Für euer aktuelles Ziel `interne Gruppe informieren` ist diese Struktur bewusst auf `WhatsApp Web` ausgelegt.

Wenn ihr später auf die offizielle `WhatsApp Business Platform` wechselt, müsst ihr vor allem das Notifier-Modul austauschen. Die Module für `KI`, State, Export und Delta-Erkennung können bestehen bleiben.
