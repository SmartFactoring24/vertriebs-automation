import type { AppConfig } from "../config.js";
import type { SalesRecord } from "../state/types.js";
import { ensureKiDesktopReady, performKiLogin } from "./ki-desktop.js";

export async function fetchSalesRecords(config: AppConfig): Promise<SalesRecord[]> {
  const state = await ensureKiDesktopReady(config);

  if (state.stage === "login") {
    await performKiLogin(config);
  }

  return collectSalesRecords(config);
}

async function collectSalesRecords(config: AppConfig): Promise<SalesRecord[]> {
  // Der installierte Smartclient ist nach den lokalen Artefakten sehr wahrscheinlich
  // ein Java/Swing-Client mit eingebettetem Chromium ueber JxBrowser.
  // Der echte KI-Connector wird deshalb nativ am Desktop arbeiten:
  // 1. smartclient.exe starten
  // 2. Login-Fenster erkennen
  // 3. Benutzerkennung/Passwort eintragen
  // 4. auf 2FA-Freigabe und das Hauptfenster warten
  // 5. danach Daten im Hauptfenster extrahieren
  void config;

  return [
    {
      businessId: "A12345",
      customerName: "Max Mustermann",
      productName: "XYZ Schutzbrief",
      status: "eingereicht",
      salesValue: 1234.56,
      submittedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      source: "KI"
    }
  ];
}
