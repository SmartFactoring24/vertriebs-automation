import { loadConfig } from "./config.js";
import { fetchSalesRecords } from "./connectors/ki.js";
import { inspectKiDesktop, performKiLogin } from "./connectors/ki-desktop.js";
import { detectChanges } from "./detectors/changes.js";
import { exportChangesToCsv, exportSnapshotToCsv } from "./exporters/csv.js";
import { exportSnapshotToXlsx } from "./exporters/xlsx.js";
import { buildWhatsAppMessage, sendWhatsAppNotification } from "./notifiers/whatsapp-web.js";
import { waitForNextRun } from "./scheduler.js";
import { StateStore } from "./state/store.js";
import { appendLog } from "./utils/log.js";

async function runCycle() {
  const config = loadConfig();
  const stateStore = new StateStore(config.stateDirectory);
  const previousState = await stateStore.load();

  await appendLog(config.logDirectory, "Starte neuen Polling-Zyklus.");

  const records = await fetchSalesRecords(config);
  const changes = detectChanges(previousState.records, records).filter(
    (change) => !previousState.sentEventIds.includes(change.eventId)
  );

  await exportSnapshotToCsv(records, config.exportDirectory);
  await exportSnapshotToXlsx(records, config.exportDirectory);
  await exportChangesToCsv(changes, config.exportDirectory);

  if (changes.length > 0) {
    const message = buildWhatsAppMessage(changes);
    await sendWhatsAppNotification(config, message);
    await appendLog(config.logDirectory, `${changes.length} Aenderungen an WhatsApp uebergeben.`);
  } else {
    await appendLog(config.logDirectory, "Keine neuen relevanten Aenderungen erkannt.");
  }

  await stateStore.save({
    records,
    sentEventIds: [...new Set([...previousState.sentEventIds, ...changes.map((change) => change.eventId)])],
    updatedAt: new Date().toISOString()
  });
}

async function main() {
  const config = loadConfig();
  const inspectKi = process.argv.includes("--inspect-ki");
  const loginKi = process.argv.includes("--login-ki");
  const runOnce = process.argv.includes("--once");

  if (inspectKi) {
    const state = await inspectKiDesktop(config);
    console.log(JSON.stringify(state, null, 2));
    return;
  }

  if (loginKi) {
    const state = await performKiLogin(config);
    console.log(JSON.stringify(state, null, 2));
    return;
  }

  do {
    await runCycle();

    if (runOnce) {
      break;
    }

    await waitForNextRun(config);
  } while (true);
}

main().catch(async (error) => {
  const config = loadConfig();
  await appendLog(config.logDirectory, `Fehler: ${error instanceof Error ? error.stack ?? error.message : String(error)}`);
  process.exitCode = 1;
});
