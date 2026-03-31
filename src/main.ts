import { loadConfig } from "./config.js";
import { fetchSalesRecords } from "./connectors/ki.js";
import { forceCloseKiProcesses, inspectKiDesktop, performKiLogin } from "./connectors/ki-desktop.js";
import { detectChanges } from "./detectors/changes.js";
import { exportChangesToCsv, exportSnapshotToCsv } from "./exporters/csv.js";
import { exportSnapshotToXlsx } from "./exporters/xlsx.js";
import { buildWhatsAppMessage, sendWhatsAppNotification } from "./notifiers/whatsapp-web.js";
import { waitForNextRun } from "./scheduler.js";
import { StateStore } from "./state/store.js";
import { appendLog, writeCrashLog } from "./utils/log.js";
import { showInfoDialog, showRecoveryPrompt } from "./utils/windows.js";

function printStartupNotice(options: {
  loginKi: boolean;
  runOnce: boolean;
  pollIntervalMinutes: number;
  modeLabel?: string;
  extraLines?: string[];
}) {
  const modeLabel = options.modeLabel ?? (options.loginKi
    ? "Login-Testlauf"
    : options.runOnce
      ? "Einmallauf"
      : "Automationsbetrieb");

  console.log("============================================================");
  console.log("Vertriebs-Automation");
  console.log(`Modus: ${modeLabel}`);
  console.log("");
  console.log("Hinweis:");
  console.log("- Für den laufenden Automationsbetrieb muss das PowerShell-Fenster geöffnet bleiben.");
  console.log(
    options.loginKi || options.runOnce
      ? "- Dieser Lauf endet nach Abschluss automatisch."
      : `- Neue Daten werden in regelmäßigen Intervallen geprüft, aktuell alle ${options.pollIntervalMinutes} Minute(n).`
  );
  console.log("- Crash- und Diagnoseprotokolle werden im Ordner ./data/logs gespeichert.");
  console.log("- Die DVAG-2FA muss während des Anmeldevorgangs manuell freigegeben werden.");
  console.log("- Bitte halten Sie Ihr Freigabegerät bereit, damit der Bot den Ablauf fortsetzen kann.");
  for (const line of options.extraLines ?? []) {
    console.log(line);
  }
  console.log("");
  console.log("Firmeninterne Nutzung:");
  console.log("- Dieses Programm wurde von Leo Mitteneder und Moritz Rolle mit Unterstützung durch Codex erstellt.");
  console.log("- Es ist ausschließlich für die firmeninterne Nutzung bestimmt.");
  console.log("- Eine Vervielfältigung oder Weitergabe ist nicht zulässig.");
  console.log("============================================================");
  console.log("");
}

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
    await appendLog(config.logDirectory, `${changes.length} Änderungen an WhatsApp übergeben.`);
  } else {
    await appendLog(config.logDirectory, "Keine neuen relevanten Änderungen erkannt.");
  }

  await stateStore.save({
    records,
    sentEventIds: [...new Set([...previousState.sentEventIds, ...changes.map((change) => change.eventId)])],
    updatedAt: new Date().toISOString()
  });
}

async function captureKiDiagnostics(config: ReturnType<typeof loadConfig>): Promise<Record<string, unknown>> {
  try {
    const state = await inspectKiDesktop(config);
    return { kiState: state };
  } catch (diagnosticError) {
    return {
      kiDiagnosticsError: diagnosticError instanceof Error ? diagnosticError.message : String(diagnosticError)
    };
  }
}

async function handleCrash(options: {
  config: ReturnType<typeof loadConfig>;
  scope: string;
  error: unknown;
  recoveryAlreadyAttempted: boolean;
}): Promise<"retry" | "abort"> {
  const diagnostics = await captureKiDiagnostics(options.config);
  const crashLogPath = await writeCrashLog(options.config.logDirectory, {
    scope: options.scope,
    error: options.error,
    diagnostics
  });
  const errorMessage = options.error instanceof Error ? options.error.message : String(options.error);

  await appendLog(options.config.logDirectory, `Crash-Log geschrieben: ${crashLogPath}`);

  if (options.recoveryAlreadyAttempted) {
    await appendLog(options.config.logDirectory, "Zweiter Fehler nach Recovery-Versuch erkannt. Anwendung wird beendet.");
    await showInfoDialog({
      title: "Vertriebs-Automation",
      message:
        "Nach dem automatischen Neustartversuch ist erneut ein Fehler aufgetreten.\n\n" +
        `Für die Fehleranalyse wurde ein Crash-Log gespeichert:\n${crashLogPath}\n\n` +
        "Bitte wenden Sie sich für das Troubleshooting direkt an den technischen Support:\n" +
        "Moritz Rolle\n" +
        "E-Mail: moritz.rolle.assistent@dvag.de\n" +
        "Telefon: 01737552159\n\n" +
        "Die Anwendung wird jetzt zusammen mit allen zugehörigen KI-Prozessen beendet."
    });
    await forceCloseKiProcesses(options.config);
    return "abort";
  }

  const shouldRetry = await showRecoveryPrompt({
    title: "Vertriebs-Automation",
    message:
      "Beim Zugriff auf KI ist ein Fehler aufgetreten.\n\n" +
      `Fehlermeldung: ${errorMessage}\n\n` +
      `Ein Diagnoseprotokoll wurde erstellt:\n${crashLogPath}\n\n` +
      "Soll die Anwendung alle zugehörigen KI-Prozesse beenden und anschließend einen automatischen Neustartversuch durchführen?"
  });

  if (!shouldRetry) {
    return "abort";
  }

  await appendLog(options.config.logDirectory, "Automatischer Recovery-Neustart nach Benutzerbestätigung wird ausgeführt.");
  await forceCloseKiProcesses(options.config);
  await new Promise((resolve) => setTimeout(resolve, 1500));
  return "retry";
}

async function runRecoveryTestMode(config: ReturnType<typeof loadConfig>): Promise<void> {
  printStartupNotice({
    loginKi: false,
    runOnce: true,
    pollIntervalMinutes: config.pollIntervalMinutes,
    modeLabel: "Recovery-Testmodus",
    extraLines: [
      "- Dieser Modus simuliert absichtlich zwei aufeinanderfolgende KI-Fehler.",
      "- So können Crash-Logs, Dialogfenster und der Recovery-Ablauf gezielt geprüft werden."
    ]
  });

  await appendLog(config.logDirectory, "Recovery-Testmodus wurde gestartet.");

  let recoveryAttempted = false;

  while (true) {
    const simulatedError = recoveryAttempted
      ? new Error("Simulierter Folgefehler nach automatischem Neustartversuch im Recovery-Testmodus.")
      : new Error("Simulierter Erstfehler im Recovery-Testmodus.");

    const action = await handleCrash({
      config,
      scope: recoveryAttempted ? "recovery-test-second-failure" : "recovery-test-first-failure",
      error: simulatedError,
      recoveryAlreadyAttempted: recoveryAttempted
    });

    if (action === "abort") {
      await appendLog(config.logDirectory, "Recovery-Testmodus wurde beendet.");
      return;
    }

    recoveryAttempted = true;
  }
}

async function runCloseKiMode(config: ReturnType<typeof loadConfig>): Promise<void> {
  await appendLog(config.logDirectory, "Manueller KI-Prozessstopp wurde gestartet.");
  await forceCloseKiProcesses(config);
  console.log("Alle zugehörigen KI-Prozesse wurden beendet.");
}

async function main() {
  const config = loadConfig();
  const inspectKi = process.argv.includes("--inspect-ki");
  const loginKi = process.argv.includes("--login-ki");
  const runOnce = process.argv.includes("--once");
  const testRecovery = process.argv.includes("--test-recovery");
  const closeKi = process.argv.includes("--close-ki");

  if (inspectKi) {
    const state = await inspectKiDesktop(config);
    console.log(JSON.stringify(state, null, 2));
    return;
  }

  if (closeKi) {
    await runCloseKiMode(config);
    return;
  }

  if (testRecovery) {
    await runRecoveryTestMode(config);
    return;
  }

  if (loginKi) {
    printStartupNotice({ loginKi: true, runOnce, pollIntervalMinutes: config.pollIntervalMinutes });
    let recoveryAttempted = false;

    while (true) {
      try {
        const state = await performKiLogin(config);
        console.log(JSON.stringify(state, null, 2));
        return;
      } catch (error) {
        const action = await handleCrash({
          config,
          scope: "login-ki",
          error,
          recoveryAlreadyAttempted: recoveryAttempted
        });
        if (action === "abort") {
          throw error;
        }
        recoveryAttempted = true;
      }
    }
  }

  printStartupNotice({ loginKi: false, runOnce, pollIntervalMinutes: config.pollIntervalMinutes });
  let recoveryAttempted = false;

  do {
    try {
      await runCycle();
      recoveryAttempted = false;
    } catch (error) {
      const action = await handleCrash({
        config,
        scope: "automation-cycle",
        error,
        recoveryAlreadyAttempted: recoveryAttempted
      });
      if (action === "abort") {
        throw error;
      }
      recoveryAttempted = true;
      continue;
    }

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
