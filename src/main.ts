import { loadConfig } from "./config.js";
import { fetchSalesRecords } from "./connectors/ki.js";
import {
  captureCurrentKiFullWindowRegion,
  captureCurrentKiHeaderRegion,
  captureCurrentKiTableRegion,
  captureCurrentKiTreeRegion,
  forceCloseKiProcesses,
  inspectKiDesktop,
  navigateKiToSubmittedUnits,
  performKiLogin
} from "./connectors/ki-desktop.js";
import { detectChanges } from "./detectors/changes.js";
import { exportChangesToCsv, exportSnapshotToCsv } from "./exporters/csv.js";
import { exportSnapshotToXlsx } from "./exporters/xlsx.js";
import { buildWhatsAppMessage, sendWhatsAppNotification } from "./notifiers/whatsapp-web.js";
import { waitForNextRun } from "./scheduler.js";
import { StateStore } from "./state/store.js";
import { appendLog, writeCrashLog } from "./utils/log.js";
import { showInfoDialog, showRecoveryPrompt } from "./utils/windows.js";
import { analyzeHeaderCapture, analyzeTreeCapture } from "./vision/analyze.js";
import { readTableCaptureWithOcr } from "./vision/ocr.js";
import { captureVisionTemplates } from "./vision/templates.js";
import { matchTemplateInImage } from "./vision/match.js";

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
  console.log("- FÃ¼r den laufenden Automationsbetrieb muss das PowerShell-Fenster geÃ¶ffnet bleiben.");
  console.log(
    options.loginKi || options.runOnce
      ? "- Dieser Lauf endet nach Abschluss automatisch."
      : `- Neue Daten werden in regelmÃ¤ÃŸigen Intervallen geprÃ¼ft, aktuell alle ${options.pollIntervalMinutes} Minute(n).`
  );
  console.log("- Crash- und Diagnoseprotokolle werden im Ordner ./data/logs gespeichert.");
  console.log("- Die DVAG-2FA muss wÃ¤hrend des Anmeldevorgangs manuell freigegeben werden.");
  console.log("- Bitte halten Sie Ihr FreigabegerÃ¤t bereit, damit der Bot den Ablauf fortsetzen kann.");
  for (const line of options.extraLines ?? []) {
    console.log(line);
  }
  console.log("");
  console.log("Firmeninterne Nutzung:");
  console.log("- Dieses Programm wurde von Leo Mitteneder und Moritz Rolle mit UnterstÃ¼tzung durch Codex erstellt.");
  console.log("- Es ist ausschlieÃŸlich fÃ¼r die firmeninterne Nutzung bestimmt.");
  console.log("- Eine VervielfÃ¤ltigung oder Weitergabe ist nicht zulÃ¤ssig.");
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

  if (changes.length === 0) {
    await appendLog(config.logDirectory, "Keine neuen relevanten Änderungen erkannt.");
    return;
  }

  await exportSnapshotToCsv(records, config.exportDirectory);
  await exportSnapshotToXlsx(records, config.exportDirectory);
  await exportChangesToCsv(changes, config.exportDirectory);

  if (config.whatsappNotificationsEnabled) {
    const message = buildWhatsAppMessage(changes);
    await sendWhatsAppNotification(config, message);
    await appendLog(config.logDirectory, `${changes.length} Änderungen an WhatsApp übergeben.`);
  } else {
    await appendLog(config.logDirectory, `${changes.length} Änderungen erkannt, WhatsApp-Versand ist deaktiviert.`);
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
        `FÃ¼r die Fehleranalyse wurde ein Crash-Log gespeichert:\n${crashLogPath}\n\n` +
        "Bitte wenden Sie sich fÃ¼r das Troubleshooting direkt an den technischen Support:\n" +
        "Moritz Rolle\n" +
        "E-Mail: moritz.rolle.assistent@dvag.de\n" +
        "Telefon: 01737552159\n\n" +
        "Die Anwendung wird jetzt zusammen mit allen zugehÃ¶rigen KI-Prozessen beendet."
    });
    await forceCloseKiProcesses(options.config);
    return "abort";
  }

  const recoveryDecision = await showRecoveryPrompt({
    title: "Vertriebs-Automation",
    message:
      "Beim Zugriff auf KI ist ein Fehler aufgetreten.\n\n" +
      `Fehlermeldung: ${errorMessage}\n\n` +
      `Ein Diagnoseprotokoll wurde erstellt:\n${crashLogPath}\n\n` +
      "Wie soll die Anwendung jetzt reagieren?"
  });

  if (recoveryDecision === "abort_and_close_ki") {
    await appendLog(options.config.logDirectory, "Benutzer hat 'Bot und KI beenden' im Recovery-Dialog gewÃ¤hlt.");
    await forceCloseKiProcesses(options.config);
    return "abort";
  }

  if (recoveryDecision === "abort_only") {
    await appendLog(options.config.logDirectory, "Benutzer hat 'Nur Bot beenden' im Recovery-Dialog gewÃ¤hlt.");
    return "abort";
  }

  await appendLog(options.config.logDirectory, "Automatischer Recovery-Neustart nach BenutzerbestÃ¤tigung wird ausgefÃ¼hrt.");
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
      "- So kÃ¶nnen Crash-Logs, Dialogfenster und der Recovery-Ablauf gezielt geprÃ¼ft werden."
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
  console.log("Alle zugehÃ¶rigen KI-Prozesse wurden beendet.");
}

async function main() {
  const config = loadConfig();
  const inspectKi = process.argv.includes("--inspect-ki");
  const loginKi = process.argv.includes("--login-ki");
  const runOnce = process.argv.includes("--once");
  const testRecovery = process.argv.includes("--test-recovery");
  const closeKi = process.argv.includes("--close-ki");
  const navigateKiSource = process.argv.includes("--navigate-ki-source");
  const captureKiTree = process.argv.includes("--capture-ki-tree");
  const captureKiHeader = process.argv.includes("--capture-ki-header");
  const captureKiTable = process.argv.includes("--capture-ki-table");
  const readKiTable = process.argv.includes("--read-ki-table");
  const analyzeKiVision = process.argv.includes("--analyze-ki-vision");
  const captureKiTemplates = process.argv.includes("--capture-ki-templates");
  const matchKiTemplates = process.argv.includes("--match-ki-templates");

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

  if (navigateKiSource) {
    printStartupNotice({
      loginKi: true,
      runOnce: true,
      pollIntervalMinutes: config.pollIntervalMinutes,
      modeLabel: "KI-Navigationstest",
      extraLines: [
        "- Dieser Modus navigiert nur bis zur Datenansicht 'Einheiten nach Sparten der Gruppe'.",
        "- Die mittlere Liste wird dabei noch nicht ausgelesen."
      ]
    });
    const state = await navigateKiToSubmittedUnits(config);
    console.log(JSON.stringify(state, null, 2));
    return;
  }

  if (captureKiTree) {
    printStartupNotice({
      loginKi: true,
      runOnce: true,
      pollIntervalMinutes: config.pollIntervalMinutes,
      modeLabel: "KI-Tree-Capture",
      extraLines: [
        "- Dieser Modus erstellt einen Screenshot nur vom linken KI-Navigationsbaum.",
        "- Das Bild dient als Grundlage fÃ¼r den geplanten OpenCV/OCR-Hybridansatz."
      ]
    });
    const artifact = await captureCurrentKiTreeRegion(config);
    console.log(JSON.stringify(artifact, null, 2));
    return;
  }

  if (captureKiHeader) {
    printStartupNotice({
      loginKi: true,
      runOnce: true,
      pollIntervalMinutes: config.pollIntervalMinutes,
      modeLabel: "KI-Header-Capture",
      extraLines: [
        "- Dieser Modus erstellt einen Screenshot vom oberen KI-Inhaltsbereich.",
        "- Das Bild dient zur Erkennung, ob sich der Bot bereits in der Gruppen-Akte befindet."
      ]
    });
    const artifact = await captureCurrentKiHeaderRegion(config);
    console.log(JSON.stringify(artifact, null, 2));
    return;
  }

  if (captureKiTable) {
    printStartupNotice({
      loginKi: true,
      runOnce: true,
      pollIntervalMinutes: config.pollIntervalMinutes,
      modeLabel: "KI-Tabellen-Capture",
      extraLines: [
        "- Dieser Modus erstellt einen Screenshot nur vom mittleren Tabellenbereich.",
        "- Das Bild dient als Grundlage fÃ¼r das anschlieÃŸende Datenauslesen."
      ]
    });
    const artifact = await captureCurrentKiTableRegion(config);
    console.log(JSON.stringify(artifact, null, 2));
    return;
  }

  if (readKiTable) {
    printStartupNotice({
      loginKi: true,
      runOnce: true,
      pollIntervalMinutes: config.pollIntervalMinutes,
      modeLabel: "KI-Tabellenlesung",
      extraLines: [
        "- Dieser Modus erstellt einen Tabellen-Capture und versucht die Inhalte per OCR zu lesen.",
        "- ZunÃ¤chst geht es nur um die technische Lesbarkeit, noch nicht um das finale Business-Parsing."
      ]
    });
    const artifact = await captureCurrentKiTableRegion(config);
    const ocr = await readTableCaptureWithOcr(config, artifact.imagePath);
    console.log(JSON.stringify({ artifact, ocr }, null, 2));
    return;
  }

  if (analyzeKiVision) {
    printStartupNotice({
      loginKi: true,
      runOnce: true,
      pollIntervalMinutes: config.pollIntervalMinutes,
      modeLabel: "KI-Vision-Analyse",
      extraLines: [
        "- Dieser Modus erstellt Header- und Tree-Captures und wertet sie lokal heuristisch aus.",
        "- Ziel ist die EinschÃ¤tzung, ob Gruppen-Akte und der Zielzustand im Tree visuell stabil erkennbar sind."
      ]
    });
    const treeArtifact = await captureCurrentKiTreeRegion(config);
    const headerArtifact = await captureCurrentKiHeaderRegion(config);
    const treeAnalysis = await analyzeTreeCapture(config, treeArtifact);
    const headerAnalysis = await analyzeHeaderCapture(config, headerArtifact);
    console.log(JSON.stringify({ treeAnalysis, headerAnalysis }, null, 2));
    return;
  }

  if (captureKiTemplates) {
    printStartupNotice({
      loginKi: true,
      runOnce: true,
      pollIntervalMinutes: config.pollIntervalMinutes,
      modeLabel: "KI-Template-Capture",
      extraLines: [
        "- Dieser Modus schneidet aus den aktuellen KI-Captures erste Bildanker fÃ¼r das spÃ¤tere Matching aus.",
        "- Bitte nur dann verwenden, wenn die gewÃ¼nschte Zielansicht gerade wirklich sichtbar ist."
      ]
    });
    const fullWindowArtifact = await captureCurrentKiFullWindowRegion(config);
    const treeArtifact = await captureCurrentKiTreeRegion(config);
    const headerArtifact = await captureCurrentKiHeaderRegion(config);
    const templates = await captureVisionTemplates({ config, fullWindowArtifact, headerArtifact, treeArtifact });
    console.log(JSON.stringify(templates, null, 2));
    return;
  }

  if (matchKiTemplates) {
    printStartupNotice({
      loginKi: true,
      runOnce: true,
      pollIntervalMinutes: config.pollIntervalMinutes,
      modeLabel: "KI-Template-Matching",
      extraLines: [
        "- Dieser Modus prÃ¼ft, ob die aktuellen KI-Bildanker live im Fensterbild wiedergefunden werden.",
        "- Er dient nur der Verifikation des geplanten bildbasierten Zustandsabgleichs."
      ]
    });
    const fullWindowArtifact = await captureCurrentKiFullWindowRegion(config);
    const treeArtifact = await captureCurrentKiTreeRegion(config);

    const templatesBase = `${config.visionDirectory}\\templates`;
    const treeGruppenAkteHeader = await matchTemplateInImage({
      sourcePath: treeArtifact.imagePath,
      templatePath: `${templatesBase}\\tree-gruppen-akte-header.png`,
      searchRegion: { x: 0, y: 0, width: treeArtifact.region.width, height: Math.round(treeArtifact.region.height * 0.16) },
      sampleCols: 14,
      sampleRows: 5,
      threshold: 28
    });
    const treeSubmittedUnitsSelected = await matchTemplateInImage({
      sourcePath: treeArtifact.imagePath,
      templatePath: `${templatesBase}\\tree-submitted-units-selected.png`,
      searchRegion: {
        x: 0,
        y: Math.round(treeArtifact.region.height * 0.24),
        width: treeArtifact.region.width,
        height: Math.round(treeArtifact.region.height * 0.30)
      },
      sampleCols: 14,
      sampleRows: 5,
      threshold: 28
    });
    const contentOpenPath = await matchTemplateInImage({
      sourcePath: fullWindowArtifact.imagePath,
      templatePath: `${templatesBase}\\content-open-path.png`,
      searchRegion: { x: 180, y: 90, width: 720, height: 90 },
      sampleCols: 16,
      sampleRows: 4,
      threshold: 28
    });

    console.log(
      JSON.stringify(
        {
          treeArtifact,
          fullWindowArtifact,
          matches: {
            treeGruppenAkteHeader,
            treeSubmittedUnitsSelected,
            contentOpenPath
          }
        },
        null,
        2
      )
    );
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
