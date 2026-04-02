import { execFile } from "node:child_process";
import { access } from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";
import type { AppConfig } from "../config.js";
import { matchTemplateInImage } from "../vision/match.js";
import { captureKiFullWindowRegion, captureKiHeaderRegion, captureKiTableRegion, captureKiTreeRegion, disposeVisionArtifact } from "../vision/tree-capture.js";
import type { VisionCaptureArtifact } from "../vision/types.js";

const execFileAsync = promisify(execFile);

type KiDesktopState = {
  processDetected: boolean;
  startupWindowDetected: boolean;
  loginWindowDetected: boolean;
  portalWindowDetected: boolean;
  mainWindowDetected: boolean;
  stage: KiDesktopStage;
  visibleWindowTitles: string[];
  windowDiagnostics: KiWindowInfo[];
};

type KiDesktopStage = "not_running" | "starting" | "login" | "portal" | "main" | "unknown_window";

type KiWindowInfo = {
  handle: number;
  processId: number;
  processName: string;
  visible: boolean;
  minimized: boolean;
  maximized: boolean;
  title: string;
  className: string;
  x: number;
  y: number;
  width: number;
  height: number;
};

type KiWindowMatchContext = {
  loginTitleHint: string;
  startupTitleHint: string;
  portalTitleHint: string;
  postPortalNewsTitleHint: string;
  mainTitleHint: string;
};

type KiWorkspaceTab = "ki" | "vbi" | "unknown";

type KiNavigationCheckpoint = "submitted_units" | "gruppen_akte" | "unknown";

type KiVisualNavigationState = {
  isTargetOpen: boolean;
  hasSelectedTreeNode: boolean;
  hasOpenPathBanner: boolean;
};

type NamedPixelSample = {
  name: string;
  r: number;
  g: number;
  b: number;
};

function escapeSingleQuotedPowerShell(value: string): string {
  return value.replace(/'/g, "''");
}

function escapeForPowerShellDoubleQuoted(value: string): string {
  return value.replace(/`/g, "``").replace(/"/g, '`"');
}

async function runPowerShell(script: string): Promise<string> {
  const { stdout } = await execFileAsync(
    "powershell.exe",
    ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
    { windowsHide: true }
  );

  return stdout.trim();
}

export async function ensureKiDesktopReady(config: AppConfig): Promise<KiDesktopState> {
  const alreadyRunning = await getKiDesktopState(config);

  if (!alreadyRunning.processDetected) {
    await startKiProcess(config);
  }

  const timeoutAt = Date.now() + config.kiLoginTimeoutSeconds * 1000;
  let latestState = alreadyRunning;

  while (Date.now() < timeoutAt) {
    latestState = await getKiDesktopState(config);

    if (latestState.stage !== "not_running" && latestState.stage !== "starting") {
      return latestState;
    }

    await new Promise((resolve) => setTimeout(resolve, 1000));
  }

  throw new Error(
    `KI hat innerhalb von ${config.kiLoginTimeoutSeconds} Sekunden weder das Login-Fenster noch das Hauptfenster erreicht.`
  );
}

export async function inspectKiDesktop(config: AppConfig): Promise<KiDesktopState> {
  return getKiDesktopState(config);
}

export async function performKiLogin(config: AppConfig): Promise<KiDesktopState> {
  if (config.kiCloseJavaDiagnosticsBeforeAutomation) {
    await closeJavaDiagnosticTools();
  }

  const state = await ensureKiDesktopReady(config);

  if (state.stage === "main") {
    return state;
  }

  if (state.stage !== "login") {
    throw new Error(`KI befindet sich nicht im erwarteten Login-Status. Aktueller Status: ${state.stage}`);
  }

  await submitKiLogin(config);
  await waitForLoginWindowToDisappear(config);

  if (config.kiPostLoginNewsEnabled) {
    await dismissKiNewsWindow(config);
  }

  const portalState = await waitForKiPortalWindow(config);
  const portalProcessId = getPortalProcessId(config, portalState);
  await clickKiOpenInPortal(config, portalState, portalProcessId);
  await dismissPostPortalConflictPopup(config, portalProcessId);

  if (config.kiPostPortalNewsEnabled) {
    await dismissPostPortalNewsWindow(config, portalProcessId);
  }

  const mainState = await waitForKiMainWindow(config);
  return normalizeKiMainWindowAfterLaunch(config, mainState);
}

export async function navigateKiToSubmittedUnits(config: AppConfig): Promise<KiDesktopState> {
  const state = await ensureKiDesktopReady(config);
  const mainState = state.stage === "main" ? state : await performKiLogin(config);
  const preparedNavigation = await prepareKiMainWindowForNavigation(config, mainState);
  const stabilizedState = preparedNavigation.state;

  const mainWindow = findMainWindow(stabilizedState.windowDiagnostics, createWindowMatchContext(config));
  if (!mainWindow) {
    throw new Error("Das KI-Hauptfenster konnte für die VBI-Navigation nicht gefunden werden.");
  }

  const visualStateBeforeNavigation = await detectKiVisualNavigationState(config, mainWindow);
  if (visualStateBeforeNavigation.isTargetOpen) {
    if (preparedNavigation.shouldReminimize) {
      await minimizeWindowByTitle(mainWindow.title);
      await new Promise((resolve) => setTimeout(resolve, 300));
    }

    return getKiDesktopState(config);
  }

  await ensureVbiTabActive(config, mainWindow);
  const visualStateAfterTab = await detectKiVisualNavigationState(config, mainWindow);
  if (visualStateAfterTab.isTargetOpen) {
    if (preparedNavigation.shouldReminimize) {
      await minimizeWindowByTitle(mainWindow.title);
      await new Promise((resolve) => setTimeout(resolve, 300));
    }

    return getKiDesktopState(config);
  }

  let checkpoint = await detectKiNavigationCheckpoint(mainWindow);

  if (checkpoint !== "submitted_units" && checkpoint !== "gruppen_akte") {
    await clickRelativeToWindow(mainWindow, config.kiVbiGruppenAkteRelX, config.kiVbiGruppenAkteRelY);
    await new Promise((resolve) => setTimeout(resolve, config.kiVbiNavigationStepDelayMs));
    checkpoint = await detectKiNavigationCheckpoint(mainWindow);
  }

  if (checkpoint !== "submitted_units") {
    await clickRelativeToWindow(mainWindow, config.kiVbiEingereichtesGeschaeftRelX, config.kiVbiEingereichtesGeschaeftRelY);
    await new Promise((resolve) => setTimeout(resolve, config.kiVbiNavigationStepDelayMs));
    checkpoint = await detectKiNavigationCheckpoint(mainWindow);
  }

  if (checkpoint !== "submitted_units") {
    await clickRelativeToWindow(mainWindow, config.kiVbiEinheitenNachSpartenRelX, config.kiVbiEinheitenNachSpartenRelY);
    await new Promise((resolve) => setTimeout(resolve, config.kiVbiNavigationStepDelayMs));
  }

  if (preparedNavigation.shouldReminimize) {
    await minimizeWindowByTitle(mainWindow.title);
    await new Promise((resolve) => setTimeout(resolve, 300));
  }

  return getKiDesktopState(config);
}

export async function captureCurrentKiTreeRegion(config: AppConfig): Promise<VisionCaptureArtifact> {
  const state = await ensureKiDesktopReady(config);
  const mainState = state.stage === "main" ? state : await performKiLogin(config);
  const preparedNavigation = await prepareKiMainWindowForNavigation(config, mainState);
  const stabilizedState = preparedNavigation.state;
  const mainWindow = findMainWindow(stabilizedState.windowDiagnostics, createWindowMatchContext(config));

  if (!mainWindow) {
    throw new Error("Das KI-Hauptfenster konnte für den Tree-Screenshot nicht gefunden werden.");
  }

  const artifact = await captureKiTreeRegion(config, mainWindow, { keepArtifact: true, fileStem: "ki-tree-debug" });

  if (preparedNavigation.shouldReminimize) {
    await minimizeWindowByTitle(mainWindow.title);
    await new Promise((resolve) => setTimeout(resolve, 300));
  }

  return artifact;
}

export async function captureCurrentKiHeaderRegion(config: AppConfig): Promise<VisionCaptureArtifact> {
  const state = await ensureKiDesktopReady(config);
  const mainState = state.stage === "main" ? state : await performKiLogin(config);
  const preparedNavigation = await prepareKiMainWindowForNavigation(config, mainState);
  const stabilizedState = preparedNavigation.state;
  const mainWindow = findMainWindow(stabilizedState.windowDiagnostics, createWindowMatchContext(config));

  if (!mainWindow) {
    throw new Error("Das KI-Hauptfenster konnte für den Header-Screenshot nicht gefunden werden.");
  }

  const artifact = await captureKiHeaderRegion(config, mainWindow, { keepArtifact: true, fileStem: "ki-header-debug" });

  if (preparedNavigation.shouldReminimize) {
    await minimizeWindowByTitle(mainWindow.title);
    await new Promise((resolve) => setTimeout(resolve, 300));
  }

  return artifact;
}

export async function captureCurrentKiFullWindowRegion(config: AppConfig): Promise<VisionCaptureArtifact> {
  const state = await ensureKiDesktopReady(config);
  const mainState = state.stage === "main" ? state : await performKiLogin(config);
  const preparedNavigation = await prepareKiMainWindowForNavigation(config, mainState);
  const stabilizedState = preparedNavigation.state;
  const mainWindow = findMainWindow(stabilizedState.windowDiagnostics, createWindowMatchContext(config));

  if (!mainWindow) {
    throw new Error("Das KI-Hauptfenster konnte für den Vollfenster-Screenshot nicht gefunden werden.");
  }

  const artifact = await captureKiFullWindowRegion(config, mainWindow, { keepArtifact: true, fileStem: "ki-full-debug" });

  if (preparedNavigation.shouldReminimize) {
    await minimizeWindowByTitle(mainWindow.title);
    await new Promise((resolve) => setTimeout(resolve, 300));
  }

  return artifact;
}

export async function captureCurrentKiTableRegion(config: AppConfig): Promise<VisionCaptureArtifact> {
  const state = await navigateKiToSubmittedUnits(config);
  const preparedNavigation = await prepareKiMainWindowForNavigation(config, state);
  const stabilizedState = preparedNavigation.state;
  const mainWindow = findMainWindow(stabilizedState.windowDiagnostics, createWindowMatchContext(config));

  if (!mainWindow) {
    throw new Error("Das KI-Hauptfenster konnte für den Tabellen-Screenshot nicht gefunden werden.");
  }

  const artifact = await captureKiTableRegion(config, mainWindow, { keepArtifact: true, fileStem: "ki-table-debug" });

  if (preparedNavigation.shouldReminimize) {
    await minimizeWindowByTitle(mainWindow.title);
    await new Promise((resolve) => setTimeout(resolve, 300));
  }

  return artifact;
}

async function detectKiVisualNavigationState(
  config: AppConfig,
  mainWindow: KiWindowInfo
): Promise<KiVisualNavigationState> {
  const templatesDir = path.join(config.visionDirectory, "templates");
  const selectedTemplatePath = path.join(templatesDir, "tree-submitted-units-selected.png");
  const pathTemplatePath = path.join(templatesDir, "content-open-path.png");

  try {
    await access(selectedTemplatePath);
    await access(pathTemplatePath);
  } catch {
    return {
      isTargetOpen: false,
      hasSelectedTreeNode: false,
      hasOpenPathBanner: false
    };
  }

  const treeArtifact = await captureKiTreeRegion(config, mainWindow, { keepArtifact: false, fileStem: `ki-tree-live-${Date.now()}` });
  const fullWindowArtifact = await captureKiFullWindowRegion(config, mainWindow, {
    keepArtifact: false,
    fileStem: `ki-full-live-${Date.now()}`
  });

  try {
    const selectedNodeMatch = await matchTemplateInImage({
      sourcePath: treeArtifact.imagePath,
      templatePath: selectedTemplatePath,
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

    const openPathMatch = await matchTemplateInImage({
      sourcePath: fullWindowArtifact.imagePath,
      templatePath: pathTemplatePath,
      searchRegion: { x: 180, y: 90, width: 720, height: 90 },
      sampleCols: 16,
      sampleRows: 4,
      threshold: 28
    });

    return {
      isTargetOpen: selectedNodeMatch.found && openPathMatch.found,
      hasSelectedTreeNode: selectedNodeMatch.found,
      hasOpenPathBanner: openPathMatch.found
    };
  } finally {
    await disposeVisionArtifact(config, treeArtifact);
    await disposeVisionArtifact(config, fullWindowArtifact);
  }
}

async function prepareKiMainWindowForNavigation(
  config: AppConfig,
  state: KiDesktopState
): Promise<{ state: KiDesktopState; shouldReminimize: boolean }> {
  const mainWindow = findMainWindow(state.windowDiagnostics, createWindowMatchContext(config));
  if (!mainWindow) {
    return { state, shouldReminimize: false };
  }

  if (!mainWindow.minimized) {
    return { state, shouldReminimize: false };
  }

  await maximizeWindowByTitle(mainWindow.title);
  await new Promise((resolve) => setTimeout(resolve, 700));
  return {
    state: await getKiDesktopState(config),
    shouldReminimize: true
  };
}

async function ensureVbiTabActive(config: AppConfig, mainWindow: KiWindowInfo): Promise<void> {
  const activeTab = await detectKiWorkspaceTab(config, mainWindow);

  if (activeTab !== "vbi") {
    await clickRelativeToWindow(mainWindow, config.kiVbiTabRelX, config.kiVbiTabRelY);
    await new Promise((resolve) => setTimeout(resolve, config.kiVbiNavigationStepDelayMs));
  }
}

async function detectKiNavigationCheckpoint(window: KiWindowInfo): Promise<KiNavigationCheckpoint> {
  const samples = await sampleWindowPixels(window, [
    { name: "photo", relX: 0.34, relY: 0.19 },
    { name: "pie", relX: 0.46, relY: 0.18 },
    { name: "linkTop", relX: 0.275, relY: 0.158 },
    { name: "linkBottom", relX: 0.275, relY: 0.176 },
    { name: "treeSelectedBackground", relX: 0.078, relY: 0.345 },
    { name: "tableBlueLink", relX: 0.282, relY: 0.235 },
    { name: "tableHeader", relX: 0.305, relY: 0.205 }
  ]);

  const photo = findPixelSample(samples, "photo");
  const pie = findPixelSample(samples, "pie");
  const linkTop = findPixelSample(samples, "linkTop");
  const linkBottom = findPixelSample(samples, "linkBottom");
  const treeSelectedBackground = findPixelSample(samples, "treeSelectedBackground");
  const tableBlueLink = findPixelSample(samples, "tableBlueLink");
  const tableHeader = findPixelSample(samples, "tableHeader");

  const looksLikeGruppenAkte =
    (photo ? isRichColoredPixel(photo) : false) ||
    (pie ? isRichColoredPixel(pie) : false);

  const looksLikeSubmittedUnits =
    (linkTop ? isBlueLinkPixel(linkTop) : false) ||
    (linkBottom ? isBlueLinkPixel(linkBottom) : false) ||
    (
      (treeSelectedBackground ? isWarmSelectionPixel(treeSelectedBackground) : false) &&
      (
        (tableBlueLink ? isBlueLinkPixel(tableBlueLink) : false) ||
        (tableHeader ? isDarkTextPixel(tableHeader) : false)
      )
    );

  if (looksLikeSubmittedUnits && !looksLikeGruppenAkte) {
    return "submitted_units";
  }

  if (looksLikeGruppenAkte) {
    return "gruppen_akte";
  }

  return "unknown";
}

export async function forceCloseKiProcesses(config: AppConfig): Promise<void> {
  const processName = escapeSingleQuotedPowerShell(config.kiProcessName);
  const installRoot = escapeSingleQuotedPowerShell(path.dirname(path.dirname(config.kiAppPath)));
  const script = `
$rootProcesses = @(Get-Process -Name '${processName}' -ErrorAction SilentlyContinue)
$allProcesses = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue)
$killIds = New-Object System.Collections.Generic.HashSet[int]
$queue = New-Object System.Collections.Generic.Queue[int]

foreach ($root in $rootProcesses) {
  if ($killIds.Add([int]$root.Id)) {
    $queue.Enqueue([int]$root.Id)
  }
}

while ($queue.Count -gt 0) {
  $parentId = $queue.Dequeue()
  $children = @($allProcesses | Where-Object { $_.ParentProcessId -eq $parentId })
  foreach ($child in $children) {
    if ($killIds.Add([int]$child.ProcessId)) {
      $queue.Enqueue([int]$child.ProcessId)
    }
  }
}

$relatedJavaProcesses = @(Get-Process -Name 'javaw' -ErrorAction SilentlyContinue | Where-Object {
  $_.Path -and $_.Path.StartsWith('${installRoot}', [System.StringComparison]::OrdinalIgnoreCase)
})

foreach ($proc in $relatedJavaProcesses) {
  [void]$killIds.Add([int]$proc.Id)
}

foreach ($processId in @($killIds)) {
  Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
}
`;

  await runPowerShell(script);
}

async function startKiProcess(config: AppConfig): Promise<void> {
  const escapedPath = escapeSingleQuotedPowerShell(config.kiAppPath);
  await runPowerShell(`Start-Process -FilePath '${escapedPath}'`);
}

async function submitKiLogin(config: AppConfig): Promise<void> {
  const loginHint = config.kiLoginWindowTitleHint || config.kiStartupWindowTitleHint;

  if (!loginHint.trim()) {
    throw new Error("Der Fenstertitel-Hinweis für das KI-Login ist nicht konfiguriert.");
  }

  const escapedTitle = escapeSingleQuotedPowerShell(loginHint);
  const escapedUsername = escapeForPowerShellDoubleQuoted(config.kiUsername);
  const escapedPassword = escapeForPowerShellDoubleQuoted(config.kiPassword);
  const passwordOnlyScript = `
Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject WScript.Shell
if (-not $wshell.AppActivate('${escapedTitle}')) {
  throw 'Das KI-Login-Fenster konnte nicht aktiviert werden.'
}
Start-Sleep -Milliseconds 500
Set-Clipboard -Value "${escapedPassword}"
[System.Windows.Forms.SendKeys]::SendWait('^a')
Start-Sleep -Milliseconds 100
[System.Windows.Forms.SendKeys]::SendWait('^v')
Start-Sleep -Milliseconds 200
[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
`;

  const fullLoginScript = `
Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject WScript.Shell
if (-not $wshell.AppActivate('${escapedTitle}')) {
  throw 'Das KI-Login-Fenster konnte nicht aktiviert werden.'
}
Start-Sleep -Milliseconds 500
Set-Clipboard -Value "${escapedUsername}"
[System.Windows.Forms.SendKeys]::SendWait('^v')
Start-Sleep -Milliseconds 200
[System.Windows.Forms.SendKeys]::SendWait('{TAB}')
Start-Sleep -Milliseconds 150
Set-Clipboard -Value "${escapedPassword}"
[System.Windows.Forms.SendKeys]::SendWait('^a')
Start-Sleep -Milliseconds 100
[System.Windows.Forms.SendKeys]::SendWait('^v')
Start-Sleep -Milliseconds 200
[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
`;

  const script = config.kiLoginStrategy === "password_only" ? passwordOnlyScript : fullLoginScript;
  await runPowerShell(script);
}

async function waitForLoginWindowToDisappear(config: AppConfig): Promise<void> {
  const timeoutAt = Date.now() + config.kiLoginTimeoutSeconds * 1000;

  while (Date.now() < timeoutAt) {
    const latestState = await getKiDesktopState(config);

    if (!latestState.loginWindowDetected) {
      return;
    }

    await new Promise((resolve) => setTimeout(resolve, 1000));
  }

  throw new Error(`Das KI-Login-Fenster wurde nach dem Anmelden nicht innerhalb von ${config.kiLoginTimeoutSeconds} Sekunden geschlossen.`);
}

async function waitForKiPortalWindow(config: AppConfig): Promise<KiDesktopState> {
  const timeoutAt = Date.now() + config.kiLoginTimeoutSeconds * 1000;

  while (Date.now() < timeoutAt) {
    const latestState = await getKiDesktopState(config);

    if (latestState.stage === "portal" || latestState.portalWindowDetected) {
      return latestState;
    }

    await new Promise((resolve) => setTimeout(resolve, 1000));
  }

  throw new Error(`KI hat das VB-Portal-Fenster nach dem Anmelden nicht innerhalb von ${config.kiLoginTimeoutSeconds} Sekunden erreicht.`);
}

async function waitForKiMainWindow(config: AppConfig): Promise<KiDesktopState> {
  const timeoutAt = Date.now() + config.kiMainWindowTimeoutSeconds * 1000;

  while (Date.now() < timeoutAt) {
    const latestState = await getKiDesktopState(config);

    if (latestState.stage === "main" || latestState.mainWindowDetected) {
      return latestState;
    }

    await new Promise((resolve) => setTimeout(resolve, 1000));
  }

  throw new Error(`KI hat das eigentliche Hauptfenster nach dem Portal-Ablauf nicht innerhalb von ${config.kiMainWindowTimeoutSeconds} Sekunden erreicht.`);
}

async function normalizeKiMainWindowAfterLaunch(
  config: AppConfig,
  state: KiDesktopState
): Promise<KiDesktopState> {
  const mainWindow = findMainWindow(state.windowDiagnostics, createWindowMatchContext(config));
  if (!mainWindow) {
    return state;
  }

  if (!mainWindow.maximized) {
    await maximizeWindowByTitle(mainWindow.title);
    await new Promise((resolve) => setTimeout(resolve, 700));
    return getKiDesktopState(config);
  }

  return state;
}

async function clickKiOpenInPortal(
  config: AppConfig,
  state: KiDesktopState,
  portalProcessId?: number
): Promise<void> {
  const portalWindow = findPortalWindow(
    state.windowDiagnostics,
    createWindowMatchContext(config)
  );

  if (!portalWindow) {
    throw new Error("Das VB-Portal-Fenster steht für den Schritt 'KI öffnen' nicht zur Verfügung.");
  }

  if (hasPostPortalActivity(state.windowDiagnostics, createWindowMatchContext(config), portalProcessId)) {
    return;
  }

  const targetX = Math.round(portalWindow.x + portalWindow.width * config.kiPortalKiButtonRelX);
  const targetY = Math.round(portalWindow.y + portalWindow.height * config.kiPortalKiButtonRelY);
  const escapedTitle = escapeSingleQuotedPowerShell(portalWindow.title);
  const script = `
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class NativeMouse {
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int X, int Y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
}
"@
$wshell = New-Object -ComObject WScript.Shell
if (-not $wshell.AppActivate('${escapedTitle}')) {
  throw 'Das VB-Portal-Fenster konnte nicht aktiviert werden.'
}
Start-Sleep -Milliseconds 700
[NativeMouse]::SetCursorPos(${targetX}, ${targetY}) | Out-Null
Start-Sleep -Milliseconds 120
[NativeMouse]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)
Start-Sleep -Milliseconds 60
[NativeMouse]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)
`;

  await runPowerShell(script);
  await new Promise((resolve) => setTimeout(resolve, config.kiPostPortalClickWaitSeconds * 1000));
}

async function dismissKiNewsWindow(config: AppConfig): Promise<void> {
  await dismissJavaWindowByTabSequence(config.kiPostLoginNewsDelayMs, config.kiPostLoginNewsTabCount, config);
}

async function dismissPostPortalNewsWindow(config: AppConfig, portalProcessId?: number): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, config.kiPostPortalNewsDelayMs));

  const timeoutAt = Date.now() + config.kiPostPortalNewsPollSeconds * 1000;

  while (Date.now() < timeoutAt) {
    const handledByTitle = await tryCloseWindowByTitle(config.kiPostPortalNewsWindowTitleHint);
    if (handledByTitle) {
      return;
    }

    await new Promise((resolve) => setTimeout(resolve, 300));
  }

  await dismissJavaWindowByTabSequence(0, config.kiPostPortalNewsTabCount, config, {
    portalProcessId,
    preferNonPortalJavaw: true
  });
}

async function clickRelativeToWindow(window: KiWindowInfo, relX: number, relY: number): Promise<void> {
  const targetX = Math.round(window.x + window.width * relX);
  const targetY = Math.round(window.y + window.height * relY);
  const escapedTitle = escapeSingleQuotedPowerShell(window.title);
  const script = `
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class NativeMouse {
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int X, int Y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
}
"@
$wshell = New-Object -ComObject WScript.Shell
if (-not $wshell.AppActivate('${escapedTitle}')) {
  throw 'Das Zielfenster konnte nicht aktiviert werden.'
}
Start-Sleep -Milliseconds 350
[NativeMouse]::SetCursorPos(${targetX}, ${targetY}) | Out-Null
Start-Sleep -Milliseconds 200
[NativeMouse]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)
Start-Sleep -Milliseconds 40
[NativeMouse]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)
`;

  await runPowerShell(script);
}

async function detectKiWorkspaceTab(config: AppConfig, window: KiWindowInfo): Promise<KiWorkspaceTab> {
  const escapedTitle = escapeSingleQuotedPowerShell(window.title);
  const kiX = Math.round(window.x + window.width * config.kiKiTabRelX);
  const kiY = Math.round(window.y + window.height * config.kiKiTabRelY);
  const vbiX = Math.round(window.x + window.width * config.kiVbiTabRelX);
  const vbiY = Math.round(window.y + window.height * config.kiVbiTabRelY);
  const script = `
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$wshell = New-Object -ComObject WScript.Shell
if (-not $wshell.AppActivate('${escapedTitle}')) {
  return '{"activeTab":"unknown"}'
}
Start-Sleep -Milliseconds 300
function Get-AverageBrightness([int]$centerX, [int]$centerY) {
  $samples = @(
    @{ X = $centerX; Y = $centerY },
    @{ X = ($centerX - 10); Y = $centerY },
    @{ X = ($centerX + 10); Y = $centerY },
    @{ X = $centerX; Y = ($centerY + 4) }
  )
  $brightnessValues = @()
  foreach ($sample in $samples) {
    $bmp = New-Object System.Drawing.Bitmap 1, 1
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    $graphics.CopyFromScreen($sample.X, $sample.Y, 0, 0, $bmp.Size)
    $pixel = $bmp.GetPixel(0, 0)
    $brightnessValues += (($pixel.R + $pixel.G + $pixel.B) / 3.0)
    $graphics.Dispose()
    $bmp.Dispose()
  }
  return ($brightnessValues | Measure-Object -Average).Average
}
$kiBrightness = Get-AverageBrightness ${kiX} ${kiY}
$vbiBrightness = Get-AverageBrightness ${vbiX} ${vbiY}
$difference = [Math]::Abs($kiBrightness - $vbiBrightness)
$activeTab = 'unknown'
if ($difference -ge 8) {
  if ($vbiBrightness -lt $kiBrightness) {
    $activeTab = 'vbi'
  } elseif ($kiBrightness -lt $vbiBrightness) {
    $activeTab = 'ki'
  }
}
[PSCustomObject]@{
  activeTab = $activeTab
  kiBrightness = [Math]::Round($kiBrightness, 2)
  vbiBrightness = [Math]::Round($vbiBrightness, 2)
  difference = [Math]::Round($difference, 2)
} | ConvertTo-Json -Compress
`;

  const raw = await runPowerShell(script);

  try {
    const parsed = JSON.parse(raw) as { activeTab?: KiWorkspaceTab };
    return parsed.activeTab ?? "unknown";
  } catch {
    return "unknown";
  }
}

async function sampleWindowPixels(
  window: KiWindowInfo,
  points: Array<{ name: string; relX: number; relY: number }>
): Promise<NamedPixelSample[]> {
  const escapedTitle = escapeSingleQuotedPowerShell(window.title);
  const pointPayload = points
    .map((point) => {
      const x = Math.round(window.x + window.width * point.relX);
      const y = Math.round(window.y + window.height * point.relY);
      return `@{ Name='${escapeSingleQuotedPowerShell(point.name)}'; X=${x}; Y=${y} }`;
    })
    .join(",\n");

  const script = `
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$wshell = New-Object -ComObject WScript.Shell
if (-not $wshell.AppActivate('${escapedTitle}')) {
  return '[]'
}
Start-Sleep -Milliseconds 200
$points = @(
${pointPayload}
)
$samples = @()
foreach ($point in $points) {
  $bmp = New-Object System.Drawing.Bitmap 1, 1
  $graphics = [System.Drawing.Graphics]::FromImage($bmp)
  $graphics.CopyFromScreen($point.X, $point.Y, 0, 0, $bmp.Size)
  $pixel = $bmp.GetPixel(0, 0)
  $samples += [PSCustomObject]@{
    name = $point.Name
    r = [int]$pixel.R
    g = [int]$pixel.G
    b = [int]$pixel.B
  }
  $graphics.Dispose()
  $bmp.Dispose()
}
$samples | ConvertTo-Json -Compress
`;

  const raw = await runPowerShell(script);

  try {
    const parsed = JSON.parse(raw) as NamedPixelSample[] | NamedPixelSample;
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch {
    return [];
  }
}

function findPixelSample(samples: NamedPixelSample[], name: string): NamedPixelSample | undefined {
  return samples.find((sample) => sample.name === name);
}

function isRichColoredPixel(sample: NamedPixelSample): boolean {
  const maxChannel = Math.max(sample.r, sample.g, sample.b);
  const minChannel = Math.min(sample.r, sample.g, sample.b);
  const average = (sample.r + sample.g + sample.b) / 3;
  return average < 245 && maxChannel - minChannel >= 35;
}

function isBlueLinkPixel(sample: NamedPixelSample): boolean {
  return sample.b >= 120 && sample.b - sample.r >= 25 && sample.b - sample.g >= 10;
}

function isWarmSelectionPixel(sample: NamedPixelSample): boolean {
  return sample.r >= 180 && sample.g >= 150 && sample.b <= 120;
}

function isDarkTextPixel(sample: NamedPixelSample): boolean {
  return sample.r <= 120 && sample.g <= 120 && sample.b <= 160;
}

async function maximizeWindowByTitle(windowTitle: string): Promise<void> {
  const escapedTitle = escapeSingleQuotedPowerShell(windowTitle);
  const script = `
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class NativeWindow {
  [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
}
"@
$process = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '${escapedTitle}' } | Select-Object -First 1
if (-not $process) {
  return
}
[void][NativeWindow]::ShowWindowAsync($process.MainWindowHandle, 3)
`;

  await runPowerShell(script);
}

async function minimizeWindowByTitle(windowTitle: string): Promise<void> {
  const escapedTitle = escapeSingleQuotedPowerShell(windowTitle);
  const script = `
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class NativeWindow {
  [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
}
"@
$process = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '${escapedTitle}' } | Select-Object -First 1
if (-not $process) {
  return
}
[void][NativeWindow]::ShowWindowAsync($process.MainWindowHandle, 6)
`;

  await runPowerShell(script);
}

async function dismissPostPortalConflictPopup(config: AppConfig, portalProcessId?: number): Promise<void> {
  const timeoutAt = Date.now() + config.kiPostPortalConflictPopupPollSeconds * 1000;

  while (Date.now() < timeoutAt) {
    const state = await getKiDesktopState(config);
    const conflictPopup = findConflictPopupWindow(state.windowDiagnostics, createWindowMatchContext(config), portalProcessId);

    if (conflictPopup) {
      await confirmWindowByProcessId(conflictPopup.processId);
      return;
    }

    await new Promise((resolve) => setTimeout(resolve, 300));
  }
}

async function dismissJavaWindowByTabSequence(
  delayMs: number,
  tabCount: number,
  config: AppConfig,
  options?: {
    portalProcessId?: number;
    preferNonPortalJavaw?: boolean;
  }
): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, delayMs));
  const state = await getKiDesktopState(config);
  const preferredWindow =
    findPreferredJavawWindow(state, config, options) ??
    state.windowDiagnostics.find((window) => window.visible && window.processName === "javaw") ??
    state.windowDiagnostics.find((window) => window.processName === "javaw") ??
    state.windowDiagnostics.find((window) => window.visible) ??
    state.windowDiagnostics[0];

  if (!preferredWindow) {
    return;
  }

  const tabs = "{TAB}".repeat(Math.max(0, tabCount));
  const script = `
Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject WScript.Shell
if (-not $wshell.AppActivate(${preferredWindow.processId})) {
  return
}
Start-Sleep -Milliseconds 500
[System.Windows.Forms.SendKeys]::SendWait('${tabs}')
Start-Sleep -Milliseconds 150
[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
`;

  await runPowerShell(script);
}

async function confirmWindowByProcessId(processId: number): Promise<void> {
  const script = `
Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject WScript.Shell
if (-not $wshell.AppActivate(${processId})) {
  return
}
Start-Sleep -Milliseconds 250
[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
`;

  await runPowerShell(script);
}

function findPreferredJavawWindow(
  state: KiDesktopState,
  config: AppConfig,
  options?: {
    portalProcessId?: number;
    preferNonPortalJavaw?: boolean;
  }
): KiWindowInfo | undefined {
  const javawWindows = state.windowDiagnostics.filter((window) => window.processName === "javaw");

  if (javawWindows.length === 0) {
    return undefined;
  }

  if (options?.preferNonPortalJavaw) {
    const nonPortalWindows = javawWindows
      .filter((window) => {
        if (options.portalProcessId && window.processId === options.portalProcessId) {
          return false;
        }

        if (window.title && config.kiPortalWindowTitleHint.trim()) {
          return !window.title.toLowerCase().includes(config.kiPortalWindowTitleHint.trim().toLowerCase());
        }

        return true;
      })
      .sort((left, right) => right.processId - left.processId);

    if (nonPortalWindows.length > 0) {
      return nonPortalWindows[0];
    }
  }

  return undefined;
}

function getPortalProcessId(config: AppConfig, state: KiDesktopState): number | undefined {
  return findPortalWindow(state.windowDiagnostics, createWindowMatchContext(config))?.processId;
}

async function closeJavaDiagnosticTools(): Promise<void> {
  await runPowerShell(
    "Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -match 'jaccess' } | Stop-Process -Force -ErrorAction SilentlyContinue"
  );
}

async function tryCloseWindowByTitle(windowTitleHint: string): Promise<boolean> {
  if (!windowTitleHint.trim()) {
    return false;
  }

  const escapedTitle = escapeSingleQuotedPowerShell(windowTitleHint);
  const script = `
Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject WScript.Shell
if (-not $wshell.AppActivate('${escapedTitle}')) {
  return 'false'
}
Start-Sleep -Milliseconds 250
[System.Windows.Forms.SendKeys]::SendWait('%{F4}')
return 'true'
`;

  const result = await runPowerShell(script);
  return result.trim().toLowerCase() === "true";
}

async function getKiDesktopState(config: AppConfig): Promise<KiDesktopState> {
  const processName = escapeSingleQuotedPowerShell(config.kiProcessName);
  const installRoot = escapeSingleQuotedPowerShell(path.dirname(path.dirname(config.kiAppPath)));
  const script = buildKiDesktopProbeScript(processName, installRoot);

  const raw = await runPowerShell(script);
  const parsed = raw
    ? (JSON.parse(raw) as {
        processDetected: boolean;
        visibleWindowTitles?: string[] | string;
        windowDiagnostics?: KiWindowInfo[] | KiWindowInfo;
      })
    : { processDetected: false, visibleWindowTitles: [], windowDiagnostics: [] };

  const visibleWindowTitles = Array.isArray(parsed.visibleWindowTitles)
    ? parsed.visibleWindowTitles
    : parsed.visibleWindowTitles
      ? [parsed.visibleWindowTitles]
      : [];

  const windowDiagnostics = Array.isArray(parsed.windowDiagnostics)
    ? parsed.windowDiagnostics
    : parsed.windowDiagnostics
      ? [parsed.windowDiagnostics]
      : [];

  const matchContext = createWindowMatchContext(config);
  const startupWindowDetected = hasStartupWindow(windowDiagnostics, matchContext);
  const loginWindowDetected = hasLoginWindow(windowDiagnostics, matchContext);
  const portalWindowDetected = hasPortalWindow(windowDiagnostics, matchContext);
  const mainWindowDetected = hasMainWindow(windowDiagnostics, matchContext);
  const stage = determineKiDesktopStage({
    processDetected: parsed.processDetected,
    windowDiagnostics,
    startupWindowDetected,
    loginWindowDetected,
    portalWindowDetected,
    mainWindowDetected
  });

  return {
    processDetected: parsed.processDetected,
    startupWindowDetected,
    portalWindowDetected,
    visibleWindowTitles,
    windowDiagnostics,
    loginWindowDetected,
    mainWindowDetected,
    stage
  };
}

function buildKiDesktopProbeScript(processName: string, installRoot: string): string {
  const script = `
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Win32Rects {
  [StructLayout(LayoutKind.Sequential)]
  public struct RECT {
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
  }
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool IsZoomed(IntPtr hWnd);
}
"@
$processes = Get-Process -Name '${processName}' -ErrorAction SilentlyContinue
$windowInfo = @()
if ($processes) {
  $candidateNames = @('${processName}', 'javaw')
  $relatedProcesses = @(
    Get-Process -Name $candidateNames -ErrorAction SilentlyContinue | Where-Object {
      $_.Path -and $_.Path.StartsWith('${installRoot}', [System.StringComparison]::OrdinalIgnoreCase)
    }
  )
  foreach ($proc in $relatedProcesses) {
    $rect = New-Object Win32Rects+RECT
    $x = 0
    $y = 0
    $width = 0
    $height = 0
    $isMinimized = $false
    $isMaximized = $false
    if ($proc.MainWindowHandle -and [Win32Rects]::GetWindowRect([IntPtr]$proc.MainWindowHandle, [ref]$rect)) {
      $x = $rect.Left
      $y = $rect.Top
      $width = $rect.Right - $rect.Left
      $height = $rect.Bottom - $rect.Top
      $isMinimized = [Win32Rects]::IsIconic([IntPtr]$proc.MainWindowHandle)
      $isMaximized = [Win32Rects]::IsZoomed([IntPtr]$proc.MainWindowHandle)
    }
    $windowInfo += [PSCustomObject]@{
      handle = [int64]$proc.MainWindowHandle
      processId = [int]$proc.Id
      processName = $proc.ProcessName
      visible = [bool]$proc.MainWindowTitle
      minimized = [bool]$isMinimized
      maximized = [bool]$isMaximized
      title = $proc.MainWindowTitle
      className = ""
      x = $x
      y = $y
      width = $width
      height = $height
    }
  }
}
$result = [PSCustomObject]@{
  processDetected = [bool]$processes
  visibleWindowTitles = @($windowInfo | Where-Object { $_.visible -and $_.title } | Select-Object -ExpandProperty title)
  windowDiagnostics = @($windowInfo)
}
$result | ConvertTo-Json -Compress
`;

  return script;
}

function determineKiDesktopStage(input: {
  processDetected: boolean;
  windowDiagnostics: KiWindowInfo[];
  startupWindowDetected: boolean;
  loginWindowDetected: boolean;
  portalWindowDetected: boolean;
  mainWindowDetected: boolean;
}): KiDesktopStage {
  if (!input.processDetected) {
    return "not_running";
  }

  if (input.mainWindowDetected) {
    return "main";
  }

  if (input.portalWindowDetected) {
    return "portal";
  }

  if (input.loginWindowDetected) {
    return "login";
  }

  if (input.startupWindowDetected) {
    return "starting";
  }

  if (input.windowDiagnostics.some((window) => window.visible && window.title.trim())) {
    return "unknown_window";
  }

  return "starting";
}

function createWindowMatchContext(config: AppConfig): KiWindowMatchContext {
  return {
    loginTitleHint: normalizeWindowTitle(config.kiLoginWindowTitleHint),
    startupTitleHint: normalizeWindowTitle(config.kiStartupWindowTitleHint),
    portalTitleHint: normalizeWindowTitle(config.kiPortalWindowTitleHint),
    postPortalNewsTitleHint: normalizeWindowTitle(config.kiPostPortalNewsWindowTitleHint),
    mainTitleHint: normalizeWindowTitle(config.kiMainWindowTitleHint)
  };
}

function normalizeWindowTitle(value: string): string {
  return value.trim().toLowerCase();
}

function isVisibleJavaWindow(window: KiWindowInfo): boolean {
  return window.processName === "javaw" && window.visible;
}

function titleIncludes(title: string, needle: string): boolean {
  return needle.length > 0 && normalizeWindowTitle(title).includes(needle);
}

function titleEquals(title: string, needle: string): boolean {
  return needle.length > 0 && normalizeWindowTitle(title) === needle;
}

function hasPortalWindow(windows: KiWindowInfo[], context: KiWindowMatchContext): boolean {
  return findPortalWindow(windows, context) !== undefined;
}

function findPortalWindow(windows: KiWindowInfo[], context: KiWindowMatchContext): KiWindowInfo | undefined {
  return windows.find((window) => isVisibleJavaWindow(window) && titleIncludes(window.title, context.portalTitleHint));
}

function hasNewsWindow(windows: KiWindowInfo[], context: KiWindowMatchContext): boolean {
  return windows.some((window) => isVisibleJavaWindow(window) && titleIncludes(window.title, context.postPortalNewsTitleHint));
}

function findConflictPopupWindow(
  windows: KiWindowInfo[],
  context: KiWindowMatchContext,
  portalProcessId?: number
): KiWindowInfo | undefined {
  return windows
    .filter((window) => {
      if (!isVisibleJavaWindow(window)) {
        return false;
      }

      if (portalProcessId && window.processId === portalProcessId) {
        return false;
      }

      if (titleIncludes(window.title, context.portalTitleHint) || titleIncludes(window.title, context.postPortalNewsTitleHint)) {
        return false;
      }

      if (isMainWindowCandidate(window, context) || isLoginWindowCandidate(window, context)) {
        return false;
      }

      return window.width > 0 && window.height > 0 && window.width <= 900 && window.height <= 800;
    })
    .sort((left, right) => right.processId - left.processId)[0];
}

function hasPostPortalActivity(
  windows: KiWindowInfo[],
  context: KiWindowMatchContext,
  portalProcessId?: number
): boolean {
  return windows.some((window) => {
    if (window.processName !== "javaw") {
      return false;
    }

    if (portalProcessId && window.processId === portalProcessId) {
      return false;
    }

    if (findConflictPopupWindow([window], context, portalProcessId)) {
      return true;
    }

    if (isVisibleJavaWindow(window) && titleIncludes(window.title, context.postPortalNewsTitleHint)) {
      return true;
    }

    return isMainWindowCandidate(window, context);
  });
}

function hasMainWindow(windows: KiWindowInfo[], context: KiWindowMatchContext): boolean {
  return windows.some((window) => isMainWindowCandidate(window, context));
}

function findMainWindow(windows: KiWindowInfo[], context: KiWindowMatchContext): KiWindowInfo | undefined {
  return windows.find((window) => isMainWindowCandidate(window, context));
}

function isMainWindowCandidate(window: KiWindowInfo, context: KiWindowMatchContext): boolean {
  if (!isVisibleJavaWindow(window)) {
    return false;
  }

  if (titleIncludes(window.title, context.portalTitleHint) || titleIncludes(window.title, context.postPortalNewsTitleHint)) {
    return false;
  }

  if (context.mainTitleHint && titleIncludes(window.title, context.mainTitleHint)) {
    return true;
  }

  const normalizedTitle = normalizeWindowTitle(window.title);
  const isGenericDvagMain = normalizedTitle.includes("dvag online-system");
  const looksLikeExactLoginWindow =
    titleEquals(window.title, context.loginTitleHint) || titleEquals(window.title, context.startupTitleHint);

  return isGenericDvagMain && !looksLikeExactLoginWindow && window.width >= 600 && window.height >= 400;
}

function hasLoginWindow(windows: KiWindowInfo[], context: KiWindowMatchContext): boolean {
  return windows.some((window) => isLoginWindowCandidate(window, context));
}

function isLoginWindowCandidate(window: KiWindowInfo, context: KiWindowMatchContext): boolean {
  if (!isVisibleJavaWindow(window)) {
    return false;
  }

  if (titleIncludes(window.title, context.portalTitleHint) || titleIncludes(window.title, context.postPortalNewsTitleHint)) {
    return false;
  }

  if (titleEquals(window.title, context.loginTitleHint)) {
    return true;
  }

  return titleEquals(window.title, context.startupTitleHint);
}

function hasStartupWindow(windows: KiWindowInfo[], context: KiWindowMatchContext): boolean {
  if (windows.every((window) => !window.visible || !window.title.trim())) {
    return true;
  }

  return windows.some((window) => {
    if (!isVisibleJavaWindow(window)) {
      return false;
    }

    return titleEquals(window.title, context.startupTitleHint) && !titleEquals(window.title, context.loginTitleHint);
  });
}

export type { KiDesktopStage, KiDesktopState, KiWindowInfo };
