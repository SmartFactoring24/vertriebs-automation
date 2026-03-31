import { execFile } from "node:child_process";
import path from "node:path";
import { promisify } from "node:util";
import type { AppConfig } from "../config.js";

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
  title: string;
  className: string;
  x: number;
  y: number;
  width: number;
  height: number;
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
  await clickKiOpenInPortal(config, portalState);

  if (config.kiPostPortalNewsEnabled) {
    await dismissPostPortalNewsWindow(config, portalProcessId);
  }

  return waitForKiMainWindow(config);
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
  const timeoutAt = Date.now() + config.kiLoginTimeoutSeconds * 1000;

  while (Date.now() < timeoutAt) {
    const latestState = await getKiDesktopState(config);

    if (latestState.stage === "main" || latestState.mainWindowDetected) {
      return latestState;
    }

    await new Promise((resolve) => setTimeout(resolve, 1000));
  }

  throw new Error(`KI hat das eigentliche Hauptfenster nach dem Portal-Ablauf nicht innerhalb von ${config.kiLoginTimeoutSeconds} Sekunden erreicht.`);
}

async function clickKiOpenInPortal(config: AppConfig, state: KiDesktopState): Promise<void> {
  const portalWindow = state.windowDiagnostics.find(
    (window) => window.visible && window.processName === "javaw" && window.title.includes(config.kiPortalWindowTitleHint)
  );

  if (!portalWindow) {
    throw new Error("Das VB-Portal-Fenster steht für den Schritt 'KI öffnen' nicht zur Verfügung.");
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
Start-Sleep -Milliseconds 220
[NativeMouse]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)
Start-Sleep -Milliseconds 180
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
  return state.windowDiagnostics.find(
    (window) => window.processName === "javaw" && window.title.includes(config.kiPortalWindowTitleHint)
  )?.processId;
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

  const startupWindowDetected = matchesAnyWindowTitle(visibleWindowTitles, config.kiStartupWindowTitleHint);
  const loginWindowDetected = matchesAnyWindowTitle(visibleWindowTitles, config.kiLoginWindowTitleHint);
  const portalWindowDetected = matchesAnyWindowTitle(visibleWindowTitles, config.kiPortalWindowTitleHint);
  const mainWindowDetected = matchesAnyWindowTitle(visibleWindowTitles, config.kiMainWindowTitleHint);
  const stage = determineKiDesktopStage({
    processDetected: parsed.processDetected,
    visibleWindowTitles,
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
    if ($proc.MainWindowHandle -and [Win32Rects]::GetWindowRect([IntPtr]$proc.MainWindowHandle, [ref]$rect)) {
      $x = $rect.Left
      $y = $rect.Top
      $width = $rect.Right - $rect.Left
      $height = $rect.Bottom - $rect.Top
    }
    $windowInfo += [PSCustomObject]@{
      handle = [int64]$proc.MainWindowHandle
      processId = [int]$proc.Id
      processName = $proc.ProcessName
      visible = [bool]$proc.MainWindowTitle
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

function matchesAnyWindowTitle(windowTitles: string[], hint: string): boolean {
  if (!hint.trim()) {
    return false;
  }

  const normalizedHint = hint.trim().toLowerCase();
  return windowTitles.some((title) => title.toLowerCase().includes(normalizedHint));
}

function determineKiDesktopStage(input: {
  processDetected: boolean;
  visibleWindowTitles: string[];
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

  if (input.visibleWindowTitles.length > 0) {
    return "unknown_window";
  }

  return "starting";
}

export type { KiDesktopStage, KiDesktopState, KiWindowInfo };
