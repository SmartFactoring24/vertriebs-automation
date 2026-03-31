import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

function escapeSingleQuotedPowerShell(value: string): string {
  return value.replace(/'/g, "''");
}

async function runPowerShell(script: string): Promise<string> {
  const { stdout } = await execFileAsync(
    "powershell.exe",
    ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
    { windowsHide: true }
  );

  return stdout.trim();
}

export async function showRecoveryPrompt(options: {
  title: string;
  message: string;
}): Promise<"restart" | "abort_only" | "abort_and_close_ki"> {
  const escapedTitle = escapeSingleQuotedPowerShell(options.title);
  const escapedMessage = escapeSingleQuotedPowerShell(options.message);
  const script = `
Add-Type -AssemblyName System.Windows.Forms
$result = [System.Windows.Forms.MessageBox]::Show(
  '${escapedMessage}',
  '${escapedTitle}',
  [System.Windows.Forms.MessageBoxButtons]::AbortRetryIgnore,
  [System.Windows.Forms.MessageBoxIcon]::Warning
)

switch ($result) {
  ([System.Windows.Forms.DialogResult]::Retry) { 'restart' }
  ([System.Windows.Forms.DialogResult]::Abort) { 'abort_and_close_ki' }
  default { 'abort_only' }
}
`;

  const result = await runPowerShell(script);
  const normalizedResult = result.trim().toLowerCase();

  if (normalizedResult === "restart" || normalizedResult === "abort_only" || normalizedResult === "abort_and_close_ki") {
    return normalizedResult;
  }

  return "abort_only";
}

export async function showInfoDialog(options: {
  title: string;
  message: string;
}): Promise<void> {
  const escapedTitle = escapeSingleQuotedPowerShell(options.title);
  const escapedMessage = escapeSingleQuotedPowerShell(options.message);
  const script = `
Add-Type -AssemblyName System.Windows.Forms
[void][System.Windows.Forms.MessageBox]::Show(
  '${escapedMessage}',
  '${escapedTitle}',
  [System.Windows.Forms.MessageBoxButtons]::OK,
  [System.Windows.Forms.MessageBoxIcon]::Information
)
`;

  await runPowerShell(script);
}
