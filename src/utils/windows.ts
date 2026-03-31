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
}): Promise<boolean> {
  const escapedTitle = escapeSingleQuotedPowerShell(options.title);
  const escapedMessage = escapeSingleQuotedPowerShell(options.message);
  const script = `
Add-Type -AssemblyName System.Windows.Forms
$result = [System.Windows.Forms.MessageBox]::Show(
  '${escapedMessage}',
  '${escapedTitle}',
  [System.Windows.Forms.MessageBoxButtons]::YesNo,
  [System.Windows.Forms.MessageBoxIcon]::Warning
)
if ($result -eq [System.Windows.Forms.DialogResult]::Yes) { 'yes' } else { 'no' }
`;

  const result = await runPowerShell(script);
  return result.trim().toLowerCase() === "yes";
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
