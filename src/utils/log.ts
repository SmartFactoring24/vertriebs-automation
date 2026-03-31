import fs from "node:fs/promises";
import path from "node:path";

export async function appendLog(logDirectory: string, message: string): Promise<void> {
  await fs.mkdir(logDirectory, { recursive: true });
  const filePath = path.join(logDirectory, "runtime.log");
  const line = `[${new Date().toISOString()}] ${message}\n`;
  await fs.appendFile(filePath, line, "utf-8");
}

export async function writeCrashLog(
  logDirectory: string,
  details: {
    scope: string;
    error: unknown;
    diagnostics?: Record<string, unknown>;
  }
): Promise<string> {
  await fs.mkdir(logDirectory, { recursive: true });
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const filePath = path.join(logDirectory, `crash-${timestamp}.txt`);
  const normalizedError = normalizeError(details.error);
  const lines = [
    `Timestamp: ${new Date().toISOString()}`,
    `Scope: ${details.scope}`,
    `Name: ${normalizedError.name}`,
    `Message: ${normalizedError.message}`,
    "Stack:",
    normalizedError.stack ?? "(no stack trace)"
  ];

  if (details.diagnostics) {
    lines.push("");
    lines.push("Diagnostics:");
    lines.push(JSON.stringify(details.diagnostics, null, 2));
  }

  await fs.writeFile(filePath, `${lines.join("\n")}\n`, "utf-8");
  return filePath;
}

function normalizeError(error: unknown): Error {
  if (error instanceof Error) {
    return error;
  }

  return new Error(typeof error === "string" ? error : JSON.stringify(error));
}
