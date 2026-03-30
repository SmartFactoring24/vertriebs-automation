import fs from "node:fs/promises";
import path from "node:path";

export async function appendLog(logDirectory: string, message: string): Promise<void> {
  await fs.mkdir(logDirectory, { recursive: true });
  const filePath = path.join(logDirectory, "runtime.log");
  const line = `[${new Date().toISOString()}] ${message}\n`;
  await fs.appendFile(filePath, line, "utf-8");
}
