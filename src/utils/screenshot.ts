import fs from "node:fs/promises";
import path from "node:path";
import type { Page } from "playwright";

export async function saveScreenshot(page: Page, screenshotDirectory: string, label: string): Promise<string> {
  await fs.mkdir(screenshotDirectory, { recursive: true });
  const safeLabel = label.replace(/[^a-zA-Z0-9-_]/g, "-");
  const filePath = path.join(screenshotDirectory, `${Date.now()}-${safeLabel}.png`);
  await page.screenshot({ path: filePath, fullPage: true });
  return filePath;
}
