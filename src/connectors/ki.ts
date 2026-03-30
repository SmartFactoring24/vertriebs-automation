import { chromium, type BrowserContext } from "playwright";
import type { AppConfig } from "../config.js";
import type { SalesRecord } from "../state/types.js";

export async function fetchSalesRecords(config: AppConfig): Promise<SalesRecord[]> {
  const context = await chromium.launchPersistentContext(config.kiBrowserProfile, {
    headless: config.headless
  });

  try {
    return await collectSalesRecords(context, config);
  } finally {
    await context.close();
  }
}

async function collectSalesRecords(context: BrowserContext, config: AppConfig): Promise<SalesRecord[]> {
  const page = context.pages()[0] ?? (await context.newPage());
  await page.goto(config.kiBaseUrl, { waitUntil: "domcontentloaded" });

  return [
    {
      businessId: "A12345",
      customerName: "Max Mustermann",
      productName: "XYZ Schutzbrief",
      status: "eingereicht",
      salesValue: 1234.56,
      submittedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      source: "KI"
    }
  ];
}
