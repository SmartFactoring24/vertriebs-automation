import { chromium, type BrowserContext } from "playwright";
import type { AppConfig } from "../config.js";
import type { SalesChange } from "../state/types.js";

function formatCurrency(value: number): string {
  return new Intl.NumberFormat("de-DE", {
    style: "currency",
    currency: "EUR"
  }).format(value);
}

export function buildWhatsAppMessage(changes: SalesChange[]): string {
  const lines = ["Update Verkaufszahlen", ""];

  for (const [index, change] of changes.entries()) {
    lines.push(
      `${index + 1}. ${change.record.businessId} | ${change.record.customerName} | ${change.record.productName} | ${formatCurrency(change.record.salesValue)} | ${change.record.status}`
    );
  }

  return lines.join("\n");
}

export async function sendWhatsAppNotification(config: AppConfig, message: string): Promise<void> {
  const context = await chromium.launchPersistentContext(config.whatsappBrowserProfile, {
    headless: config.headless
  });

  try {
    await postMessageToGroup(context, config, message);
  } finally {
    await context.close();
  }
}

async function postMessageToGroup(context: BrowserContext, config: AppConfig, message: string): Promise<void> {
  const page = context.pages()[0] ?? (await context.newPage());
  await page.goto(config.whatsappWebUrl, { waitUntil: "domcontentloaded" });

  await page.getByRole("textbox", { name: /search/i }).fill(config.whatsappGroupName);
  await page.getByText(config.whatsappGroupName, { exact: true }).click();

  const chatHeader = page.getByRole("button", { name: config.whatsappGroupName });
  await chatHeader.waitFor({ timeout: 10000 });

  const messageBox = page.locator("div[contenteditable='true']").last();
  await messageBox.fill(message);
  await page.keyboard.press("Enter");
}
