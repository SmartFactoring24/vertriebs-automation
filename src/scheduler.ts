import type { AppConfig } from "./config.js";

export async function waitForNextRun(config: AppConfig): Promise<void> {
  const milliseconds = config.pollIntervalMinutes * 60 * 1000;
  await new Promise((resolve) => setTimeout(resolve, milliseconds));
}
