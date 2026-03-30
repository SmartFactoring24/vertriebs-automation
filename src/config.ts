import "dotenv/config";
import path from "node:path";

type AppConfig = {
  kiBaseUrl: string;
  whatsappWebUrl: string;
  whatsappGroupName: string;
  pollIntervalMinutes: number;
  exportDirectory: string;
  logDirectory: string;
  screenshotDirectory: string;
  stateDirectory: string;
  kiBrowserProfile: string;
  whatsappBrowserProfile: string;
  headless: boolean;
};

function requireEnv(name: string, fallback?: string): string {
  const value = process.env[name] ?? fallback;
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

function resolveLocalPath(value: string): string {
  return path.resolve(process.cwd(), value);
}

export function loadConfig(): AppConfig {
  return {
    kiBaseUrl: requireEnv("KI_BASE_URL"),
    whatsappWebUrl: requireEnv("WHATSAPP_WEB_URL", "https://web.whatsapp.com"),
    whatsappGroupName: requireEnv("WHATSAPP_GROUP_NAME"),
    pollIntervalMinutes: Number.parseInt(requireEnv("POLL_INTERVAL_MINUTES", "5"), 10),
    exportDirectory: resolveLocalPath(requireEnv("EXPORT_DIRECTORY", "./data/exports")),
    logDirectory: resolveLocalPath(requireEnv("LOG_DIRECTORY", "./data/logs")),
    screenshotDirectory: resolveLocalPath(requireEnv("SCREENSHOT_DIRECTORY", "./data/screenshots")),
    stateDirectory: resolveLocalPath(requireEnv("STATE_DIRECTORY", "./data/state")),
    kiBrowserProfile: resolveLocalPath(requireEnv("KI_BROWSER_PROFILE", "./browser-profiles/ki-profile")),
    whatsappBrowserProfile: resolveLocalPath(requireEnv("WHATSAPP_BROWSER_PROFILE", "./browser-profiles/whatsapp-profile")),
    headless: requireEnv("HEADLESS", "false").toLowerCase() === "true"
  };
}

export type { AppConfig };

