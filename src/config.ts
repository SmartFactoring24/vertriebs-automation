import "dotenv/config";
import path from "node:path";

type AppConfig = {
  kiAppPath: string;
  kiUsername: string;
  kiPassword: string;
  kiLoginTimeoutSeconds: number;
  kiProcessName: string;
  kiStartupWindowTitleHint: string;
  kiLoginWindowTitleHint: string;
  kiPortalWindowTitleHint: string;
  kiMainWindowTitleHint: string;
  kiLoginStrategy: "password_only" | "full_login";
  kiPostLoginNewsEnabled: boolean;
  kiPostLoginNewsDelayMs: number;
  kiPostLoginNewsTabCount: number;
  kiPostPortalNewsEnabled: boolean;
  kiPostPortalNewsWindowTitleHint: string;
  kiPostPortalNewsDelayMs: number;
  kiPostPortalConflictPopupPollSeconds: number;
  kiPostPortalNewsPollSeconds: number;
  kiPostPortalNewsTabCount: number;
  kiPortalKiButtonRelX: number;
  kiPortalKiButtonRelY: number;
  kiPostPortalClickWaitSeconds: number;
  kiCloseJavaDiagnosticsBeforeAutomation: boolean;
  whatsappWebUrl: string;
  whatsappGroupName: string;
  pollIntervalMinutes: number;
  exportDirectory: string;
  logDirectory: string;
  screenshotDirectory: string;
  stateDirectory: string;
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

function optionalEnv(name: string, fallback = ""): string {
  return process.env[name] ?? fallback;
}

function resolveLocalPath(value: string): string {
  return path.resolve(process.cwd(), value);
}

export function loadConfig(): AppConfig {
  return {
    kiAppPath: requireEnv("KI_APP_PATH"),
    kiUsername: requireEnv("KI_USERNAME"),
    kiPassword: requireEnv("KI_PASSWORD"),
    kiLoginTimeoutSeconds: Number.parseInt(requireEnv("KI_LOGIN_TIMEOUT_SECONDS", "120"), 10),
    kiProcessName: requireEnv("KI_PROCESS_NAME", "smartclient"),
    kiStartupWindowTitleHint: optionalEnv("KI_STARTUP_WINDOW_TITLE_HINT"),
    kiLoginWindowTitleHint: optionalEnv("KI_LOGIN_WINDOW_TITLE_HINT"),
    kiPortalWindowTitleHint: optionalEnv("KI_PORTAL_WINDOW_TITLE_HINT"),
    kiMainWindowTitleHint: optionalEnv("KI_MAIN_WINDOW_TITLE_HINT"),
    kiLoginStrategy: (requireEnv("KI_LOGIN_STRATEGY", "password_only") as "password_only" | "full_login"),
    kiPostLoginNewsEnabled: requireEnv("KI_POST_LOGIN_NEWS_ENABLED", "true").toLowerCase() === "true",
    kiPostLoginNewsDelayMs: Number.parseInt(requireEnv("KI_POST_LOGIN_NEWS_DELAY_MS", "1500"), 10),
    kiPostLoginNewsTabCount: Number.parseInt(requireEnv("KI_POST_LOGIN_NEWS_TAB_COUNT", "4"), 10),
    kiPostPortalNewsEnabled: requireEnv("KI_POST_PORTAL_NEWS_ENABLED", "true").toLowerCase() === "true",
    kiPostPortalNewsWindowTitleHint: optionalEnv("KI_POST_PORTAL_NEWS_WINDOW_TITLE_HINT", "Mitteilungen"),
    kiPostPortalNewsDelayMs: Number.parseInt(requireEnv("KI_POST_PORTAL_NEWS_DELAY_MS", "2500"), 10),
    kiPostPortalConflictPopupPollSeconds: Number.parseInt(requireEnv("KI_POST_PORTAL_CONFLICT_POPUP_POLL_SECONDS", "5"), 10),
    kiPostPortalNewsPollSeconds: Number.parseInt(requireEnv("KI_POST_PORTAL_NEWS_POLL_SECONDS", "15"), 10),
    kiPostPortalNewsTabCount: Number.parseInt(requireEnv("KI_POST_PORTAL_NEWS_TAB_COUNT", "4"), 10),
    kiPortalKiButtonRelX: Number.parseFloat(requireEnv("KI_PORTAL_KI_BUTTON_REL_X", "0.56375")),
    kiPortalKiButtonRelY: Number.parseFloat(requireEnv("KI_PORTAL_KI_BUTTON_REL_Y", "0.42154")),
    kiPostPortalClickWaitSeconds: Number.parseInt(requireEnv("KI_POST_PORTAL_CLICK_WAIT_SECONDS", "20"), 10),
    kiCloseJavaDiagnosticsBeforeAutomation:
      requireEnv("KI_CLOSE_JAVA_DIAGNOSTICS_BEFORE_AUTOMATION", "false").toLowerCase() === "true",
    whatsappWebUrl: requireEnv("WHATSAPP_WEB_URL", "https://web.whatsapp.com"),
    whatsappGroupName: requireEnv("WHATSAPP_GROUP_NAME"),
    pollIntervalMinutes: Number.parseInt(requireEnv("POLL_INTERVAL_MINUTES", "5"), 10),
    exportDirectory: resolveLocalPath(requireEnv("EXPORT_DIRECTORY", "./data/exports")),
    logDirectory: resolveLocalPath(requireEnv("LOG_DIRECTORY", "./data/logs")),
    screenshotDirectory: resolveLocalPath(requireEnv("SCREENSHOT_DIRECTORY", "./data/screenshots")),
    stateDirectory: resolveLocalPath(requireEnv("STATE_DIRECTORY", "./data/state")),
    whatsappBrowserProfile: resolveLocalPath(requireEnv("WHATSAPP_BROWSER_PROFILE", "./browser-profiles/whatsapp-profile")),
    headless: requireEnv("HEADLESS", "false").toLowerCase() === "true"
  };
}

export type { AppConfig };
