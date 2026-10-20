import { existsSync } from "node:fs";
import { defineConfig } from "@playwright/test";
export default defineConfig({
  testDir: "tests/browser",
  use: {
    baseURL: "http://127.0.0.1:5174",
    launchOptions: {
      executablePath:
        process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH ||
        (existsSync("/usr/bin/chromium") ? "/usr/bin/chromium" : undefined),
      args: ["--no-sandbox"],
    },
  },
  webServer: {
    command:
      process.env.PLAYWRIGHT_SERVER_COMMAND ||
      "npm run dev -- --port 5174 --strictPort",
    url: `http://127.0.0.1:5174${process.env.PLAYWRIGHT_APP_PATH || "/"}`,
    reuseExistingServer: false,
  },
  workers: 1,
});
