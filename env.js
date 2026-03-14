// FILE: env.js
// Loads .env from the server directory BEFORE any other imports.
// Import this as the very first thing in server.js.

import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";

const __filename = fileURLToPath(import.meta.url);
const __dirname  = path.dirname(__filename);

const envPath = path.join(__dirname, ".env");

if (!fs.existsSync(envPath)) {
  console.warn("[env] WARNING: .env not found at:", envPath);
} else {
  const raw = fs.readFileSync(envPath, "utf8");
  let loaded = 0;
  for (const rawLine of raw.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;
    const eqIdx = line.indexOf("=");
    if (eqIdx < 1) continue;
    const key = line.slice(0, eqIdx).trim();
    let val = line.slice(eqIdx + 1).trim();
    // Strip inline comments (only "space + #", not # inside URLs)
    val = val.replace(/\s+#[^"']*$/, "");
    // Strip surrounding quotes
    if ((val.startsWith('"') && val.endsWith('"')) ||
        (val.startsWith("'") && val.endsWith("'"))) {
      val = val.slice(1, -1);
    }
    if (key && !(key in process.env)) {
      process.env[key] = val;
      loaded++;
    }
  }
  console.log(`[env] Loaded ${loaded} vars from: ${envPath}`);
  console.log(`[env] HELP_PAY_STK_PROCESS = ${process.env.HELP_PAY_STK_PROCESS || "NOT SET"}`);
}