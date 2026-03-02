
// import "dotenv/config";
// import express from "express";
// import multer from "multer";
// import fs from "fs";
// import path from "path";
// import os from "os";
// import { spawn } from "child_process";
// import { google } from "googleapis";

// import session from "express-session";
// import passport from "./auth/passport.js";
// import authRoutes from "./auth/authRoutes.js";
// import { requireAuth } from "./middleware/requireAuth.js";

// import { fileURLToPath } from "url";
// const __filename = fileURLToPath(import.meta.url);
// const __dirname = path.dirname(__filename);

// const app = express();

// // =========================
// // JOBS (in-memory)
// // =========================
// const JOBS = new Map(); // jobId -> job

// function makeJobId() {
//   return "job_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
// }
// function nowISO() {
//   return new Date().toISOString();
// }
// function userKey(req) {
//   return req.user?.email || req.user?.id || "anon";
// }
// function createJob({ req, project, runId }) {
//   const jobId = makeJobId();
//   const job = {
//     jobId,
//     userKey: userKey(req),
//     project,
//     runId,
//     status: "queued", // queued | running | uploading | done | error
//     pct: 0,
//     hint: "Queued…",
//     logs: [],
//     outputs: [],
//     error: null,
//     createdAt: nowISO(),
//     updatedAt: nowISO(),
//   };
//   JOBS.set(jobId, job);
//   return job;
// }
// function setJob(job, patch = {}) {
//   Object.assign(job, patch);
//   job.updatedAt = nowISO();
// }
// function addLog(job, line) {
//   const s = String(line ?? "").trimEnd();
//   if (!s) return;
//   job.logs.push(s);
//   if (job.logs.length > 600) job.logs = job.logs.slice(-600);
//   job.updatedAt = nowISO();
// }

// // -------------------------
// // BASIC MIDDLEWARES
// // -------------------------
// app.use(express.urlencoded({ extended: true }));
// app.use(express.json());

// // -------------------------
// // SESSION + PASSPORT
// // -------------------------
// app.use(
//   session({
//     name: "sid",
//     secret: process.env.SESSION_SECRET || "dev_secret_change_me",
//     resave: false,
//     saveUninitialized: false,
//     cookie: {
//       httpOnly: true,
//       sameSite: "lax",
//       secure: false,
//       maxAge: 7 * 24 * 60 * 60 * 1000,
//     },
//   })
// );

// app.use(passport.initialize());
// app.use(passport.session());
// app.use("/auth", authRoutes);

// // -------------------------
// // STATIC FILES
// // -------------------------
// app.use(express.static(path.join(__dirname, "public"), { index: false }));

// // -------------------------
// // PAGES
// // -------------------------
// app.get("/login", (req, res) => {
//   res.sendFile(path.join(__dirname, "public", "login.html"));
// });

// app.get("/", requireAuth, (req, res) => {
//   res.sendFile(path.join(__dirname, "public", "index.html"));
// });

// app.get("/me", requireAuth, (req, res) => {
//   res.json({ ok: true, user: req.user });
// });

// // -------------------------
// // MULTER
// // -------------------------
// const upload = multer({ dest: path.join(os.tmpdir(), "uploads") });

// const DRIVE_ROOT_FOLDER_ID = process.env.DRIVE_ROOT_FOLDER_ID;
// const PYTHON_BIN = process.env.PYTHON_BIN || "python";

// // -------------------------
// // PROJECT CONFIG
// // -------------------------
// const PROJECTS = {
//   "Payable Stock": {
//     script: "agents/Payable_STK.py",
//     mode: "single",
//     outputName: "Payable_STK.xlsx",
//     supportsOutputNameArg: true,
//     needsMonthYear: false,
//   },
//   "Payable Trade Discount": {
//     script: "agents/payable_TD.py",
//     mode: "single",
//     outputName: "Payable_TD.xlsx",
//     supportsOutputNameArg: true,
//     needsMonthYear: false,
//   },
//   "Receivable Stock": {
//     script: "agents/rec_stk.py",
//     mode: "multi",
//     outputName: "REC_STK_MASTER.xlsx",
//     supportsOutputNameArg: true,
//     needsMonthYear: true,
//   },
//   "Receivable Trade Discount": {
//     script: "agents/rec_TD.py",
//     mode: "multi",
//     outputName: "MASTER_OUTPUT.xlsx",
//     supportsOutputNameArg: true,
//     needsMonthYear: true,
//   },
//   "Primary Sale": {
//     script: "agents/primary_sale.py",
//     mode: "single",
//     outputName: "Output.xlsx",
//     supportsOutputNameArg: true,
//     needsMonthYear: false,
//   },
// };

// // -------------------------
// // REPORT SHEET CONFIG
// // -------------------------
// const REPORT_SHEET_TABS = {
//   "Payable Stock":             process.env.PAY_STK_TAB_NAME   || "PAY_STK",
//   "Payable Trade Discount":    process.env.PAY_TD_TAB_NAME    || "PAY_TD",
//   "Receivable Stock":          process.env.REC_STK_TAB_NAME   || "REC_STK",
//   "Receivable Trade Discount": process.env.REC_TD_TAB_NAME    || "REC_TD",
//   "Primary Sale":              process.env.PRIM_SALE_TAB_NAME || "PRIM_SALE",
// };

// const REPORT_SHEET_ID = process.env.REPORT_SHEET_ID;

// const DRIVE_OPTS = {
//   supportsAllDrives: true,
//   includeItemsFromAllDrives: true,
// };

// // -------------------------
// // HELPERS
// // -------------------------
// function makeRunIdIST() {
//   const now = new Date();
//   const parts = new Intl.DateTimeFormat("en-GB", {
//     timeZone: "Asia/Kolkata",
//     year: "numeric",
//     month: "2-digit",
//     day: "2-digit",
//     hour: "2-digit",
//     minute: "2-digit",
//     second: "2-digit",
//     hour12: false,
//   }).formatToParts(now);
//   const get = (t) => parts.find((p) => p.type === t)?.value;
//   const ms = String(now.getMilliseconds()).padStart(3, "0");
//   return `${get("year")}-${get("month")}-${get("day")}T${get("hour")}-${get("minute")}-${get("second")}-${ms}+05-30`;
// }

// function getAuthClient() {
//   const b64 = process.env.GOOGLE_SA_JSON_B64;
//   if (!b64) throw new Error("Missing GOOGLE_SA_JSON_B64 in .env");
//   const sa = JSON.parse(Buffer.from(b64, "base64").toString("utf8"));
//   return new google.auth.GoogleAuth({
//     credentials: sa,
//     scopes: [
//       "https://www.googleapis.com/auth/drive",
//       "https://www.googleapis.com/auth/spreadsheets.readonly",
//     ],
//   });
// }

// function getDriveClient() {
//   return google.drive({ version: "v3", auth: getAuthClient() });
// }

// async function ensureFolder(drive, parentId, name) {
//   const q = `'${parentId}' in parents and name='${name}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
//   const res = await drive.files.list({ q, fields: "files(id,name)", ...DRIVE_OPTS });
//   if (res.data.files?.length) return res.data.files[0].id;
//   const created = await drive.files.create({
//     requestBody: { name, mimeType: "application/vnd.google-apps.folder", parents: [parentId] },
//     fields: "id",
//     ...DRIVE_OPTS,
//   });
//   return created.data.id;
// }

// async function uploadToDrive(drive, parentId, filePath, driveName) {
//   const res = await drive.files.create({
//     requestBody: { name: driveName, parents: [parentId] },
//     media: { body: fs.createReadStream(filePath) },
//     fields: "id,name",
//     ...DRIVE_OPTS,
//   });
//   return res.data;
// }

// // ✅ UPDATED: UTF-8 safe + streams logs to JOB
// function runPython(scriptAbsPath, args, job) {
//   return new Promise((resolve, reject) => {
//     const p = spawn(PYTHON_BIN, [scriptAbsPath, ...args], {
//       stdio: "pipe",
//       windowsHide: true,
//       env: {
//         ...process.env,
//         PYTHONIOENCODING: "utf-8",
//         PYTHONUTF8: "1",
//       },
//     });

//     let out = "", err = "";

//     p.stdout.on("data", (d) => {
//       const s = d.toString("utf8");
//       out += s;
//       if (job) addLog(job, s);
//     });

//     p.stderr.on("data", (d) => {
//       const s = d.toString("utf8");
//       err += s;
//       if (job) addLog(job, s);
//     });

//     p.on("close", (code) => {
//       if (code === 0) resolve(out);
//       else reject(new Error(err || out || `Python failed with code ${code}`));
//     });
//   });
// }

// function safeCleanup(dir) {
//   try { fs.rmSync(dir, { recursive: true, force: true }); } catch {}
// }

// function cleanMonthYear(v) {
//   return String(v ?? "").trim();
// }

// // -------------------------
// // API — PROJECTS LIST
// // -------------------------
// app.get("/projects", requireAuth, (req, res) => res.json(PROJECTS));

// // -------------------------
// // API — JOB STATUS (polling)
// // -------------------------
// app.get("/api/job-status", requireAuth, (req, res) => {
//   const jobId = String(req.query.jobId || "").trim();
//   if (!jobId) return res.status(400).json({ ok: false, error: "jobId required" });

//   const job = JOBS.get(jobId);
//   if (!job) return res.status(404).json({ ok: false, error: "Job not found" });

//   if (job.userKey !== userKey(req)) {
//     return res.status(403).json({ ok: false, error: "Forbidden" });
//   }

//   return res.json({
//     ok: true,
//     jobId: job.jobId,
//     project: job.project,
//     runId: job.runId,
//     status: job.status,
//     pct: job.pct,
//     hint: job.hint,
//     outputs: job.outputs,
//     error: job.error,
//     logs: job.logs.slice(-200),
//     updatedAt: job.updatedAt,
//   });
// });

// // -------------------------
// // API — REPORT DATA
// // -------------------------
// app.get("/api/report-data", async (req, res) => {
//   res.setHeader("Content-Type", "application/json");

//   if (!req.isAuthenticated || !req.isAuthenticated()) {
//     return res
//       .status(401)
//       .json({ ok: false, error: "Not authenticated. Please refresh the page and log in again." });
//   }

//   try {
//     const { agent, month, year } = req.query;

//     if (!agent) return res.status(400).json({ ok: false, error: "agent param required" });
//     if (!REPORT_SHEET_ID) return res.status(500).json({ ok: false, error: "REPORT_SHEET_ID not set in .env" });

//     const tabName = REPORT_SHEET_TABS[agent];
//     if (!tabName) return res.status(400).json({ ok: false, error: `No tab configured for agent: ${agent}` });

//     const b64 = process.env.GOOGLE_SA_JSON_B64;
//     const sa = JSON.parse(Buffer.from(b64, "base64").toString("utf8"));

//     const authClient = new google.auth.GoogleAuth({
//       credentials: {
//         client_email: sa.client_email,
//         private_key: sa.private_key,
//       },
//       scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
//     });

//     const client = await authClient.getClient();
//     const sheets = google.sheets({ version: "v4", auth: client });

//     const response = await sheets.spreadsheets.values.get({
//       spreadsheetId: REPORT_SHEET_ID,
//       range: tabName,
//     });

//     const rows = response.data.values || [];
//     if (rows.length < 2) {
//       return res.json({ ok: true, headers: [], rows: [], total: 0 });
//     }

//     const headers = rows[0].map((h) => String(h || "").trim());
//     let dataRows = rows.slice(1);

//     const monthIdx = headers.findIndex((h) => h.toUpperCase() === "MONTH");
//     const yearIdx = headers.findIndex((h) => h.toUpperCase() === "YEAR");

//     const filterMonth = (month || "ALL").trim().toUpperCase();
//     const filterYear = (year || "ALL").trim().toUpperCase();

//     if (filterMonth !== "ALL" && monthIdx !== -1) {
//       dataRows = dataRows.filter((r) => (r[monthIdx] || "").trim().toUpperCase() === filterMonth);
//     }
//     if (filterYear !== "ALL" && yearIdx !== -1) {
//       dataRows = dataRows.filter((r) => (r[yearIdx] || "").trim() === filterYear);
//     }

//     return res.json({ ok: true, headers, rows: dataRows, total: dataRows.length });
//   } catch (e) {
//     return res.status(500).json({ ok: false, error: e.message });
//   }
// });

// // -------------------------
// // API — RUN AGENT (ASYNC JOB)
// // -------------------------
// app.post("/run/:project", requireAuth, upload.array("files"), async (req, res) => {
//   let baseTmp = null;

//   try {
//     const project = req.params.project;
//     const cfg = PROJECTS[project];

//     if (!cfg) return res.status(400).json({ ok: false, error: "Unknown project" });
//     if (!DRIVE_ROOT_FOLDER_ID) throw new Error("Missing DRIVE_ROOT_FOLDER_ID in .env");
//     if (!req.files || req.files.length === 0) throw new Error("No files uploaded");

//     const month = cleanMonthYear(req.body?.month).toUpperCase();
//     const year = cleanMonthYear(req.body?.year);

//     if (cfg.needsMonthYear && (!month || !year)) {
//       return res.status(400).json({ ok: false, error: "Month and Year are required for this project." });
//     }

//     const runId = makeRunIdIST();
//     const job = createJob({ req, project, runId });

//     res.json({ ok: true, jobId: job.jobId, runId, project });

//     (async () => {
//       try {
//         setJob(job, { status: "running", pct: 5, hint: "Preparing…" });
//         addLog(job, `→ Agent: ${project}`);
//         if (cfg.needsMonthYear) addLog(job, `→ Period: ${month} / ${year}`);
//         addLog(job, `→ Files: ${req.files.length} file(s)`);

//         const drive = getDriveClient();

//         setJob(job, { pct: 10, hint: "Creating Drive folders…" });

//         const projectFolderId = await ensureFolder(drive, DRIVE_ROOT_FOLDER_ID, project);
//         const inputRootId = await ensureFolder(drive, projectFolderId, "input");
//         const outputRootId = await ensureFolder(drive, projectFolderId, "output");
//         const inputRunId = await ensureFolder(drive, inputRootId, runId);
//         const outputRunId = await ensureFolder(drive, outputRootId, runId);

//         baseTmp = fs.mkdtempSync(path.join(os.tmpdir(), "job-"));
//         const inDir = path.join(baseTmp, "input");
//         const outDir = path.join(baseTmp, "output");
//         fs.mkdirSync(inDir);
//         fs.mkdirSync(outDir);

//         setJob(job, { pct: 18, hint: "Uploading inputs…" });

//         for (const f of req.files) {
//           const dest = path.join(inDir, f.originalname);
//           fs.renameSync(f.path, dest);
//           await uploadToDrive(drive, inputRunId, dest, f.originalname);
//         }

//         const scriptAbs = path.join(__dirname, cfg.script);
//         if (!fs.existsSync(scriptAbs)) throw new Error(`Script not found: ${scriptAbs}`);

//         const wantsOutputName = Boolean(cfg.outputName && cfg.supportsOutputNameArg);

//         setJob(job, { pct: 30, hint: "Running agent…" });
//         addLog(job, "→ Processing started…");

//         if (cfg.mode === "single") {
//           const inputFile = path.join(inDir, req.files[0].originalname);
//           const args = ["--input", inputFile, "--output_dir", outDir];
//           if (wantsOutputName) args.push("--output_name", cfg.outputName);
//           if (cfg.needsMonthYear) args.push("--month", month, "--year", year);
//           await runPython(scriptAbs, args, job);
//         } else {
//           const args = ["--input_dir", inDir, "--output_dir", outDir];
//           if (wantsOutputName) args.push("--output_name", cfg.outputName);
//           if (cfg.needsMonthYear) args.push("--month", month, "--year", year);
//           await runPython(scriptAbs, args, job);
//         }

//         setJob(job, { status: "uploading", pct: 90, hint: "Uploading outputs…" });

//         const outputs = fs.readdirSync(outDir);
//         const uploaded = [];

//         for (const name of outputs) {
//           const fpath = path.join(outDir, name);
//           const meta = await uploadToDrive(drive, outputRunId, fpath, name);
//           uploaded.push(meta);
//         }

//         job.outputs = uploaded;
//         setJob(job, { status: "done", pct: 100, hint: "Completed" });
//         addLog(job, "✓ Completed");
//       } catch (e) {
//         job.error = e.message;
//         setJob(job, { status: "error", hint: "Failed" });
//         addLog(job, "✗ Error: " + e.message);
//       } finally {
//         if (baseTmp) safeCleanup(baseTmp);
//       }
//     })();
//   } catch (e) {
//     if (baseTmp) safeCleanup(baseTmp);
//     return res.status(500).json({ ok: false, error: e.message });
//   }
// });

// const PORT = process.env.PORT || 3000;
// app.listen(PORT, () => console.log(`http://localhost:${PORT}`));




// server.js (UPDATED for Render + GitHub deploy, no Docker)
//
// Key changes vs your file:
// 1) app.set("trust proxy", 1)  -> required behind Render proxy (cookies, OAuth)
// 2) cookie.secure is true ONLY in production (Render). Local stays false.
// 3) (optional) basic health route /healthz
//
// Everything else is same as your code.

import "dotenv/config";
import express from "express";
import multer from "multer";
import fs from "fs";
import path from "path";
import os from "os";
import { spawn } from "child_process";
import { google } from "googleapis";

import session from "express-session";
import passport from "./auth/passport.js";
import authRoutes from "./auth/authRoutes.js";
import { requireAuth } from "./middleware/requireAuth.js";

import { fileURLToPath } from "url";
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();

// ✅ IMPORTANT for Render (behind proxy). Needed for secure cookies + OAuth redirects.
app.set("trust proxy", 1);

// =========================
// JOBS (in-memory)
// =========================
const JOBS = new Map(); // jobId -> job

function makeJobId() {
  return "job_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
}
function nowISO() {
  return new Date().toISOString();
}
function userKey(req) {
  return req.user?.email || req.user?.id || "anon";
}
function createJob({ req, project, runId }) {
  const jobId = makeJobId();
  const job = {
    jobId,
    userKey: userKey(req),
    project,
    runId,
    status: "queued", // queued | running | uploading | done | error
    pct: 0,
    hint: "Queued…",
    logs: [],
    outputs: [],
    error: null,
    createdAt: nowISO(),
    updatedAt: nowISO(),
  };
  JOBS.set(jobId, job);
  return job;
}
function setJob(job, patch = {}) {
  Object.assign(job, patch);
  job.updatedAt = nowISO();
}
function addLog(job, line) {
  const s = String(line ?? "").trimEnd();
  if (!s) return;
  job.logs.push(s);
  if (job.logs.length > 600) job.logs = job.logs.slice(-600);
  job.updatedAt = nowISO();
}

// -------------------------
// BASIC MIDDLEWARES
// -------------------------
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Optional: health check (helps on Render / debugging)
app.get("/healthz", (req, res) => res.status(200).send("ok"));

// -------------------------
// SESSION + PASSPORT
// -------------------------
const isProd = process.env.NODE_ENV === "production";

app.use(
  session({
    name: "sid",
    secret: process.env.SESSION_SECRET || "dev_secret_change_me",
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: "lax",
      // ✅ In Render (HTTPS), secure cookies should be true
      // ✅ Locally (http://localhost), this must be false or session cookie won't set
      secure: isProd,
      maxAge: 7 * 24 * 60 * 60 * 1000,
    },
  })
);

app.use(passport.initialize());
app.use(passport.session());
app.use("/auth", authRoutes);

// -------------------------
// STATIC FILES
// -------------------------
app.use(express.static(path.join(__dirname, "public"), { index: false }));

// -------------------------
// PAGES
// -------------------------
app.get("/login", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "login.html"));
});

app.get("/", requireAuth, (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.get("/me", requireAuth, (req, res) => {
  res.json({ ok: true, user: req.user });
});

// -------------------------
// MULTER
// -------------------------
const upload = multer({ dest: path.join(os.tmpdir(), "uploads") });

const DRIVE_ROOT_FOLDER_ID = process.env.DRIVE_ROOT_FOLDER_ID;
// ✅ Set PYTHON_BIN=python3 in Render env vars (recommended)
const PYTHON_BIN = process.env.PYTHON_BIN || "python";

// -------------------------
// PROJECT CONFIG
// -------------------------
const PROJECTS = {
  "Payable Stock": {
    script: "agents/Payable_STK.py",
    mode: "single",
    outputName: "Payable_STK.xlsx",
    supportsOutputNameArg: true,
    needsMonthYear: false,
  },
  "Payable Trade Discount": {
    script: "agents/payable_TD.py",
    mode: "single",
    outputName: "Payable_TD.xlsx",
    supportsOutputNameArg: true,
    needsMonthYear: false,
  },
  "Receivable Stock": {
    script: "agents/rec_stk.py",
    mode: "multi",
    outputName: "REC_STK_MASTER.xlsx",
    supportsOutputNameArg: true,
    needsMonthYear: true,
  },
  "Receivable Trade Discount": {
    script: "agents/rec_TD.py",
    mode: "multi",
    outputName: "MASTER_OUTPUT.xlsx",
    supportsOutputNameArg: true,
    needsMonthYear: true,
  },
  "Primary Sale": {
    script: "agents/primary_sale.py",
    mode: "single",
    outputName: "Output.xlsx",
    supportsOutputNameArg: true,
    needsMonthYear: false,
  },
};

// -------------------------
// REPORT SHEET CONFIG
// -------------------------
const REPORT_SHEET_TABS = {
  "Payable Stock": process.env.PAY_STK_TAB_NAME || "PAY_STK",
  "Payable Trade Discount": process.env.PAY_TD_TAB_NAME || "PAY_TD",
  "Receivable Stock": process.env.REC_STK_TAB_NAME || "REC_STK",
  "Receivable Trade Discount": process.env.REC_TD_TAB_NAME || "REC_TD",
  "Primary Sale": process.env.PRIM_SALE_TAB_NAME || "PRIM_SALE",
};

const REPORT_SHEET_ID = process.env.REPORT_SHEET_ID;

const DRIVE_OPTS = {
  supportsAllDrives: true,
  includeItemsFromAllDrives: true,
};

// -------------------------
// HELPERS
// -------------------------
function makeRunIdIST() {
  const now = new Date();
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Asia/Kolkata",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  }).formatToParts(now);
  const get = (t) => parts.find((p) => p.type === t)?.value;
  const ms = String(now.getMilliseconds()).padStart(3, "0");
  return `${get("year")}-${get("month")}-${get("day")}T${get("hour")}-${get("minute")}-${get("second")}-${ms}+05-30`;
}

function getAuthClient() {
  const b64 = process.env.GOOGLE_SA_JSON_B64;
  if (!b64) throw new Error("Missing GOOGLE_SA_JSON_B64 in env");
  const sa = JSON.parse(Buffer.from(b64, "base64").toString("utf8"));
  return new google.auth.GoogleAuth({
    credentials: sa,
    scopes: [
      "https://www.googleapis.com/auth/drive",
      "https://www.googleapis.com/auth/spreadsheets.readonly",
    ],
  });
}

function getDriveClient() {
  return google.drive({ version: "v3", auth: getAuthClient() });
}

async function ensureFolder(drive, parentId, name) {
  const q = `'${parentId}' in parents and name='${name}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
  const res = await drive.files.list({ q, fields: "files(id,name)", ...DRIVE_OPTS });
  if (res.data.files?.length) return res.data.files[0].id;
  const created = await drive.files.create({
    requestBody: { name, mimeType: "application/vnd.google-apps.folder", parents: [parentId] },
    fields: "id",
    ...DRIVE_OPTS,
  });
  return created.data.id;
}

async function uploadToDrive(drive, parentId, filePath, driveName) {
  const res = await drive.files.create({
    requestBody: { name: driveName, parents: [parentId] },
    media: { body: fs.createReadStream(filePath) },
    fields: "id,name",
    ...DRIVE_OPTS,
  });
  return res.data;
}

// ✅ UTF-8 safe + streams logs to JOB
function runPython(scriptAbsPath, args, job) {
  return new Promise((resolve, reject) => {
    const p = spawn(PYTHON_BIN, [scriptAbsPath, ...args], {
      stdio: "pipe",
      windowsHide: true,
      env: {
        ...process.env,
        PYTHONIOENCODING: "utf-8",
        PYTHONUTF8: "1",
      },
    });

    let out = "",
      err = "";

    p.stdout.on("data", (d) => {
      const s = d.toString("utf8");
      out += s;
      if (job) addLog(job, s);
    });

    p.stderr.on("data", (d) => {
      const s = d.toString("utf8");
      err += s;
      if (job) addLog(job, s);
    });

    p.on("close", (code) => {
      if (code === 0) resolve(out);
      else reject(new Error(err || out || `Python failed with code ${code}`));
    });
  });
}

function safeCleanup(dir) {
  try {
    fs.rmSync(dir, { recursive: true, force: true });
  } catch {}
}

function cleanMonthYear(v) {
  return String(v ?? "").trim();
}

// -------------------------
// API — PROJECTS LIST
// -------------------------
app.get("/projects", requireAuth, (req, res) => res.json(PROJECTS));

// -------------------------
// API — JOB STATUS (polling)
// -------------------------
app.get("/api/job-status", requireAuth, (req, res) => {
  const jobId = String(req.query.jobId || "").trim();
  if (!jobId) return res.status(400).json({ ok: false, error: "jobId required" });

  const job = JOBS.get(jobId);
  if (!job) return res.status(404).json({ ok: false, error: "Job not found" });

  if (job.userKey !== userKey(req)) {
    return res.status(403).json({ ok: false, error: "Forbidden" });
  }

  return res.json({
    ok: true,
    jobId: job.jobId,
    project: job.project,
    runId: job.runId,
    status: job.status,
    pct: job.pct,
    hint: job.hint,
    outputs: job.outputs,
    error: job.error,
    logs: job.logs.slice(-200),
    updatedAt: job.updatedAt,
  });
});

// -------------------------
// API — REPORT DATA
// -------------------------
app.get("/api/report-data", async (req, res) => {
  res.setHeader("Content-Type", "application/json");

  if (!req.isAuthenticated || !req.isAuthenticated()) {
    return res
      .status(401)
      .json({ ok: false, error: "Not authenticated. Please refresh the page and log in again." });
  }

  try {
    const { agent, month, year } = req.query;

    if (!agent) return res.status(400).json({ ok: false, error: "agent param required" });
    if (!REPORT_SHEET_ID)
      return res.status(500).json({ ok: false, error: "REPORT_SHEET_ID not set in env" });

    const tabName = REPORT_SHEET_TABS[agent];
    if (!tabName) return res.status(400).json({ ok: false, error: `No tab configured for agent: ${agent}` });

    const b64 = process.env.GOOGLE_SA_JSON_B64;
    const sa = JSON.parse(Buffer.from(b64, "base64").toString("utf8"));

    const authClient = new google.auth.GoogleAuth({
      credentials: {
        client_email: sa.client_email,
        private_key: sa.private_key,
      },
      scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
    });

    const client = await authClient.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: REPORT_SHEET_ID,
      range: tabName,
    });

    const rows = response.data.values || [];
    if (rows.length < 2) {
      return res.json({ ok: true, headers: [], rows: [], total: 0 });
    }

    const headers = rows[0].map((h) => String(h || "").trim());
    let dataRows = rows.slice(1);

    const monthIdx = headers.findIndex((h) => h.toUpperCase() === "MONTH");
    const yearIdx = headers.findIndex((h) => h.toUpperCase() === "YEAR");

    const filterMonth = (month || "ALL").trim().toUpperCase();
    const filterYear = (year || "ALL").trim().toUpperCase();

    if (filterMonth !== "ALL" && monthIdx !== -1) {
      dataRows = dataRows.filter((r) => (r[monthIdx] || "").trim().toUpperCase() === filterMonth);
    }
    if (filterYear !== "ALL" && yearIdx !== -1) {
      dataRows = dataRows.filter((r) => (r[yearIdx] || "").trim() === filterYear);
    }

    return res.json({ ok: true, headers, rows: dataRows, total: dataRows.length });
  } catch (e) {
    return res.status(500).json({ ok: false, error: e.message });
  }
});

// -------------------------
// API — RUN AGENT (ASYNC JOB)
// -------------------------
app.post("/run/:project", requireAuth, upload.array("files"), async (req, res) => {
  let baseTmp = null;

  try {
    const project = req.params.project;
    const cfg = PROJECTS[project];

    if (!cfg) return res.status(400).json({ ok: false, error: "Unknown project" });
    if (!DRIVE_ROOT_FOLDER_ID) throw new Error("Missing DRIVE_ROOT_FOLDER_ID in env");
    if (!req.files || req.files.length === 0) throw new Error("No files uploaded");

    const month = cleanMonthYear(req.body?.month).toUpperCase();
    const year = cleanMonthYear(req.body?.year);

    if (cfg.needsMonthYear && (!month || !year)) {
      return res.status(400).json({ ok: false, error: "Month and Year are required for this project." });
    }

    const runId = makeRunIdIST();
    const job = createJob({ req, project, runId });

    // respond immediately
    res.json({ ok: true, jobId: job.jobId, runId, project });

    (async () => {
      try {
        setJob(job, { status: "running", pct: 5, hint: "Preparing…" });
        addLog(job, `→ Agent: ${project}`);
        if (cfg.needsMonthYear) addLog(job, `→ Period: ${month} / ${year}`);
        addLog(job, `→ Files: ${req.files.length} file(s)`);

        const drive = getDriveClient();

        setJob(job, { pct: 10, hint: "Creating Drive folders…" });

        const projectFolderId = await ensureFolder(drive, DRIVE_ROOT_FOLDER_ID, project);
        const inputRootId = await ensureFolder(drive, projectFolderId, "input");
        const outputRootId = await ensureFolder(drive, projectFolderId, "output");
        const inputRunId = await ensureFolder(drive, inputRootId, runId);
        const outputRunId = await ensureFolder(drive, outputRootId, runId);

        baseTmp = fs.mkdtempSync(path.join(os.tmpdir(), "job-"));
        const inDir = path.join(baseTmp, "input");
        const outDir = path.join(baseTmp, "output");
        fs.mkdirSync(inDir);
        fs.mkdirSync(outDir);

        setJob(job, { pct: 18, hint: "Uploading inputs…" });

        for (const f of req.files) {
          const dest = path.join(inDir, f.originalname);
          fs.renameSync(f.path, dest);
          await uploadToDrive(drive, inputRunId, dest, f.originalname);
        }

        const scriptAbs = path.join(__dirname, cfg.script);
        if (!fs.existsSync(scriptAbs)) throw new Error(`Script not found: ${scriptAbs}`);

        const wantsOutputName = Boolean(cfg.outputName && cfg.supportsOutputNameArg);

        setJob(job, { pct: 30, hint: "Running agent…" });
        addLog(job, "→ Processing started…");

        if (cfg.mode === "single") {
          const inputFile = path.join(inDir, req.files[0].originalname);
          const args = ["--input", inputFile, "--output_dir", outDir];
          if (wantsOutputName) args.push("--output_name", cfg.outputName);
          if (cfg.needsMonthYear) args.push("--month", month, "--year", year);
          await runPython(scriptAbs, args, job);
        } else {
          const args = ["--input_dir", inDir, "--output_dir", outDir];
          if (wantsOutputName) args.push("--output_name", cfg.outputName);
          if (cfg.needsMonthYear) args.push("--month", month, "--year", year);
          await runPython(scriptAbs, args, job);
        }

        setJob(job, { status: "uploading", pct: 90, hint: "Uploading outputs…" });

        const outputs = fs.readdirSync(outDir);
        const uploaded = [];

        for (const name of outputs) {
          const fpath = path.join(outDir, name);
          const meta = await uploadToDrive(drive, outputRunId, fpath, name);
          uploaded.push(meta);
        }

        job.outputs = uploaded;
        setJob(job, { status: "done", pct: 100, hint: "Completed" });
        addLog(job, "✓ Completed");
      } catch (e) {
        job.error = e.message;
        setJob(job, { status: "error", hint: "Failed" });
        addLog(job, "✗ Error: " + e.message);
      } finally {
        if (baseTmp) safeCleanup(baseTmp);
      }
    })();
  } catch (e) {
    if (baseTmp) safeCleanup(baseTmp);
    return res.status(500).json({ ok: false, error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`http://localhost:${PORT}`));