// import "./env.js";

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

// import archiver from "archiver";

// const app = express();

// // IMPORTANT for Render (behind proxy). Needed for secure cookies + OAuth redirects.
// app.set("trust proxy", 1);

// // =========================
// // JOBS — file-persisted (survives server restarts on Render)
// // =========================
// const JOBS_FILE = path.join(os.tmpdir(), "sor_jobs.json");

// function _loadJobs() {
//   try {
//     if (fs.existsSync(JOBS_FILE)) {
//       const raw = fs.readFileSync(JOBS_FILE, "utf8");
//       if (!raw || !raw.trim()) return new Map();
//       const obj = JSON.parse(raw);
//       if (typeof obj !== "object" || Array.isArray(obj)) return new Map();
//       const m = new Map();
//       Object.entries(obj).forEach(([k, v]) => {
//         if (v && typeof v === "object" && v.jobId && v.status) m.set(k, v);
//       });
//       console.log("[jobs] Loaded " + m.size + " jobs from disk");
//       return m;
//     }
//   } catch (e) {
//     console.error("[jobs] Corrupt jobs file, resetting:", e.message);
//     try { fs.unlinkSync(JOBS_FILE); } catch {}
//   }
//   return new Map();
// }

// function _saveJobs(jobs) {
//   try {
//     const obj = {};
//     jobs.forEach((v, k) => { obj[k] = v; });
//     fs.writeFileSync(JOBS_FILE, JSON.stringify(obj), "utf8");
//   } catch (e) {
//     console.error("[jobs] Failed to save to disk:", e.message);
//   }
// }

// const JOBS = _loadJobs();

// ;(function cleanOldJobs() {
//   const cutoff = Date.now() - 86400000;
//   let n = 0;
//   JOBS.forEach((v, k) => {
//     if (new Date(v.createdAt).getTime() < cutoff) { JOBS.delete(k); n++; }
//   });
//   if (n) { _saveJobs(JOBS); console.log(`[jobs] Cleaned ${n} old jobs`); }
// })();

// function makeJobId() {
//   return "job_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
// }
// function nowISO() { return new Date().toISOString(); }
// function userKey(req) { return req.user?.email || req.user?.id || "anon"; }

// function createJob({ req, project, runId }) {
//   const jobId = makeJobId();
//   const job = {
//     jobId, userKey: userKey(req), project, runId,
//     status: "queued", pct: 0, hint: "Queued…",
//     logs: [], outputs: [], error: null,
//     createdAt: nowISO(), updatedAt: nowISO(),
//   };
//   JOBS.set(jobId, job);
//   _saveJobs(JOBS);
//   return job;
// }
// function setJob(job, patch = {}) {
//   Object.assign(job, patch);
//   job.updatedAt = nowISO();
//   _saveJobs(JOBS);
// }
// function addLog(job, line) {
//   const s = String(line ?? "").trimEnd();
//   if (!s) return;
//   job.logs.push(s);
//   if (job.logs.length > 600) job.logs = job.logs.slice(-600);
//   job.updatedAt = nowISO();
//   _saveJobs(JOBS);
// }

// // -------------------------
// // BASIC MIDDLEWARES
// // -------------------------
// app.use(express.urlencoded({ extended: true }));
// app.use(express.json());
// app.get("/healthz", (req, res) => res.status(200).send("ok"));

// // -------------------------
// // SESSION + PASSPORT
// // -------------------------
// const isProd = process.env.NODE_ENV === "production";

// app.use(session({
//   name: "sid",
//   secret: process.env.SESSION_SECRET || "dev_secret_change_me",
//   resave: false,
//   saveUninitialized: false,
//   cookie: {
//     httpOnly: true, sameSite: "lax", secure: isProd,
//     maxAge: 7 * 24 * 60 * 60 * 1000,
//   },
// }));

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
// app.get("/login", (req, res) => res.sendFile(path.join(__dirname, "public", "login.html")));
// app.get("/", requireAuth, (req, res) => res.sendFile(path.join(__dirname, "public", "index.html")));
// app.get("/me", requireAuth, (req, res) => res.json({ ok: true, user: req.user }));

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
//   "Payable Stock":            { script: "agents/Payable_STK.py", mode: "single", outputName: "Payable_STK.xlsx",    supportsOutputNameArg: true, needsMonthYear: false },
//   "Payable Trade Discount":   { script: "agents/payable_TD.py",  mode: "single", outputName: "Payable_TD.xlsx",     supportsOutputNameArg: true, needsMonthYear: false },
//   "Receivable Stock":         { script: "agents/rec_stk.py",     mode: "multi",  outputName: "REC_STK_MASTER.xlsx", supportsOutputNameArg: true, needsMonthYear: true  },
//   "Receivable Trade Discount":{ script: "agents/rec_TD.py",      mode: "multi",  outputName: "MASTER_OUTPUT.xlsx",  supportsOutputNameArg: true, needsMonthYear: true  },
//   "Primary Sale":             { script: "agents/primary_sale.py",mode: "single", outputName: "Output.xlsx",         supportsOutputNameArg: true, needsMonthYear: false },
// };

// // -------------------------
// // REPORT SHEET CONFIG
// // -------------------------
// const REPORT_SHEET_TABS = {
//   "Payable Stock":            process.env.PAY_STK_TAB_NAME   || "PAY_STK",
//   "Payable Trade Discount":   process.env.PAY_TD_TAB_NAME    || "PAY_TD",
//   "Receivable Stock":         process.env.REC_STK_TAB_NAME   || "REC_STK",
//   "Receivable Trade Discount":process.env.REC_TD_TAB_NAME    || "REC_TD",
//   "Primary Sale":             process.env.PRIM_SALE_TAB_NAME || "PRIM_SALE",
// };
// const REPORT_SHEET_ID = process.env.REPORT_SHEET_ID;
// const DRIVE_OPTS = { supportsAllDrives: true, includeItemsFromAllDrives: true };

// // -------------------------
// // HELPERS
// // -------------------------
// function makeRunIdIST() {
//   const now = new Date();
//   const parts = new Intl.DateTimeFormat("en-GB", {
//     timeZone: "Asia/Kolkata", year: "numeric", month: "2-digit", day: "2-digit",
//     hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false,
//   }).formatToParts(now);
//   const get = (t) => parts.find((p) => p.type === t)?.value;
//   const ms = String(now.getMilliseconds()).padStart(3, "0");
//   return `${get("year")}-${get("month")}-${get("day")}T${get("hour")}-${get("minute")}-${get("second")}-${ms}+05-30`;
// }

// function getAuthClient() {
//   const b64 = process.env.GOOGLE_SA_JSON_B64;
//   if (!b64) throw new Error("Missing GOOGLE_SA_JSON_B64 in env");
//   const sa = JSON.parse(Buffer.from(b64, "base64").toString("utf8"));
//   return new google.auth.GoogleAuth({
//     credentials: sa,
//     scopes: ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets.readonly"],
//   });
// }

// function getDriveClient() { return google.drive({ version: "v3", auth: getAuthClient() }); }

// async function ensureFolder(drive, parentId, name) {
//   const q = `'${parentId}' in parents and name='${name}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
//   const res = await drive.files.list({ q, fields: "files(id,name)", ...DRIVE_OPTS });
//   if (res.data.files?.length) return res.data.files[0].id;
//   const created = await drive.files.create({
//     requestBody: { name, mimeType: "application/vnd.google-apps.folder", parents: [parentId] },
//     fields: "id", ...DRIVE_OPTS,
//   });
//   return created.data.id;
// }

// async function uploadToDrive(drive, parentId, filePath, driveName) {
//   const res = await drive.files.create({
//     requestBody: { name: driveName, parents: [parentId] },
//     media: { body: fs.createReadStream(filePath) },
//     fields: "id,name", ...DRIVE_OPTS,
//   });
//   return res.data;
// }

// function runPython(scriptAbsPath, args, job) {
//   return new Promise((resolve, reject) => {
//     const p = spawn(PYTHON_BIN, [scriptAbsPath, ...args], {
//       stdio: "pipe", windowsHide: true,
//       env: { ...process.env, PYTHONIOENCODING: "utf-8", PYTHONUTF8: "1" },
//     });
//     let out = "", err = "";
//     p.stdout.on("data", (d) => { const s = d.toString("utf8"); out += s; if (job) addLog(job, s); });
//     p.stderr.on("data", (d) => { const s = d.toString("utf8"); err += s; if (job) addLog(job, s); });
//     p.on("close", (code) => {
//       if (code === 0) resolve(out);
//       else reject(new Error(err || out || `Python failed with code ${code}`));
//     });
//   });
// }

// function safeCleanup(dir) { try { fs.rmSync(dir, { recursive: true, force: true }); } catch {} }
// function cleanMonthYear(v) { return String(v ?? "").trim(); }

// function extractPeriodFromLogs(logs) {
//   for (const line of logs) {
//     const m = String(line).match(/PERIOD\s*:\s*([A-Z]{3}-\d{4})/i);
//     if (m) return m[1].toUpperCase();
//   }
//   return null;
// }

// function makeRunFolderName(runId, period) {
//   return period ? `${runId} (${period})` : runId;
// }

// // -------------------------
// // API — PROJECTS LIST
// // -------------------------
// app.get("/projects", requireAuth, (req, res) => {
//   const out = {};
//   Object.entries(PROJECTS).forEach(([name, cfg]) => {
//     out[name] = { mode: cfg.mode, outputName: cfg.outputName, needsMonthYear: cfg.needsMonthYear };
//   });
//   res.json(out);
// });

// // -------------------------
// // API — JOB STATUS
// // -------------------------
// app.get("/api/job-status", requireAuth, (req, res) => {
//   const jobId = String(req.query.jobId || "").trim();
//   if (!jobId) return res.status(400).json({ ok: false, error: "jobId required" });
//   const job = JOBS.get(jobId);
//   if (!job) return res.status(404).json({ ok: false, error: "Job not found" });
//   if (job.userKey !== userKey(req)) return res.status(403).json({ ok: false, error: "Forbidden" });
//   return res.json({
//     ok: true, jobId: job.jobId, project: job.project, runId: job.runId,
//     status: job.status, pct: job.pct, hint: job.hint,
//     outputs: job.outputs, error: job.error,
//     logs: job.logs.slice(-200), updatedAt: job.updatedAt,
//   });
// });

// // -------------------------
// // API — REPORT DATA
// // -------------------------
// app.get("/api/report-data", async (req, res) => {
//   res.setHeader("Content-Type", "application/json");
//   if (!req.isAuthenticated || !req.isAuthenticated()) {
//     return res.status(401).json({ ok: false, error: "Not authenticated." });
//   }
//   try {
//     const { agent, month, year } = req.query;
//     if (!agent) return res.status(400).json({ ok: false, error: "agent param required" });
//     if (!REPORT_SHEET_ID) return res.status(500).json({ ok: false, error: "REPORT_SHEET_ID not set in env" });
//     const tabName = REPORT_SHEET_TABS[agent];
//     if (!tabName) return res.status(400).json({ ok: false, error: `No tab configured for agent: ${agent}` });
//     const b64 = process.env.GOOGLE_SA_JSON_B64;
//     const sa = JSON.parse(Buffer.from(b64, "base64").toString("utf8"));
//     const authClient = new google.auth.GoogleAuth({
//       credentials: { client_email: sa.client_email, private_key: sa.private_key },
//       scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
//     });
//     const client = await authClient.getClient();
//     const sheets = google.sheets({ version: "v4", auth: client });
//     const response = await sheets.spreadsheets.values.get({ spreadsheetId: REPORT_SHEET_ID, range: tabName });
//     const rows = response.data.values || [];
//     if (rows.length < 2) return res.json({ ok: true, headers: [], rows: [], total: 0 });
//     const headers = rows[0].map((h) => String(h || "").trim());
//     let dataRows = rows.slice(1);
//     const monthIdx = headers.findIndex((h) => h.toUpperCase() === "MONTH");
//     const yearIdx  = headers.findIndex((h) => h.toUpperCase() === "YEAR");
//     const filterMonth = (month || "ALL").trim().toUpperCase();
//     const filterYear  = (year  || "ALL").trim().toUpperCase();
//     if (filterMonth !== "ALL" && monthIdx !== -1)
//       dataRows = dataRows.filter((r) => (r[monthIdx] || "").trim().toUpperCase() === filterMonth);
//     if (filterYear !== "ALL" && yearIdx !== -1)
//       dataRows = dataRows.filter((r) => (r[yearIdx] || "").trim() === filterYear);
//     return res.json({ ok: true, headers, rows: dataRows, total: dataRows.length });
//   } catch (e) {
//     return res.status(500).json({ ok: false, error: e.message });
//   }
// });

// // -------------------------
// // API — HELP LINKS
// // -------------------------
// app.get("/api/help-links", requireAuth, (req, res) => {
//   const agents = ["PAY_STK", "PAY_TD", "REC_STK", "REC_TD", "PRI_SALE"];
//   const links = {};
//   for (const key of agents) {
//     links[key] = {
//       process: process.env[`HELP_${key}_PROCESS`] || null,
//       format:  process.env[`HELP_${key}_FORMAT`]  || null,
//       master:  process.env[`HELP_${key}_MASTER`]  || null,
//     };
//   }
//   res.json({ ok: true, links });
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
//     if (!DRIVE_ROOT_FOLDER_ID) throw new Error("Missing DRIVE_ROOT_FOLDER_ID in env");
//     if (!req.files || req.files.length === 0) throw new Error("No files uploaded");
//     const month = cleanMonthYear(req.body?.month).toUpperCase();
//     const year  = cleanMonthYear(req.body?.year);
//     if (cfg.needsMonthYear && (!month || !year)) {
//       return res.status(400).json({ ok: false, error: "Month and Year are required for this project." });
//     }
//     const runId = makeRunIdIST();
//     const job   = createJob({ req, project, runId });
//     res.json({ ok: true, jobId: job.jobId, runId, project });

//     (async () => {
//       try {
//         setJob(job, { status: "running", pct: 5, hint: "Preparing…" });
//         addLog(job, `-> Agent: ${project}`);
//         if (cfg.needsMonthYear) addLog(job, `-> Period: ${month} / ${year}`);
//         addLog(job, `-> Files: ${req.files.length} file(s)`);
//         const drive = getDriveClient();
//         setJob(job, { pct: 10, hint: "Creating Drive folders…" });
//         const projectFolderId = await ensureFolder(drive, DRIVE_ROOT_FOLDER_ID, project);
//         const inputRootId     = await ensureFolder(drive, projectFolderId, "input");
//         const outputRootId    = await ensureFolder(drive, projectFolderId, "output");
//         const inputRunId      = await ensureFolder(drive, inputRootId, runId);
//         baseTmp = fs.mkdtempSync(path.join(os.tmpdir(), "job-"));
//         const inDir  = path.join(baseTmp, "input");
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
//         addLog(job, "-> Processing started…");
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
//         const period = extractPeriodFromLogs(job.logs);
//         addLog(job, period ? `-> Data period detected: ${period}` : "-> Data period not detected, using plain runId");
//         const outputFolderName = makeRunFolderName(runId, period);
//         const outputRunId      = await ensureFolder(drive, outputRootId, outputFolderName);
//         setJob(job, { status: "uploading", pct: 90, hint: "Uploading outputs…" });
//         const outputFiles = fs.readdirSync(outDir);
//         const uploaded    = [];
//         for (const name of outputFiles) {
//           const fpath = path.join(outDir, name);
//           const meta  = await uploadToDrive(drive, outputRunId, fpath, name);
//           uploaded.push(meta);
//         }
//         job.outputs = uploaded;
//         setJob(job, { status: "done", pct: 100, hint: "Completed" });
//         addLog(job, "-> Completed");
//       } catch (e) {
//         job.error = e.message;
//         setJob(job, { status: "error", hint: "Failed" });
//         addLog(job, "-> Error: " + e.message);
//       } finally {
//         if (baseTmp) safeCleanup(baseTmp);
//       }
//     })();
//   } catch (e) {
//     if (baseTmp) safeCleanup(baseTmp);
//     return res.status(500).json({ ok: false, error: e.message });
//   }
// });

// // -------------------------
// // BASE DUMP ROUTES
// // -------------------------
// const BASE_DUMP_FOLDER_ID = process.env.BASE_DUMP_FOLDER_ID;

// async function listDriveFolder(drive, folderId) {
//   const q   = `'${folderId}' in parents and trashed=false`;
//   const res = await drive.files.list({
//     q, fields: "files(id,name,mimeType,size,modifiedTime)",
//     orderBy: "name desc", pageSize: 200, ...DRIVE_OPTS,
//   });
//   return res.data.files || [];
// }

// app.get("/api/basedump/runs", requireAuth, async (req, res) => {
//   try {
//     if (!BASE_DUMP_FOLDER_ID)
//       return res.status(500).json({ ok: false, error: "BASE_DUMP_FOLDER_ID not set in env" });
//     const { agent, month, year } = req.query;
//     if (!agent) return res.status(400).json({ ok: false, error: "agent param required" });
//     const drive = getDriveClient();
//     const agentQ   = `'${BASE_DUMP_FOLDER_ID}' in parents and name='${agent}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
//     const agentRes = await drive.files.list({ q: agentQ, fields: "files(id,name)", ...DRIVE_OPTS });
//     const agentFolders = agentRes.data.files || [];
//     if (!agentFolders.length) return res.json({ ok: true, runs: [], total: 0 });
//     const agentFolderId = agentFolders[0].id;
//     const outQ   = `'${agentFolderId}' in parents and name='output' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
//     const outRes = await drive.files.list({ q: outQ, fields: "files(id,name)", ...DRIVE_OPTS });
//     const outFolders = outRes.data.files || [];
//     if (!outFolders.length) return res.json({ ok: true, runs: [], total: 0 });
//     const outputRootId = outFolders[0].id;
//     const runFolders   = await listDriveFolder(drive, outputRootId);

//     function parseRunFolder(name) {
//       const match = name.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2})-(\d{2})-(\d{2})/);
//       if (!match) return null;
//       const periodMatch = name.match(/\(([A-Z]{3}-\d{4})\)/i);
//       const period = periodMatch ? periodMatch[1].toUpperCase() : null;
//       const label = period
//         ? `${match[3]}/${match[2]}/${match[1]}  ${match[4]}:${match[5]}:${match[6]} IST  [${period}]`
//         : `${match[3]}/${match[2]}/${match[1]}  ${match[4]}:${match[5]}:${match[6]} IST`;
//       return { year: match[1], month: match[2], day: match[3], hour: match[4], minute: match[5], second: match[6], period, label };
//     }

//     let runs = runFolders
//       .filter((f) => f.mimeType === "application/vnd.google-apps.folder")
//       .map((f) => ({ id: f.id, name: f.name, parsed: parseRunFolder(f.name) }))
//       .filter((r) => r.parsed !== null);
//     const fm = (month || "").trim();
//     const fy = (year  || "").trim();
//     if (fm) runs = runs.filter((r) => r.parsed.month === fm.padStart(2, "0"));
//     if (fy) runs = runs.filter((r) => r.parsed.year  === fy);
//     return res.json({ ok: true, runs, total: runs.length });
//   } catch (e) {
//     return res.status(500).json({ ok: false, error: e.message });
//   }
// });

// app.get("/api/basedump/zip", requireAuth, async (req, res) => {
//   try {
//     const { folderId, folderName } = req.query;
//     if (!folderId) return res.status(400).json({ ok: false, error: "folderId required" });
//     const drive     = getDriveClient();
//     const files     = await listDriveFolder(drive, folderId);
//     const fileItems = files.filter((f) => f.mimeType !== "application/vnd.google-apps.folder");
//     if (!fileItems.length)
//       return res.status(404).json({ ok: false, error: "No files found in this run folder" });
//     const zipName = (folderName || folderId) + ".zip";
//     res.setHeader("Content-Type", "application/zip");
//     res.setHeader("Content-Disposition", `attachment; filename="${zipName}"`);
//     const archive = archiver("zip", { zlib: { level: 6 } });
//     archive.pipe(res);
//     for (const file of fileItems) {
//       const stream = await drive.files.get(
//         { fileId: file.id, alt: "media", ...DRIVE_OPTS },
//         { responseType: "stream" }
//       );
//       archive.append(stream.data, { name: file.name });
//     }
//     archive.finalize();
//   } catch (e) {
//     if (!res.headersSent) res.status(500).json({ ok: false, error: e.message });
//   }
// });

// app.get("/api/basedump/debug", requireAuth, async (req, res) => {
//   try {
//     const drive = getDriveClient();
//     const q     = `'${process.env.BASE_DUMP_FOLDER_ID}' in parents and trashed=false`;
//     const r     = await drive.files.list({ q, fields: "files(id,name,mimeType)", ...DRIVE_OPTS });
//     res.json({ root: process.env.BASE_DUMP_FOLDER_ID, files: r.data.files });
//   } catch (e) {
//     res.status(500).json({ error: e.message });
//   }
// });

// // -------------------------
// // START SERVER
// // -------------------------
// const PORT = process.env.PORT || 3000;
// app.listen(PORT, () => {
//   console.log(`[server] Running at http://localhost:${PORT}`);
// });

import "./env.js";
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
import archiver from "archiver";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();

// Important for Render / reverse proxies
app.set("trust proxy", 1);

// -------------------------
// BASIC CONFIG
// -------------------------
const isProd = process.env.NODE_ENV === "production";
const PORT = process.env.PORT || 3000;
const PYTHON_BIN = process.env.PYTHON_BIN || "python3";
const DRIVE_ROOT_FOLDER_ID = process.env.DRIVE_ROOT_FOLDER_ID;
const REPORT_SHEET_ID = process.env.REPORT_SHEET_ID;
const BASE_DUMP_FOLDER_ID = process.env.BASE_DUMP_FOLDER_ID;

const DRIVE_OPTS = {
  supportsAllDrives: true,
  includeItemsFromAllDrives: true,
};

if (isProd && !process.env.SESSION_SECRET) {
  throw new Error("SESSION_SECRET is required in production");
}

// -------------------------
// JOBS — temp file cache only
// not guaranteed to survive restart/redeploy
// -------------------------
const JOBS_FILE = path.join(os.tmpdir(), "sor_jobs.json");

function _loadJobs() {
  try {
    if (fs.existsSync(JOBS_FILE)) {
      const raw = fs.readFileSync(JOBS_FILE, "utf8");
      if (!raw || !raw.trim()) return new Map();

      const obj = JSON.parse(raw);
      if (typeof obj !== "object" || Array.isArray(obj)) return new Map();

      const m = new Map();
      Object.entries(obj).forEach(([k, v]) => {
        if (v && typeof v === "object" && v.jobId && v.status) {
          m.set(k, v);
        }
      });

      console.log("[jobs] Loaded " + m.size + " jobs from temp storage");
      return m;
    }
  } catch (e) {
    console.error("[jobs] Corrupt jobs file, resetting:", e.message);
    try {
      fs.unlinkSync(JOBS_FILE);
    } catch {}
  }

  return new Map();
}

function _saveJobs(jobs) {
  try {
    const obj = {};
    jobs.forEach((v, k) => {
      obj[k] = v;
    });
    fs.writeFileSync(JOBS_FILE, JSON.stringify(obj), "utf8");
  } catch (e) {
    console.error("[jobs] Failed to save to temp storage:", e.message);
  }
}

const JOBS = _loadJobs();

(function cleanOldJobs() {
  const cutoff = Date.now() - 86400000; // 24 hours
  let n = 0;

  JOBS.forEach((v, k) => {
    if (new Date(v.createdAt).getTime() < cutoff) {
      JOBS.delete(k);
      n++;
    }
  });

  if (n) {
    _saveJobs(JOBS);
    console.log(`[jobs] Cleaned ${n} old jobs`);
  }
})();

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
    status: "queued",
    pct: 0,
    hint: "Queued…",
    logs: [],
    outputs: [],
    error: null,
    createdAt: nowISO(),
    updatedAt: nowISO(),
  };

  JOBS.set(jobId, job);
  _saveJobs(JOBS);
  return job;
}

function setJob(job, patch = {}) {
  Object.assign(job, patch);
  job.updatedAt = nowISO();
  _saveJobs(JOBS);
}

function addLog(job, line) {
  const s = String(line ?? "").trimEnd();
  if (!s) return;

  job.logs.push(s);
  if (job.logs.length > 150) {
    job.logs = job.logs.slice(-150);
  }

  job.updatedAt = nowISO();
  _saveJobs(JOBS);
}

function hasRunningJob() {
  for (const [, job] of JOBS) {
    if (job.status === "queued" || job.status === "running" || job.status === "uploading") {
      return true;
    }
  }
  return false;
}

function logMemory(prefix = "") {
  const m = process.memoryUsage();
  console.log(
    `[mem] ${prefix} rss=${Math.round(m.rss / 1024 / 1024)}MB ` +
      `heapUsed=${Math.round(m.heapUsed / 1024 / 1024)}MB ` +
      `heapTotal=${Math.round(m.heapTotal / 1024 / 1024)}MB ` +
      `external=${Math.round(m.external / 1024 / 1024)}MB`
  );
}

// -------------------------
// BASIC MIDDLEWARES
// -------------------------
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

app.get("/healthz", (req, res) => {
  res.status(200).send("ok");
});

// -------------------------
// SESSION + PASSPORT
// -------------------------
app.use(
  session({
    name: "sid",
    secret: process.env.SESSION_SECRET || "dev_secret_change_me",
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: "lax",
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
app.use(
  express.static(path.join(__dirname, "public"), {
    index: false,
  })
);

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
const upload = multer({
  dest: path.join(os.tmpdir(), "uploads"),
  limits: {
    fileSize: 100 * 1024 * 1024, // 100 MB per file
    files: 10, // max 10 files
  },
});

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

  return `${get("year")}-${get("month")}-${get("day")}T${get("hour")}-${get("minute")}-${get(
    "second"
  )}-${ms}+05-30`;
}

function getServiceAccountJson() {
  const b64 = process.env.GOOGLE_SA_JSON_B64;
  if (!b64) {
    throw new Error("Missing GOOGLE_SA_JSON_B64 in environment");
  }

  try {
    return JSON.parse(Buffer.from(b64, "base64").toString("utf8"));
  } catch (e) {
    throw new Error("Invalid GOOGLE_SA_JSON_B64 format");
  }
}

function getAuthClient() {
  const sa = getServiceAccountJson();
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

  const res = await drive.files.list({
    q,
    fields: "files(id,name)",
    ...DRIVE_OPTS,
  });

  if (res.data.files?.length) {
    return res.data.files[0].id;
  }

  const created = await drive.files.create({
    requestBody: {
      name,
      mimeType: "application/vnd.google-apps.folder",
      parents: [parentId],
    },
    fields: "id",
    ...DRIVE_OPTS,
  });

  return created.data.id;
}

async function uploadToDrive(drive, parentId, filePath, driveName) {
  const res = await drive.files.create({
    requestBody: {
      name: driveName,
      parents: [parentId],
    },
    media: {
      body: fs.createReadStream(filePath),
    },
    fields: "id,name",
    ...DRIVE_OPTS,
  });

  return res.data;
}

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

    let out = "";
    let err = "";

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

    p.on("error", (spawnErr) => {
      reject(new Error(`Failed to start Python process: ${spawnErr.message}`));
    });

    p.on("close", (code) => {
      if (code === 0) {
        resolve(out);
      } else {
        reject(new Error(err || out || `Python failed with code ${code}`));
      }
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

function extractPeriodFromLogs(logs) {
  for (const line of logs) {
    const m = String(line).match(/PERIOD\s*:\s*([A-Z]{3}-\d{4})/i);
    if (m) return m[1].toUpperCase();
  }
  return null;
}

function makeRunFolderName(runId, period) {
  return period ? `${runId} (${period})` : runId;
}

// -------------------------
// API — PROJECTS LIST
// -------------------------
app.get("/projects", requireAuth, (req, res) => {
  const out = {};
  Object.entries(PROJECTS).forEach(([name, cfg]) => {
    out[name] = {
      mode: cfg.mode,
      outputName: cfg.outputName,
      needsMonthYear: cfg.needsMonthYear,
    };
  });
  res.json(out);
});

// -------------------------
// API — JOB STATUS
// -------------------------
app.get("/api/job-status", requireAuth, (req, res) => {
  const jobId = String(req.query.jobId || "").trim();
  if (!jobId) {
    return res.status(400).json({ ok: false, error: "jobId required" });
  }

  const job = JOBS.get(jobId);
  if (!job) {
    return res.status(404).json({ ok: false, error: "Job not found" });
  }

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
    logs: job.logs.slice(-100),
    updatedAt: job.updatedAt,
  });
});

// -------------------------
// API — REPORT DATA
// -------------------------
app.get("/api/report-data", async (req, res) => {
  res.setHeader("Content-Type", "application/json");

  if (!req.isAuthenticated || !req.isAuthenticated()) {
    return res.status(401).json({ ok: false, error: "Not authenticated." });
  }

  try {
    const { agent, month, year } = req.query;

    if (!agent) {
      return res.status(400).json({ ok: false, error: "agent param required" });
    }

    if (!REPORT_SHEET_ID) {
      return res.status(500).json({ ok: false, error: "Report sheet is not configured" });
    }

    const tabName = REPORT_SHEET_TABS[agent];
    if (!tabName) {
      return res.status(400).json({ ok: false, error: `No tab configured for agent: ${agent}` });
    }

    const sa = getServiceAccountJson();
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

    return res.json({
      ok: true,
      headers,
      rows: dataRows,
      total: dataRows.length,
    });
  } catch (e) {
    return res.status(500).json({ ok: false, error: e.message });
  }
});

// -------------------------
// API — HELP LINKS
// -------------------------
app.get("/api/help-links", requireAuth, (req, res) => {
  const agents = ["PAY_STK", "PAY_TD", "REC_STK", "REC_TD", "PRI_SALE"];
  const links = {};

  for (const key of agents) {
    links[key] = {
      process: process.env[`HELP_${key}_PROCESS`] || null,
      format: process.env[`HELP_${key}_FORMAT`] || null,
      master: process.env[`HELP_${key}_MASTER`] || null,
    };
  }

  res.json({ ok: true, links });
});

// -------------------------
// API — RUN AGENT (ASYNC JOB)
// -------------------------
app.post("/run/:project", requireAuth, upload.array("files"), async (req, res) => {
  let baseTmp = null;

  try {
    const project = req.params.project;
    const cfg = PROJECTS[project];

    if (!cfg) {
      return res.status(400).json({ ok: false, error: "Unknown project" });
    }

    if (!DRIVE_ROOT_FOLDER_ID) {
      throw new Error("Drive root folder is not configured");
    }

    if (!req.files || req.files.length === 0) {
      throw new Error("No files uploaded");
    }

    if (cfg.mode === "single" && req.files.length !== 1) {
      return res.status(400).json({
        ok: false,
        error: "Exactly one file is required for this project.",
      });
    }

    if (hasRunningJob()) {
      return res.status(429).json({
        ok: false,
        error: "Another job is already running. Please wait until it finishes.",
      });
    }

    const month = cleanMonthYear(req.body?.month).toUpperCase();
    const year = cleanMonthYear(req.body?.year);

    if (cfg.needsMonthYear && (!month || !year)) {
      return res.status(400).json({
        ok: false,
        error: "Month and Year are required for this project.",
      });
    }

    const runId = makeRunIdIST();
    const job = createJob({ req, project, runId });

    res.json({ ok: true, jobId: job.jobId, runId, project });

    (async () => {
      try {
        setJob(job, { status: "running", pct: 5, hint: "Preparing…" });
        addLog(job, `-> Agent: ${project}`);

        if (cfg.needsMonthYear) {
          addLog(job, `-> Period: ${month} / ${year}`);
        }

        addLog(job, `-> Files: ${req.files.length} file(s)`);
        logMemory("before drive init");

        const drive = getDriveClient();

        setJob(job, { pct: 10, hint: "Creating Drive folders…" });

        const projectFolderId = await ensureFolder(drive, DRIVE_ROOT_FOLDER_ID, project);
        const inputRootId = await ensureFolder(drive, projectFolderId, "input");
        const outputRootId = await ensureFolder(drive, projectFolderId, "output");
        const inputRunId = await ensureFolder(drive, inputRootId, runId);

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
        if (!fs.existsSync(scriptAbs)) {
          throw new Error(`Script not found: ${scriptAbs}`);
        }

        const wantsOutputName = Boolean(cfg.outputName && cfg.supportsOutputNameArg);

        setJob(job, { pct: 30, hint: "Running agent…" });
        addLog(job, "-> Processing started…");
        logMemory("before python");

        if (cfg.mode === "single") {
          const inputFile = path.join(inDir, req.files[0].originalname);
          const args = ["--input", inputFile, "--output_dir", outDir];

          if (wantsOutputName) {
            args.push("--output_name", cfg.outputName);
          }

          if (cfg.needsMonthYear) {
            args.push("--month", month, "--year", year);
          }

          await runPython(scriptAbs, args, job);
        } else {
          const args = ["--input_dir", inDir, "--output_dir", outDir];

          if (wantsOutputName) {
            args.push("--output_name", cfg.outputName);
          }

          if (cfg.needsMonthYear) {
            args.push("--month", month, "--year", year);
          }

          await runPython(scriptAbs, args, job);
        }

        logMemory("after python");

        const period = extractPeriodFromLogs(job.logs);
        addLog(
          job,
          period ? `-> Data period detected: ${period}` : "-> Data period not detected, using plain runId"
        );

        const outputFolderName = makeRunFolderName(runId, period);
        const outputRunId = await ensureFolder(drive, outputRootId, outputFolderName);

        setJob(job, { status: "uploading", pct: 90, hint: "Uploading outputs…" });
        logMemory("before output upload");

        const outputFiles = fs.readdirSync(outDir);
        const uploaded = [];

        for (const name of outputFiles) {
          const fpath = path.join(outDir, name);
          const meta = await uploadToDrive(drive, outputRunId, fpath, name);
          uploaded.push(meta);
        }

        job.outputs = uploaded;

        logMemory("after output upload");

        setJob(job, { status: "done", pct: 100, hint: "Completed" });
        addLog(job, "-> Completed");
      } catch (e) {
        job.error = e.message;
        setJob(job, { status: "error", hint: "Failed" });
        addLog(job, "-> Error: " + e.message);
      } finally {
        if (baseTmp) {
          safeCleanup(baseTmp);
        }
      }
    })();
  } catch (e) {
    if (baseTmp) {
      safeCleanup(baseTmp);
    }
    return res.status(500).json({ ok: false, error: e.message });
  }
});

// -------------------------
// UPLOAD ERROR HANDLER
// -------------------------
app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    if (err.code === "LIMIT_FILE_SIZE") {
      return res.status(400).json({
        ok: false,
        error: "One or more files exceed the 100 MB limit.",
      });
    }

    if (err.code === "LIMIT_FILE_COUNT") {
      return res.status(400).json({
        ok: false,
        error: "Too many files uploaded. Maximum allowed is 10 files.",
      });
    }

    return res.status(400).json({
      ok: false,
      error: `Upload error: ${err.message}`,
    });
  }

  return next(err);
});

// -------------------------
// BASE DUMP ROUTES
// -------------------------
async function listDriveFolder(drive, folderId) {
  const q = `'${folderId}' in parents and trashed=false`;

  const res = await drive.files.list({
    q,
    fields: "files(id,name,mimeType,size,modifiedTime)",
    orderBy: "name desc",
    pageSize: 200,
    ...DRIVE_OPTS,
  });

  return res.data.files || [];
}

app.get("/api/basedump/runs", requireAuth, async (req, res) => {
  try {
    if (!BASE_DUMP_FOLDER_ID) {
      return res.status(500).json({ ok: false, error: "Base dump folder is not configured" });
    }

    const { agent, month, year } = req.query;

    if (!agent) {
      return res.status(400).json({ ok: false, error: "agent param required" });
    }

    const drive = getDriveClient();

    const agentQ = `'${BASE_DUMP_FOLDER_ID}' in parents and name='${agent}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
    const agentRes = await drive.files.list({
      q: agentQ,
      fields: "files(id,name)",
      ...DRIVE_OPTS,
    });

    const agentFolders = agentRes.data.files || [];
    if (!agentFolders.length) {
      return res.json({ ok: true, runs: [], total: 0 });
    }

    const agentFolderId = agentFolders[0].id;

    const outQ = `'${agentFolderId}' in parents and name='output' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
    const outRes = await drive.files.list({
      q: outQ,
      fields: "files(id,name)",
      ...DRIVE_OPTS,
    });

    const outFolders = outRes.data.files || [];
    if (!outFolders.length) {
      return res.json({ ok: true, runs: [], total: 0 });
    }

    const outputRootId = outFolders[0].id;
    const runFolders = await listDriveFolder(drive, outputRootId);

    function parseRunFolder(name) {
      const match = name.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2})-(\d{2})-(\d{2})/);
      if (!match) return null;

      const periodMatch = name.match(/\(([A-Z]{3}-\d{4})\)/i);
      const period = periodMatch ? periodMatch[1].toUpperCase() : null;

      const label = period
        ? `${match[3]}/${match[2]}/${match[1]} ${match[4]}:${match[5]}:${match[6]} IST [${period}]`
        : `${match[3]}/${match[2]}/${match[1]} ${match[4]}:${match[5]}:${match[6]} IST`;

      return {
        year: match[1],
        month: match[2],
        day: match[3],
        hour: match[4],
        minute: match[5],
        second: match[6],
        period,
        label,
      };
    }

    let runs = runFolders
      .filter((f) => f.mimeType === "application/vnd.google-apps.folder")
      .map((f) => ({
        id: f.id,
        name: f.name,
        parsed: parseRunFolder(f.name),
      }))
      .filter((r) => r.parsed !== null);

    const fm = (month || "").trim();
    const fy = (year || "").trim();

    if (fm) {
      runs = runs.filter((r) => r.parsed.month === fm.padStart(2, "0"));
    }

    if (fy) {
      runs = runs.filter((r) => r.parsed.year === fy);
    }

    return res.json({ ok: true, runs, total: runs.length });
  } catch (e) {
    return res.status(500).json({ ok: false, error: e.message });
  }
});

app.get("/api/basedump/zip", requireAuth, async (req, res) => {
  try {
    const { folderId, folderName } = req.query;

    if (!folderId) {
      return res.status(400).json({ ok: false, error: "folderId required" });
    }

    const drive = getDriveClient();
    const files = await listDriveFolder(drive, folderId);
    const fileItems = files.filter((f) => f.mimeType !== "application/vnd.google-apps.folder");

    if (!fileItems.length) {
      return res.status(404).json({ ok: false, error: "No files found in this run folder" });
    }

    const zipName = (folderName || folderId) + ".zip";

    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="${zipName}"`);

    const archive = archiver("zip", { zlib: { level: 6 } });
    archive.pipe(res);

    for (const file of fileItems) {
      const stream = await drive.files.get(
        {
          fileId: file.id,
          alt: "media",
          ...DRIVE_OPTS,
        },
        { responseType: "stream" }
      );

      archive.append(stream.data, { name: file.name });
    }

    archive.finalize();
  } catch (e) {
    if (!res.headersSent) {
      res.status(500).json({ ok: false, error: e.message });
    }
  }
});

app.get("/api/basedump/debug", requireAuth, async (req, res) => {
  try {
    if (!BASE_DUMP_FOLDER_ID) {
      return res.status(500).json({ ok: false, error: "Base dump folder is not configured" });
    }

    const drive = getDriveClient();
    const q = `'${BASE_DUMP_FOLDER_ID}' in parents and trashed=false`;

    const r = await drive.files.list({
      q,
      fields: "files(id,name,mimeType)",
      ...DRIVE_OPTS,
    });

    res.json({
      root: BASE_DUMP_FOLDER_ID,
      files: r.data.files,
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// -------------------------
// START SERVER
// -------------------------
app.listen(PORT, "0.0.0.0", () => {
  console.log(`[server] Running on port ${PORT}`);
});