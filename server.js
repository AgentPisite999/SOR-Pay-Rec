

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
// import archiver from "archiver";

// const __filename = fileURLToPath(import.meta.url);
// const __dirname = path.dirname(__filename);

// const app = express();

// // Important for Render / reverse proxies
// app.set("trust proxy", 1);

// // -------------------------
// // BASIC CONFIG
// // -------------------------
// const isProd = process.env.NODE_ENV === "production";
// const PORT = process.env.PORT || 3000;
// const PYTHON_BIN = process.env.PYTHON_BIN || "python3";
// const DRIVE_ROOT_FOLDER_ID = process.env.DRIVE_ROOT_FOLDER_ID;
// const REPORT_SHEET_ID = process.env.REPORT_SHEET_ID;
// const BASE_DUMP_FOLDER_ID = process.env.BASE_DUMP_FOLDER_ID;

// const DRIVE_OPTS = {
//   supportsAllDrives: true,
//   includeItemsFromAllDrives: true,
// };

// if (isProd && !process.env.SESSION_SECRET) {
//   throw new Error("SESSION_SECRET is required in production");
// }

// // -------------------------
// // JOBS — temp file cache only
// // not guaranteed to survive restart/redeploy
// // -------------------------
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
//         if (v && typeof v === "object" && v.jobId && v.status) {
//           m.set(k, v);
//         }
//       });

//       console.log("[jobs] Loaded " + m.size + " jobs from temp storage");
//       return m;
//     }
//   } catch (e) {
//     console.error("[jobs] Corrupt jobs file, resetting:", e.message);
//     try {
//       fs.unlinkSync(JOBS_FILE);
//     } catch {}
//   }

//   return new Map();
// }

// function _saveJobs(jobs) {
//   try {
//     const obj = {};
//     jobs.forEach((v, k) => {
//       obj[k] = v;
//     });
//     fs.writeFileSync(JOBS_FILE, JSON.stringify(obj), "utf8");
//   } catch (e) {
//     console.error("[jobs] Failed to save to temp storage:", e.message);
//   }
// }

// const JOBS = _loadJobs();

// (function cleanOldJobs() {
//   const cutoff = Date.now() - 86400000; // 24 hours
//   let n = 0;

//   JOBS.forEach((v, k) => {
//     if (new Date(v.createdAt).getTime() < cutoff) {
//       JOBS.delete(k);
//       n++;
//     }
//   });

//   if (n) {
//     _saveJobs(JOBS);
//     console.log(`[jobs] Cleaned ${n} old jobs`);
//   }
// })();

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
//     status: "queued",
//     pct: 0,
//     hint: "Queued…",
//     logs: [],
//     outputs: [],
//     error: null,
//     createdAt: nowISO(),
//     updatedAt: nowISO(),
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
//   if (job.logs.length > 150) {
//     job.logs = job.logs.slice(-150);
//   }

//   job.updatedAt = nowISO();
//   _saveJobs(JOBS);
// }

// function hasRunningJob() {
//   for (const [, job] of JOBS) {
//     if (job.status === "queued" || job.status === "running" || job.status === "uploading") {
//       return true;
//     }
//   }
//   return false;
// }

// function logMemory(prefix = "") {
//   const m = process.memoryUsage();
//   console.log(
//     `[mem] ${prefix} rss=${Math.round(m.rss / 1024 / 1024)}MB ` +
//       `heapUsed=${Math.round(m.heapUsed / 1024 / 1024)}MB ` +
//       `heapTotal=${Math.round(m.heapTotal / 1024 / 1024)}MB ` +
//       `external=${Math.round(m.external / 1024 / 1024)}MB`
//   );
// }

// // -------------------------
// // BASIC MIDDLEWARES
// // -------------------------
// app.use(express.urlencoded({ extended: true }));
// app.use(express.json());

// app.get("/healthz", (req, res) => {
//   res.status(200).send("ok");
// });

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
//       secure: isProd,
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
// app.use(
//   express.static(path.join(__dirname, "public"), {
//     index: false,
//   })
// );

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
// const upload = multer({
//   dest: path.join(os.tmpdir(), "uploads"),
//   limits: {
//     fileSize: 100 * 1024 * 1024, // 100 MB per file
//     files: 10, // max 10 files
//   },
// });

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
//   // -------------------------
//   // NEW: GRC Report
//   // -------------------------
//   "GRC Report": {
//     script: "agents/grc_report.py",
//     mode: "single",
//     outputName: "GRC_OUTPUT.xlsx",
//     supportsOutputNameArg: true,
//     needsMonthYear: false,        // month/year auto-extracted from file
//     driveFolderName: "GRC FILE", // custom Drive folder name
//   },
// };

// // -------------------------
// // REPORT SHEET CONFIG
// // -------------------------
// const REPORT_SHEET_TABS = {
//   "Payable Stock": process.env.PAY_STK_TAB_NAME || "PAY_STK",
//   "Payable Trade Discount": process.env.PAY_TD_TAB_NAME || "PAY_TD",
//   "Receivable Stock": process.env.REC_STK_TAB_NAME || "REC_STK",
//   "Receivable Trade Discount": process.env.REC_TD_TAB_NAME || "REC_TD",
//   "Primary Sale": process.env.PRIM_SALE_TAB_NAME || "PRIM_SALE",
//   "GRC Report": process.env.GRC_TAB_NAME || "GRC",
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

//   return `${get("year")}-${get("month")}-${get("day")}T${get("hour")}-${get("minute")}-${get(
//     "second"
//   )}-${ms}+05-30`;
// }

// function getServiceAccountJson() {
//   const b64 = process.env.GOOGLE_SA_JSON_B64;
//   if (!b64) {
//     throw new Error("Missing GOOGLE_SA_JSON_B64 in environment");
//   }

//   try {
//     return JSON.parse(Buffer.from(b64, "base64").toString("utf8"));
//   } catch (e) {
//     throw new Error("Invalid GOOGLE_SA_JSON_B64 format");
//   }
// }

// function getAuthClient() {
//   const sa = getServiceAccountJson();
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

//   const res = await drive.files.list({
//     q,
//     fields: "files(id,name)",
//     ...DRIVE_OPTS,
//   });

//   if (res.data.files?.length) {
//     return res.data.files[0].id;
//   }

//   const created = await drive.files.create({
//     requestBody: {
//       name,
//       mimeType: "application/vnd.google-apps.folder",
//       parents: [parentId],
//     },
//     fields: "id",
//     ...DRIVE_OPTS,
//   });

//   return created.data.id;
// }

// async function uploadToDrive(drive, parentId, filePath, driveName) {
//   const res = await drive.files.create({
//     requestBody: {
//       name: driveName,
//       parents: [parentId],
//     },
//     media: {
//       body: fs.createReadStream(filePath),
//     },
//     fields: "id,name",
//     ...DRIVE_OPTS,
//   });

//   return res.data;
// }

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

//     let out = "";
//     let err = "";

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

//     p.on("error", (spawnErr) => {
//       reject(new Error(`Failed to start Python process: ${spawnErr.message}`));
//     });

//     p.on("close", (code) => {
//       if (code === 0) {
//         resolve(out);
//       } else {
//         reject(new Error(err || out || `Python failed with code ${code}`));
//       }
//     });
//   });
// }

// function safeCleanup(dir) {
//   try {
//     fs.rmSync(dir, { recursive: true, force: true });
//   } catch {}
// }

// function cleanMonthYear(v) {
//   return String(v ?? "").trim();
// }

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
//     out[name] = {
//       mode: cfg.mode,
//       outputName: cfg.outputName,
//       needsMonthYear: cfg.needsMonthYear,
//     };
//   });
//   res.json(out);
// });

// // -------------------------
// // API — JOB STATUS
// // -------------------------
// app.get("/api/job-status", requireAuth, (req, res) => {
//   const jobId = String(req.query.jobId || "").trim();
//   if (!jobId) {
//     return res.status(400).json({ ok: false, error: "jobId required" });
//   }

//   const job = JOBS.get(jobId);
//   if (!job) {
//     return res.status(404).json({ ok: false, error: "Job not found" });
//   }

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
//     logs: job.logs.slice(-100),
//     updatedAt: job.updatedAt,
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

//     if (!agent) {
//       return res.status(400).json({ ok: false, error: "agent param required" });
//     }

//     // GRC Report uses its own dedicated sheet (GRC_SHEET_ID), not REPORT_SHEET_ID
//     const isGRC = agent === "GRC Report";
//     const sheetId = isGRC
//       ? process.env.GRC_SHEET_ID
//       : REPORT_SHEET_ID;

//     if (!sheetId) {
//       return res.status(500).json({ ok: false, error: "Report sheet is not configured" });
//     }

//     const tabName = REPORT_SHEET_TABS[agent];
//     if (!tabName) {
//       return res.status(400).json({ ok: false, error: `No tab configured for agent: ${agent}` });
//     }

//     const sa = getServiceAccountJson();
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
//       spreadsheetId: sheetId,
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

//     return res.json({
//       ok: true,
//       headers,
//       rows: dataRows,
//       total: dataRows.length,
//     });
//   } catch (e) {
//     return res.status(500).json({ ok: false, error: e.message });
//   }
// });

// // -------------------------
// // API — HELP LINKS
// // -------------------------
// app.get("/api/help-links", requireAuth, (req, res) => {
//   const agents = ["PAY_STK", "PAY_TD", "REC_STK", "REC_TD", "PRI_SALE", "GRC"];
//   const links = {};

//   for (const key of agents) {
//     links[key] = {
//       process: process.env[`HELP_${key}_PROCESS`] || null,
//       format: process.env[`HELP_${key}_FORMAT`] || null,
//       master: process.env[`HELP_${key}_MASTER`] || null,
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

//     if (!cfg) {
//       return res.status(400).json({ ok: false, error: "Unknown project" });
//     }

//     if (!DRIVE_ROOT_FOLDER_ID) {
//       throw new Error("Drive root folder is not configured");
//     }

//     if (!req.files || req.files.length === 0) {
//       throw new Error("No files uploaded");
//     }

//     if (cfg.mode === "single" && req.files.length !== 1) {
//       return res.status(400).json({
//         ok: false,
//         error: "Exactly one file is required for this project.",
//       });
//     }

//     if (hasRunningJob()) {
//       return res.status(429).json({
//         ok: false,
//         error: "Another job is already running. Please wait until it finishes.",
//       });
//     }

//     const month = cleanMonthYear(req.body?.month).toUpperCase();
//     const year = cleanMonthYear(req.body?.year);

//     if (cfg.needsMonthYear && (!month || !year)) {
//       return res.status(400).json({
//         ok: false,
//         error: "Month and Year are required for this project.",
//       });
//     }

//     const runId = makeRunIdIST();
//     const job = createJob({ req, project, runId });

//     res.json({ ok: true, jobId: job.jobId, runId, project });

//     (async () => {
//       try {
//         setJob(job, { status: "running", pct: 5, hint: "Preparing…" });
//         addLog(job, `-> Agent: ${project}`);

//         if (cfg.needsMonthYear) {
//           addLog(job, `-> Period: ${month} / ${year}`);
//         }

//         addLog(job, `-> Files: ${req.files.length} file(s)`);
//         logMemory("before drive init");

//         const drive = getDriveClient();

//         setJob(job, { pct: 10, hint: "Creating Drive folders…" });

//         // Use custom driveFolderName if set (e.g. "GRC FILE"), else use project name
//         const rootFolderName = cfg.driveFolderName || project;
//         const projectFolderId = await ensureFolder(drive, DRIVE_ROOT_FOLDER_ID, rootFolderName);
//         const inputRootId = await ensureFolder(drive, projectFolderId, "input");
//         const outputRootId = await ensureFolder(drive, projectFolderId, "output");
//         const inputRunId = await ensureFolder(drive, inputRootId, runId);

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
//         if (!fs.existsSync(scriptAbs)) {
//           throw new Error(`Script not found: ${scriptAbs}`);
//         }

//         const wantsOutputName = Boolean(cfg.outputName && cfg.supportsOutputNameArg);

//         setJob(job, { pct: 30, hint: "Running agent…" });
//         addLog(job, "-> Processing started…");
//         logMemory("before python");

//         if (cfg.mode === "single") {
//           const inputFile = path.join(inDir, req.files[0].originalname);
//           const args = ["--input", inputFile, "--output_dir", outDir];

//           if (wantsOutputName) {
//             args.push("--output_name", cfg.outputName);
//           }

//           if (cfg.needsMonthYear) {
//             args.push("--month", month, "--year", year);
//           }

//           await runPython(scriptAbs, args, job);
//         } else {
//           const args = ["--input_dir", inDir, "--output_dir", outDir];

//           if (wantsOutputName) {
//             args.push("--output_name", cfg.outputName);
//           }

//           if (cfg.needsMonthYear) {
//             args.push("--month", month, "--year", year);
//           }

//           await runPython(scriptAbs, args, job);
//         }

//         logMemory("after python");

//         const period = extractPeriodFromLogs(job.logs);
//         addLog(
//           job,
//           period ? `-> Data period detected: ${period}` : "-> Data period not detected, using plain runId"
//         );

//         const outputFolderName = makeRunFolderName(runId, period);
//         const outputRunId = await ensureFolder(drive, outputRootId, outputFolderName);

//         setJob(job, { status: "uploading", pct: 90, hint: "Uploading outputs…" });
//         logMemory("before output upload");

//         const outputFiles = fs.readdirSync(outDir);
//         const uploaded = [];

//         for (const name of outputFiles) {
//           const fpath = path.join(outDir, name);
//           const meta = await uploadToDrive(drive, outputRunId, fpath, name);
//           uploaded.push(meta);
//         }

//         job.outputs = uploaded;

//         logMemory("after output upload");

//         setJob(job, { status: "done", pct: 100, hint: "Completed" });
//         addLog(job, "-> Completed");
//       } catch (e) {
//         job.error = e.message;
//         setJob(job, { status: "error", hint: "Failed" });
//         addLog(job, "-> Error: " + e.message);
//       } finally {
//         if (baseTmp) {
//           safeCleanup(baseTmp);
//         }
//       }
//     })();
//   } catch (e) {
//     if (baseTmp) {
//       safeCleanup(baseTmp);
//     }
//     return res.status(500).json({ ok: false, error: e.message });
//   }
// });

// // -------------------------
// // UPLOAD ERROR HANDLER
// // -------------------------
// app.use((err, req, res, next) => {
//   if (err instanceof multer.MulterError) {
//     if (err.code === "LIMIT_FILE_SIZE") {
//       return res.status(400).json({
//         ok: false,
//         error: "One or more files exceed the 100 MB limit.",
//       });
//     }

//     if (err.code === "LIMIT_FILE_COUNT") {
//       return res.status(400).json({
//         ok: false,
//         error: "Too many files uploaded. Maximum allowed is 10 files.",
//       });
//     }

//     return res.status(400).json({
//       ok: false,
//       error: `Upload error: ${err.message}`,
//     });
//   }

//   return next(err);
// });

// // -------------------------
// // BASE DUMP ROUTES
// // -------------------------
// async function listDriveFolder(drive, folderId) {
//   const q = `'${folderId}' in parents and trashed=false`;

//   const res = await drive.files.list({
//     q,
//     fields: "files(id,name,mimeType,size,modifiedTime)",
//     orderBy: "name desc",
//     pageSize: 200,
//     ...DRIVE_OPTS,
//   });

//   return res.data.files || [];
// }

// app.get("/api/basedump/runs", requireAuth, async (req, res) => {
//   try {
//     if (!BASE_DUMP_FOLDER_ID) {
//       return res.status(500).json({ ok: false, error: "Base dump folder is not configured" });
//     }

//     const { agent, month, year } = req.query;

//     if (!agent) {
//       return res.status(400).json({ ok: false, error: "agent param required" });
//     }

//     const drive = getDriveClient();

//     // Map agent name → actual Drive folder name (e.g. "GRC Report" → "GRC FILE")
//     const cfg = PROJECTS[agent];
//     const agentFolderName = cfg?.driveFolderName || agent;

//     const agentQ = `'${BASE_DUMP_FOLDER_ID}' in parents and name='${agentFolderName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
//     const agentRes = await drive.files.list({
//       q: agentQ,
//       fields: "files(id,name)",
//       ...DRIVE_OPTS,
//     });

//     const agentFolders = agentRes.data.files || [];
//     if (!agentFolders.length) {
//       return res.json({ ok: true, runs: [], total: 0 });
//     }

//     const agentFolderId = agentFolders[0].id;

//     const outQ = `'${agentFolderId}' in parents and name='output' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
//     const outRes = await drive.files.list({
//       q: outQ,
//       fields: "files(id,name)",
//       ...DRIVE_OPTS,
//     });

//     const outFolders = outRes.data.files || [];
//     if (!outFolders.length) {
//       return res.json({ ok: true, runs: [], total: 0 });
//     }

//     const outputRootId = outFolders[0].id;
//     const runFolders = await listDriveFolder(drive, outputRootId);

//     function parseRunFolder(name) {
//       const match = name.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2})-(\d{2})-(\d{2})/);
//       if (!match) return null;

//       const periodMatch = name.match(/\(([A-Z]{3}-\d{4})\)/i);
//       const period = periodMatch ? periodMatch[1].toUpperCase() : null;

//       const label = period
//         ? `${match[3]}/${match[2]}/${match[1]} ${match[4]}:${match[5]}:${match[6]} IST [${period}]`
//         : `${match[3]}/${match[2]}/${match[1]} ${match[4]}:${match[5]}:${match[6]} IST`;

//       return {
//         year: match[1],
//         month: match[2],
//         day: match[3],
//         hour: match[4],
//         minute: match[5],
//         second: match[6],
//         period,
//         label,
//       };
//     }

//     let runs = runFolders
//       .filter((f) => f.mimeType === "application/vnd.google-apps.folder")
//       .map((f) => ({
//         id: f.id,
//         name: f.name,
//         parsed: parseRunFolder(f.name),
//       }))
//       .filter((r) => r.parsed !== null);

//     const fm = (month || "").trim();
//     const fy = (year || "").trim();

//     if (fm) {
//       runs = runs.filter((r) => r.parsed.month === fm.padStart(2, "0"));
//     }

//     if (fy) {
//       runs = runs.filter((r) => r.parsed.year === fy);
//     }

//     return res.json({ ok: true, runs, total: runs.length });
//   } catch (e) {
//     return res.status(500).json({ ok: false, error: e.message });
//   }
// });

// app.get("/api/basedump/zip", requireAuth, async (req, res) => {
//   try {
//     const { folderId, folderName } = req.query;

//     if (!folderId) {
//       return res.status(400).json({ ok: false, error: "folderId required" });
//     }

//     const drive = getDriveClient();
//     const files = await listDriveFolder(drive, folderId);
//     const fileItems = files.filter((f) => f.mimeType !== "application/vnd.google-apps.folder");

//     if (!fileItems.length) {
//       return res.status(404).json({ ok: false, error: "No files found in this run folder" });
//     }

//     const zipName = (folderName || folderId) + ".zip";

//     res.setHeader("Content-Type", "application/zip");
//     res.setHeader("Content-Disposition", `attachment; filename="${zipName}"`);

//     const archive = archiver("zip", { zlib: { level: 6 } });
//     archive.pipe(res);

//     for (const file of fileItems) {
//       const stream = await drive.files.get(
//         {
//           fileId: file.id,
//           alt: "media",
//           ...DRIVE_OPTS,
//         },
//         { responseType: "stream" }
//       );

//       archive.append(stream.data, { name: file.name });
//     }

//     archive.finalize();
//   } catch (e) {
//     if (!res.headersSent) {
//       res.status(500).json({ ok: false, error: e.message });
//     }
//   }
// });

// app.get("/api/basedump/debug", requireAuth, async (req, res) => {
//   try {
//     if (!BASE_DUMP_FOLDER_ID) {
//       return res.status(500).json({ ok: false, error: "Base dump folder is not configured" });
//     }

//     const drive = getDriveClient();
//     const q = `'${BASE_DUMP_FOLDER_ID}' in parents and trashed=false`;

//     const r = await drive.files.list({
//       q,
//       fields: "files(id,name,mimeType)",
//       ...DRIVE_OPTS,
//     });

//     res.json({
//       root: BASE_DUMP_FOLDER_ID,
//       files: r.data.files,
//     });
//   } catch (e) {
//     res.status(500).json({ error: e.message });
//   }
// });

// // -------------------------
// // START SERVER
// // -------------------------
// app.listen(PORT, "0.0.0.0", () => {
//   console.log(`[server] Running on port ${PORT}`);
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
// JOBS
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
        if (v && typeof v === "object" && v.jobId && v.status) m.set(k, v);
      });
      console.log("[jobs] Loaded " + m.size + " jobs from temp storage");
      return m;
    }
  } catch (e) {
    console.error("[jobs] Corrupt jobs file, resetting:", e.message);
    try { fs.unlinkSync(JOBS_FILE); } catch {}
  }
  return new Map();
}

function _saveJobs(jobs) {
  try {
    const obj = {};
    jobs.forEach((v, k) => { obj[k] = v; });
    fs.writeFileSync(JOBS_FILE, JSON.stringify(obj), "utf8");
  } catch (e) {
    console.error("[jobs] Failed to save to temp storage:", e.message);
  }
}

const JOBS = _loadJobs();

(function cleanOldJobs() {
  const cutoff = Date.now() - 86400000;
  let n = 0;
  JOBS.forEach((v, k) => {
    if (new Date(v.createdAt).getTime() < cutoff) { JOBS.delete(k); n++; }
  });
  if (n) { _saveJobs(JOBS); console.log(`[jobs] Cleaned ${n} old jobs`); }
})();

function makeJobId() {
  return "job_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
}
function nowISO() { return new Date().toISOString(); }
function userKey(req) { return req.user?.email || req.user?.id || "anon"; }

function createJob({ req, project, runId }) {
  const jobId = makeJobId();
  const job = {
    jobId, userKey: userKey(req), project, runId,
    status: "queued", pct: 0, hint: "Queued…",
    logs: [], outputs: [], error: null,
    createdAt: nowISO(), updatedAt: nowISO(),
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
  if (job.logs.length > 150) job.logs = job.logs.slice(-150);
  job.updatedAt = nowISO();
  _saveJobs(JOBS);
}

function hasRunningJob() {
  for (const [, job] of JOBS) {
    if (job.status === "queued" || job.status === "running" || job.status === "uploading") return true;
  }
  return false;
}

function logMemory(prefix = "") {
  const m = process.memoryUsage();
  console.log(
    `[mem] ${prefix} rss=${Math.round(m.rss/1024/1024)}MB ` +
    `heapUsed=${Math.round(m.heapUsed/1024/1024)}MB ` +
    `heapTotal=${Math.round(m.heapTotal/1024/1024)}MB ` +
    `external=${Math.round(m.external/1024/1024)}MB`
  );
}

// -------------------------
// BASIC MIDDLEWARES
// -------------------------
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

app.get("/healthz", (req, res) => res.status(200).send("ok"));

// -------------------------
// SESSION + PASSPORT
// -------------------------
app.use(session({
  name: "sid",
  secret: process.env.SESSION_SECRET || "dev_secret_change_me",
  resave: false,
  saveUninitialized: false,
  cookie: { httpOnly: true, sameSite: "lax", secure: isProd, maxAge: 7 * 24 * 60 * 60 * 1000 },
}));
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
app.get("/login", (req, res) => res.sendFile(path.join(__dirname, "public", "login.html")));
app.get("/", requireAuth, (req, res) => res.sendFile(path.join(__dirname, "public", "index.html")));
app.get("/me", requireAuth, (req, res) => res.json({ ok: true, user: req.user }));

// -------------------------
// MULTER
// -------------------------
const upload = multer({
  dest: path.join(os.tmpdir(), "uploads"),
  limits: { fileSize: 100 * 1024 * 1024, files: 10 },
});

// -------------------------
// PROJECT CONFIG
// -------------------------
const PROJECTS = {
  "Payable Stock": {
    script: "agents/Payable_STK.py", mode: "single",
    outputName: "Payable_STK.xlsx", supportsOutputNameArg: true, needsMonthYear: false,
  },
  "Payable Trade Discount": {
    script: "agents/payable_TD.py", mode: "single",
    outputName: "Payable_TD.xlsx", supportsOutputNameArg: true, needsMonthYear: false,
  },
  "Receivable Stock": {
    script: "agents/rec_stk.py", mode: "multi",
    outputName: "REC_STK_MASTER.xlsx", supportsOutputNameArg: true, needsMonthYear: true,
  },
  "Receivable Trade Discount": {
    script: "agents/rec_TD.py", mode: "multi",
    outputName: "MASTER_OUTPUT.xlsx", supportsOutputNameArg: true, needsMonthYear: true,
  },
  "Primary Sale": {
    script: "agents/primary_sale.py", mode: "single",
    outputName: "Output.xlsx", supportsOutputNameArg: true, needsMonthYear: false,
  },
  "GRC Report": {
    script: "agents/grc_report.py", mode: "single",
    outputName: "GRC_OUTPUT.xlsx", supportsOutputNameArg: true,
    needsMonthYear: false, driveFolderName: "GRC FILE",
  },
};

// -------------------------
// REPORT SHEET CONFIG
// -------------------------
const REPORT_SHEET_TABS = {
  "Payable Stock":            process.env.PAY_STK_TAB_NAME    || "PAY_STK",
  "Payable Trade Discount":   process.env.PAY_TD_TAB_NAME     || "PAY_TD",
  "Receivable Stock":         process.env.REC_STK_TAB_NAME    || "REC_STK",
  "Receivable Trade Discount":process.env.REC_TD_TAB_NAME     || "REC_TD",
  "Primary Sale":             process.env.PRIM_SALE_TAB_NAME  || "PRIM_SALE",
  "GRC Report":               process.env.GRC_TAB_NAME        || "GRC",
};

// -------------------------
// HELPERS
// -------------------------
function makeRunIdIST() {
  const now = new Date();
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Asia/Kolkata", year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false,
  }).formatToParts(now);
  const get = (t) => parts.find((p) => p.type === t)?.value;
  const ms = String(now.getMilliseconds()).padStart(3, "0");
  return `${get("year")}-${get("month")}-${get("day")}T${get("hour")}-${get("minute")}-${get("second")}-${ms}+05-30`;
}

function getServiceAccountJson() {
  const b64 = process.env.GOOGLE_SA_JSON_B64;
  if (!b64) throw new Error("Missing GOOGLE_SA_JSON_B64 in environment");
  try { return JSON.parse(Buffer.from(b64, "base64").toString("utf8")); }
  catch (e) { throw new Error("Invalid GOOGLE_SA_JSON_B64 format"); }
}

function getAuthClient() {
  const sa = getServiceAccountJson();
  return new google.auth.GoogleAuth({
    credentials: sa,
    scopes: ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets.readonly"],
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
    fields: "id", ...DRIVE_OPTS,
  });
  return created.data.id;
}

async function uploadToDrive(drive, parentId, filePath, driveName) {
  const res = await drive.files.create({
    requestBody: { name: driveName, parents: [parentId] },
    media: { body: fs.createReadStream(filePath) },
    fields: "id,name", ...DRIVE_OPTS,
  });
  return res.data;
}

function runPython(scriptAbsPath, args, job) {
  return new Promise((resolve, reject) => {
    const p = spawn(PYTHON_BIN, [scriptAbsPath, ...args], {
      stdio: "pipe", windowsHide: true,
      env: { ...process.env, PYTHONIOENCODING: "utf-8", PYTHONUTF8: "1" },
    });
    let out = "", err = "";
    p.stdout.on("data", (d) => { const s = d.toString("utf8"); out += s; if (job) addLog(job, s); });
    p.stderr.on("data", (d) => { const s = d.toString("utf8"); err += s; if (job) addLog(job, s); });
    p.on("error", (spawnErr) => reject(new Error(`Failed to start Python process: ${spawnErr.message}`)));
    p.on("close", (code) => {
      if (code === 0) resolve(out);
      else reject(new Error(err || out || `Python failed with code ${code}`));
    });
  });
}

function safeCleanup(dir) { try { fs.rmSync(dir, { recursive: true, force: true }); } catch {} }
function cleanMonthYear(v) { return String(v ?? "").trim(); }

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
    out[name] = { mode: cfg.mode, outputName: cfg.outputName, needsMonthYear: cfg.needsMonthYear };
  });
  res.json(out);
});

// -------------------------
// API — JOB STATUS
// -------------------------
app.get("/api/job-status", requireAuth, (req, res) => {
  const jobId = String(req.query.jobId || "").trim();
  if (!jobId) return res.status(400).json({ ok: false, error: "jobId required" });
  const job = JOBS.get(jobId);
  if (!job) return res.status(404).json({ ok: false, error: "Job not found" });
  if (job.userKey !== userKey(req)) return res.status(403).json({ ok: false, error: "Forbidden" });
  return res.json({
    ok: true, jobId: job.jobId, project: job.project, runId: job.runId,
    status: job.status, pct: job.pct, hint: job.hint, outputs: job.outputs,
    error: job.error, logs: job.logs.slice(-100), updatedAt: job.updatedAt,
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
    if (!agent) return res.status(400).json({ ok: false, error: "agent param required" });

    const isGRC = agent === "GRC Report";
    const sheetId = isGRC ? process.env.GRC_SHEET_ID : REPORT_SHEET_ID;
    if (!sheetId) return res.status(500).json({ ok: false, error: "Report sheet is not configured" });

    const tabName = REPORT_SHEET_TABS[agent];
    if (!tabName) return res.status(400).json({ ok: false, error: `No tab configured for agent: ${agent}` });

    const sa = getServiceAccountJson();
    const authClient = new google.auth.GoogleAuth({
      credentials: { client_email: sa.client_email, private_key: sa.private_key },
      scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
    });
    const client = await authClient.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });
    const response = await sheets.spreadsheets.values.get({ spreadsheetId: sheetId, range: tabName });

    const rows = response.data.values || [];
    if (rows.length < 2) return res.json({ ok: true, headers: [], rows: [], total: 0 });

    const headers = rows[0].map((h) => String(h || "").trim());
    let dataRows = rows.slice(1);

    const monthIdx = headers.findIndex((h) => h.toUpperCase() === "MONTH");
    const yearIdx  = headers.findIndex((h) => h.toUpperCase() === "YEAR");
    const filterMonth = (month || "ALL").trim().toUpperCase();
    const filterYear  = (year  || "ALL").trim().toUpperCase();

    if (filterMonth !== "ALL" && monthIdx !== -1)
      dataRows = dataRows.filter((r) => (r[monthIdx] || "").trim().toUpperCase() === filterMonth);
    if (filterYear !== "ALL" && yearIdx !== -1)
      dataRows = dataRows.filter((r) => (r[yearIdx] || "").trim() === filterYear);

    return res.json({ ok: true, headers, rows: dataRows, total: dataRows.length });
  } catch (e) {
    return res.status(500).json({ ok: false, error: e.message });
  }
});

// -------------------------
// API — HELP LINKS
// -------------------------
app.get("/api/help-links", requireAuth, (req, res) => {
  const agents = ["PAY_STK", "PAY_TD", "REC_STK", "REC_TD", "PRI_SALE", "GRC"];
  const links = {};

  // Standard agents — master, format, process
  for (const key of agents) {
    links[key] = {
      process: process.env[`HELP_${key}_PROCESS`] || null,
      format:  process.env[`HELP_${key}_FORMAT`]  || null,
      master:  process.env[`HELP_${key}_MASTER`]  || null,
    };
  }

  // GRC and MOVEMENT share 3 cards:
  // format   → HELP_GRC_SALE_FORMAT   (Input Format)
  // process  → HELP_GRC_SALE_PROCESS  (GRC Process Doc)
  // process2 → HELP_GRC2_SALE_PROCESS (Movement Logic Doc)
  const grcMovement = {
    format:   process.env.HELP_GRC_SALE_FORMAT   || null,
    process:  process.env.HELP_GRC_SALE_PROCESS  || null,
    process2: process.env.HELP_GRC2_SALE_PROCESS || null,
  };
  links["GRC"] = grcMovement;

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
    if (!cfg) return res.status(400).json({ ok: false, error: "Unknown project" });
    if (!DRIVE_ROOT_FOLDER_ID) throw new Error("Drive root folder is not configured");
    if (!req.files || req.files.length === 0) throw new Error("No files uploaded");
    if (cfg.mode === "single" && req.files.length !== 1)
      return res.status(400).json({ ok: false, error: "Exactly one file is required for this project." });
    if (hasRunningJob())
      return res.status(429).json({ ok: false, error: "Another job is already running. Please wait until it finishes." });

    const month = cleanMonthYear(req.body?.month).toUpperCase();
    const year  = cleanMonthYear(req.body?.year);
    if (cfg.needsMonthYear && (!month || !year))
      return res.status(400).json({ ok: false, error: "Month and Year are required for this project." });

    const runId = makeRunIdIST();
    const job   = createJob({ req, project, runId });
    res.json({ ok: true, jobId: job.jobId, runId, project });

    (async () => {
      try {
        setJob(job, { status: "running", pct: 5, hint: "Preparing…" });
        addLog(job, `-> Agent: ${project}`);
        if (cfg.needsMonthYear) addLog(job, `-> Period: ${month} / ${year}`);
        addLog(job, `-> Files: ${req.files.length} file(s)`);
        logMemory("before drive init");

        const drive = getDriveClient();
        setJob(job, { pct: 10, hint: "Creating Drive folders…" });

        const rootFolderName  = cfg.driveFolderName || project;
        const projectFolderId = await ensureFolder(drive, DRIVE_ROOT_FOLDER_ID, rootFolderName);
        const inputRootId     = await ensureFolder(drive, projectFolderId, "input");
        const outputRootId    = await ensureFolder(drive, projectFolderId, "output");
        const inputRunId      = await ensureFolder(drive, inputRootId, runId);

        baseTmp = fs.mkdtempSync(path.join(os.tmpdir(), "job-"));
        const inDir  = path.join(baseTmp, "input");
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
        addLog(job, "-> Processing started…");
        logMemory("before python");

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

        logMemory("after python");
        const period = extractPeriodFromLogs(job.logs);
        addLog(job, period ? `-> Data period detected: ${period}` : "-> Data period not detected, using plain runId");

        const outputFolderName = makeRunFolderName(runId, period);
        const outputRunId      = await ensureFolder(drive, outputRootId, outputFolderName);

        setJob(job, { status: "uploading", pct: 90, hint: "Uploading outputs…" });
        logMemory("before output upload");

        const outputFiles = fs.readdirSync(outDir);
        const uploaded = [];
        for (const name of outputFiles) {
          const fpath = path.join(outDir, name);
          const meta  = await uploadToDrive(drive, outputRunId, fpath, name);
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
        if (baseTmp) safeCleanup(baseTmp);
      }
    })();
  } catch (e) {
    if (baseTmp) safeCleanup(baseTmp);
    return res.status(500).json({ ok: false, error: e.message });
  }
});

// -------------------------
// UPLOAD ERROR HANDLER
// -------------------------
app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    if (err.code === "LIMIT_FILE_SIZE")
      return res.status(400).json({ ok: false, error: "One or more files exceed the 100 MB limit." });
    if (err.code === "LIMIT_FILE_COUNT")
      return res.status(400).json({ ok: false, error: "Too many files uploaded. Maximum allowed is 10 files." });
    return res.status(400).json({ ok: false, error: `Upload error: ${err.message}` });
  }
  return next(err);
});

// -------------------------
// BASE DUMP ROUTES
// -------------------------
async function listDriveFolder(drive, folderId) {
  const q = `'${folderId}' in parents and trashed=false`;
  const res = await drive.files.list({
    q, fields: "files(id,name,mimeType,size,modifiedTime)", orderBy: "name desc", pageSize: 200, ...DRIVE_OPTS,
  });
  return res.data.files || [];
}

app.get("/api/basedump/runs", requireAuth, async (req, res) => {
  try {
    if (!BASE_DUMP_FOLDER_ID)
      return res.status(500).json({ ok: false, error: "Base dump folder is not configured" });

    const { agent, month, year } = req.query;
    if (!agent) return res.status(400).json({ ok: false, error: "agent param required" });

    const drive = getDriveClient();

    // Map agent name → actual Drive folder name (e.g. "GRC Report" → "GRC FILE")
    const cfg = PROJECTS[agent];
    const agentFolderName = cfg?.driveFolderName || agent;

    const agentQ = `'${BASE_DUMP_FOLDER_ID}' in parents and name='${agentFolderName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
    const agentRes = await drive.files.list({ q: agentQ, fields: "files(id,name)", ...DRIVE_OPTS });
    const agentFolders = agentRes.data.files || [];
    if (!agentFolders.length) return res.json({ ok: true, runs: [], total: 0 });

    const agentFolderId = agentFolders[0].id;
    const outQ = `'${agentFolderId}' in parents and name='output' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
    const outRes = await drive.files.list({ q: outQ, fields: "files(id,name)", ...DRIVE_OPTS });
    const outFolders = outRes.data.files || [];
    if (!outFolders.length) return res.json({ ok: true, runs: [], total: 0 });

    const outputRootId = outFolders[0].id;
    const runFolders   = await listDriveFolder(drive, outputRootId);

    function parseRunFolder(name) {
      const match = name.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2})-(\d{2})-(\d{2})/);
      if (!match) return null;
      const periodMatch = name.match(/\(([A-Z]{3}-\d{4})\)/i);
      const period = periodMatch ? periodMatch[1].toUpperCase() : null;
      const label  = period
        ? `${match[3]}/${match[2]}/${match[1]} ${match[4]}:${match[5]}:${match[6]} IST [${period}]`
        : `${match[3]}/${match[2]}/${match[1]} ${match[4]}:${match[5]}:${match[6]} IST`;
      return { year: match[1], month: match[2], day: match[3], hour: match[4], minute: match[5], second: match[6], period, label };
    }

    let runs = runFolders
      .filter((f) => f.mimeType === "application/vnd.google-apps.folder")
      .map((f) => ({ id: f.id, name: f.name, parsed: parseRunFolder(f.name) }))
      .filter((r) => r.parsed !== null);

    const fm = (month || "").trim();
    const fy = (year  || "").trim();
    if (fm) runs = runs.filter((r) => r.parsed.month === fm.padStart(2, "0"));
    if (fy) runs = runs.filter((r) => r.parsed.year  === fy);

    return res.json({ ok: true, runs, total: runs.length });
  } catch (e) {
    return res.status(500).json({ ok: false, error: e.message });
  }
});

app.get("/api/basedump/zip", requireAuth, async (req, res) => {
  try {
    const { folderId, folderName } = req.query;
    if (!folderId) return res.status(400).json({ ok: false, error: "folderId required" });

    const drive = getDriveClient();
    const files = await listDriveFolder(drive, folderId);
    const fileItems = files.filter((f) => f.mimeType !== "application/vnd.google-apps.folder");
    if (!fileItems.length) return res.status(404).json({ ok: false, error: "No files found in this run folder" });

    const zipName = (folderName || folderId) + ".zip";
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="${zipName}"`);

    const archive = archiver("zip", { zlib: { level: 6 } });
    archive.pipe(res);
    for (const file of fileItems) {
      const stream = await drive.files.get(
        { fileId: file.id, alt: "media", ...DRIVE_OPTS },
        { responseType: "stream" }
      );
      archive.append(stream.data, { name: file.name });
    }
    archive.finalize();
  } catch (e) {
    if (!res.headersSent) res.status(500).json({ ok: false, error: e.message });
  }
});

app.get("/api/basedump/debug", requireAuth, async (req, res) => {
  try {
    if (!BASE_DUMP_FOLDER_ID)
      return res.status(500).json({ ok: false, error: "Base dump folder is not configured" });
    const drive = getDriveClient();
    const r = await drive.files.list({
      q: `'${BASE_DUMP_FOLDER_ID}' in parents and trashed=false`,
      fields: "files(id,name,mimeType)", ...DRIVE_OPTS,
    });
    res.json({ root: BASE_DUMP_FOLDER_ID, files: r.data.files });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// -------------------------
// DASHBOARD — partner normalizer
// -------------------------
const DASHBOARD_PARTNER_MAP = [
  { patterns: ["shoppers stop ltd", "shoppers stop", "shoppers", "shopper"], key: "Shoppers" },
  { patterns: ["lifestyle international pvt", "lifestyle international", "lifestyle"], key: "Lifestyle" },
  { patterns: ["v retail ltd", "v retail", "v-retail", "v - retail"], key: "V Retail" },
  { patterns: ["sabharwals nirankar venture", "sabharwal nirankar venture", "sabharwal nirankar", "nirankar venture", "nirankar"], key: "Sabharwal Nirankar" },
  { patterns: ["rrl-centro", "rrl centro", "reliance retail ltd (rrl-centro)", "reliance centro", "relience centro", "centro"], key: "Reliance Centro" },
  { patterns: ["kora retail private", "kora retail", "kora retails", "kora"], key: "Kora Retail" },
  { patterns: ["leayan global pvt", "leayan global", "leayan", "zuup"], key: "Leayan_Zuup" },
  { patterns: ["myntra jabong india pvt", "myntra jabong india", "myntra jabong", "myntra"], key: "Myntra" },
  { patterns: ["flipkart india pvt", "flipkart india", "flipkart"], key: "Flipkart" },
  { patterns: ["reliance retail ltd ajio sor", "reliance retail ltd_ajio sor", "reliance retail ajio sor", "reliance retail ajio", "ajio sor", "ajio"], key: "Reliance Ajio" },
];

function normPartner(name) {
  const s = String(name || "").toLowerCase().replace(/[-_]+/g, " ").trim();
  for (const entry of DASHBOARD_PARTNER_MAP) {
    if (entry.patterns.some(p => s.includes(p))) return entry.key;
  }
  return null;
}

function getPrevMonth(month, year) {
  const MONTHS = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
  const idx = MONTHS.indexOf(month.toUpperCase());
  if (idx === -1) return null;
  if (idx === 0) return { month: "DEC", year: String(parseInt(year) - 1) };
  return { month: MONTHS[idx - 1], year: String(year) };
}

// -------------------------
// CORE CALCULATION FUNCTION
// Calculates matrix for a given month/year from pre-loaded tab data
// -------------------------
function calcMatrix(recStkRows, recTdRows, grcRows, primRows, M, Y) {

  const PARTNER_KEYS = [
    "Shoppers", "Lifestyle", "V Retail", "Sabharwal Nirankar",
    "Reliance Centro", "Kora Retail", "Leayan_Zuup",
    "Myntra", "Flipkart", "Reliance Ajio",
  ];

  function toNum(v) {
    const n = parseFloat(String(v || "").replace(/,/g, ""));
    return isNaN(n) ? 0 : n;
  }

  // Find FIRST matching row (for REC_STK, REC_TD)
  function findRow(rows, monthCol, yearCol, partnerCol, targetMonth, targetYear, targetPartnerKey) {
    return rows.find(r => {
      const m = String(r[monthCol] || "").trim().toUpperCase();
      const y = String(r[yearCol]  || "").trim();
      const p = normPartner(r[partnerCol] || "");
      return m === targetMonth.toUpperCase() && y === targetYear && p === targetPartnerKey;
    });
  }

  // Sum rows from PRIM_SALE — only SIS Customer group rows (not RTV/return rows)
  // "Groups" column = "SIS Customer" identifies the main sale rows
  // Flexible column matching — trims whitespace from keys
  function sumRows(rows, monthCol, yearCol, partnerCol, targetMonth, targetYear, targetPartnerKey, valueCol) {
    const getVal = (r, col) => {
      if (r[col] !== undefined) return String(r[col] || "").trim();
      const found = Object.keys(r).find(k => k.trim() === col);
      return found ? String(r[found] || "").trim() : "";
    };

    const matched = rows.filter(r => {
      const m = getVal(r, monthCol).toUpperCase();
      const y = getVal(r, yearCol);
      const p = normPartner(getVal(r, partnerCol));

      // Only include SIS Customer group rows (main sale, not RTV/return rows)
      const grp = getVal(r, "Groups").toLowerCase();
      const isSIS = grp === "sis customer";

      return m === targetMonth.toUpperCase() && y === targetYear && p === targetPartnerKey && isSIS;
    });

    return matched.reduce((acc, r) => {
      let raw = r[valueCol] !== undefined ? String(r[valueCol] || "").trim() : "";
      if (!raw) {
        const found = Object.keys(r).find(k => k.trim() === valueCol);
        raw = found ? String(r[found] || "").trim() : "";
      }
      if (raw === "-" || raw === "") return acc;
      return acc + toNum(raw);
    }, 0);
  }

  const prev = getPrevMonth(M, Y);

  const bv = {
    opening_stock: {}, opening_sec_diff: {}, primary: {},
    rate_diff: {}, secondary_sales: {}, markdown: {},
    closing_stock: {}, sales_adj: {}, cogs: {}, cogs_pct: {},
  };

  const qty = {
    opening_stock: {}, opening_sec_diff: {}, primary: {},
    rate_diff: {}, secondary_sales: {}, markdown: {},
    closing_stock: {}, sales_adj: {},
  };

  PARTNER_KEYS.forEach(pKey => {

    // Shared row lookups
    const prevStkRow = prev
      ? findRow(recStkRows, "Month", "Year", "Customer Name", prev.month, prev.year, pKey)
      : null;
    const currStkRow = findRow(recStkRows, "Month", "Year", "Customer Name", M, Y, pKey);
    const tdRow      = findRow(recTdRows,  "Month", "Year", "Partner Name",  M, Y, pKey);
    const grcRow     = grcRows.find(r => {
      const rm = String(r["Month"]      || "").trim().toUpperCase();
      const ry = String(r["Year"]       || "").trim();
      const rp = normPartner(r["Row Labels"] || "");
      return rm === M && ry === Y && rp === pKey;
    });

    // ── BASIC VALUE ──

    // 1. Opening Stock → REC_STK prev month Basic Value
    bv.opening_stock[pKey] = prevStkRow ? toNum(prevStkRow["Basic Value"]) : 0;

    // 2. Opening Secondary Difference → filled from prev month Sales-Adj SOR
    //    (calculated separately and injected after prev month calc)
    bv.opening_sec_diff[pKey] = 0; // placeholder, overwritten below

    // 3. Primary → PRIM_SALE sum of Net col for ALL matching rows
    // Primary → PRIM_SALE sum of Net col (try "Net" and "Net " with trailing space)
    const primNetCol = primRows.length > 0 && (Object.keys(primRows[0]).find(k => k === "Net") || Object.keys(primRows[0]).find(k => k.replace(/_\d+$/,"") === "Net")) || "Net";
    bv.primary[pKey] = sumRows(primRows, "Month", "Year", "Customer Name", M, Y, pKey, primNetCol);

    // 4. Rate Difference_Sales → GRC Sum of NET_AMOUNT
    bv.rate_diff[pKey] = grcRow ? toNum(grcRow["Sum of NET_AMOUNT"]) : 0;

    // 5. Secondary Sales → REC_TD Basic Value (Payment )
    bv.secondary_sales[pKey] = tdRow ? toNum(tdRow["Basic Value (Payment )"]) : 0;

    // 6. Markdown → REC_TD Discount CN
    bv.markdown[pKey] = tdRow ? toNum(tdRow["Discount CN"]) : 0;

    // 7. Closing Stock → REC_STK current month Basic Value
    bv.closing_stock[pKey] = currStkRow ? toNum(currStkRow["Basic Value"]) : 0;

    // 8. Sales Adj SOR (calculated) — will be recalculated after opening_sec_diff is set
    bv.sales_adj[pKey] = 0; // placeholder

    // 9. COGS % → from REC_STK current month
    // Overhead  = Qty × 90.8  (Qty col may be "Qty." or "Qty" or "Qty ")
    // Total Cost = COGS Value + Overhead
    // COGS %    = Total Cost / Basic Value
    let cogsPct = 0;
    if (currStkRow) {
      // Find qty column — try all variations
      const qtyKey = Object.keys(currStkRow).find(k =>
        k.trim().replace(/\.$/,"").toLowerCase() === "qty"
      ) || "Qty.";
      const cogsKey = Object.keys(currStkRow).find(k =>
        k.trim().toLowerCase() === "cogs value"
      ) || "COGS Value";
      const bvKey = Object.keys(currStkRow).find(k =>
        k.trim().toLowerCase() === "basic value"
      ) || "Basic Value";

      const cogsVal   = toNum(currStkRow[cogsKey]);
      const basicVal  = toNum(currStkRow[bvKey]);
      const qtyVal    = toNum(currStkRow[qtyKey]);
      const overhead  = qtyVal * 90.8;
      const totalCost = cogsVal + overhead;
      cogsPct = basicVal !== 0 ? totalCost / basicVal : 0;
    }
    bv.cogs_pct[pKey] = cogsPct;

    // 10. COGS — calculated after sales_adj is set
    bv.cogs[pKey] = 0; // placeholder

    // ── QTY ──

    // 1. Opening Stock → REC_STK prev month Qty.
    // Opening Stock Qty
    const qtyColName = currStkRow
      ? (Object.keys(currStkRow).find(k => k.trim().replace(/\.$/,"").toLowerCase() === "qty") || "Qty.")
      : "Qty.";
    qty.opening_stock[pKey] = prevStkRow ? toNum(prevStkRow[qtyColName]) : 0;

    // 2. Opening Secondary Difference → placeholder
    qty.opening_sec_diff[pKey] = 0;

    // 3. Primary → PRIM_SALE sum of Qty. col for ALL matching rows
    // Primary Qty → PRIM_SALE sum of Qty. col
    const primQtyCol = primRows.length > 0 && (Object.keys(primRows[0]).find(k => k === "Qty.") || Object.keys(primRows[0]).find(k => k.replace(/_\d+$/,"") === "Qty.")) || "Qty.";
    qty.primary[pKey] = sumRows(primRows, "Month", "Year", "Customer Name", M, Y, pKey, primQtyCol);

    // 4. Rate Difference → GRC Sum of Purchase Qty & Purchase Return Qty
    qty.rate_diff[pKey] = grcRow
      ? toNum(grcRow["Sum of Purchase Qty & Purchase Return Qty"])
      : 0;

    // 5. Secondary Sales → REC_TD Qty (Sales)
    qty.secondary_sales[pKey] = tdRow ? toNum(tdRow["Qty (Sales)"]) : 0;

    // 6. Markdown qty → 0 (no direct qty for markdown in REC_TD)
    qty.markdown[pKey] = 0;

    // 7. Closing Stock → REC_STK current month Qty.
    qty.closing_stock[pKey] = currStkRow ? toNum(currStkRow[qtyColName]) : 0;

    // 8. Sales Adj SOR — placeholder
    qty.sales_adj[pKey] = 0;
  });

  return { bv, qty, partnerKeys: PARTNER_KEYS };
}

// Finalize sales_adj and cogs after opening_sec_diff is set
function finalizeMatrix(bv, qty, PARTNER_KEYS) {
  const hasQty = qty && qty.sales_adj;
  PARTNER_KEYS.forEach(pKey => {
    // BV Sales Adj = SUM(Add rows) - SUM(Less rows)
    // Add: Opening Stock + Opening_Diff + Primary
    // Less: Rate_Diff + Secondary + Markdown + Closing
    bv.sales_adj[pKey] =
      (bv.opening_stock[pKey] + bv.opening_sec_diff[pKey] + bv.primary[pKey]) -
      (bv.rate_diff[pKey] + bv.secondary_sales[pKey] + bv.markdown[pKey] + bv.closing_stock[pKey]);

    // COGS = Sales_Adj * COGS%
    bv.cogs[pKey] = bv.sales_adj[pKey] * bv.cogs_pct[pKey];

    // QTY Sales Adj (only if qty matrix was provided)
    // Add: Opening + Opening_Diff + Primary
    // Less: Rate_Diff + Secondary + Markdown + Closing
    if (hasQty) {
      qty.sales_adj[pKey] =
        (qty.opening_stock[pKey] + qty.opening_sec_diff[pKey] + qty.primary[pKey]) -
        (qty.rate_diff[pKey] + qty.secondary_sales[pKey] + qty.markdown[pKey] + qty.closing_stock[pKey]);
    }
  });
}

// -------------------------
// API — DASHBOARD
// -------------------------
app.get("/api/dashboard", requireAuth, async (req, res) => {
  try {
    const { month, year } = req.query;
    if (!month || !year)
      return res.status(400).json({ ok: false, error: "month and year required" });

    const SHEET_ID = process.env.GRC_SHEET_ID;
    const STORE_TAB = "DASHBOARD_STORE";
    if (!SHEET_ID)
      return res.status(500).json({ ok: false, error: "GRC_SHEET_ID not configured" });

    const sa = getServiceAccountJson();
    const authClient = new google.auth.GoogleAuth({
      credentials: { client_email: sa.client_email, private_key: sa.private_key },
      scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
    });
    const client = await authClient.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });

    // Write client (needs write scope)
    const writeAuthClient = new google.auth.GoogleAuth({
      credentials: { client_email: sa.client_email, private_key: sa.private_key },
      scopes: ["https://www.googleapis.com/auth/spreadsheets"],
    });
    const writeClient = await writeAuthClient.getClient();
    const sheetsWrite = google.sheets({ version: "v4", auth: writeClient });

    async function readTab(tabName) {
      const r = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: tabName });
      const rows = r.data.values || [];
      if (rows.length < 2) return [];
      // Strip surrounding quotes from headers (Google Sheets sometimes wraps in quotes)
      const seenKeys = {};
      const headers = rows[0].map(h => {
        let k = String(h || "").trim().replace(/^"|"$/g, "").trim();
        // Handle duplicate column names (PRIM_SALE has Qty./Net/etc 3 times)
        if (seenKeys[k] !== undefined) {
          seenKeys[k]++;
          k = k + "_" + seenKeys[k];
        } else {
          seenKeys[k] = 0;
        }
        return k;
      });
      return rows.slice(1).map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = (row[i] === undefined ? "" : String(row[i]).trim()); });
        return obj;
      });
    }

    const M = month.toUpperCase();
    const Y = String(year);
    const prev = getPrevMonth(M, Y);

    // Read all 4 tabs ONCE (reused for both current and prev month calc)
    const [recStkRows, recTdRows, grcRows, primRows] = await Promise.all([
      readTab(process.env.REC_STK_TAB_NAME   || "REC_STK"),
      readTab(process.env.REC_TD_TAB_NAME    || "REC_TD"),
      readTab(process.env.GRC_TAB_NAME       || "GRC"),
      readTab(process.env.PRIM_SALE_TAB_NAME || "PRIM_SALE"),
    ]);

    // ── Step 0: Read stored Sales-Adj SOR from DASHBOARD_STORE ──
    // Key: Year + Month + Partner
    let storedSalesAdj = {}; // { "Shoppers": 2850983.44, ... }
    try {
      const storeRows = await readTab(STORE_TAB);
      storeRows.forEach(r => {
        const ry = String(r["Year"]    || "").trim();
        const rm = String(r["Month"]   || "").trim().toUpperCase();
        const rp = String(r["Partner"] || "").trim();
        const rv = parseFloat(String(r["Sales_Adj_SOR_BV"] || "").replace(/,/g,""));
        if (ry === Y && rm === M && rp && !isNaN(rv)) {
          storedSalesAdj[rp] = rv;
        }
      });
      console.log(`[dashboard] Loaded ${Object.keys(storedSalesAdj).length} stored Sales-Adj values for ${M} ${Y}`);
    } catch(e) {
      console.log("[dashboard] No stored data found, will calculate:", e.message);
    }

    // ── Step 1: Calculate CURRENT month matrix ──
    const { bv, qty, partnerKeys: PARTNER_KEYS } = calcMatrix(recStkRows, recTdRows, grcRows, primRows, M, Y);

    // ── Step 2: Get Opening_Secondary_Diff for current month ──
    // Priority: 1) stored value from DASHBOARD_STORE, 2) calculated from prev month
    if (prev) {
      // Check if prev month values are stored
      const prevStoreRows = await readTab(STORE_TAB).catch(() => []);
      const prevStored = {};
      const prevStoredQty = {};
      prevStoreRows.forEach(r => {
        const ry  = String(r["Year"]            || "").trim();
        const rm  = String(r["Month"]           || "").trim().toUpperCase();
        const rp  = String(r["Partner"]         || "").trim();
        const rv  = parseFloat(String(r["Sales_Adj_SOR_BV"]  || "").replace(/,/g,""));
        const rq  = parseFloat(String(r["Sales_Adj_SOR_QTY"] || "").replace(/,/g,""));
        if (ry === prev.year && rm === prev.month.toUpperCase() && rp) {
          if (!isNaN(rv)) prevStored[rp]    = rv;
          if (!isNaN(rq)) prevStoredQty[rp] = rq;
        }
      });

      const hasPrevStored = Object.keys(prevStored).length > 0;

      if (hasPrevStored) {
        // ✅ Use stored values — 100% accurate
        console.log(`[dashboard] Using stored Opening_Diff from ${prev.month} ${prev.year}`);
        PARTNER_KEYS.forEach(pKey => {
          bv.opening_sec_diff[pKey]  = prevStored[pKey]    || 0;
          qty.opening_sec_diff[pKey] = prevStoredQty[pKey] || 0;  // use QTY stored value
        });
      } else {
        // ⚠️ Fallback: calculate prev month (2 levels deep)
        console.log(`[dashboard] No stored data for ${prev.month} ${prev.year}, calculating...`);
        const prevPrev = getPrevMonth(prev.month, prev.year);
        let prevPrevSalesAdj = {};

        if (prevPrev) {
          const { bv: ppBv, partnerKeys: ppKeys } = calcMatrix(recStkRows, recTdRows, grcRows, primRows, prevPrev.month, prevPrev.year);
          finalizeMatrix(ppBv, null, ppKeys);
          ppKeys.forEach(pKey => { prevPrevSalesAdj[pKey] = ppBv.sales_adj[pKey] || 0; });
        }

        const { bv: prevBv, partnerKeys: prevKeys } = calcMatrix(recStkRows, recTdRows, grcRows, primRows, prev.month, prev.year);
        prevKeys.forEach(pKey => {
          prevBv.opening_sec_diff[pKey] = prevPrevSalesAdj[pKey] || 0;
        });
        finalizeMatrix(prevBv, null, prevKeys);

        PARTNER_KEYS.forEach(pKey => {
          bv.opening_sec_diff[pKey]  = prevBv.sales_adj[pKey] || 0;
          qty.opening_sec_diff[pKey] = prevBv.sales_adj[pKey] || 0;
        });
      }
    }

    // ── Step 3: Finalize current month (now opening_sec_diff is set) ──
    finalizeMatrix(bv, qty, PARTNER_KEYS);

    // ── Auto-save Sales-Adj SOR to DASHBOARD_STORE ──
    const PARTNER_KEYS_SAVE = [
      "Shoppers", "Lifestyle", "V Retail", "Sabharwal Nirankar",
      "Reliance Centro", "Kora Retail", "Leayan_Zuup",
      "Myntra", "Flipkart", "Reliance Ajio",
    ];

    try {
      // Read existing store data
      let existingRows = [];
      let storeHeaders = ["Year", "Month", "Partner", "Sales_Adj_SOR_BV", "Sales_Adj_SOR_QTY"];
      try {
        const sr = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: "DASHBOARD_STORE" });
        const srows = sr.data.values || [];
        if (srows.length > 0) {
          storeHeaders = srows[0].map(h => String(h || "").trim().replace(/^"|"$/g, "").trim());
          existingRows = srows.slice(1);
        }
      } catch(e) { /* tab empty */ }

      // Build upsert map
      const rowMap = {};
      existingRows.forEach((row, i) => {
        const key = `${String(row[0]||"").trim()}|${String(row[1]||"").trim().toUpperCase()}|${String(row[2]||"").trim()}`;
        rowMap[key] = i;
      });

      // Upsert each partner
      PARTNER_KEYS_SAVE.forEach(pKey => {
        const bvAdj  = bv.sales_adj  ? (bv.sales_adj[pKey]  || 0) : 0;
        const qtyAdj = qty.sales_adj ? (qty.sales_adj[pKey] || 0) : 0;
        const key = `${Y}|${M}|${pKey}`;
        const newRow = [Y, M, pKey, bvAdj, qtyAdj];
        if (rowMap[key] !== undefined) {
          existingRows[rowMap[key]] = newRow;
        } else {
          existingRows.push(newRow);
        }
      });

      // Write back
      await sheetsWrite.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: "DASHBOARD_STORE!A1",
        valueInputOption: "RAW",
        requestBody: { values: [storeHeaders, ...existingRows] },
      });
      console.log(`[dashboard] Auto-saved Sales-Adj SOR for ${M} ${Y}`);
    } catch(saveErr) {
      console.error("[dashboard] Auto-save failed:", saveErr.message);
      // Don't fail the response — still return the matrix
    }

    return res.json({ ok: true, bv, qty, month: M, year: Y });

  } catch (e) {
    return res.status(500).json({ ok: false, error: e.message });
  }
});



// -------------------------
// API — DASHBOARD DEBUG (temp)
// -------------------------
app.get("/api/dashboard-debug", requireAuth, async (req, res) => {
  try {
    const SHEET_ID = process.env.GRC_SHEET_ID;
    const STORE_TAB = "DASHBOARD_STORE";
    const sa = getServiceAccountJson();
    const authClient = new google.auth.GoogleAuth({
      credentials: { client_email: sa.client_email, private_key: sa.private_key },
      scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
    });
    const client = await authClient.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });

    async function getHeaders(tabName) {
      const r = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: tabName + "!1:2",
      });
      const rows = r.data.values || [];
      return {
        headers: rows[0] ? rows[0].map(h => JSON.stringify(h)) : [],
        sample: rows[1] ? rows[1].slice(0, 5) : [],
      };
    }

    const [recStk, recTd, grc, prim] = await Promise.all([
      getHeaders(process.env.REC_STK_TAB_NAME   || "REC_STK"),
      getHeaders(process.env.REC_TD_TAB_NAME    || "REC_TD"),
      getHeaders(process.env.GRC_TAB_NAME       || "GRC"),
      getHeaders(process.env.PRIM_SALE_TAB_NAME || "PRIM_SALE"),
    ]);

    return res.json({ ok: true, REC_STK: recStk, REC_TD: recTd, GRC: grc, PRIM_SALE: prim });
  } catch (e) {
    return res.status(500).json({ ok: false, error: e.message });
  }
});


// -------------------------
// API — DASHBOARD SAVE
// Saves Sales-Adj SOR per partner to DASHBOARD_STORE tab (upsert by Year+Month+Partner)
// -------------------------
app.post("/api/dashboard-save", requireAuth, async (req, res) => {
  try {
    const { month, year, bv } = req.body;
    if (!month || !year || !bv)
      return res.status(400).json({ ok: false, error: "month, year, bv required" });

    const SHEET_ID = process.env.GRC_SHEET_ID;
    if (!SHEET_ID)
      return res.status(500).json({ ok: false, error: "GRC_SHEET_ID not configured" });

    const STORE_TAB = "DASHBOARD_STORE";
    const PARTNER_KEYS = [
      "Shoppers", "Lifestyle", "V Retail", "Sabharwal Nirankar",
      "Reliance Centro", "Kora Retail", "Leayan_Zuup",
      "Myntra", "Flipkart", "Reliance Ajio",
    ];

    const sa = getServiceAccountJson();
    const authClient = new google.auth.GoogleAuth({
      credentials: { client_email: sa.client_email, private_key: sa.private_key },
      scopes: ["https://www.googleapis.com/auth/spreadsheets"],
    });
    const client = await authClient.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });

    const M = month.toUpperCase();
    const Y = String(year);

    // Read existing data
    let existingRows = [];
    let headers = ["Year", "Month", "Partner", "Sales_Adj_SOR_BV", "Sales_Adj_SOR_QTY"];
    try {
      const r = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: STORE_TAB,
      });
      const rows = r.data.values || [];
      if (rows.length > 0) {
        headers = rows[0].map(h => String(h || "").trim().replace(/^"|"$/g, "").trim());
        existingRows = rows.slice(1);
      }
    } catch(e) {
      console.log("[dashboard-save] Tab empty or not found, will create headers");
    }

    // Build upsert map — key: Year|Month|Partner → row index
    const rowMap = {};
    existingRows.forEach((row, i) => {
      const ry = String(row[0] || "").trim();
      const rm = String(row[1] || "").trim().toUpperCase();
      const rp = String(row[2] || "").trim();
      rowMap[`${ry}|${rm}|${rp}`] = i;
    });

    // Prepare new/updated rows (BV + QTY)
    const qty = req.body.qty || {};
    PARTNER_KEYS.forEach(pKey => {
      const salesAdjBV  = bv.sales_adj  ? (bv.sales_adj[pKey]  || 0) : 0;
      const salesAdjQTY = qty.sales_adj ? (qty.sales_adj[pKey] || 0) : 0;
      const key = `${Y}|${M}|${pKey}`;
      const newRow = [Y, M, pKey, salesAdjBV, salesAdjQTY];

      if (rowMap[key] !== undefined) {
        existingRows[rowMap[key]] = newRow; // update
      } else {
        existingRows.push(newRow); // insert
      }
    });

    // Write back full tab
    const allData = [headers, ...existingRows];
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: STORE_TAB + "!A1",
      valueInputOption: "RAW",
      requestBody: { values: allData },
    });

    console.log(`[dashboard-save] Saved ${PARTNER_KEYS.length} rows for ${M} ${Y}`);
    return res.json({ ok: true, saved: PARTNER_KEYS.length, month: M, year: Y });

  } catch(e) {
    return res.status(500).json({ ok: false, error: e.message });
  }
});

// -------------------------
// START SERVER
// -------------------------
app.listen(PORT, "0.0.0.0", () => {
  console.log(`[server] Running on port ${PORT}`);
});