// =========================
// FILE: public/js/app.js  (FULL UPDATED FINAL)
// ✅ Fix: DONE/ERROR handled only once (no "Completed." spam)
// ✅ Fix: stop polling immediately on terminal states
// ✅ Fix: resume polling without duplicating old logs
// ✅ Adds: "— Select Agent —" placeholder
// ✅ Keeps: lock overlay + polling survives tab switching
// ✅ After DONE/ERROR: keeps agent selected, clears files/month/year for next run
// ✅ NEW: Top-right Reset does FULL CONSOLE refresh (page reload) when not running
// =========================

window.__sectionInit = window.__sectionInit || {};
window.__consoleRunning = false;

// ✅ global job state survives section switches
// { jobId, agentName, terminalHandled:boolean, lastLogCount:number }
window.__activeJob = window.__activeJob || null;
window.__jobPollTimer = window.__jobPollTimer || null;

// ─── TOAST ───
window.toast =
  window.toast ||
  function (msg, type) {
    const el = document.getElementById("toast");
    if (!el) return;
    el.textContent = msg;
    el.className = "toast show" + (type ? " toast-" + type : "");
    clearTimeout(window.__toastTimer);
    window.__toastTimer = setTimeout(() => el.classList.remove("show"), 2400);
  };

// ─── USER CHIP ───
async function hydrateUserChip() {
  const nameEl = document.getElementById("userName");
  const emailEl = document.getElementById("userEmail");
  const avatarEl = document.getElementById("userAvatar");
  if (!nameEl || !emailEl || !avatarEl) return;

  try {
    const res = await fetch("/me", { cache: "no-store", credentials: "include" });
    if (!res.ok) return;
    const ct = res.headers.get("content-type") || "";
    if (!ct.includes("application/json")) return;

    const data = await res.json();
    if (!data?.ok || !data?.user) return;

    const u = data.user;
    const dn = u.displayName || u.name || u.email || "User";
    const em = u.email || "—";

    nameEl.textContent = dn;
    emailEl.textContent = em;
    avatarEl.textContent = String(dn).trim().charAt(0).toUpperCase() || "U";
  } catch {}
}
document.addEventListener("DOMContentLoaded", hydrateUserChip);

// ─────────────────────────────────────────────────────────────
// ✅ Polling helpers
// ─────────────────────────────────────────────────────────────
function stopJobPolling() {
  if (window.__jobPollTimer) clearInterval(window.__jobPollTimer);
  window.__jobPollTimer = null;
}

async function fetchJobStatus(jobId) {
  const res = await fetch(`/api/job-status?jobId=${encodeURIComponent(jobId)}`, {
    cache: "no-store",
    credentials: "include",
  });

  const ct = res.headers.get("content-type") || "";
  if (!ct.includes("application/json")) {
    const txt = await res.text();
    throw new Error(`job-status not JSON (HTTP ${res.status})\n${txt.slice(0, 200)}`);
  }

  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "job-status ok:false");
  return data;
}

function startJobPolling({ jobId, onTick }) {
  stopJobPolling();

  const tick = async () => {
    try {
      const st = await fetchJobStatus(jobId);
      onTick && onTick(null, st);
    } catch (e) {
      onTick && onTick(e, null);
    }
  };

  tick();
  window.__jobPollTimer = setInterval(tick, 1200);
}

// ─────────────────────────────────────────────────────────────
// ✅ Status normalization
// ─────────────────────────────────────────────────────────────
function normStatus(s) {
  return String(s || "").trim().toLowerCase();
}
function isRunningStatus(s) {
  const v = normStatus(s);
  return v === "running" || v === "uploading" || v === "queued" || v === "processing";
}
function isDoneStatus(s) {
  const v = normStatus(s);
  return v === "done" || v === "completed" || v === "complete" || v === "success" || v === "finished";
}
function isErrorStatus(s) {
  const v = normStatus(s);
  return v === "error" || v === "failed" || v === "fail";
}

// ─── CONSOLE SECTION ───
window.__sectionInit.console = function () {
  const logContentEl = document.getElementById("logContent");
  const outListEl = document.getElementById("outList");
  const progressPctEl = document.getElementById("progressPct");
  const progressFillEl = document.getElementById("progressFill");
  const progressHintEl = document.getElementById("progressHint");

  const agentSelect = document.getElementById("agentSelect");
  const fileInput = document.getElementById("files");
  const fileZone = document.getElementById("fileZone");
  const btnRun = document.getElementById("btnRun");
  const btnClear = document.getElementById("btnClear");
  const btnClearAll = document.getElementById("btnClearAll"); // ✅ top-right Reset
  const fileStatus = document.getElementById("fileStatus");
  const lockEl = document.getElementById("consoleLock");

  if (!agentSelect || !fileInput || !btnRun || !logContentEl) return;

  const safe = (s) =>
    String(s || "").replace(/[&<>"']/g, (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" }[c])
    );

  const logLine = (t, type) => {
    const prefix = type === "ok" ? "✓ " : type === "err" ? "✗ " : type === "info" ? "→ " : "  ";
    logContentEl.textContent += prefix + t + "\n";
    logContentEl.scrollTop = logContentEl.scrollHeight;
  };

  const setProgress = (pct, hint) => {
    const v = Math.max(0, Math.min(100, Number(pct) || 0));
    if (progressPctEl) progressPctEl.textContent = v + "%";
    if (progressFillEl) progressFillEl.style.width = v + "%";
    if (progressHintEl) progressHintEl.textContent = hint || "";
  };

  const needsMonthYear = (name) => name === "Receivable Stock" || name === "Receivable Trade Discount";

  let PROJECTS = {};
  let running = false;

  function syncRunButton() {
    const agentOk = Boolean(agentSelect.value);
    const filesOk = Boolean(fileInput.files && fileInput.files.length);
    const locked = Boolean(window.__consoleRunning || window.__activeJob?.jobId);
    btnRun.disabled = locked || !agentOk || !filesOk;
  }

  function lockInputs(locked) {
    if (lockEl) lockEl.style.display = locked ? "flex" : "none";

    agentSelect.disabled = locked;
    fileInput.disabled = locked;

    const monthEl = document.getElementById("month");
    const yearEl = document.getElementById("year");
    if (monthEl) monthEl.disabled = locked;
    if (yearEl) yearEl.disabled = locked;

    if (btnClear) btnClear.disabled = locked;
    if (btnClearAll) btnClearAll.disabled = locked;

    if (fileZone) fileZone.classList.toggle("locked", locked);

    syncRunButton();
  }

  function applyAgentUI(agentName) {
    const pillMode = document.getElementById("pillMode");
    const pillOutput = document.getElementById("pillOutput");

    if (!agentName || !PROJECTS[agentName]) {
      if (pillMode) pillMode.textContent = "MODE: —";
      if (pillOutput) pillOutput.textContent = "OUT: —";
      const myBox = document.getElementById("monthYearBox");
      if (myBox) myBox.classList.add("hidden");
      return;
    }

    const cfg = PROJECTS[agentName];
    if (pillMode) pillMode.textContent = `MODE: ${String(cfg.mode || "").toUpperCase()}`;
    if (pillOutput) pillOutput.textContent = `OUT: ${cfg.outputName || "—"}`;

    const isSingle = cfg.mode === "single";
    fileInput.multiple = !isSingle;

    const fileZoneLabel = document.getElementById("fileZoneLabel");
    const fileZoneSub = document.getElementById("fileZoneSub");

    if (fileZoneLabel)
      fileZoneLabel.innerHTML = isSingle ? `Drop file here or <span>browse</span>` : `Drop files here or <span>browse</span>`;

    if (fileZoneSub)
      fileZoneSub.textContent = isSingle ? "Single file · .xlsx / .csv" : "Multiple files · .xlsx / .csv";

    const myBox = document.getElementById("monthYearBox");
    if (myBox) {
      if (needsMonthYear(agentName)) {
        myBox.classList.remove("hidden");
        const y = document.getElementById("year");
        if (y && !y.value) y.value = String(new Date().getFullYear());
      } else {
        myBox.classList.add("hidden");
      }
    }
  }

  function renderOutputs(outputs) {
    if (!outListEl) return;
    if (!outputs || !outputs.length) {
      outListEl.innerHTML = `<div class="outputs-empty">No output files found.</div>`;
      return;
    }
    outListEl.innerHTML = outputs
      .map((f) => {
        const name = safe(f.name);
        const url = `https://drive.google.com/file/d/${f.id}/view`;
        return `<a class="out-link" target="_blank" rel="noreferrer" href="${url}">${name}</a>`;
      })
      .join("");
  }

  function clearActiveJob() {
    stopJobPolling();
    window.__activeJob = null;
    window.__consoleRunning = false;
    running = false;
  }

  function resetForNextRunKeepAgent() {
    // ✅ keep agent, clear only inputs
    try {
      fileInput.value = "";
    } catch {}

    if (fileZone) fileZone.classList.remove("has-files");
    if (fileStatus) fileStatus.textContent = "";

    const monthEl = document.getElementById("month");
    const yearEl = document.getElementById("year");
    if (monthEl) monthEl.value = "";
    if (yearEl) yearEl.value = "";

    applyAgentUI(agentSelect.value);
    syncRunButton();
  }

  async function initAgents() {
    let res;
    try {
      res = await fetch("/projects", { cache: "no-store", credentials: "include" });
    } catch (e) {
      logLine("Network error: " + e.message, "err");
      return;
    }

    if (!res.ok) {
      logLine(`Failed to load agents — HTTP ${res.status}`, "err");
      return;
    }

    const ct = res.headers.get("content-type") || "";
    if (!ct.includes("application/json")) {
      logLine("Session expired — please reload and login.", "err");
      return;
    }

    PROJECTS = await res.json();
    const keys = Object.keys(PROJECTS || {});
    agentSelect.innerHTML = "";

    // ✅ placeholder
    const ph = document.createElement("option");
    ph.value = "";
    ph.textContent = "— Select Agent —";
    agentSelect.appendChild(ph);

    keys.forEach((name) => {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      agentSelect.appendChild(opt);
    });

    // keep running agent selected if exists
    if (window.__activeJob?.agentName) {
      agentSelect.value = window.__activeJob.agentName;
      applyAgentUI(agentSelect.value);
    } else {
      applyAgentUI("");
    }

    syncRunButton();
  }

  // ✅ Attach polling for a jobId
  function attachPolling(jobId, { resume = false } = {}) {
    // Ensure we have a job object
    if (!window.__activeJob || window.__activeJob.jobId !== jobId) {
      window.__activeJob = {
        jobId,
        agentName: agentSelect.value || window.__activeJob?.agentName || "",
        terminalHandled: false,
        lastLogCount: 0,
      };
    }

    startJobPolling({
      jobId,
      onTick: (err, st) => {
        const job = window.__activeJob;
        if (!job || job.jobId !== jobId) return;

        // If already handled terminal, ignore any further ticks
        if (job.terminalHandled) return;

        if (err) {
          const curPct = Number((progressPctEl?.textContent || "0").replace("%", "")) || 0;
          setProgress(curPct, "Connecting…");
          return;
        }

        const status = normStatus(st.status);
        const hint = st.hint || (status ? status.toUpperCase() : "");

        // progress
        if (typeof st.pct === "number") {
          setProgress(st.pct, hint);
        } else {
          const curPct = Number((progressPctEl?.textContent || "0").replace("%", "")) || 0;
          setProgress(curPct, hint);
        }

        // logs (no duplication on resume)
        const logs = Array.isArray(st.logs) ? st.logs : [];

        if (resume && job.lastLogCount === 0) {
          job.lastLogCount = logs.length;
        }

        if (logs.length > job.lastLogCount) {
          logs.slice(job.lastLogCount).forEach((line) => logLine(line, "info"));
          job.lastLogCount = logs.length;
        }

        // running
        if (isRunningStatus(status)) {
          running = true;
          window.__consoleRunning = true;
          lockInputs(true);
          return;
        }

        // done
        if (isDoneStatus(status)) {
          job.terminalHandled = true;
          stopJobPolling(); // ✅ stop immediately

          setProgress(100, "Completed");
          lockInputs(false);

          if (st.outputs) renderOutputs(st.outputs);

          logLine("Completed.", "ok");
          window.toast("✓ Agent completed");

          clearActiveJob();
          resetForNextRunKeepAgent();
          return;
        }

        // error
        if (isErrorStatus(status) || st.error) {
          job.terminalHandled = true;
          stopJobPolling(); // ✅ stop immediately

          lockInputs(false);
          setProgress(Math.min(100, Number(st.pct) || 0), "Failed");

          logLine("Error: " + (st.error || "Unknown error"), "err");
          window.toast("Agent failed", "err");

          clearActiveJob();
          resetForNextRunKeepAgent();
          return;
        }
      },
    });
  }

  async function runAgent() {
    if (window.__consoleRunning || window.__activeJob?.jobId) return;

    const agentName = agentSelect.value;
    if (!agentName) {
      window.toast("Select an agent first", "err");
      return;
    }

    const files = fileInput.files;
    if (!files || !files.length) {
      window.toast("Select file(s) first", "err");
      return;
    }

    const month = (document.getElementById("month")?.value || "").trim();
    const year = (document.getElementById("year")?.value || "").trim();

    if (needsMonthYear(agentName) && (!month || !year)) {
      logLine("Month and Year are required for this agent.", "err");
      window.toast("Month/Year required", "err");
      return;
    }

    // reset UI for a fresh run
    logContentEl.textContent = "";
    if (outListEl) outListEl.innerHTML = `<div class="outputs-empty">— Running… —</div>`;
    setProgress(2, "Uploading…");

    const fd = new FormData();
    for (const f of files) fd.append("files", f);
    if (needsMonthYear(agentName)) {
      fd.append("month", month);
      fd.append("year", year);
    }

    running = true;
    window.__consoleRunning = true;
    lockInputs(true);

    logLine(`Client: Inc.5 Shoes`, "info");
    logLine(`Agent: ${agentName}`, "info");
    if (needsMonthYear(agentName)) logLine(`Period: ${month} / ${year}`, "info");
    logLine(`Files: ${files.length} file(s)`, "info");
    logLine(`Starting...`, "info");

    let res;
    try {
      res = await fetch(`/run/${encodeURIComponent(agentName)}`, {
        method: "POST",
        body: fd,
        credentials: "include",
      });
    } catch (e) {
      lockInputs(false);
      setProgress(0, "Failed");
      logLine("Network error: " + e.message, "err");
      window.toast("Network error", "err");
      clearActiveJob();
      resetForNextRunKeepAgent();
      return;
    }

    let data;
    try {
      data = await res.json();
    } catch {
      lockInputs(false);
      setProgress(0, "Failed");
      logLine("Server returned invalid JSON.", "err");
      window.toast("Response error", "err");
      clearActiveJob();
      resetForNextRunKeepAgent();
      return;
    }

    if (!data.ok) {
      lockInputs(false);
      setProgress(0, "Failed");
      logLine("Error: " + (data.error || "Unknown error"), "err");
      window.toast("Agent failed", "err");
      clearActiveJob();
      resetForNextRunKeepAgent();
      return;
    }

    const jobId = data.jobId;
    if (!jobId) {
      lockInputs(false);
      setProgress(0, "Failed");
      logLine("Server did not return jobId. Update server.js to support background job + /api/job-status.", "err");
      window.toast("Missing jobId", "err");
      clearActiveJob();
      resetForNextRunKeepAgent();
      return;
    }

    window.__activeJob = {
      jobId,
      agentName,
      terminalHandled: false,
      lastLogCount: 0,
    };

    logLine(`Job started: ${jobId}`, "ok");
    setProgress(5, "Queued…");

    attachPolling(jobId, { resume: false });
    syncRunButton();
  }

  // ── File zone drag & drop ──
  if (fileZone && fileInput) {
    fileZone.addEventListener("dragover", (e) => {
      if (fileZone.classList.contains("locked")) return;
      e.preventDefault();
      fileZone.classList.add("drag-over");
    });

    fileZone.addEventListener("dragleave", () => fileZone.classList.remove("drag-over"));

    fileZone.addEventListener("drop", (e) => {
      if (fileZone.classList.contains("locked")) return;
      e.preventDefault();
      fileZone.classList.remove("drag-over");

      if (e.dataTransfer?.files?.length) {
        const dt = new DataTransfer();
        Array.from(e.dataTransfer.files).forEach((f) => dt.items.add(f));
        fileInput.files = dt.files;
        updateFileStatus();
      }
    });

    fileInput.addEventListener("change", updateFileStatus);
  }

  function updateFileStatus() {
    const files = fileInput.files;

    if (!files || !files.length) {
      if (fileZone) fileZone.classList.remove("has-files");
      if (fileStatus) fileStatus.textContent = "";
      syncRunButton();
      return;
    }

    if (fileZone) fileZone.classList.add("has-files");

    const names = Array.from(files).map((f) => f.name).join(", ");
    if (fileStatus) fileStatus.textContent = files.length === 1 ? names : `${files.length} files selected`;

    const fileZoneLabel = document.getElementById("fileZoneLabel");
    if (fileZoneLabel)
      fileZoneLabel.innerHTML =
        files.length === 1
          ? `<span style="color:var(--green)">${safe(names)}</span>`
          : `<span style="color:var(--green)">${files.length} files selected</span>`;

    syncRunButton();
  }

  // buttons
  btnRun.onclick = runAgent;

  // ✅ Clear button: reset fields/logs WITHOUT reload
  const clearConsole = () => {
    if (window.__consoleRunning || window.__activeJob?.jobId) {
      window.toast("Agent is running. Wait to finish.", "err");
      return;
    }

    logContentEl.textContent = "";
    if (outListEl) outListEl.innerHTML = `<div class="outputs-empty">— No outputs yet —</div>`;
    setProgress(0, "Idle");
    resetForNextRunKeepAgent();
    lockInputs(false);
  };

  // ✅ TOP-RIGHT Reset: FULL refresh (like browser refresh)
  const hardResetConsole = () => {
    if (window.__consoleRunning || window.__activeJob?.jobId) {
      window.toast("Agent is running. Wait to finish.", "err");
      return;
    }

    // clear timers & state, then reload
    stopJobPolling();
    window.__activeJob = null;
    window.__consoleRunning = false;

    window.location.reload();
  };

  if (btnClear) btnClear.onclick = clearConsole;
  if (btnClearAll) btnClearAll.onclick = hardResetConsole;

  agentSelect.onchange = () => {
    if (window.__consoleRunning || window.__activeJob?.jobId) return;

    applyAgentUI(agentSelect.value);

    // clear file selection on agent change
    try {
      fileInput.value = "";
    } catch {}
    if (fileZone) fileZone.classList.remove("has-files");
    if (fileStatus) fileStatus.textContent = "";

    setProgress(0, "Idle");
    syncRunButton();
  };

  // ✅ resume polling if job already running
  if (window.__activeJob?.jobId) {
    running = true;
    window.__consoleRunning = true;

    // select the running agent if known
    if (window.__activeJob.agentName) {
      agentSelect.value = window.__activeJob.agentName;
      applyAgentUI(agentSelect.value);
    }

    lockInputs(true);
    logLine(`Resuming job: ${window.__activeJob.jobId}`, "info");

    attachPolling(window.__activeJob.jobId, { resume: true });
  } else {
    lockInputs(false);
  }

  initAgents();
};