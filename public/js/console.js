// FILE: public/js/console.js
window.__sectionInit = window.__sectionInit || {};

window.__sectionInit.console = function () {
  const logEl = document.getElementById("log");
  const outListEl = document.getElementById("outList");
  const progressPctEl = document.getElementById("progressPct");
  const progressBarEl = document.getElementById("progressBar");
  const progressHintEl = document.getElementById("progressHint");

  const logLine = (t) => {
    if (!logEl) return;
    logEl.textContent += t + "\n";
    logEl.scrollTop = logEl.scrollHeight;
  };

  const setProgress = (pct, hint) => {
    const v = Math.max(0, Math.min(100, Number(pct) || 0));
    if (progressPctEl) progressPctEl.textContent = String(v);
    if (progressBarEl) progressBarEl.style.width = v + "%";
    if (progressHintEl) progressHintEl.textContent = hint || "";
  };

  const resetConsole = () => {
    if (logEl) logEl.textContent = "";
    if (outListEl) {
      outListEl.textContent = "No outputs yet.";
      outListEl.classList.add("muted");
    }
    setProgress(0, "Idle");
    running = false;
    if (progressTimer) clearInterval(progressTimer);
    progressTimer = null;
  };

  const needsMonthYear = (name) =>
    name === "Receivable Stock" || name === "Receivable Trade Discount";

  let PROJECTS = {};
  let running = false;
  let progressTimer = null;

  const startProgressSim = () => {
    if (progressTimer) clearInterval(progressTimer);
    let pct = 1;
    progressTimer = setInterval(() => {
      if (!running) return;
      if (pct < 92) {
        pct += Math.floor(Math.random() * 4) + 1;
        pct = Math.min(pct, 92);
        setProgress(pct, "Processing…");
      }
    }, 650);
  };

  const applyAgentUI = (agentName) => {
    const cfg = PROJECTS[agentName];
    if (!cfg) return;

    document.getElementById("pillMode").textContent = `Mode: ${cfg.mode}`;
    document.getElementById("pillOutput").textContent = `Output: ${cfg.outputName || "—"}`;

    const filesEl = document.getElementById("files");
    const isSingle = cfg.mode === "single";
    filesEl.multiple = !isSingle;

    document.getElementById("fileLabel").textContent = isSingle ? "Upload file" : "Upload files";
    document.getElementById("fileHint").textContent = isSingle ? "Select a single file." : "You can select multiple files.";

    const myBox = document.getElementById("monthYearBox");
    if (needsMonthYear(agentName)) {
      myBox.classList.remove("hidden");
      const y = document.getElementById("year");
      if (!y.value) y.value = String(new Date().getFullYear());
    } else {
      myBox.classList.add("hidden");
    }
  };

  async function initAgents() {
    // ✅ IMPORTANT: send cookies/session
    const res = await fetch("/projects", { cache: "no-store", credentials: "include" });

    if (!res.ok) {
      logLine(`Error loading agents: /projects returned ${res.status}`);
      logLine("If you see 401/302, login session is not active.");
      return;
    }

    PROJECTS = await res.json();

    const sel = document.getElementById("agentSelect");
    sel.innerHTML = "";

    Object.keys(PROJECTS).forEach((name) => {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      sel.appendChild(opt);
    });

    sel.addEventListener("change", () => {
      applyAgentUI(sel.value);
      resetConsole();
    });

    applyAgentUI(sel.value);
  }

  async function runAgent() {
    if (running) return;

    const agentName = document.getElementById("agentSelect").value;
    const files = document.getElementById("files").files;

    resetConsole();

    if (!files || !files.length) return logLine("Please select file(s).");

    const month = (document.getElementById("month")?.value || "").trim();
    const year = (document.getElementById("year")?.value || "").trim();

    if (needsMonthYear(agentName) && (!month || !year)) {
      return logLine("Please select Month and Year for this agent.");
    }

    const fd = new FormData();
    for (const f of files) fd.append("files", f);
    if (needsMonthYear(agentName)) {
      fd.append("month", month);
      fd.append("year", year);
    }

    running = true;
    setProgress(1, "Uploading…");
    startProgressSim();

    logLine("Client: Inc.5 Shoes");
    logLine("Agent: " + agentName);
    logLine("Uploading files and starting processing...");

    const res = await fetch(`/run/${encodeURIComponent(agentName)}`, {
      method: "POST",
      body: fd,
      credentials: "include", // ✅ IMPORTANT
    });

    const data = await res.json();

    running = false;
    if (progressTimer) clearInterval(progressTimer);

    if (!data.ok) {
      setProgress(0, "Failed");
      return logLine("Error: " + (data.error || "Unknown error"));
    }

    setProgress(100, "Completed");

    logLine("Done. RunId: " + data.runId);

    const outputs = data.outputs || [];
    if (!outputs.length) {
      outListEl.textContent = "No output files found.";
      outListEl.classList.add("muted");
    } else {
      outListEl.classList.remove("muted");
      outListEl.innerHTML = outputs
        .map((f) => `<a class="out-link" target="_blank" rel="noreferrer" href="https://drive.google.com/file/d/${f.id}/view">Open ${f.name}</a>`)
        .join("");
    }
  }

  document.getElementById("btnRun").addEventListener("click", runAgent);
  document.getElementById("btnClear").addEventListener("click", resetConsole);
  document.getElementById("btnClearAll").addEventListener("click", resetConsole);

  resetConsole();
  initAgents();
};