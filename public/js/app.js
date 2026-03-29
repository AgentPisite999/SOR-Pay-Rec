
// // FILE: public/js/app.js

// window.__sectionInit = window.__sectionInit || {};
// window.__job = window.__job || { id: null, agent: null };
// window.__pollTimer = window.__pollTimer || null;
// window.__consoleRunning = window.__consoleRunning || false;
// window.toast = window.toast || function(msg){ try { console.log(msg); } catch(e) {} };

// window.toast = window.toast || function (msg, type) {
//   var el = document.getElementById("toast");
//   if (!el) return;
//   el.textContent = msg;
//   el.className = "toast show" + (type ? " toast-" + type : "");
//   clearTimeout(window.__toastT);
//   window.__toastT = setTimeout(function () {
//     el.classList.remove("show");
//   }, 2800);
// };

// async function hydrateUserChip() {
//   try {
//     var res = await fetch("/me", { cache: "no-store", credentials: "include" });
//     if (!res.ok) return;
//     var d = await res.json();
//     if (!d.ok || !d.user) return;
//     var dn = d.user.displayName || d.user.name || d.user.email || "User";
//     var e;
//     e = document.getElementById("userName");
//     if (e) e.textContent = dn;
//     e = document.getElementById("userEmail");
//     if (e) e.textContent = d.user.email || "—";
//     e = document.getElementById("userAvatar");
//     if (e) e.textContent = dn.trim()[0].toUpperCase();
//   } catch (e) {}
// }
// document.addEventListener("DOMContentLoaded", hydrateUserChip);

// function stopPolling() {
//   if (window.__pollTimer) {
//     clearInterval(window.__pollTimer);
//     window.__pollTimer = null;
//   }
// }

// window.resetConsoleUI = function () {
//   if (window.__job && window.__job.id) {
//     window.toast && window.toast("Agent is running", "err");
//     return;
//   }

//   stopPolling();
//   window.__job = { id: null, agent: null };
//   window.__consoleRunning = false;

//   var g = function (id) { return document.getElementById(id); };

//   var lc = g("logContent");
//   if (lc) lc.textContent = "";

//   var ol = g("outList");
//   if (ol) ol.innerHTML = '<div class="outputs-empty">— No outputs yet —</div>';

//   try {
//     var fi = g("files");
//     if (fi) fi.value = "";
//   } catch (e) {}

//   var fz = g("fileZone");
//   if (fz) {
//     fz.classList.remove("has-files");
//     fz.classList.remove("locked");
//   }

//   var fs = g("fileStatus");
//   if (fs) fs.textContent = "";

//   var sel = g("agentSelect");
//   if (sel) sel.value = "";

//   var month = g("month");
//   if (month) month.value = "";

//   var year = g("year");
//   if (year) year.value = "";

//   var pm = g("pillMode");
//   if (pm) pm.textContent = "MODE: —";

//   var po = g("pillOutput");
//   if (po) po.textContent = "OUT: —";

//   var myb = g("monthYearBox");
//   if (myb) myb.classList.add("hidden");

//   var pct = g("progressPct");
//   if (pct) pct.textContent = "0%";

//   var fill = g("progressFill");
//   if (fill) fill.style.width = "0%";

//   var hint = g("progressHint");
//   if (hint) hint.textContent = "Idle";

//   var ov = g("consoleLock");
//   if (ov) {
//     ov.classList.remove("show");
//     ov.setAttribute("aria-hidden", "true");
//   }

//   ["agentSelect", "files", "month", "year", "btnClear", "btnClearAll"].forEach(function (id) {
//     var el = g(id);
//     if (el) el.disabled = false;
//   });

//   var runBtn = g("btnRun");
//   if (runBtn) runBtn.disabled = true;

//   window.toast && window.toast("Console reset");
// };

// // ═══════════════════════════════════════════════════════════════
// window.__sectionInit.console = function () {
// // ═══════════════════════════════════════════════════════════════

//   var PROJECTS = {};
//   var finished = false;

//   function g(id) {
//     return document.getElementById(id);
//   }

//   function safe(s) {
//     return String(s || "").replace(/[&<>"']/g, function (c) {
//       return {
//         "&": "&amp;",
//         "<": "&lt;",
//         ">": "&gt;",
//         '"': "&quot;",
//         "'": "&#039;"
//       }[c];
//     });
//   }

//   function showOverlay() {
//     var ov = g("consoleLock");
//     if (!ov) return;
//     ov.classList.add("show");
//     ov.setAttribute("aria-hidden", "false");
//   }

//   function hideOverlay() {
//     var ov = g("consoleLock");
//     if (!ov) return;
//     ov.classList.remove("show");
//     ov.setAttribute("aria-hidden", "true");
//   }

//   function setProgress(pct, hint) {
//     pct = Math.max(0, Math.min(100, Number(pct) || 0));
//     var p = g("progressPct");
//     if (p) p.textContent = pct + "%";
//     var f = g("progressFill");
//     if (f) f.style.width = pct + "%";
//     var h = g("progressHint");
//     if (h) h.textContent = hint || "";
//   }

//   function logLine(txt, type) {
//     var box = g("logContent");
//     if (!box) return;
//     var pre = type === "ok" ? "✓ " : type === "err" ? "✗ " : "→ ";
//     box.textContent += pre + txt + "\n";
//     box.scrollTop = box.scrollHeight;
//   }

//   function renderOutputs(outputs) {
//     var ol = g("outList");
//     if (!ol) return;

//     if (!outputs || !outputs.length) {
//       ol.innerHTML = '<div class="outputs-empty">Run complete — check Base Dump for files.</div>';
//       return;
//     }

//     ol.innerHTML = outputs.map(function (f) {
//       var url = "https://drive.google.com/file/d/" + f.id + "/view";
//       return '<a class="out-link" target="_blank" rel="noreferrer" href="' + url + '">' + safe(f.name) + "</a>";
//     }).join("");
//   }

//   function lockAll() {
//     window.__consoleRunning = true;

//     ["agentSelect", "files", "month", "year", "btnClear"].forEach(function (id) {
//       var e = g(id);
//       if (e) e.disabled = true;
//     });

//     var bRun = g("btnRun");
//     if (bRun) bRun.disabled = true;

//     var bReset = g("btnClearAll");
//     if (bReset) bReset.disabled = false;

//     var fz = g("fileZone");
//     if (fz) fz.classList.add("locked");

//     showOverlay();
//   }

//   function unlockAll() {
//     stopPolling();

//     window.__job = { id: null, agent: null };
//     window.__consoleRunning = false;

//     ["agentSelect", "files", "month", "year", "btnClear"].forEach(function (id) {
//       var e = g(id);
//       if (e) e.disabled = false;
//     });

//     var bRun = g("btnRun");
//     if (bRun) bRun.disabled = false;

//     var bReset = g("btnClearAll");
//     if (bReset) bReset.disabled = false;

//     var fz = g("fileZone");
//     if (fz) fz.classList.remove("locked");

//     hideOverlay();
//     syncBtn();
//   }

//   function syncBtn() {
//     var bRun = g("btnRun");
//     if (!bRun) return;

//     var agOk = !!(g("agentSelect") && g("agentSelect").value);
//     var fi = g("files");
//     var fiOk = !!(fi && fi.files && fi.files.length);

//     bRun.disabled = !!(window.__job.id) || !agOk || !fiOk;
//   }

//   function norm(s) {
//     return String(s || "").trim().toLowerCase();
//   }

//   function isRunning(s) {
//     var v = norm(s);
//     return v === "running" || v === "queued" || v === "processing" || v === "started" || v === "preparing";
//   }

//   function isUploading(s) {
//     return norm(s) === "uploading";
//   }

//   function isDone(s) {
//     var v = norm(s);
//     return v === "done" || v === "completed" || v === "complete" || v === "success" || v === "finished";
//   }

//   function isError(s) {
//     var v = norm(s);
//     return v === "error" || v === "failed" || v === "fail" || v === "err";
//   }

//   function finishRun(wasError, outputs) {
//     if (finished) return;
//     finished = true;

//     stopPolling();
//     window.__consoleRunning = false;
//     hideOverlay();

//     if (wasError) {
//       setProgress(0, "Failed");
//       logLine("Run failed.", "err");
//       window.toast("Agent failed", "err");
//     } else {
//       setProgress(100, "Completed");
//       renderOutputs(outputs || []);
//       logLine("Done! Ready for next run.", "ok");
//       window.toast("✓ Run complete");
//     }

//     try {
//       var fi = g("files");
//       if (fi) fi.value = "";
//     } catch (e) {}

//     var fz = g("fileZone");
//     if (fz) fz.classList.remove("has-files");

//     var fs = g("fileStatus");
//     if (fs) fs.textContent = "";

//     unlockAll();
//   }

//   function startWatching(jobId) {
//     stopPolling();

//     var logsSeen = 0;
//     var doneAt = null;
//     var errCount = 0;

//     window.__pollTimer = setInterval(function () {
//       if (finished) {
//         stopPolling();
//         return;
//       }

//       if (!window.__job.id) {
//         stopPolling();
//         return;
//       }

//       fetch("/api/job-status?jobId=" + encodeURIComponent(jobId), {
//         cache: "no-store",
//         credentials: "include"
//       })
//         .then(function (r) {
//           if (finished) return;

//           if (r.status === 404) {
//             var logText = ((g("logContent") || {}).textContent || "").toLowerCase();
//             var ok = logText.includes("✓ completed") ||
//                      logText.includes("upsert complete") ||
//                      logText.includes("run completed") ||
//                      logText.includes("done! ready for next run.");
//             finishRun(!ok, []);
//             return Promise.reject("stop");
//           }

//           return r.json();
//         })
//         .then(function (data) {
//           if (finished) return;
//           if (!data) return;

//           if (!data.ok) {
//             var msg = String(data.error || "").toLowerCase();
//             if (msg.includes("not found") || msg.includes("job")) {
//               var logText = ((g("logContent") || {}).textContent || "").toLowerCase();
//               var ok = logText.includes("✓ completed") ||
//                        logText.includes("upsert complete") ||
//                        logText.includes("run completed") ||
//                        logText.includes("done! ready for next run.");
//               finishRun(!ok, []);
//             }
//             return;
//           }

//           errCount = 0;
//           var status = String(data.status || "").trim();

//           var logs = Array.isArray(data.logs) ? data.logs : [];
//           if (logs.length > logsSeen) {
//             logs.slice(logsSeen).forEach(function (line) {
//               logLine(line, "info");
//             });
//             logsSeen = logs.length;
//           }

//           if (typeof data.pct === "number") {
//             setProgress(data.pct, data.hint || "");
//           }

//           if (isDone(status)) {
//             finishRun(false, data.outputs);
//             return;
//           }

//           if (isError(status) || data.error) {
//             finishRun(true, []);
//             return;
//           }

//           if (isRunning(status) || isUploading(status)) {
//             return;
//           }

//           var lastLog = logs.length ? String(logs[logs.length - 1] || "").toLowerCase() : "";
//           var logDone =
//             lastLog.includes("✓ completed") ||
//             lastLog.includes("upsert complete") ||
//             lastLog.includes("run completed") ||
//             lastLog.includes("run finished") ||
//             lastLog.includes("done! ready for next run.");

//           if (logDone && !doneAt) doneAt = Date.now();
//           if (doneAt && (Date.now() - doneAt) > 3000) {
//             finishRun(false, data.outputs || []);
//           }
//         })
//         .catch(function (e) {
//           if (e === "stop" || finished) return;

//           errCount++;
//           if (errCount >= 5) {
//             var logText = ((g("logContent") || {}).textContent || "").toLowerCase();
//             var ok = logText.includes("✓ completed") ||
//                      logText.includes("upsert complete") ||
//                      logText.includes("done! ready for next run.");
//             finishRun(!ok, []);
//           } else {
//             setProgress(0, "Reconnecting…");
//           }
//         });
//     }, 1500);
//   }

//   function needsMY(name) {
//     return name === "Receivable Stock" || name === "Receivable Trade Discount";
//   }

//   function applyAgent(name) {
//     var pm = g("pillMode");
//     var po = g("pillOutput");
//     var myb = g("monthYearBox");

//     if (!name || !PROJECTS[name]) {
//       if (pm) pm.textContent = "MODE: —";
//       if (po) po.textContent = "OUT: —";
//       if (myb) myb.classList.add("hidden");
//       return;
//     }

//     var cfg = PROJECTS[name];
//     var single = cfg.mode === "single";

//     if (pm) pm.textContent = "MODE: " + String(cfg.mode || "").toUpperCase();
//     if (po) po.textContent = "OUT: " + (cfg.outputName || "—");

//     var fi = g("files");
//     if (fi) fi.multiple = !single;

//     var lbl = g("fileZoneLabel");
//     if (lbl) {
//       lbl.innerHTML = single
//         ? "Drop file here or <span>browse</span>"
//         : "Drop files here or <span>browse</span>";
//     }

//     var sub = g("fileZoneSub");
//     if (sub) {
//       sub.textContent = single
//         ? "Single file · .xlsx / .csv"
//         : "Multiple files · .xlsx / .csv";
//     }

//     if (myb) {
//       if (needsMY(name)) {
//         myb.classList.remove("hidden");
//         var yr = g("year");
//         if (yr && !yr.value) yr.value = String(new Date().getFullYear());
//       } else {
//         myb.classList.add("hidden");
//       }
//     }
//   }

//   async function loadAgents() {
//     try {
//       var res = await fetch("/projects", { cache: "no-store", credentials: "include" });
//       if (!res.ok) {
//         logLine("Failed to load agents: HTTP " + res.status, "err");
//         return;
//       }

//       PROJECTS = await res.json();

//       var sel = g("agentSelect");
//       if (!sel) return;

//       sel.innerHTML = "";

//       var ph = document.createElement("option");
//       ph.value = "";
//       ph.textContent = "— Select Agent —";
//       sel.appendChild(ph);

//       Object.keys(PROJECTS).forEach(function (name) {
//         var opt = document.createElement("option");
//         opt.value = name;
//         opt.textContent = name;
//         sel.appendChild(opt);
//       });

//       if (window.__job.agent) {
//         sel.value = window.__job.agent;
//         applyAgent(sel.value);
//       } else {
//         applyAgent("");
//       }

//       syncBtn();
//     } catch (e) {
//       logLine("Network error: " + e.message, "err");
//     }
//   }

//   async function runAgent() {
//     if (window.__job.id) return;

//     var sel = g("agentSelect");
//     if (!sel || !sel.value) {
//       window.toast("Select an agent", "err");
//       return;
//     }

//     var fi = g("files");
//     if (!fi || !fi.files || !fi.files.length) {
//       window.toast("Select a file", "err");
//       return;
//     }

//     var agentName = sel.value;
//     var month = ((g("month") && g("month").value) || "").trim();
//     var year = ((g("year") && g("year").value) || "").trim();

//     if (needsMY(agentName) && (!month || !year)) {
//       window.toast("Month/Year required", "err");
//       return;
//     }

//     finished = false;

//     var lc = g("logContent");
//     if (lc) lc.textContent = "";

//     var ol = g("outList");
//     if (ol) ol.innerHTML = '<div class="outputs-empty">— Running… —</div>';

//     setProgress(2, "Uploading…");
//     lockAll();

//     logLine("Agent: " + agentName, "info");
//     if (needsMY(agentName)) logLine("Period: " + month + " / " + year, "info");
//     logLine("Files: " + fi.files.length + " file(s)", "info");
//     logLine("Starting...", "info");

//     var fd = new FormData();
//     for (var i = 0; i < fi.files.length; i++) {
//       fd.append("files", fi.files[i]);
//     }

//     if (needsMY(agentName)) {
//       fd.append("month", month);
//       fd.append("year", year);
//     }

//     try {
//       var res = await fetch("/run/" + encodeURIComponent(agentName), {
//         method: "POST",
//         body: fd,
//         credentials: "include"
//       });

//       var data = await res.json();

//       if (!data.ok || !data.jobId) {
//         logLine("Error: " + (data.error || "No jobId"), "err");
//         window.toast("Failed to start", "err");
//         finished = true;
//         unlockAll();
//         return;
//       }

//       window.__job = { id: data.jobId, agent: agentName };
//       window.__consoleRunning = true;

//       logLine("Job started: " + data.jobId, "ok");
//       setProgress(5, "Queued…");
//       startWatching(data.jobId);
//     } catch (e) {
//       logLine("Network error: " + e.message, "err");
//       window.toast("Network error", "err");
//       finished = true;
//       unlockAll();
//     }
//   }

//   var fzEl = g("fileZone");
//   var fiEl = g("files");

//   if (fzEl && fiEl) {
//     fzEl.addEventListener("dragover", function (e) {
//       if (!fzEl.classList.contains("locked")) {
//         e.preventDefault();
//         fzEl.classList.add("drag-over");
//       }
//     });

//     fzEl.addEventListener("dragleave", function () {
//       fzEl.classList.remove("drag-over");
//     });

//     fzEl.addEventListener("drop", function (e) {
//       if (fzEl.classList.contains("locked")) return;

//       e.preventDefault();
//       fzEl.classList.remove("drag-over");

//       if (e.dataTransfer && e.dataTransfer.files.length) {
//         var dt = new DataTransfer();
//         Array.from(e.dataTransfer.files).forEach(function (f) {
//           dt.items.add(f);
//         });
//         fiEl.files = dt.files;
//         updateFileUI();
//       }
//     });

//     fiEl.addEventListener("change", updateFileUI);
//   }

//   function updateFileUI() {
//     var fi = g("files");
//     var fz = g("fileZone");
//     var fs = g("fileStatus");
//     var lbl = g("fileZoneLabel");

//     if (!fi || !fi.files || !fi.files.length) {
//       if (fz) fz.classList.remove("has-files");
//       if (fs) fs.textContent = "";
//       applyAgent((g("agentSelect") || {}).value || "");
//       syncBtn();
//       return;
//     }

//     if (fz) fz.classList.add("has-files");

//     var names = Array.from(fi.files).map(function (f) {
//       return f.name;
//     }).join(", ");

//     if (fs) fs.textContent = fi.files.length === 1 ? names : fi.files.length + " files selected";

//     if (lbl) {
//       lbl.innerHTML = fi.files.length === 1
//         ? '<span style="color:var(--green)">' + safe(names) + "</span>"
//         : '<span style="color:var(--green)">' + fi.files.length + " files selected</span>";
//     }

//     syncBtn();
//   }

//   var bRun = g("btnRun");
//   if (bRun) bRun.onclick = runAgent;

//   var bClear = g("btnClear");
//   if (bClear) {
//     bClear.onclick = function () {
//       if (window.__job.id) {
//         window.toast("Agent is running", "err");
//         return;
//       }

//       var lc = g("logContent");
//       if (lc) lc.textContent = "";

//       var ol = g("outList");
//       if (ol) ol.innerHTML = '<div class="outputs-empty">— No outputs yet —</div>';

//       setProgress(0, "Idle");

//       try {
//         var fi = g("files");
//         if (fi) fi.value = "";
//       } catch (e) {}

//       var fz = g("fileZone");
//       if (fz) fz.classList.remove("has-files");

//       var fs = g("fileStatus");
//       if (fs) fs.textContent = "";

//       applyAgent((g("agentSelect") || {}).value || "");
//       finished = true;
//       unlockAll();
//     };
//   }

//   var agSel = g("agentSelect");
//   if (agSel) {
//     agSel.onchange = function () {
//       if (window.__job.id) return;

//       applyAgent(agSel.value);

//       try {
//         var fi = g("files");
//         if (fi) fi.value = "";
//       } catch (e) {}

//       var fz = g("fileZone");
//       if (fz) fz.classList.remove("has-files");

//       var fs = g("fileStatus");
//       if (fs) fs.textContent = "";

//       setProgress(0, "Idle");
//       syncBtn();
//     };
//   }

//   if (window.__job && window.__job.id) {
//     finished = false;
//     window.__consoleRunning = true;
//     lockAll();
//     logLine("Resuming job: " + window.__job.id, "info");
//     startWatching(window.__job.id);
//   } else {
//     window.__job = { id: null, agent: null };
//     window.__consoleRunning = false;
//     finished = true;
//     hideOverlay();
//     setProgress(0, "Idle");
//   }

//   if (window.__consoleRunning) showOverlay();
//   else hideOverlay();

//   loadAgents();
// };

// FILE: public/js/app.js

window.__sectionInit = window.__sectionInit || {};
window.__job = window.__job || { id: null, agent: null };
window.__pollTimer = window.__pollTimer || null;
window.__consoleRunning = window.__consoleRunning || false;
window.toast = window.toast || function(msg){ try { console.log(msg); } catch(e) {} };

window.toast = window.toast || function (msg, type) {
  var el = document.getElementById("toast");
  if (!el) return;
  el.textContent = msg;
  el.className = "toast show" + (type ? " toast-" + type : "");
  clearTimeout(window.__toastT);
  window.__toastT = setTimeout(function () {
    el.classList.remove("show");
  }, 2800);
};

async function hydrateUserChip() {
  try {
    var res = await fetch("/me", { cache: "no-store", credentials: "include" });
    if (!res.ok) return;
    var d = await res.json();
    if (!d.ok || !d.user) return;
    var dn = d.user.displayName || d.user.name || d.user.email || "User";
    var e;
    e = document.getElementById("userName");
    if (e) e.textContent = dn;
    e = document.getElementById("userEmail");
    if (e) e.textContent = d.user.email || "—";
    e = document.getElementById("userAvatar");
    if (e) e.textContent = dn.trim()[0].toUpperCase();
  } catch (e) {}
}
document.addEventListener("DOMContentLoaded", hydrateUserChip);

function stopPolling() {
  if (window.__pollTimer) {
    clearInterval(window.__pollTimer);
    window.__pollTimer = null;
  }
}

window.resetConsoleUI = function () {
  if (window.__job && window.__job.id) {
    window.toast && window.toast("Agent is running", "err");
    return;
  }

  stopPolling();
  window.__job = { id: null, agent: null };
  window.__consoleRunning = false;

  var g = function (id) { return document.getElementById(id); };

  var lc = g("logContent");
  if (lc) lc.textContent = "";

  var ol = g("outList");
  if (ol) ol.innerHTML = '<div class="outputs-empty">— No outputs yet —</div>';

  try {
    var fi = g("files");
    if (fi) fi.value = "";
  } catch (e) {}

  var fz = g("fileZone");
  if (fz) {
    fz.classList.remove("has-files");
    fz.classList.remove("locked");
  }

  var fs = g("fileStatus");
  if (fs) fs.textContent = "";

  var sel = g("agentSelect");
  if (sel) sel.value = "";

  var month = g("month");
  if (month) month.value = "";

  var year = g("year");
  if (year) year.value = "";

  var pm = g("pillMode");
  if (pm) pm.textContent = "MODE: —";

  var po = g("pillOutput");
  if (po) po.textContent = "OUT: —";

  var myb = g("monthYearBox");
  if (myb) myb.classList.add("hidden");

  var pct = g("progressPct");
  if (pct) pct.textContent = "0%";

  var fill = g("progressFill");
  if (fill) fill.style.width = "0%";

  var hint = g("progressHint");
  if (hint) hint.textContent = "Idle";

  var ov = g("consoleLock");
  if (ov) {
    ov.classList.remove("show");
    ov.setAttribute("aria-hidden", "true");
  }

  ["agentSelect", "files", "month", "year", "btnClear", "btnClearAll"].forEach(function (id) {
    var el = g(id);
    if (el) el.disabled = false;
  });

  var runBtn = g("btnRun");
  if (runBtn) runBtn.disabled = true;

  window.toast && window.toast("Console reset");
};

// ═══════════════════════════════════════════════════════════════
window.__sectionInit.console = function () {
// ═══════════════════════════════════════════════════════════════

  var PROJECTS = {};
  var finished = false;

  function g(id) {
    return document.getElementById(id);
  }

  function safe(s) {
    return String(s || "").replace(/[&<>"']/g, function (c) {
      return {
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#039;"
      }[c];
    });
  }

  function showOverlay() {
    var ov = g("consoleLock");
    if (!ov) return;
    ov.classList.add("show");
    ov.setAttribute("aria-hidden", "false");
  }

  function hideOverlay() {
    var ov = g("consoleLock");
    if (!ov) return;
    ov.classList.remove("show");
    ov.setAttribute("aria-hidden", "true");
  }

  function setProgress(pct, hint) {
    pct = Math.max(0, Math.min(100, Number(pct) || 0));
    var p = g("progressPct");
    if (p) p.textContent = pct + "%";
    var f = g("progressFill");
    if (f) f.style.width = pct + "%";
    var h = g("progressHint");
    if (h) h.textContent = hint || "";
  }

  function logLine(txt, type) {
    var box = g("logContent");
    if (!box) return;
    var pre = type === "ok" ? "✓ " : type === "err" ? "✗ " : "→ ";
    box.textContent += pre + txt + "\n";
    box.scrollTop = box.scrollHeight;
  }

  function renderOutputs(outputs) {
    var ol = g("outList");
    if (!ol) return;

    if (!outputs || !outputs.length) {
      ol.innerHTML = '<div class="outputs-empty">Run complete — check Base Dump for files.</div>';
      return;
    }

    ol.innerHTML = outputs.map(function (f) {
      var url = "https://drive.google.com/file/d/" + f.id + "/view";
      return '<a class="out-link" target="_blank" rel="noreferrer" href="' + url + '">' + safe(f.name) + "</a>";
    }).join("");
  }

  function lockAll() {
    window.__consoleRunning = true;

    ["agentSelect", "files", "month", "year", "btnClear"].forEach(function (id) {
      var e = g(id);
      if (e) e.disabled = true;
    });

    var bRun = g("btnRun");
    if (bRun) bRun.disabled = true;

    var bReset = g("btnClearAll");
    if (bReset) bReset.disabled = false;

    var fz = g("fileZone");
    if (fz) fz.classList.add("locked");

    showOverlay();
  }

  function unlockAll() {
    stopPolling();

    window.__job = { id: null, agent: null };
    window.__consoleRunning = false;

    ["agentSelect", "files", "month", "year", "btnClear"].forEach(function (id) {
      var e = g(id);
      if (e) e.disabled = false;
    });

    var bRun = g("btnRun");
    if (bRun) bRun.disabled = false;

    var bReset = g("btnClearAll");
    if (bReset) bReset.disabled = false;

    var fz = g("fileZone");
    if (fz) fz.classList.remove("locked");

    hideOverlay();
    syncBtn();
  }

  function syncBtn() {
    var bRun = g("btnRun");
    if (!bRun) return;

    var agOk = !!(g("agentSelect") && g("agentSelect").value);
    var fi = g("files");
    var fiOk = !!(fi && fi.files && fi.files.length);

    bRun.disabled = !!(window.__job.id) || !agOk || !fiOk;
  }

  function norm(s) {
    return String(s || "").trim().toLowerCase();
  }

  function isRunning(s) {
    var v = norm(s);
    return v === "running" || v === "queued" || v === "processing" || v === "started" || v === "preparing";
  }

  function isUploading(s) {
    return norm(s) === "uploading";
  }

  function isDone(s) {
    var v = norm(s);
    return v === "done" || v === "completed" || v === "complete" || v === "success" || v === "finished";
  }

  function isError(s) {
    var v = norm(s);
    return v === "error" || v === "failed" || v === "fail" || v === "err";
  }

  function finishRun(wasError, outputs) {
    if (finished) return;
    finished = true;

    stopPolling();
    window.__consoleRunning = false;
    hideOverlay();

    if (wasError) {
      setProgress(0, "Failed");
      logLine("Run failed.", "err");
      window.toast("Agent failed", "err");
    } else {
      setProgress(100, "Completed");
      renderOutputs(outputs || []);
      logLine("Done! Ready for next run.", "ok");
      window.toast("✓ Run complete");
    }

    try {
      var fi = g("files");
      if (fi) fi.value = "";
    } catch (e) {}

    var fz = g("fileZone");
    if (fz) fz.classList.remove("has-files");

    var fs = g("fileStatus");
    if (fs) fs.textContent = "";

    unlockAll();
  }

  function startWatching(jobId) {
    stopPolling();

    var logsSeen = 0;
    var doneAt = null;
    var errCount = 0;

    window.__pollTimer = setInterval(function () {
      if (finished) {
        stopPolling();
        return;
      }

      if (!window.__job.id) {
        stopPolling();
        return;
      }

      fetch("/api/job-status?jobId=" + encodeURIComponent(jobId), {
        cache: "no-store",
        credentials: "include"
      })
        .then(function (r) {
          if (finished) return;

          if (r.status === 404) {
            var logText = ((g("logContent") || {}).textContent || "").toLowerCase();
            var ok = logText.includes("✓ completed") ||
                     logText.includes("upsert complete") ||
                     logText.includes("run completed") ||
                     logText.includes("done! ready for next run.");
            finishRun(!ok, []);
            return Promise.reject("stop");
          }

          return r.json();
        })
        .then(function (data) {
          if (finished) return;
          if (!data) return;

          if (!data.ok) {
            var msg = String(data.error || "").toLowerCase();
            if (msg.includes("not found") || msg.includes("job")) {
              var logText = ((g("logContent") || {}).textContent || "").toLowerCase();
              var ok = logText.includes("✓ completed") ||
                       logText.includes("upsert complete") ||
                       logText.includes("run completed") ||
                       logText.includes("done! ready for next run.");
              finishRun(!ok, []);
            }
            return;
          }

          errCount = 0;
          var status = String(data.status || "").trim();

          var logs = Array.isArray(data.logs) ? data.logs : [];
          if (logs.length > logsSeen) {
            logs.slice(logsSeen).forEach(function (line) {
              logLine(line, "info");
            });
            logsSeen = logs.length;
          }

          if (typeof data.pct === "number") {
            setProgress(data.pct, data.hint || "");
          }

          if (isDone(status)) {
            finishRun(false, data.outputs);
            return;
          }

          if (isError(status) || data.error) {
            finishRun(true, []);
            return;
          }

          if (isRunning(status) || isUploading(status)) {
            return;
          }

          var lastLog = logs.length ? String(logs[logs.length - 1] || "").toLowerCase() : "";
          var logDone =
            lastLog.includes("✓ completed") ||
            lastLog.includes("upsert complete") ||
            lastLog.includes("run completed") ||
            lastLog.includes("run finished") ||
            lastLog.includes("done! ready for next run.");

          if (logDone && !doneAt) doneAt = Date.now();
          if (doneAt && (Date.now() - doneAt) > 3000) {
            finishRun(false, data.outputs || []);
          }
        })
        .catch(function (e) {
          if (e === "stop" || finished) return;

          errCount++;
          if (errCount >= 5) {
            var logText = ((g("logContent") || {}).textContent || "").toLowerCase();
            var ok = logText.includes("✓ completed") ||
                     logText.includes("upsert complete") ||
                     logText.includes("done! ready for next run.");
            finishRun(!ok, []);
          } else {
            setProgress(0, "Reconnecting…");
          }
        });
    }, 1500);
  }

  // ✅ UPDATED: dynamic check using PROJECTS config instead of hardcoded names
  function needsMY(name) {
    return !!(PROJECTS[name] && PROJECTS[name].needsMonthYear);
  }

  function applyAgent(name) {
    var pm = g("pillMode");
    var po = g("pillOutput");
    var myb = g("monthYearBox");

    if (!name || !PROJECTS[name]) {
      if (pm) pm.textContent = "MODE: —";
      if (po) po.textContent = "OUT: —";
      if (myb) myb.classList.add("hidden");
      return;
    }

    var cfg = PROJECTS[name];
    var single = cfg.mode === "single";

    if (pm) pm.textContent = "MODE: " + String(cfg.mode || "").toUpperCase();
    if (po) po.textContent = "OUT: " + (cfg.outputName || "—");

    var fi = g("files");
    if (fi) fi.multiple = !single;

    var lbl = g("fileZoneLabel");
    if (lbl) {
      lbl.innerHTML = single
        ? "Drop file here or <span>browse</span>"
        : "Drop files here or <span>browse</span>";
    }

    var sub = g("fileZoneSub");
    if (sub) {
      sub.textContent = single
        ? "Single file · .xlsx / .csv"
        : "Multiple files · .xlsx / .csv";
    }

    if (myb) {
      if (needsMY(name)) {
        myb.classList.remove("hidden");
        var yr = g("year");
        if (yr && !yr.value) yr.value = String(new Date().getFullYear());
      } else {
        myb.classList.add("hidden");
      }
    }
  }

  async function loadAgents() {
    try {
      var res = await fetch("/projects", { cache: "no-store", credentials: "include" });
      if (!res.ok) {
        logLine("Failed to load agents: HTTP " + res.status, "err");
        return;
      }

      PROJECTS = await res.json();

      var sel = g("agentSelect");
      if (!sel) return;

      sel.innerHTML = "";

      var ph = document.createElement("option");
      ph.value = "";
      ph.textContent = "— Select Agent —";
      sel.appendChild(ph);

      Object.keys(PROJECTS).forEach(function (name) {
        var opt = document.createElement("option");
        opt.value = name;
        opt.textContent = name;
        sel.appendChild(opt);
      });

      if (window.__job.agent) {
        sel.value = window.__job.agent;
        applyAgent(sel.value);
      } else {
        applyAgent("");
      }

      syncBtn();
    } catch (e) {
      logLine("Network error: " + e.message, "err");
    }
  }

  async function runAgent() {
    if (window.__job.id) return;

    var sel = g("agentSelect");
    if (!sel || !sel.value) {
      window.toast("Select an agent", "err");
      return;
    }

    var fi = g("files");
    if (!fi || !fi.files || !fi.files.length) {
      window.toast("Select a file", "err");
      return;
    }

    var agentName = sel.value;
    var month = ((g("month") && g("month").value) || "").trim();
    var year = ((g("year") && g("year").value) || "").trim();

    if (needsMY(agentName) && (!month || !year)) {
      window.toast("Month/Year required", "err");
      return;
    }

    finished = false;

    var lc = g("logContent");
    if (lc) lc.textContent = "";

    var ol = g("outList");
    if (ol) ol.innerHTML = '<div class="outputs-empty">— Running… —</div>';

    setProgress(2, "Uploading…");
    lockAll();

    logLine("Agent: " + agentName, "info");
    if (needsMY(agentName)) logLine("Period: " + month + " / " + year, "info");
    logLine("Files: " + fi.files.length + " file(s)", "info");
    logLine("Starting...", "info");

    var fd = new FormData();
    for (var i = 0; i < fi.files.length; i++) {
      fd.append("files", fi.files[i]);
    }

    if (needsMY(agentName)) {
      fd.append("month", month);
      fd.append("year", year);
    }

    try {
      var res = await fetch("/run/" + encodeURIComponent(agentName), {
        method: "POST",
        body: fd,
        credentials: "include"
      });

      var data = await res.json();

      if (!data.ok || !data.jobId) {
        logLine("Error: " + (data.error || "No jobId"), "err");
        window.toast("Failed to start", "err");
        finished = true;
        unlockAll();
        return;
      }

      window.__job = { id: data.jobId, agent: agentName };
      window.__consoleRunning = true;

      logLine("Job started: " + data.jobId, "ok");
      setProgress(5, "Queued…");
      startWatching(data.jobId);
    } catch (e) {
      logLine("Network error: " + e.message, "err");
      window.toast("Network error", "err");
      finished = true;
      unlockAll();
    }
  }

  var fzEl = g("fileZone");
  var fiEl = g("files");

  if (fzEl && fiEl) {
    fzEl.addEventListener("dragover", function (e) {
      if (!fzEl.classList.contains("locked")) {
        e.preventDefault();
        fzEl.classList.add("drag-over");
      }
    });

    fzEl.addEventListener("dragleave", function () {
      fzEl.classList.remove("drag-over");
    });

    fzEl.addEventListener("drop", function (e) {
      if (fzEl.classList.contains("locked")) return;

      e.preventDefault();
      fzEl.classList.remove("drag-over");

      if (e.dataTransfer && e.dataTransfer.files.length) {
        var dt = new DataTransfer();
        Array.from(e.dataTransfer.files).forEach(function (f) {
          dt.items.add(f);
        });
        fiEl.files = dt.files;
        updateFileUI();
      }
    });

    fiEl.addEventListener("change", updateFileUI);
  }

  function updateFileUI() {
    var fi = g("files");
    var fz = g("fileZone");
    var fs = g("fileStatus");
    var lbl = g("fileZoneLabel");

    if (!fi || !fi.files || !fi.files.length) {
      if (fz) fz.classList.remove("has-files");
      if (fs) fs.textContent = "";
      applyAgent((g("agentSelect") || {}).value || "");
      syncBtn();
      return;
    }

    if (fz) fz.classList.add("has-files");

    var names = Array.from(fi.files).map(function (f) {
      return f.name;
    }).join(", ");

    if (fs) fs.textContent = fi.files.length === 1 ? names : fi.files.length + " files selected";

    if (lbl) {
      lbl.innerHTML = fi.files.length === 1
        ? '<span style="color:var(--green)">' + safe(names) + "</span>"
        : '<span style="color:var(--green)">' + fi.files.length + " files selected</span>";
    }

    syncBtn();
  }

  var bRun = g("btnRun");
  if (bRun) bRun.onclick = runAgent;

  var bClear = g("btnClear");
  if (bClear) {
    bClear.onclick = function () {
      if (window.__job.id) {
        window.toast("Agent is running", "err");
        return;
      }

      var lc = g("logContent");
      if (lc) lc.textContent = "";

      var ol = g("outList");
      if (ol) ol.innerHTML = '<div class="outputs-empty">— No outputs yet —</div>';

      setProgress(0, "Idle");

      try {
        var fi = g("files");
        if (fi) fi.value = "";
      } catch (e) {}

      var fz = g("fileZone");
      if (fz) fz.classList.remove("has-files");

      var fs = g("fileStatus");
      if (fs) fs.textContent = "";

      applyAgent((g("agentSelect") || {}).value || "");
      finished = true;
      unlockAll();
    };
  }

  var agSel = g("agentSelect");
  if (agSel) {
    agSel.onchange = function () {
      if (window.__job.id) return;

      applyAgent(agSel.value);

      try {
        var fi = g("files");
        if (fi) fi.value = "";
      } catch (e) {}

      var fz = g("fileZone");
      if (fz) fz.classList.remove("has-files");

      var fs = g("fileStatus");
      if (fs) fs.textContent = "";

      setProgress(0, "Idle");
      syncBtn();
    };
  }

  if (window.__job && window.__job.id) {
    finished = false;
    window.__consoleRunning = true;
    lockAll();
    logLine("Resuming job: " + window.__job.id, "info");
    startWatching(window.__job.id);
  } else {
    window.__job = { id: null, agent: null };
    window.__consoleRunning = false;
    finished = true;
    hideOverlay();
    setProgress(0, "Idle");
  }

  if (window.__consoleRunning) showOverlay();
  else hideOverlay();

  loadAgents();
};