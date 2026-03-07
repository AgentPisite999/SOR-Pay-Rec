// =========================
// FILE: public/js/basedump.js
// Base Dump section logic
// =========================

window.__sectionInit = window.__sectionInit || {};

window.__sectionInit.basedump = function () {
  const bdAgent   = document.getElementById("bdAgent");
  const bdMonth   = document.getElementById("bdMonth");
  const bdYear    = document.getElementById("bdYear");
  const bdBrowse  = document.getElementById("bdBrowse");
  const bdResults = document.getElementById("bdResults");
  const bdCount   = document.getElementById("bdCount");
  const bdResultTitle = document.getElementById("bdResultTitle");

  if (!bdAgent || !bdBrowse || !bdResults) return;

  // Set default year to current
  if (bdYear && !bdYear.value) bdYear.value = String(new Date().getFullYear());

  const MONTH_NAMES = {
    "01":"January","02":"February","03":"March","04":"April",
    "05":"May","06":"June","07":"July","08":"August",
    "09":"September","10":"October","11":"November","12":"December"
  };

  function safe(s) {
    return String(s || "").replace(/[&<>"']/g, c =>
      ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[c])
    );
  }

  function renderEmpty(msg, sub) {
    bdResults.innerHTML = `
      <div class="empty-state" style="padding:48px 24px">
        <div class="empty-icon">📂</div>
        <div class="empty-title">${msg}</div>
        <div class="empty-sub">${sub || ""}</div>
      </div>`;
    bdCount.textContent = "";
  }

  function renderLoading() {
    bdResults.innerHTML = `
      <div class="empty-state" style="padding:48px 24px">
        <div class="loader-ring" style="margin:0 auto 12px"></div>
        <div class="empty-title">Fetching runs…</div>
      </div>`;
    bdCount.textContent = "";
  }

  function renderRuns(runs, agent) {
    if (!runs.length) {
      renderEmpty("No runs found", "Try a different month / year filter");
      return;
    }

    bdCount.textContent = `${runs.length} run${runs.length === 1 ? "" : "s"}`;
    if (bdResultTitle) bdResultTitle.textContent = `Output Runs — ${agent}`;

    const rows = runs.map(r => {
      const p = r.parsed;
      const monthName = MONTH_NAMES[p.month] || p.month;
      return `
        <tr>
          <td>
            <div class="bd-folder-name">${safe(r.name)}</div>
          </td>
          <td>
            <div class="bd-date-label">${safe(p.day)} ${monthName} ${safe(p.year)}</div>
            <div class="bd-date-label" style="opacity:0.6">${safe(p.hour)}:${safe(p.minute)}:${safe(p.second)} IST</div>
          </td>
          <td>
            <span class="bd-badge">${monthName} ${safe(p.year)}</span>
          </td>
          <td style="text-align:right">
            <button
              class="btn-zip"
              data-folder-id="${safe(r.id)}"
              data-folder-name="${safe(r.name)}"
              type="button"
            >
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
              Download ZIP
            </button>
          </td>
        </tr>`;
    }).join("");

    bdResults.innerHTML = `
      <table class="bd-table">
        <thead>
          <tr>
            <th>Run Folder</th>
            <th>Date / Time</th>
            <th>Period</th>
            <th style="text-align:right">Action</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>`;

    // Attach download handlers
    bdResults.querySelectorAll(".btn-zip").forEach(btn => {
      btn.addEventListener("click", () => downloadZip(btn));
    });
  }

  async function downloadZip(btn) {
    const folderId   = btn.dataset.folderId;
    const folderName = btn.dataset.folderName;
    if (!folderId) return;

    btn.classList.add("loading");
    btn.disabled = true;
    btn.innerHTML = `<span class="bd-spinner"></span> Preparing…`;

    try {
      const url = `/api/basedump/zip?folderId=${encodeURIComponent(folderId)}&folderName=${encodeURIComponent(folderName)}`;
      const res = await fetch(url, { credentials: "include" });

      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || `HTTP ${res.status}`);
      }

      // Stream download via blob
      const blob = await res.blob();
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = folderName + ".zip";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(a.href);

      window.toast && window.toast("✓ ZIP downloaded");
    } catch (e) {
      window.toast && window.toast("Download failed: " + e.message, "err");
    } finally {
      btn.classList.remove("loading");
      btn.disabled = false;
      btn.innerHTML = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Download ZIP`;
    }
  }

  async function browseRuns() {
    const agent = bdAgent.value;
    if (!agent) {
      window.toast && window.toast("Select an agent first", "err");
      return;
    }

    const month = (bdMonth?.value || "").trim();
    const year  = (bdYear?.value  || "").trim();

    bdBrowse.disabled = true;
    renderLoading();

    try {
      const params = new URLSearchParams({ agent });
      if (month) params.set("month", month);
      if (year)  params.set("year",  year);

      const res  = await fetch(`/api/basedump/runs?${params}`, { credentials: "include", cache: "no-store" });
      const data = await res.json();

      if (!data.ok) throw new Error(data.error || "Failed to fetch runs");

      renderRuns(data.runs, agent);
    } catch (e) {
      renderEmpty("Error loading runs", e.message);
      window.toast && window.toast("Error: " + e.message, "err");
    } finally {
      bdBrowse.disabled = false;
    }
  }

  bdBrowse.addEventListener("click", browseRuns);

  // Also browse on Enter in year field
  bdYear?.addEventListener("keydown", e => {
    if (e.key === "Enter") browseRuns();
  });
};