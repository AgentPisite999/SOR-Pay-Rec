// =========================
// FILE: public/js/report.js
// Inc.5 SOR Portal — Reports
// ✅ Multi-select checkboxes for Month & Year
// ✅ Auto-loads filters when agent already selected on init
// ✅ All / None quick actions + badge count
// ✅ Client-side filtering + XLSX export
// =========================

window.__sectionInit = window.__sectionInit || {};

window.__sectionInit.report = function () {

  var MO = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];

  var agSel  = document.getElementById("rpt-agent");
  if (!agSel) return;
  if (agSel.dataset.bound === "1") return;
  agSel.dataset.bound = "1";

  var showB  = document.getElementById("rpt-show");
  var refB   = document.getElementById("rpt-refresh");
  var dlB    = document.getElementById("rpt-dl");
  var retryB = document.getElementById("s-retry");
  var pill   = document.getElementById("rpt-pill");
  var cnt    = document.getElementById("rpt-count");
  var sIdle  = document.getElementById("s-idle");
  var sLF    = document.getElementById("s-lf");
  var sLD    = document.getElementById("s-ld");
  var sErr   = document.getElementById("s-err");
  var sTable = document.getElementById("s-table");
  var lfName = document.getElementById("s-lf-name");
  var ldDesc = document.getElementById("s-ld-desc");
  var errMsg = document.getElementById("s-err-msg");
  var thead  = document.getElementById("rpt-thead");
  var tbody  = document.getElementById("rpt-tbody");

  var curAgent = "", curH = [], curR = [];

  // excluded = items user unchecked; empty set = all selected
  var excluded  = { month: new Set(), year: new Set() };
  var available = { month: [], year: [] };

  // ─── Utilities ────────────────────────────────────────────
  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/[&<>"']/g, function(c) {
        return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[c];
      });
  }
  function isNum(s) {
    return s !== "-" && s !== "" && !isNaN(parseFloat(s)) && isFinite(s);
  }

  // ─── Panel switcher ───────────────────────────────────────
  function panel(active) {
    [sIdle, sLF, sLD, sErr, sTable].forEach(function(el) {
      if (el) el.style.display = "none";
    });
    if (active) {
      active.style.display = (active === sTable) ? "block" : "flex";
    }
    if (pill) pill.style.display = (active === sTable) ? "" : "none";
    if (dlB)  dlB.style.display  = (active === sTable) ? "" : "none";
  }

  // ─── API helper ───────────────────────────────────────────
  function callAPI(url, cb) {
    fetch(url, { credentials: "include", cache: "no-store" })
      .then(function(r) {
        var ct = r.headers.get("content-type") || "";
        if (!ct.includes("application/json")) {
          return r.text().then(function(t) {
            throw new Error("HTTP " + r.status + " — not JSON. " + t.slice(0, 200));
          });
        }
        return r.json();
      })
      .then(function(j) {
        if (!j.ok) throw new Error(j.error || "ok:false");
        cb(null, j);
      })
      .catch(function(e) { cb(e, null); });
  }

  // ─── Multi-select: build checkboxes ───────────────────────
  function buildList(type) {
    var listEl = document.getElementById("ms-" + type + "-list");
    if (!listEl) return;
    listEl.innerHTML = "";

    var items = available[type];
    if (!items.length) {
      listEl.innerHTML = '<div style="padding:10px 12px;font-family:\'JetBrains Mono\',monospace;font-size:11px;color:rgba(240,237,232,.3)">No data available</div>';
      return;
    }

    items.forEach(function(val) {
      var row = document.createElement("div");
      row.className = "ms-item";

      var cb = document.createElement("input");
      cb.type      = "checkbox";
      cb.className = "ms-cb";
      cb.value     = val;
      cb.checked   = !excluded[type].has(val); // checked if NOT excluded

      var lbl = document.createElement("span");
      lbl.className   = "ms-item-label";
      lbl.textContent = val;

      row.appendChild(cb);
      row.appendChild(lbl);

      // clicking anywhere on the row toggles
      row.addEventListener("mousedown", function(e) {
        e.preventDefault(); // prevent losing focus
        if (e.target === cb) return; // cb handles itself
        cb.checked = !cb.checked;
        syncExcluded(type, val, cb.checked);
      });
      cb.addEventListener("change", function() {
        syncExcluded(type, val, cb.checked);
      });

      listEl.appendChild(row);
    });
  }

  function syncExcluded(type, val, isChecked) {
    if (isChecked) {
      excluded[type].delete(val);
    } else {
      excluded[type].add(val);
    }
    refreshTriggerLabel(type);
  }

  function refreshTriggerLabel(type) {
    var btn   = document.getElementById("ms-" + type + "-btn");
    var label = document.getElementById("ms-" + type + "-label");
    if (!btn || !label) return;

    // remove old badge
    var old = btn.querySelector(".ms-badge");
    if (old) old.remove();

    var total    = available[type].length;
    var numExcl  = excluded[type].size;
    var numSel   = total - numExcl;

    if (numExcl === 0 || numSel === total) {
      label.textContent = type === "month" ? "All Months" : "All Years";
    } else if (numSel === 0) {
      label.textContent = "None selected";
    } else {
      var sel = available[type].filter(function(v) { return !excluded[type].has(v); });
      var txt = sel.join(", ");
      label.textContent = txt.length > 18 ? txt.slice(0, 18) + "…" : txt;

      var badge       = document.createElement("span");
      badge.className = "ms-badge";
      badge.textContent = numSel;
      btn.insertBefore(badge, btn.querySelector(".ms-chevron"));
    }
  }

  function setAllChecked(type, check) {
    var listEl = document.getElementById("ms-" + type + "-list");
    if (!listEl) return;

    if (check) {
      excluded[type].clear();
    } else {
      available[type].forEach(function(v) { excluded[type].add(v); });
    }
    listEl.querySelectorAll(".ms-cb").forEach(function(cb) {
      cb.checked = check;
    });
    refreshTriggerLabel(type);
  }

  // ─── Dropdown open / close ────────────────────────────────
  function openDD(type) {
    var btn = document.getElementById("ms-" + type + "-btn");
    var dd  = document.getElementById("ms-" + type + "-dd");
    if (btn) btn.classList.add("open");
    if (dd)  { dd.style.display = "block"; dd.style.animation = "ddFadeIn .15s ease both"; }
  }
  function closeDD(type) {
    var btn = document.getElementById("ms-" + type + "-btn");
    var dd  = document.getElementById("ms-" + type + "-dd");
    if (btn) btn.classList.remove("open");
    if (dd)  dd.style.display = "none";
  }
  function toggleDD(type) {
    var dd = document.getElementById("ms-" + type + "-dd");
    if (!dd) return;
    if (dd.style.display === "none" || !dd.style.display) openDD(type);
    else closeDD(type);
  }
  function setMsDisabled(type, val) {
    var btn = document.getElementById("ms-" + type + "-btn");
    if (btn) btn.disabled = val;
    if (val) closeDD(type);
  }

  // close on outside click
  document.addEventListener("mousedown", function(e) {
    ["month","year"].forEach(function(t) {
      var wrap = document.getElementById("ms-" + t + "-wrap");
      if (wrap && !wrap.contains(e.target)) closeDD(t);
    });
  });

  // toggle buttons
  ["month","year"].forEach(function(t) {
    var btn = document.getElementById("ms-" + t + "-btn");
    if (!btn) return;
    btn.addEventListener("click", function(e) {
      e.stopPropagation();
      if (!btn.disabled) toggleDD(t);
    });
  });

  // All / None buttons
  document.querySelectorAll(".ms-act-btn").forEach(function(b) {
    b.addEventListener("mousedown", function(e) {
      e.stopPropagation();
      e.preventDefault();
      setAllChecked(b.dataset.ms, b.dataset.act === "all");
    });
  });

  // ─── Get filter arrays for filtering ─────────────────────
  // returns null = no filter, [] = none selected, [..] = specific
  function getFilter(type) {
    var excl = excluded[type];
    if (excl.size === 0) return null; // all
    return available[type].filter(function(v) { return !excl.has(v); });
  }

  // ─── Load filters from API ────────────────────────────────
  function loadFilters(agent, silent) {
    setMsDisabled("month", true);
    setMsDisabled("year",  true);
    if (showB) showB.disabled = true;
    if (refB)  refB.disabled  = true;

    if (!silent) {
      available.month = []; available.year = [];
      excluded.month  = new Set(); excluded.year = new Set();
      if (lfName) lfName.textContent = agent;
      panel(sLF);
    } else {
      if (refB) refB.classList.add("rpt-spin");
    }

    callAPI(
      "/api/report-data?agent=" + encodeURIComponent(agent) + "&month=ALL&year=ALL",
      function(err, data) {
        if (refB) refB.classList.remove("rpt-spin");

        if (err) {
          if (errMsg) errMsg.textContent = err.message;
          panel(sErr);
          return;
        }

        // find MONTH and YEAR column indices
        var hdrs = (data.headers || []).map(function(h) { return String(h).trim().toUpperCase(); });
        var mi = hdrs.indexOf("MONTH");
        var yi = hdrs.indexOf("YEAR");

        var ms = {}, ys = {};
        (data.rows || []).forEach(function(r) {
          if (mi >= 0) { var m = String(r[mi]||"").trim().toUpperCase(); if (m) ms[m] = 1; }
          if (yi >= 0) { var y = String(r[yi]||"").trim();               if (y) ys[y] = 1; }
        });

        available.month = Object.keys(ms).sort(function(a,b) {
          var ia = MO.indexOf(a), ib = MO.indexOf(b);
          return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
        });
        available.year = Object.keys(ys).sort(function(a,b) {
          return Number(b) - Number(a);
        });

        // reset excluded (all checked by default)
        excluded.month = new Set();
        excluded.year  = new Set();

        // build checkbox lists
        buildList("month");
        buildList("year");
        refreshTriggerLabel("month");
        refreshTriggerLabel("year");

        setMsDisabled("month", false);
        setMsDisabled("year",  false);
        if (showB) showB.disabled = false;
        if (refB)  refB.disabled  = false;
        panel(sIdle);
      }
    );
  }

  // ─── Fetch & render report ────────────────────────────────
  function fetchReport() {
    var agent = agSel.value;
    if (!agent) return;
    curAgent = agent;

    var fM = getFilter("month");
    var fY = getFilter("year");
    var mDesc = fM === null ? "All Months" : (fM.length ? fM.join(", ") : "None");
    var yDesc = fY === null ? "All Years"  : (fY.length ? fY.join(", ") : "None");
    if (ldDesc) ldDesc.textContent = agent + " · " + mDesc + " · " + yDesc;
    panel(sLD);

    callAPI(
      "/api/report-data?agent=" + encodeURIComponent(agent) + "&month=ALL&year=ALL",
      function(err, data) {
        if (err) { if (errMsg) errMsg.textContent = err.message; panel(sErr); return; }

        curH = data.headers || [];
        var rows = data.rows || [];

        var hdrs = curH.map(function(h) { return String(h).trim().toUpperCase(); });
        var mi   = hdrs.indexOf("MONTH");
        var yi   = hdrs.indexOf("YEAR");

        curR = rows.filter(function(r) {
          if (fM !== null && mi >= 0) {
            var m = String(r[mi]||"").trim().toUpperCase();
            if (fM.indexOf(m) === -1) return false;
          }
          if (fY !== null && yi >= 0) {
            var y = String(r[yi]||"").trim();
            if (fY.indexOf(y) === -1) return false;
          }
          return true;
        });

        renderTable();
        panel(sTable);
        if (cnt) cnt.textContent = curR.length.toLocaleString();
      }
    );
  }

  function renderTable() {
    if (thead) {
      thead.innerHTML = "<tr>" + curH.map(function(h) {
        return "<th>" + esc(h) + "</th>";
      }).join("") + "</tr>";
    }

    if (!curR.length) {
      if (tbody) tbody.innerHTML =
        '<tr><td colspan="' + curH.length + '" style="text-align:center;padding:48px;' +
        'color:rgba(240,237,232,.22);font-family:\'JetBrains Mono\',monospace;font-size:11px">' +
        'No data for selected filters.</td></tr>';
      if (cnt) cnt.textContent = "0";
      return;
    }

    if (tbody) tbody.innerHTML = curR.map(function(row) {
      return "<tr>" + curH.map(function(h, i) {
        var raw = row[i];
        var val = (raw === undefined || raw === null || String(raw).trim() === "") ? "-" : String(raw);
        var hl  = h.trim().toLowerCase();
        var cls = (hl === "month" || hl === "year") ? "b" : (val === "-" ? "e" : (isNum(val) ? "n" : ""));
        return '<td class="' + cls + '">' + esc(val) + '</td>';
      }).join("") + "</tr>";
    }).join("");
  }

  // ─── XLSX export ──────────────────────────────────────────
  function doXLSX() {
    if (!curH.length) return;
    function run() {
      var ws = window.XLSX.utils.aoa_to_sheet([curH].concat(curR));
      ws["!cols"] = curH.map(function(h, i) {
        var mx = Math.max(h.length, Math.max.apply(null,
          curR.slice(0, 300).map(function(r) { return String(r[i]||"").length; })
        ));
        return { wch: Math.min(52, Math.max(10, mx + 2)) };
      });
      var fM = getFilter("month");
      var fY = getFilter("year");
      var mv = fM && fM.length ? "_" + fM.join("-") : "";
      var yv = fY && fY.length ? "_" + fY.join("-") : "";
      var wb = window.XLSX.utils.book_new();
      window.XLSX.utils.book_append_sheet(wb, ws, curAgent.slice(0, 31));
      window.XLSX.writeFile(wb, curAgent.replace(/\s+/g,"_") + mv + yv + ".xlsx");
      window.toast && window.toast("✓ Downloaded");
    }
    if (window.XLSX) { run(); return; }
    var s   = document.createElement("script");
    s.src   = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    s.onload = run;
    document.head.appendChild(s);
  }

  // ─── Events ───────────────────────────────────────────────
  agSel.addEventListener("change", function() {
    closeDD("month"); closeDD("year");
    if (!agSel.value) {
      setMsDisabled("month", true);
      setMsDisabled("year",  true);
      if (showB) showB.disabled = true;
      if (refB)  refB.disabled  = true;
      panel(sIdle);
      return;
    }
    loadFilters(agSel.value, false);
  });

  if (refB)   refB.addEventListener("click",   function() { if (agSel.value) loadFilters(agSel.value, true); });
  if (showB)  showB.addEventListener("click",  fetchReport);
  if (dlB)    dlB.addEventListener("click",    doXLSX);
  if (retryB) retryB.addEventListener("click", function() {
    if (agSel.value && showB && !showB.disabled) fetchReport();
    else if (agSel.value) loadFilters(agSel.value, false);
    else panel(sIdle);
  });

  // ─── ✅ AUTO-LOAD if agent already has a value on init ────
  if (agSel.value) {
    loadFilters(agSel.value, false);
  } else {
    panel(sIdle);
  }

};