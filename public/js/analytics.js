// =========================
// FILE: public/js/analytics.js
// Secondary Movement Dashboard — Basic Value + Qty matrices
// =========================

window.__sectionInit = window.__sectionInit || {};

window.__sectionInit.analytics = function () {

  var monthSel = document.getElementById("dash-month");
  var yearInp  = document.getElementById("dash-year");
  var genBtn   = document.getElementById("dash-generate");
  var dlBtn    = document.getElementById("dash-dl");
  var outEl    = document.getElementById("dash-output");
  var errEl    = document.getElementById("dash-err");
  var errMsg   = document.getElementById("dash-err-msg");
  var loadEl   = document.getElementById("dash-loading");
  var idleEl   = document.getElementById("dash-idle");

  if (!monthSel || !genBtn) return;
  if (genBtn.dataset.bound === "1") return;
  genBtn.dataset.bound = "1";

  var lastData = null;

  var PARTNERS = [
    { key: "Shoppers",           label: "Shoppers" },
    { key: "Lifestyle",          label: "Lifestyle" },
    { key: "V Retail",           label: "V Retail" },
    { key: "Sabharwal Nirankar", label: "Sabharwal\nNirankar" },
    { key: "Reliance Centro",    label: "Reliance Centro" },
    { key: "Kora Retail",        label: "Kora Retail" },
    { key: "Leayan_Zuup",        label: "Leayan_Zuup" },
    { key: "__total1__",         label: "Total",       isTotal: true },
    { key: "Myntra",             label: "Myntra" },
    { key: "Flipkart",           label: "Flipkart" },
    { key: "Reliance Ajio",      label: "Reliance Ajio" },
    { key: "__total2__",         label: "Total",       isTotal: true },
    { key: "__grand__",          label: "Grand Total", isGrand: true },
  ];

  var SOR_KEYS    = ["Shoppers","Lifestyle","V Retail","Sabharwal Nirankar","Reliance Centro","Kora Retail","Leayan_Zuup"];
  var ONLINE_KEYS = ["Myntra","Flipkart","Reliance Ajio"];

  var ROWS_BV = [
    { key: "opening_stock",    label: "Opening Stock",                           bold: false },
    { key: "opening_sec_diff", label: "Opening_Secondary Difference",            bold: true  },
    { key: "primary",          label: "Primary (Including RTV)",                 bold: false },
    { key: "rate_diff",        label: "Rate Difference_Sales (RTV in Purchase)", bold: true  },
    { key: "secondary_sales",  label: "Secondary Sales",                         bold: false },
    { key: "markdown",         label: "Markdown",                                bold: false },
    { key: "closing_stock",    label: "Closing Stock",                           bold: false },
    { key: "sales_adj",        label: "Sales - Secondary Adjustment SOR",        bold: true,  isCalc: true },
    { key: "cogs",             label: "COGS",                                    bold: true,  isCalc: true },
    { key: "cogs_pct",         label: "COGS %",                                  bold: true,  isPct:  true },
  ];

  var ROWS_QTY = [
    { key: "opening_stock",    label: "Opening Stock",                           bold: false },
    { key: "opening_sec_diff", label: "Opening_Secondary Difference",            bold: false },
    { key: "primary",          label: "Primary (Including RTV)",                 bold: false },
    { key: "rate_diff",        label: "Rate Difference_Sales (RTV in Purchase)", bold: false },
    { key: "secondary_sales",  label: "Secondary Sales",                         bold: false },
    { key: "markdown",         label: "Markdown",                                bold: false },
    { key: "closing_stock",    label: "Closing Stock",                           bold: false },
    { key: "sales_adj",        label: "Sales - Secondary Adjustment SOR",        bold: true,  isCalc: true },
  ];

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function(c) {
      return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[c];
    });
  }

  function fmt(v, isPct) {
    if (v === null || v === undefined || v === "") return "-";
    if (isPct) {
      var n = parseFloat(v);
      return isNaN(n) ? "-" : (n * 100).toFixed(2) + "%";
    }
    var n = parseFloat(v);
    if (isNaN(n)) return String(v);
    return n.toLocaleString("en-IN", { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  }

  function showPanel(which) {
    [outEl, errEl, loadEl, idleEl].forEach(function(el) { if (el) el.style.display = "none"; });
    if (which) which.style.display = "";
    if (dlBtn)   dlBtn.style.display   = (which === outEl) ? "" : "none";
  }

  function addTotals(matrix, rows) {
    rows.forEach(function(row) {
      if (row.isPct) return;
      var t1 = 0, t2 = 0;
      SOR_KEYS.forEach(function(p)    { t1 += parseFloat(matrix[row.key]?.[p] || 0); });
      ONLINE_KEYS.forEach(function(p) { t2 += parseFloat(matrix[row.key]?.[p] || 0); });
      if (!matrix[row.key]) matrix[row.key] = {};
      matrix[row.key]["__total1__"] = t1;
      matrix[row.key]["__total2__"] = t2;
      matrix[row.key]["__grand__"]  = t1 + t2;
    });
    return matrix;
  }

  function buildTableHTML(matrix, rows, title, subtitle) {
    var html = "";
    html += '<div class="dash-title">' + esc(title) + "</div>";
    html += '<div class="dash-subtitle">' + esc(subtitle) + "</div>";
    html += '<div class="dash-table-wrap"><table class="dash-table">';

    // Header
    html += "<thead><tr>";
    html += '<th class="dash-th-part">Particulars</th>';
    PARTNERS.forEach(function(p) {
      var cls = p.isGrand ? "dash-th-grand" : (p.isTotal ? "dash-th-total" : "dash-th");
      html += '<th class="' + cls + '">' + esc(p.label).replace("\n", "<br>") + "</th>";
    });
    html += "</tr></thead><tbody>";

    rows.forEach(function(row) {
      var trCls = row.isCalc ? "dash-tr-calc" : (row.isPct ? "dash-tr-pct" : "");
      html += '<tr class="' + trCls + '">';
      html += '<td class="dash-td-label' + (row.bold ? " dash-bold" : "") + '">' + esc(row.label) + "</td>";
      PARTNERS.forEach(function(p) {
        var val  = matrix[row.key] ? matrix[row.key][p.key] : undefined;
        var disp = fmt(val, row.isPct);
        var cls  = p.isGrand ? "dash-td-grand" : (p.isTotal ? "dash-td-total" : "dash-td");
        if (row.isCalc || row.isPct) cls += " dash-calc-val";
        html += '<td class="' + cls + '">' + esc(disp) + "</td>";
      });
      html += "</tr>";
    });

    html += "</tbody></table></div>";
    return html;
  }

  function renderTables(data) {
    var bv  = addTotals(data.bv,  ROWS_BV);
    var qty = addTotals(data.qty, ROWS_QTY);
    var M   = data.month;
    var Y   = data.year;
    var title = "Secondary Movement for the Month of " + M + " " + Y;

    // COGS % — use server-calculated values directly
    // For Total and Grand Total columns, calculate from summed COGS / summed Basic Value
    if (!bv["cogs_pct"]) bv["cogs_pct"] = {};
    ["__total1__", "__total2__", "__grand__"].forEach(function(k) {
      var cogs = parseFloat(bv["cogs"] ? (bv["cogs"][k] || 0) : 0);
      var bvv  = parseFloat(bv["closing_stock"] ? (bv["closing_stock"][k] || 0) : 0);
      bv["cogs_pct"][k] = (bvv !== 0) ? (cogs / bvv) : 0;
    });
    // Individual partner COGS% comes directly from server (already in bv.cogs_pct)

    var html = buildTableHTML(bv, ROWS_BV, title, "Basic Value");
    html += '<div style="height:32px"></div>';
    html += buildTableHTML(qty, ROWS_QTY, title, "Qty.");

    outEl.innerHTML = html;
    showPanel(outEl);
    lastData = { bv: bv, qty: qty, month: M, year: Y };
  }

  async function generate() {
    var month = monthSel.value;
    var year  = (yearInp.value || "").trim();

    if (!month || !year) {
      window.toast && window.toast("Select month and year", "err");
      return;
    }

    showPanel(loadEl);
    genBtn.disabled = true;

    try {
      var res  = await fetch(
        "/api/dashboard?month=" + encodeURIComponent(month) + "&year=" + encodeURIComponent(year),
        { credentials: "include", cache: "no-store" }
      );
      var data = await res.json();
      if (!data.ok) throw new Error(data.error || "Failed");
      renderTables(data);
    } catch (e) {
      if (errMsg) errMsg.textContent = e.message;
      showPanel(errEl);
    } finally {
      genBtn.disabled = false;
    }
  }

  function doExport() {
    if (!lastData) return;

    function run() {
      var wb = window.XLSX.utils.book_new();

      function matrixToSheet(matrix, rows) {
        var headers = ["Particulars"].concat(PARTNERS.map(function(p) { return p.label.replace("\n", " "); }));
        var sheetRows = [headers];
        rows.forEach(function(row) {
          var r = [row.label];
          PARTNERS.forEach(function(p) {
            var val = matrix[row.key] ? matrix[row.key][p.key] : undefined;
            if (val === null || val === undefined || val === "") { r.push("-"); return; }
            if (row.isPct) {
              var n = parseFloat(val);
              r.push(isNaN(n) ? "-" : parseFloat((n * 100).toFixed(2)));
            } else {
              var n = parseFloat(val);
              r.push(isNaN(n) ? val : parseFloat(n.toFixed(2)));
            }
          });
          sheetRows.push(r);
        });
        var ws = window.XLSX.utils.aoa_to_sheet(sheetRows);
        ws["!cols"] = headers.map(function(h, i) { return { wch: i === 0 ? 44 : 16 }; });
        return ws;
      }

      var M = lastData.month, Y = lastData.year;
      window.XLSX.utils.book_append_sheet(wb, matrixToSheet(lastData.bv,  ROWS_BV),  "Basic Value");
      window.XLSX.utils.book_append_sheet(wb, matrixToSheet(lastData.qty, ROWS_QTY), "Qty");
      window.XLSX.writeFile(wb, "Secondary_Movement_" + M + "_" + Y + ".xlsx");
      window.toast && window.toast("✓ Downloaded");
    }

    if (window.XLSX) { run(); return; }
    var s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    s.onload = run;
    document.head.appendChild(s);
  }

  genBtn.addEventListener("click", generate);
  if (dlBtn) dlBtn.addEventListener("click", doExport);
  if (yearInp && !yearInp.value) yearInp.value = String(new Date().getFullYear());

  showPanel(idleEl);
};