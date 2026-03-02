// =========================
// FILE: public/js/report.js  (FULL UPDATED)
// Inc.5 SOR Portal — Reports Section Init
// ✅ Works with sections.js loader (window.__sectionInit.report)
// ✅ No ES module export (because index.html uses normal <script> tags)
// ✅ Prevents duplicate event listeners on re-open
// =========================

window.__sectionInit = window.__sectionInit || {};

window.__sectionInit.report = function () {
  initReportPage(document);
};

function initReportPage(root = document) {
  var MO = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];

  var agSel  = root.getElementById("rpt-agent");
  if (!agSel) return;

  // ✅ prevent duplicate event binding when switching tabs
  if (agSel.dataset.bound === "1") return;
  agSel.dataset.bound = "1";

  var moSel  = root.getElementById("rpt-month");
  var yrSel  = root.getElementById("rpt-year");
  var showB  = root.getElementById("rpt-show");
  var refB   = root.getElementById("rpt-refresh");
  var dlB    = root.getElementById("rpt-dl");
  var retryB = root.getElementById("s-retry");
  var pill   = root.getElementById("rpt-pill");
  var cnt    = root.getElementById("rpt-count");

  var sIdle  = root.getElementById("s-idle");
  var sLF    = root.getElementById("s-lf");
  var sLD    = root.getElementById("s-ld");
  var sErr   = root.getElementById("s-err");
  var sTable = root.getElementById("s-table");
  var lfName = root.getElementById("s-lf-name");
  var ldDesc = root.getElementById("s-ld-desc");
  var errMsg = root.getElementById("s-err-msg");
  var thead  = root.getElementById("rpt-thead");
  var tbody  = root.getElementById("rpt-tbody");

  var curAgent = "", curH = [], curR = [];

  function panel(p) {
    [sIdle,sLF,sLD,sErr,sTable].forEach(function(el){ el.style.display="none"; });
    if(p) p.style.display = (p===sTable) ? "block" : "flex";
    pill.style.display = (p===sTable) ? "" : "none";
    dlB.style.display  = (p===sTable) ? "" : "none";
  }

  function esc(s){
    return String(s).replace(/[&<>"']/g,function(c){
      return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[c];
    });
  }

  function isN(s){ return s!=="-"&&s!==""&&!isNaN(parseFloat(s))&&isFinite(s); }

  function callAPI(url, done) {
    fetch(url, { credentials:"include", cache:"no-store" })
      .then(function(r){
        var ct = r.headers.get("content-type")||"";
        if(!ct.includes("application/json")){
          return r.text().then(function(t){
            throw new Error("HTTP "+r.status+" — not JSON. Session may have expired.\n"+t.slice(0,200));
          });
        }
        return r.json();
      })
      .then(function(j){
        if(!j.ok) throw new Error(j.error||"ok:false");
        done(null,j);
      })
      .catch(function(e){ done(e,null); });
  }

  function loadFilters(agent, silent) {
    moSel.disabled = yrSel.disabled = showB.disabled = refB.disabled = true;

    if(!silent){
      moSel.innerHTML = "<option value='ALL'>Loading…</option>";
      yrSel.innerHTML = "<option value='ALL'>Loading…</option>";
      lfName.textContent = agent;
      panel(sLF);
    } else {
      refB.classList.add("rpt-spin");
    }

    callAPI("/api/report-data?agent="+encodeURIComponent(agent)+"&month=ALL&year=ALL", function(err,data){
      refB.classList.remove("rpt-spin");

      if(err){
        moSel.disabled = yrSel.disabled = true;
        moSel.innerHTML = "<option value='ALL'>All Months</option>";
        yrSel.innerHTML = "<option value='ALL'>All Years</option>";
        errMsg.textContent = err.message;
        panel(sErr);
        return;
      }

      moSel.disabled = yrSel.disabled = false;

      var hdrs = (data.headers||[]).map(function(h){ return String(h).trim(); });
      var mi = -1, yi = -1;
      hdrs.forEach(function(h,i){
        if (h.toUpperCase() === "MONTH") mi = i;
        if (h.toUpperCase() === "YEAR")  yi = i;
      });

      var ms = {}, ys = {};
      (data.rows||[]).forEach(function(r){
        if (mi >= 0) { var m = String(r[mi]||"").trim().toUpperCase(); if(m) ms[m]=1; }
        if (yi >= 0) { var y = String(r[yi]||"").trim();               if(y) ys[y]=1; }
      });

      var months = Object.keys(ms).sort(function(a,b){
        return (MO.indexOf(a)<0?99:MO.indexOf(a)) - (MO.indexOf(b)<0?99:MO.indexOf(b));
      });
      var years = Object.keys(ys).sort(function(a,b){ return Number(b)-Number(a); });

      moSel.innerHTML = "<option value='ALL'>All Months</option>";
      months.forEach(function(m){
        var o = document.createElement("option");
        o.value = o.textContent = m;
        moSel.appendChild(o);
      });

      yrSel.innerHTML = "<option value='ALL'>All Years</option>";
      years.forEach(function(y){
        var o = document.createElement("option");
        o.value = o.textContent = y;
        yrSel.appendChild(o);
      });

      showB.disabled = refB.disabled = false;
      panel(sIdle);
    });
  }

  function fetchReport(){
    var agent = agSel.value, month = moSel.value, year = yrSel.value;
    if(!agent) return;

    curAgent = agent;
    ldDesc.textContent = agent+" · "+(month!=="ALL"?month:"All Months")+" · "+(year!=="ALL"?year:"All Years");
    panel(sLD);

    callAPI(
      "/api/report-data?agent="+encodeURIComponent(agent)+"&month="+encodeURIComponent(month)+"&year="+encodeURIComponent(year),
      function(err,data){
        if(err){ errMsg.textContent = err.message; panel(sErr); return; }
        curH = data.headers || [];
        curR = data.rows || [];
        renderTable();
        panel(sTable);
        cnt.textContent = curR.length.toLocaleString();
      }
    );
  }

  function renderTable(){
    thead.innerHTML = "<tr>"+curH.map(function(h){ return "<th>"+esc(h)+"</th>"; }).join("")+"</tr>";

    if(!curR.length){
      tbody.innerHTML =
        '<tr><td colspan="'+curH.length+'" style="text-align:center;padding:48px;color:rgba(240,237,232,.22);font-family:\'JetBrains Mono\',monospace;font-size:11px">No data for selected filters.</td></tr>';
      cnt.textContent = "0";
      return;
    }

    tbody.innerHTML = curR.map(function(row){
      return "<tr>"+curH.map(function(h,i){
        var raw = row[i];
        var val = (raw===undefined||raw===null||String(raw).trim()==="") ? "-" : String(raw);
        var hl  = h.trim().toLowerCase();
        var cls = (hl==="month"||hl==="year") ? "b" : (val==="-" ? "e" : (isN(val) ? "n" : ""));
        return '<td class="'+cls+'">'+esc(val)+'</td>';
      }).join("")+"</tr>";
    }).join("");
  }

  function doXLSX(){
    if(!curH.length) return;

    function run(){
      var ws = window.XLSX.utils.aoa_to_sheet([curH].concat(curR));
      ws["!cols"] = curH.map(function(h,i){
        var mx = Math.max(
          h.length,
          Math.max.apply(null, curR.slice(0,300).map(function(r){ return String(r[i]||"").length; }))
        );
        return { wch: Math.min(52, Math.max(10, mx+2)) };
      });

      var wb = window.XLSX.utils.book_new();
      var mv = moSel.value!=="ALL" ? "_"+moSel.value : "";
      var yv = yrSel.value!=="ALL" ? "_"+yrSel.value : "";

      window.XLSX.utils.book_append_sheet(wb, ws, curAgent.slice(0,31));
      window.XLSX.writeFile(wb, curAgent.replace(/\s+/g,"_")+mv+yv+".xlsx");
      window.toast && window.toast("✓ Downloaded");
    }

    if(window.XLSX){ run(); return; }
    var s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    s.onload = run;
    document.head.appendChild(s);
  }

  // events
  agSel.addEventListener("change", function(){
    var a = agSel.value;
    if(!a){
      moSel.disabled = yrSel.disabled = showB.disabled = refB.disabled = true;
      moSel.innerHTML = "<option value='ALL'>All Months</option>";
      yrSel.innerHTML = "<option value='ALL'>All Years</option>";
      panel(sIdle);
      return;
    }
    loadFilters(a, false);
  });

  refB.addEventListener("click", function(){
    if(agSel.value) loadFilters(agSel.value, true);
  });

  showB.addEventListener("click", fetchReport);
  dlB.addEventListener("click", doXLSX);

  retryB.addEventListener("click", function(){
    if(agSel.value && !showB.disabled) fetchReport();
    else if(agSel.value) loadFilters(agSel.value, false);
    else panel(sIdle);
  });

  panel(sIdle);
}