(function () {
  "use strict";

  var AGENT_NAMES = {
    PAY_STK:  "Payable Stock",
    PAY_TD:   "Payable Trade Discount",
    REC_STK:  "Receivable Stock",
    REC_TD:   "Receivable Trade Discount",
    PRI_SALE: "Primary Sale",
    GRC:      "GRC & Movement",
  };

  // Default cards for standard agents (master + format + process)
  var CARDS_DEFAULT = [
    {
      type:  "master",
      icon:  "🗂️",
      title: "Update Master Data",
      desc:  "Open the master Google Sheet used by this agent — Margin, HSN, COGS, deductions and other reference data.",
      label: "Open Master Sheet",
      key:   "master",
    },
    {
      type:  "format",
      icon:  "📋",
      title: "Download Input Format",
      desc:  "Download the blank Excel / CSV template with the correct headers and column structure for uploading to this agent.",
      label: "Download Template",
      key:   "format",
    },
    {
      type:  "process",
      icon:  "📖",
      title: "Process Document",
      desc:  "Step-by-step SOP for preparing input files, running this agent, verifying outputs, and resolving common errors.",
      label: "Open Process Doc",
      key:   "process",
    },
  ];

  // GRC + Movement: 3 cards — Input Format + GRC Process Doc + Movement Logic Doc
  var CARDS_GRC_MOVEMENT = [
    {
      type:  "format",
      icon:  "📋",
      title: "Download Input Format",
      desc:  "Download the blank Excel template with the correct headers and column structure for the GRC Purchase Register upload.",
      label: "Download Template",
      key:   "format",
    },
    {
      type:  "process",
      icon:  "📖",
      title: "GRC Agent — Process Document",
      desc:  "Step-by-step SOP for preparing the GRC Purchase Register file, running the agent, verifying outputs and resolving common errors.",
      label: "Open GRC Process Doc",
      key:   "process",
    },
    {
      type:  "process2",
      icon:  "📈",
      title: "Movement Matrix — Logic Document",
      desc:  "Detailed logic document covering all rows of the Secondary Movement Matrix — data sources, formulas, and calculation steps for Basic Value and Qty.",
      label: "Open Movement Doc",
      key:   "process2",
    },
  ];

  var CARD_CONFIG = {
    PAY_STK:  CARDS_DEFAULT,
    PAY_TD:   CARDS_DEFAULT,
    REC_STK:  CARDS_DEFAULT,
    REC_TD:   CARDS_DEFAULT,
    PRI_SALE: CARDS_DEFAULT,
    GRC:      CARDS_GRC_MOVEMENT,
  };

  var SVG_ARROW =
    '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" ' +
    'stroke-linecap="round" stroke-linejoin="round">' +
    '<line x1="7" y1="17" x2="17" y2="7"/><polyline points="7 7 17 7 17 17"/></svg>';

  var _links = {};
  var _activeAgent = "PAY_STK";

  function makeCard(def, url, idx) {
    var hasUrl = typeof url === "string" && url.trim() !== "";
    var el = document.createElement(hasUrl ? "a" : "div");
    el.className = "hrc-card hrc-" + def.type + " " + (hasUrl ? "hrc-link" : "hrc-nolink");
    if (hasUrl) {
      el.href = url.trim();
      el.target = "_blank";
      el.rel = "noopener";
    }
    el.innerHTML =
      (hasUrl ? "" : '<span class="hrc-not-set-badge">Not configured</span>') +
      '<div class="hrc-icon-wrap">' + def.icon + "</div>" +
      '<div><div class="hrc-title">' + def.title + '</div><div class="hrc-desc">' + def.desc + "</div></div>" +
      '<div class="hrc-footer">' +
        '<span class="hrc-action-label">' + (hasUrl ? def.label : "Not configured") + "</span>" +
        '<span class="hrc-arrow-btn">' + SVG_ARROW + "</span>" +
      "</div>";
    setTimeout(function () {
      el.classList.add("hrc-in");
    }, 30 + idx * 80);
    return el;
  }

  function renderCards(agentKey) {
    var grid = document.getElementById("helpResourceCards");
    var label = document.getElementById("helpResourcesLabel");
    if (!grid) return;
    if (label) label.textContent = "Resources — " + (AGENT_NAMES[agentKey] || agentKey);
    grid.innerHTML = "";

    var cards = CARD_CONFIG[agentKey] || CARDS_DEFAULT;
    var agentLinks = _links && _links[agentKey] ? _links[agentKey] : {};

    cards.forEach(function (def, i) {
      grid.appendChild(makeCard(def, agentLinks[def.key] || "", i));
    });
  }

  function selectAgent(key) {
    _activeAgent = key;
    document.querySelectorAll("#helpAgentSelector .has-btn").forEach(function (btn) {
      btn.classList.toggle("active", btn.dataset.key === key);
    });
    renderCards(key);
  }

  function attachButtons() {
    document.querySelectorAll("#helpAgentSelector .has-btn").forEach(function (btn) {
      btn.addEventListener("click", function () {
        selectAgent(btn.dataset.key);
      });
    });
  }

  function initHelp() {
    attachButtons();
    selectAgent("PAY_STK");

    fetch("/api/help-links", { credentials: "include" })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (data && data.ok && data.links) {
          _links = data.links;
          renderCards(_activeAgent);
        }
      })
      .catch(function (e) {
        console.warn("[help] fetch failed:", e.message);
      });
  }

  window.__sectionInit = window.__sectionInit || {};
  window.__sectionInit.help = initHelp;

  if (document.getElementById("helpAgentSelector")) {
    initHelp();
  }

})();