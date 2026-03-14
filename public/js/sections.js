// =========================
// FILE: public/js/sections.js
// Inc.5 SOR Portal — Section Loader
// Supports: console, report, analytics, help, basedump
// =========================

const mainEl = document.getElementById("main");
const VALID_SECTIONS = ["console", "report", "analytics", "help", "basedump"];

function setActive(page) {
  document.querySelectorAll(".sb-item").forEach((b) => {
    b.classList.toggle("active", b.dataset.page === page);
  });
}

// innerHTML does NOT execute <script> tags — this re-executes them safely
function runScriptsIn(container) {
  try {
    Array.from(container.querySelectorAll("script")).forEach((oldScript) => {
      try {
        const newScript = document.createElement("script");
        Array.from(oldScript.attributes).forEach((attr) => {
          newScript.setAttribute(attr.name, attr.value);
        });
        newScript.textContent = oldScript.textContent;
        oldScript.parentNode.replaceChild(newScript, oldScript);
      } catch (scriptErr) {
        console.warn("[sections] Script execution error:", scriptErr);
      }
    });
  } catch (e) {
    console.warn("[sections] runScriptsIn error:", e);
  }
}

async function loadSection(name, { pushHash = true } = {}) {
  setActive(name);
  if (pushHash) history.replaceState(null, "", `#${name}`);

  mainEl.innerHTML = `
    <div class="page-loader" id="pageLoader">
      <div class="loader-ring"></div>
      <div class="loader-text">Loading ${name}...</div>
    </div>
  `;

  try {
    const res = await fetch(`/sections/${name}.html`, {
      cache: "no-store",
      credentials: "include",
    });

    if (!res.ok) throw new Error(`Section not found: ${name} (HTTP ${res.status})`);

    mainEl.innerHTML = await res.text();

    // Execute <script> tags injected via innerHTML (browser doesn't run them automatically)
    runScriptsIn(mainEl);

    // Call the section's registered init function
    const fn = window.__sectionInit?.[name];
    if (typeof fn === "function") fn();

  } catch (e) {
    mainEl.innerHTML = `
      <div class="page-wrap">
        <div class="card" style="margin-top:24px">
          <div class="card-body">
            <div class="empty-state">
              <div class="empty-icon">⚠</div>
              <div class="empty-title">Failed to load section</div>
              <div class="empty-sub">${String(e.message || e)}</div>
            </div>
          </div>
        </div>
      </div>
    `;
  }
}

function initSidebar() {
  document.querySelectorAll(".sb-item").forEach((b) => {
    b.addEventListener("click", () => loadSection(b.dataset.page));
  });
}

function getDefaultSection() {
  const h = (location.hash || "").replace("#", "").trim();
  if (VALID_SECTIONS.includes(h)) return h;
  return "console";
}

document.addEventListener("DOMContentLoaded", () => {
  initSidebar();
  loadSection(getDefaultSection(), { pushHash: false });

  window.addEventListener("hashchange", () => {
    loadSection(getDefaultSection(), { pushHash: false });
  });
});