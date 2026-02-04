/* global XLSX */
const DEFAULT_SHEET_URL =
  "https://docs.google.com/spreadsheets/d/1wpp99NP1l83r_hDi5klWy3NKmHn9n-mX/edit?usp=drive_link&ouid=109090514128015175660&rtpof=true&sd=true";

const STORAGE_KEY = "amigos_nc_tracker_v2";

const els = {
  sheetUrl: document.getElementById("sheetUrl"),
  btnLoadSheet: document.getElementById("btnLoadSheet"),
  fileInput: document.getElementById("fileInput"),

  btnExportCSV: document.getElementById("btnExportCSV"),
  btnExportJSON: document.getElementById("btnExportJSON"),
  btnReset: document.getElementById("btnReset"),

  toggleOnlyNC: document.getElementById("toggleOnlyNC"),
  toggleRequireEvidence: document.getElementById("toggleRequireEvidence"),

  reportMeta: document.getElementById("reportMeta"),

  kpiTotal: document.getElementById("kpiTotal"),
  kpiNC: document.getElementById("kpiNC"),
  kpiClosed: document.getElementById("kpiClosed"),
  kpiOpen: document.getElementById("kpiOpen"),

  ringPct: document.getElementById("ringPct"),
  barFill: document.getElementById("barFill"),
  progressText: document.getElementById("progressText"),

  searchInput: document.getElementById("searchInput"),
  filterSeverity: document.getElementById("filterSeverity"),
  filterState: document.getElementById("filterState"),

  dashboardList: document.getElementById("dashboardList"),
  ncList: document.getElementById("ncList"),

  navItems: Array.from(document.querySelectorAll(".nav-item")),
  views: Array.from(document.querySelectorAll(".view")),
};

const state = {
  reportName: "Compliance Report AMIGOS",
  loadedAt: null,
  source: null,

  // parsed rows
  items: [],

  // local progress per finding id
  progress: loadProgress(),
};

init();

function init() {
  els.sheetUrl.value = DEFAULT_SHEET_URL;

  els.navItems.forEach(btn => {
    btn.addEventListener("click", () => {
      els.navItems.forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      const viewId = btn.dataset.view;
      els.views.forEach(v => v.classList.toggle("hidden", v.id !== viewId));
    });
  });

  els.btnLoadSheet.addEventListener("click", async () => {
    const url = els.sheetUrl.value.trim();
    if (!url) return;
    setEmpty("Loading report from Google Sheets…");
    try {
      await loadFromGoogleSheetUrl(url);
    } catch (e) {
      console.error(e);
      setEmpty("Failed to load from Google Sheets. Upload the XLSX/CSV instead (sheet may be private/blocked).");
      els.reportMeta.textContent = "Load failed from Google Sheets (likely private/blocked). Upload XLSX/CSV instead.";
    }
  });

  els.fileInput.addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setEmpty(`Loading file: ${file.name} …`);
    try {
      await loadFromLocalFile(file);
    } catch (err) {
      console.error(err);
      setEmpty("Could not parse the file. Please upload a valid XLSX/CSV.");
    } finally {
      els.fileInput.value = "";
    }
  });

  els.toggleOnlyNC.addEventListener("change", () => render());
  els.toggleRequireEvidence.addEventListener("change", () => render());

  els.searchInput.addEventListener("input", () => render());
  els.filterSeverity.addEventListener("change", () => render());
  els.filterState.addEventListener("change", () => render());

  els.btnExportCSV.addEventListener("click", exportCSV);
  els.btnExportJSON.addEventListener("click", exportJSON);

  els.btnReset.addEventListener("click", () => {
    if (!confirm("Reset saved checklist state on this browser?")) return;
    localStorage.removeItem(STORAGE_KEY);
    state.progress = {};
    render();
  });

  render();
}

/**
 * Item model (normalized)
 * id, category, finding, recommendation, severity, isNC,
 * reqMorocco, reqInditex, rawStatus
 */

function setEmpty(message) {
  const html = `<div class="empty">${escapeHtml(message)}</div>`;
  els.dashboardList.innerHTML = html;
  els.ncList.innerHTML = html;
}

async function loadFromGoogleSheetUrl(sheetUrl) {
  const sheetId = extractSheetId(sheetUrl);
  if (!sheetId) throw new Error("Could not extract sheetId from URL.");

  const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
  const res = await fetch(exportUrl);
  if (!res.ok) throw new Error(`Fetch failed: ${res.status}`);
  const ab = await res.arrayBuffer();

  const { items, meta } = parseXlsx(ab);
  state.items = items;
  state.loadedAt = new Date();
  state.source = meta;
  els.reportMeta.textContent = `${state.reportName} • Loaded • ${state.items.length} rows`;
  render();
}

async function loadFromLocalFile(file) {
  const ext = file.name.toLowerCase().split(".").pop();
  if (ext === "csv") {
    const text = await file.text();
    const rows = parseCsv(text);
    const items = mapRowsToItems(rows);
    state.items = items;
    state.loadedAt = new Date();
    state.source = { type: "csv", name: file.name };
    els.reportMeta.textContent = `${state.reportName} • Loaded CSV • ${items.length} rows`;
    render();
    return;
  }

  const ab = await file.arrayBuffer();
  const { items, meta } = parseXlsx(ab);
  state.items = items;
  state.loadedAt = new Date();
  state.source = { ...meta, name: file.name };
  els.reportMeta.textContent = `${state.reportName} • Loaded XLSX • ${items.length} rows`;
  render();
}

function parseXlsx(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const name = wb.SheetNames[0];
  const ws = wb.Sheets[name];
  const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
  const items = mapRowsToItems(json);
  return { items, meta: { type: "xlsx", sheet: name } };
}

/**
 * Key goal: show checklist for NON‑CONFORMITIES only.
 * We detect NC using status/conformity columns if present, else keyword heuristics.
 */
function mapRowsToItems(rows) {
  if (!rows || !rows.length) return [];

  // array-of-arrays -> header-based objects
  if (Array.isArray(rows[0])) {
    const header = rows[0].map(h => String(h ?? ""));
    rows = rows.slice(1).map(r => {
      const obj = {};
      header.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });
  }

  const items = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || {};
    const keys = Object.keys(r);

    const anyVal = keys.some(k => String(r[k] ?? "").trim() !== "");
    if (!anyVal) continue;

    const get = (candidates) => {
      const k = findKey(keys, candidates);
      return k ? String(r[k] ?? "").trim() : "";
    };

    const idRaw = get(["id","ref","#","no","n","index","finding id","findingid"]);
    const category = get(["category","section","domain","area","topic"]);
    const finding = get(["finding","non conformity","nonconformity","issue","observation","gap","problem"]);
    const recommendation = get(["recommendation","corrective action","action","remediation","proposed action"]);
    const severity = normalizeSeverity(get(["severity","risk","priority","criticality","rating"]));

    const rawStatus = get(["status","conformity","compliance result","result","nc status","state"]);
    const reqMorocco = get(["morocco law","legal reference","law reference","moroccan law","maroc law","code du travail"]);
    const reqInditex = get(["inditex","inditex reference","inditex requirement","code of conduct","ics","social audit"]);

    const id = idRaw || `F-${String(i + 1).padStart(3,"0")}`;

    const isNC = detectNonConformity({ rawStatus, finding, recommendation });

    items.push({
      id,
      category: category || "General",
      finding: finding || summarizeFallback(r),
      recommendation,
      severity,
      rawStatus,
      reqMorocco,
      reqInditex,
      isNC,
    });

    // Ensure progress state exists for NC
    if (!state.progress[id]) {
      state.progress[id] = defaultProgressForItem();
    }
  }

  saveProgress(state.progress);
  return items;
}

function defaultProgressForItem() {
  return {
    checklist: {
      containment: false,
      rootCause: false,
      correctiveAction: false,
      preventiveAction: false,
      evidence: false,
      verification: false,
      managementSignoff: false
    },
    evidenceLink: "",
    evidenceNote: "",
    owner: "",
    dueDate: "",
    comment: "",
    updatedAtISO: new Date().toISOString(),
  };
}

function detectNonConformity({ rawStatus, finding, recommendation }) {
  const s = String(rawStatus ?? "").toLowerCase();
  const f = `${finding ?? ""} ${recommendation ?? ""}`.toLowerCase();

  // Strong signals from status-like fields
  const ncWords = ["non conform", "non-conform", "nonconform", "nc", "nok", "no", "fail", "not compliant", "non compliant"];
  const okWords = ["conform", "ok", "yes", "pass", "compliant", "closed"];

  if (s) {
    if (ncWords.some(w => s === w || s.includes(w))) return true;
    if (okWords.some(w => s === w || s.includes(w))) return false;
  }

  // Heuristic fallback (weak): if text clearly says non-compliant
  if (["non conform", "non-compliant", "not compliant", "non compliant", "violation"].some(w => f.includes(w))) {
    return true;
  }

  // Default: treat as NC? No. Default to false to avoid false positives.
  return false;
}

function render() {
  renderKPIs();
  renderLists();
}

function getVisibleItems() {
  const q = els.searchInput.value.trim().toLowerCase();
  const sev = els.filterSeverity.value;
  const st = els.filterState.value;
  const onlyNC = els.toggleOnlyNC.checked;

  let items = state.items.slice();

  if (onlyNC) items = items.filter(x => x.isNC);

  if (q) {
    items = items.filter(x => {
      const hay = `${x.id} ${x.category} ${x.finding} ${x.recommendation} ${x.reqMorocco} ${x.reqInditex}`.toLowerCase();
      return hay.includes(q);
    });
  }

  if (sev !== "all") items = items.filter(x => (x.severity || "info") === sev);

  if (st !== "all") {
    items = items.filter(x => getClosureState(x) === st);
  }

  return items;
}

function renderKPIs() {
  const total = state.items.length;
  const nc = state.items.filter(x => x.isNC).length;

  const ncItems = state.items.filter(x => x.isNC);
  const closed = ncItems.filter(x => getClosureState(x) === "closed").length;
  const open = ncItems.length - closed;

  els.kpiTotal.textContent = String(total || 0);
  els.kpiNC.textContent = String(nc || 0);
  els.kpiClosed.textContent = String(closed || 0);
  els.kpiOpen.textContent = String(open || 0);

  const pct = ncItems.length ? Math.round((closed / ncItems.length) * 100) : 0;
  els.ringPct.textContent = ncItems.length ? `${pct}%` : "—";
  document.querySelector(".ring").style.setProperty("--ring", `${pct}%`);

  els.barFill.style.width = `${pct}%`;
  els.progressText.textContent = ncItems.length
    ? `${closed} closed out of ${ncItems.length} Non‑Conformities (${pct}%)`
    : "No Non‑Conformities detected yet (check your sheet columns / status values).";
}

function renderLists() {
  if (!state.items.length) {
    setEmpty("Load a report to see Non‑Conformities.");
    return;
  }

  const visible = getVisibleItems();
  if (!visible.length) {
    els.dashboardList.innerHTML = `<div class="empty">No items match filters.</div>`;
    els.ncList.innerHTML = `<div class="empty">No items match filters.</div>`;
    return;
  }

  // Dashboard shows a compact list (still cards)
  els.dashboardList.innerHTML = visible.slice(0, 12).map(renderCard).join("");
  // NC view shows full list
  els.ncList.innerHTML = visible.map(renderCard).join("");

  bindCardEvents();
}

function renderCard(item) {
  const p = state.progress[item.id] || defaultProgressForItem();
  const stateKey = getClosureState(item);
  const resultLabel = stateKey === "closed" ? "Conform (Closed)" : "Non‑Conform (Open)";
  const sev = (item.severity || "info");

  const checklist = p.checklist;

  return `
    <article class="nc-card" data-id="${escapeAttr(item.id)}">
      <div class="nc-card-head">
        <div>
          <div class="nc-title">${escapeHtml(item.id)} — ${escapeHtml(item.category)}</div>
          <div class="nc-sub">${escapeHtml(item.finding)}</div>
        </div>

        <div class="pills">
          <span class="pill sev-${escapeAttr(sev)}">Severity: ${escapeHtml(capitalize(sev))}</span>
          <span class="pill state-${stateKey === "closed" ? "closed" : "open"}">Result: ${escapeHtml(resultLabel)}</span>
          ${item.isNC ? `<span class="pill">NC</span>` : `<span class="pill">Info</span>`}
        </div>
      </div>

      <div class="nc-card-body">
        <div class="nc-grid">
          <div class="block">
            <div class="block-title">Requirement references</div>

            <div class="field">
              <label>Moroccan law reference (as used by your compliance team)</label>
              <input class="input js-moroccoRef" type="text" placeholder="e.g., Labor Code article / decree / internal mapping"
                value="${escapeAttr(item.reqMorocco || "")}">
            </div>

            <div class="field">
              <label>Inditex requirement reference</label>
              <input class="input js-inditexRef" type="text" placeholder="e.g., Inditex CoC / ICS clause / audit requirement"
                value="${escapeAttr(item.reqInditex || "")}">
            </div>

            <div class="field">
              <label>Recommendation / corrective action requested</label>
              <textarea class="input js-reco" placeholder="Action requested / recommendation">${escapeHtml(item.recommendation || "")}</textarea>
              <div class="hint">You can edit this text locally to match your CAPA wording.</div>
            </div>
          </div>

          <div class="block">
            <div class="block-title">Checklist (closure conditions)</div>

            <div class="checklist">
              ${checkRow("containment", "Immediate containment applied (stop risk / temporary measure)", checklist.containment)}
              ${checkRow("rootCause", "Root cause analysis completed", checklist.rootCause)}
              ${checkRow("correctiveAction", "Corrective action implemented (fix the NC)", checklist.correctiveAction)}
              ${checkRow("preventiveAction", "Preventive action implemented (avoid recurrence)", checklist.preventiveAction)}
              ${checkRow("evidence", "Evidence attached (link or note)", checklist.evidence)}
              ${checkRow("verification", "Internal verification performed", checklist.verification)}
              ${checkRow("managementSignoff", "Management sign‑off / approval", checklist.managementSignoff)}
            </div>

            <div class="field">
              <label>Evidence link (ticket / drive / photo / document)</label>
              <input class="input js-evidenceLink" type="url" placeholder="https://..."
                value="${escapeAttr(p.evidenceLink || "")}">
            </div>

            <div class="field">
              <label>Evidence note (if no link)</label>
              <textarea class="input js-evidenceNote" placeholder="Describe the proof: what changed, where, who validated…">${escapeHtml(p.evidenceNote || "")}</textarea>
            </div>

            <div class="field">
              <label>Owner / Responsible</label>
              <input class="input js-owner" type="text" placeholder="Name / Department" value="${escapeAttr(p.owner || "")}">
            </div>

            <div class="row">
              <div class="field" style="flex:1">
                <label>Due date</label>
                <input class="input js-dueDate" type="date" value="${escapeAttr(p.dueDate || "")}">
              </div>
              <div class="field" style="flex:1">
                <label>Last update</label>
                <input class="input" type="text" value="${escapeAttr(formatDateTime(p.updatedAtISO))}" disabled>
              </div>
            </div>

            <div class="field">
              <label>Reviewer comment</label>
              <textarea class="input js-comment" placeholder="CAPA notes, auditor response, follow-up…">${escapeHtml(p.comment || "")}</textarea>
            </div>

            <div class="hint">
              Result becomes <b>Conform (Closed)</b> only when all checklist items are checked + verification is done
              ${els.toggleRequireEvidence?.checked ? " + evidence is provided." : "."}
            </div>
          </div>
        </div>
      </div>
    </article>
  `;
}

function checkRow(key, label, checked) {
  return `
    <label class="check">
      <input class="js-check" data-key="${escapeAttr(key)}" type="checkbox" ${checked ? "checked" : ""}>
      <span>${escapeHtml(label)}</span>
    </label>
  `;
}

function bindCardEvents() {
  const cards = document.querySelectorAll(".nc-card");
  cards.forEach(card => {
    const id = card.getAttribute("data-id");
    const item = state.items.find(x => x.id === id);
    if (!item) return;

    const p = state.progress[id] || defaultProgressForItem();

    // checklist toggles
    card.querySelectorAll(".js-check").forEach(cb => {
      cb.addEventListener("change", () => {
        const k = cb.dataset.key;
        p.checklist[k] = cb.checked;

        // If evidence required, auto-toggle evidence when link/note exists
        if (k === "evidence" && cb.checked === false) {
          // allow manual, do nothing
        }

        touchProgress(id, p);
        saveProgress(state.progress);
        renderKPIs();
        // update pills only (simple re-render for correctness)
        renderLists();
      });
    });

    // evidence fields
    const evidenceLink = card.querySelector(".js-evidenceLink");
    evidenceLink?.addEventListener("input", () => {
      p.evidenceLink = evidenceLink.value.trim();
      autoSetEvidenceChecklist(p);
      touchProgress(id, p);
      saveProgress(state.progress);
      renderKPIs();
    });

    const evidenceNote = card.querySelector(".js-evidenceNote");
    evidenceNote?.addEventListener("input", () => {
      p.evidenceNote = evidenceNote.value;
      autoSetEvidenceChecklist(p);
      touchProgress(id, p);
      saveProgress(state.progress);
      renderKPIs();
    });

    // other fields
    const owner = card.querySelector(".js-owner");
    owner?.addEventListener("input", () => {
      p.owner = owner.value;
      touchProgress(id, p);
      saveProgress(state.progress);
    });

    const dueDate = card.querySelector(".js-dueDate");
    dueDate?.addEventListener("change", () => {
      p.dueDate = dueDate.value;
      touchProgress(id, p);
      saveProgress(state.progress);
    });

    const comment = card.querySelector(".js-comment");
    comment?.addEventListener("input", () => {
      p.comment = comment.value;
      touchProgress(id, p);
      saveProgress(state.progress);
    });

    // Requirement refs & reco are shown and editable but not persisted into sheet data;
    // you can persist them too if you want after you send the real structure.
    const moroccoRef = card.querySelector(".js-moroccoRef");
    moroccoRef?.addEventListener("input", () => {
      item.reqMorocco = moroccoRef.value;
    });
    const inditexRef = card.querySelector(".js-inditexRef");
    inditexRef?.addEventListener("input", () => {
      item.reqInditex = inditexRef.value;
    });
    const reco = card.querySelector(".js-reco");
    reco?.addEventListener("input", () => {
      item.recommendation = reco.value;
    });
  });
}

function autoSetEvidenceChecklist(p) {
  const requireEvidence = els.toggleRequireEvidence.checked;
  if (!requireEvidence) return;
  const has = hasEvidence(p);
  p.checklist.evidence = has;
}

function hasEvidence(p) {
  const linkOk = (p.evidenceLink || "").trim().length > 8;
  const noteOk = (p.evidenceNote || "").trim().length >= 20;
  return linkOk || noteOk;
}

/**
 * Closure states:
 * - open: nothing done
 * - progress: partial checklist
 * - ready: checklist all done but verification/signoff maybe missing
 * - closed: all checklist + verification + signoff (+ evidence if required)
 */
function getClosureState(item) {
  if (!item.isNC) return "closed"; // info items are not treated as NC

  const p = state.progress[item.id] || defaultProgressForItem();
  const c = p.checklist || {};

  const requireEvidence = els.toggleRequireEvidence.checked;

  const requiredKeys = ["containment", "rootCause", "correctiveAction", "preventiveAction", "verification", "managementSignoff"];
  const baseAllDone = requiredKeys.every(k => !!c[k]);

  const evidenceOk = !requireEvidence || (c.evidence && hasEvidence(p));
  const allDone = baseAllDone && evidenceOk;

  const anyDone = Object.values(c).some(Boolean);

  if (!anyDone) return "open";
  if (allDone) return "closed";

  // If mostly done but missing verification / signoff / evidence etc.
  const missing = requiredKeys.filter(k => !c[k]);
  if (missing.length === 1 && missing[0] === "managementSignoff") return "ready";
  if (missing.length === 1 && missing[0] === "verification") return "ready";
  if (requireEvidence && !evidenceOk && baseAllDone) return "ready";

  return "progress";
}

function touchProgress(id, p) {
  p.updatedAtISO = new Date().toISOString();
  state.progress[id] = p;
}

/* ---------------- Exports ---------------- */

function exportCSV() {
  if (!state.items.length) return alert("No report loaded.");

  const headers = [
    "ID","IsNC","Category","Finding","Severity",
    "MoroccoLawRef","InditexRef","Recommendation",
    "State","Result",
    "Owner","DueDate",
    "EvidenceLink","EvidenceNote",
    "Containment","RootCause","CorrectiveAction","PreventiveAction","Evidence","Verification","ManagementSignoff",
    "UpdatedAt"
  ];

  const lines = [headers.join(",")];

  for (const it of state.items) {
    const p = state.progress[it.id] || defaultProgressForItem();
    const st = getClosureState(it);
    const result = it.isNC ? (st === "closed" ? "CONFORM" : "NON-CONFORM") : "CONFORM";

    const row = [
      it.id,
      it.isNC ? "YES" : "NO",
      it.category,
      it.finding,
      it.severity,
      it.reqMorocco || "",
      it.reqInditex || "",
      it.recommendation || "",
      st,
      result,
      p.owner || "",
      p.dueDate || "",
      p.evidenceLink || "",
      p.evidenceNote || "",
      bool(p.checklist.containment),
      bool(p.checklist.rootCause),
      bool(p.checklist.correctiveAction),
      bool(p.checklist.preventiveAction),
      bool(p.checklist.evidence),
      bool(p.checklist.verification),
      bool(p.checklist.managementSignoff),
      p.updatedAtISO || ""
    ].map(csvCell);

    lines.push(row.join(","));
  }

  downloadBlob(new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" }),
    `AMIGOS_NC_Corrections_${new Date().toISOString().slice(0,10)}.csv`
  );
}

function exportJSON() {
  if (!state.items.length) return alert("No report loaded.");
  const payload = {
    reportName: state.reportName,
    loadedAt: state.loadedAt?.toISOString() || null,
    source: state.source,
    requireEvidenceForClosure: els.toggleRequireEvidence.checked,
    items: state.items,
    progress: state.progress
  };

  downloadBlob(
    new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" }),
    `AMIGOS_NC_Corrections_${new Date().toISOString().slice(0,10)}.json`
  );
}

/* ---------------- Utilities ---------------- */

function extractSheetId(url) {
  const m = String(url).match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : null;
}

function normalizeSeverity(sev) {
  const s = String(sev ?? "").toLowerCase().trim();
  if (!s) return "info";
  if (s.includes("crit")) return "critical";
  if (s.includes("high")) return "high";
  if (s.includes("med")) return "medium";
  if (s.includes("low")) return "low";
  if (s.includes("info")) return "info";
  if (["1","p1"].includes(s)) return "critical";
  if (["2","p2"].includes(s)) return "high";
  if (["3","p3"].includes(s)) return "medium";
  if (["4","p4"].includes(s)) return "low";
  return "info";
}

function summarizeFallback(obj) {
  const pairs = Object.entries(obj)
    .map(([k,v]) => [String(k).trim(), String(v ?? "").trim()])
    .filter(([k,v]) => k && v);
  return pairs.slice(0, 3).map(([k,v]) => `${k}: ${v}`).join(" • ") || "—";
}

function findKey(keys, candidates) {
  const normalized = new Map(keys.map(k => [norm(k), k]));
  for (const c of candidates) {
    const hit = normalized.get(norm(c));
    if (hit) return hit;
  }
  // partial fallback
  for (const c of candidates) {
    const nc = norm(c);
    for (const k of keys) {
      const nk = norm(k);
      if (nk.includes(nc) || nc.includes(nk)) return k;
    }
  }
  return null;
}

function norm(s) {
  return String(s ?? "").toLowerCase().replace(/[\s\-_()]/g, "").trim();
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];

    if (ch === '"' && inQuotes && next === '"') { cell += '"'; i++; continue; }
    if (ch === '"') { inQuotes = !inQuotes; continue; }

    if (ch === "," && !inQuotes) { row.push(cell); cell = ""; continue; }
    if ((ch === "\n" || ch === "\r") && !inQuotes) {
      if (ch === "\r" && next === "\n") i++;
      row.push(cell); rows.push(row);
      row = []; cell = "";
      continue;
    }
    cell += ch;
  }

  row.push(cell); rows.push(row);
  return rows.filter(r => r.some(c => String(c).trim() !== ""));
}

function loadProgress() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function saveProgress(obj) {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(obj || {})); } catch {}
}

function csvCell(v) {
  const s = String(v ?? "");
  if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
function escapeAttr(s) { return escapeHtml(s).replaceAll("`", "&#096;"); }

function capitalize(s){ return s ? s.charAt(0).toUpperCase() + s.slice(1) : ""; }
function bool(x){ return x ? "YES" : "NO"; }

function formatDateTime(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleString();
}
