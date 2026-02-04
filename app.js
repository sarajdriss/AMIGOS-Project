/* global XLSX */

const DEFAULT_SHEET_URL =
  "https://docs.google.com/spreadsheets/d/1wpp99NP1l83r_hDi5klWy3NKmHn9n-mX/edit?usp=drive_link&ouid=109090514128015175660&rtpof=true&sd=true";

const STORAGE_KEY = "amigos_compliance_corrections_v1";

const els = {
  sheetUrl: document.getElementById("sheetUrl"),
  btnLoad: document.getElementById("btnLoad"),
  fileInput: document.getElementById("fileInput"),
  btnExport: document.getElementById("btnExport"),
  btnReset: document.getElementById("btnReset"),

  kpiTotal: document.getElementById("kpiTotal"),
  kpiCorrected: document.getElementById("kpiCorrected"),
  kpiRemaining: document.getElementById("kpiRemaining"),
  kpiProgress: document.getElementById("kpiProgress"),

  reportMeta: document.getElementById("reportMeta"),

  searchInput: document.getElementById("searchInput"),
  filterStatus: document.getElementById("filterStatus"),
  filterSeverity: document.getElementById("filterSeverity"),

  progressFill: document.getElementById("progressFill"),
  progressText: document.getElementById("progressText"),

  tbody: document.getElementById("tbody"),
  resultsBox: document.getElementById("resultsBox"),
};

const state = {
  reportName: "Compliance Report AMIGOS",
  loadedAt: null,
  source: null,
  findings: /** @type {Finding[]} */ ([]),
  corrections: loadCorrections(), // { [id]: { corrected:boolean, note:string, dateISO:string } }
};

init();

function init() {
  els.sheetUrl.value = DEFAULT_SHEET_URL;

  els.btnLoad.addEventListener("click", async () => {
    const url = els.sheetUrl.value.trim();
    if (!url) return;

    setBusyTable("Loading report from Google Sheets…");
    try {
      await loadFromGoogleSheetUrl(url);
    } catch (err) {
      console.error(err);
      setBusyTable("Could not load from Google Sheets. Try uploading the exported XLSX/CSV file instead.");
      els.reportMeta.textContent =
        "Load failed from Google Sheets (sheet may be private / blocked). Upload XLSX/CSV instead.";
    }
  });

  els.fileInput.addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setBusyTable(`Loading file: ${file.name} …`);
    try {
      await loadFromLocalFile(file);
    } catch (err) {
      console.error(err);
      setBusyTable("Could not parse file. Please upload a valid XLSX or CSV.");
    } finally {
      els.fileInput.value = "";
    }
  });

  els.searchInput.addEventListener("input", () => render());
  els.filterStatus.addEventListener("change", () => render());
  els.filterSeverity.addEventListener("change", () => render());

  els.btnExport.addEventListener("click", () => exportResultsCSV());
  els.btnReset.addEventListener("click", () => {
    if (!confirm("Reset saved correction state on this device?")) return;
    localStorage.removeItem(STORAGE_KEY);
    state.corrections = {};
    render();
  });

  render();
}

/** Types (for clarity)
 * @typedef {Object} Finding
 * @property {string} id
 * @property {string} category
 * @property {string} requirement
 * @property {string} finding
 * @property {string} severity
 * @property {string} recommendation
 * @property {string} owner
 * @property {string} dueDate
 * @property {string} rawStatus
 */

function setBusyTable(message) {
  els.tbody.innerHTML = `<tr><td colspan="6" class="empty">${escapeHtml(message)}</td></tr>`;
  els.resultsBox.innerHTML = `<div class="empty">${escapeHtml(message)}</div>`;
}

async function loadFromGoogleSheetUrl(sheetUrl) {
  const sheetId = extractSheetId(sheetUrl);
  if (!sheetId) throw new Error("Could not extract sheetId from URL.");

  const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
  const res = await fetch(exportUrl);
  if (!res.ok) throw new Error(`Fetch failed: ${res.status}`);
  const ab = await res.arrayBuffer();

  const { findings, meta } = parseXlsxArrayBuffer(ab);
  state.findings = findings;
  state.loadedAt = new Date();
  state.source = meta;

  els.reportMeta.textContent = `${state.reportName} • Loaded from Google Sheets • ${state.findings.length} findings`;
  render();
}

async function loadFromLocalFile(file) {
  const ext = file.name.toLowerCase().split(".").pop();

  if (ext === "csv") {
    const text = await file.text();
    const rows = parseCsv(text);
    const findings = mapRowsToFindings(rows);
    state.findings = findings;
    state.loadedAt = new Date();
    state.source = { type: "csv", name: file.name };
    els.reportMeta.textContent = `${state.reportName} • Loaded from CSV • ${findings.length} findings`;
    render();
    return;
  }

  const ab = await file.arrayBuffer();
  const { findings, meta } = parseXlsxArrayBuffer(ab);
  state.findings = findings;
  state.loadedAt = new Date();
  state.source = { ...meta, name: file.name };
  els.reportMeta.textContent = `${state.reportName} • Loaded from XLSX • ${findings.length} findings`;
  render();
}

function parseXlsxArrayBuffer(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });

  const firstSheetName = wb.SheetNames[0];
  const ws = wb.Sheets[firstSheetName];

  // Read as JSON objects using the first row as headers.
  const json = XLSX.utils.sheet_to_json(ws, { defval: "" });

  // If the sheet is "flat table", json is already usable.
  // If it isn't, user can export a clean table; still we try.
  const findings = mapRowsToFindings(json);

  return { findings, meta: { type: "xlsx", sheet: firstSheetName } };
}

/**
 * Attempts to map arbitrary column names into a standard Finding object.
 * @param {Array<Object>|Array<Array<string>>} rows
 * @returns {Finding[]}
 */
function mapRowsToFindings(rows) {
  if (!rows || rows.length === 0) return [];

  // If rows are array-of-arrays, treat first row as header.
  if (Array.isArray(rows[0])) {
    const header = rows[0].map((h) => String(h ?? ""));
    const out = [];
    for (let i = 1; i < rows.length; i++) {
      const obj = {};
      for (let c = 0; c < header.length; c++) obj[header[c]] = rows[i][c];
      out.push(obj);
    }
    rows = out;
  }

  const mapped = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || {};
    const keys = Object.keys(r);

    // Skip empty lines
    const hasAnyValue = keys.some((k) => String(r[k] ?? "").trim() !== "");
    if (!hasAnyValue) continue;

    const get = (candidates) => {
      const k = findKey(keys, candidates);
      return k ? String(r[k] ?? "").trim() : "";
    };

    const idRaw = get(["id", "no", "n", "#", "index", "findingid", "ref", "reference"]);
    const category = get(["category", "section", "domain", "control", "area"]);
    const requirement = get(["requirement", "clause", "controlrequirement", "criteria", "standard"]);
    const findingText = get(["finding", "issue", "noncompliance", "observation", "gap"]);
    const severity = normalizeSeverity(get(["severity", "risk", "priority", "criticality"]));
    const recommendation = get(["recommendation", "remediation", "action", "correctiveaction", "mitigation"]);
    const owner = get(["owner", "responsible", "assignee"]);
    const dueDate = get(["duedate", "targetdate", "deadline"]);
    const rawStatus = get(["status", "state", "corrected", "closed"]);

    const id = idRaw || `F-${String(i + 1).padStart(3, "0")}`;

    mapped.push({
      id,
      category: category || "General",
      requirement,
      finding: findingText || summarizeRowFallback(r),
      severity,
      recommendation,
      owner,
      dueDate,
      rawStatus,
    });
  }

  // Initialize corrections from any "Status" column if local state missing
  for (const f of mapped) {
    if (state.corrections[f.id]) continue;
    const inferred = statusLooksCorrected(f.rawStatus);
    if (inferred) {
      state.corrections[f.id] = {
        corrected: true,
        note: "Marked corrected based on report status.",
        dateISO: new Date().toISOString().slice(0, 10),
      };
    }
  }

  saveCorrections(state.corrections);
  return mapped;
}

function summarizeRowFallback(rowObj) {
  // Fallback: concatenate the first few meaningful fields
  const pairs = Object.entries(rowObj)
    .map(([k, v]) => [String(k).trim(), String(v ?? "").trim()])
    .filter(([k, v]) => k && v);
  return pairs.slice(0, 3).map(([k, v]) => `${k}: ${v}`).join(" • ") || "Finding (no details provided)";
}

function findKey(keys, candidates) {
  const normalized = new Map(keys.map((k) => [normalizeKey(k), k]));
  for (const c of candidates) {
    const hit = normalized.get(normalizeKey(c));
    if (hit) return hit;
  }
  // Partial match fallback
  for (const c of candidates) {
    const nc = normalizeKey(c);
    for (const k of keys) {
      const nk = normalizeKey(k);
      if (nk.includes(nc) || nc.includes(nk)) return k;
    }
  }
  return null;
}

function normalizeKey(s) {
  return String(s ?? "")
    .toLowerCase()
    .replace(/[\s\-_()]/g, "")
    .trim();
}

function normalizeSeverity(sev) {
  const s = String(sev ?? "").toLowerCase().trim();
  if (!s) return "info";
  if (s.includes("crit")) return "critical";
  if (s.includes("high")) return "high";
  if (s.includes("med")) return "medium";
  if (s.includes("low")) return "low";
  if (s.includes("info")) return "info";
  // if numeric:
  if (["1", "p1"].includes(s)) return "critical";
  if (["2", "p2"].includes(s)) return "high";
  if (["3", "p3"].includes(s)) return "medium";
  if (["4", "p4"].includes(s)) return "low";
  return "info";
}

function statusLooksCorrected(status) {
  const s = String(status ?? "").toLowerCase().trim();
  if (!s) return false;
  return ["closed", "done", "corrected", "fixed", "complete", "completed", "yes", "true", "1"].some((w) =>
    s === w || s.includes(w)
  );
}

function extractSheetId(url) {
  // Matches: https://docs.google.com/spreadsheets/d/<ID>/edit
  const m = String(url).match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : null;
}

function render() {
  renderKpis();
  renderTable();
  renderResults();
}

function getFilteredFindings() {
  const q = els.searchInput.value.trim().toLowerCase();
  const statusFilter = els.filterStatus.value;
  const sevFilter = els.filterSeverity.value;

  return state.findings.filter((f) => {
    const corr = !!state.corrections[f.id]?.corrected;

    if (statusFilter === "corrected" && !corr) return false;
    if (statusFilter === "open" && corr) return false;

    if (sevFilter !== "all" && (f.severity || "info") !== sevFilter) return false;

    if (q) {
      const hay = `${f.id} ${f.category} ${f.requirement} ${f.finding} ${f.recommendation} ${f.owner}`.toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

function renderKpis() {
  const total = state.findings.length;
  const corrected = state.findings.reduce((acc, f) => acc + (state.corrections[f.id]?.corrected ? 1 : 0), 0);
  const remaining = Math.max(0, total - corrected);
  const pct = total ? Math.round((corrected / total) * 100) : 0;

  els.kpiTotal.textContent = String(total);
  els.kpiCorrected.textContent = String(corrected);
  els.kpiRemaining.textContent = String(remaining);
  els.kpiProgress.textContent = `${pct}%`;

  els.progressFill.style.width = `${pct}%`;
  els.progressText.textContent = total
    ? `${corrected} corrected out of ${total} findings (${pct}%)`
    : "—";
}

function renderTable() {
  if (!state.findings.length) {
    els.tbody.innerHTML = `<tr><td colspan="6" class="empty">Load a report to see findings.</td></tr>`;
    return;
  }

  const filtered = getFilteredFindings();
  if (!filtered.length) {
    els.tbody.innerHTML = `<tr><td colspan="6" class="empty">No findings match your filters.</td></tr>`;
    return;
  }

  els.tbody.innerHTML = filtered.map((f) => rowHtml(f)).join("");

  // Bind events for checkboxes and notes
  for (const f of filtered) {
    const cb = document.getElementById(`cb-${cssId(f.id)}`);
    const note = document.getElementById(`note-${cssId(f.id)}`);
    const date = document.getElementById(`date-${cssId(f.id)}`);

    if (!cb || !note || !date) continue;

    cb.addEventListener("change", () => {
      const entry = state.corrections[f.id] || { corrected: false, note: "", dateISO: "" };
      entry.corrected = cb.checked;
      if (entry.corrected && !entry.dateISO) entry.dateISO = new Date().toISOString().slice(0, 10);
      if (!entry.corrected) {
        // Keep note/date (optional). If you prefer to clear, uncomment:
        // entry.note = "";
        // entry.dateISO = "";
      }
      state.corrections[f.id] = entry;
      saveCorrections(state.corrections);
      render(); // refresh KPIs/results
    });

    note.addEventListener("input", () => {
      const entry = state.corrections[f.id] || { corrected: false, note: "", dateISO: "" };
      entry.note = note.value;
      state.corrections[f.id] = entry;
      saveCorrections(state.corrections);
      // no full render needed
    });

    date.addEventListener("change", () => {
      const entry = state.corrections[f.id] || { corrected: false, note: "", dateISO: "" };
      entry.dateISO = date.value;
      state.corrections[f.id] = entry;
      saveCorrections(state.corrections);
      renderResults();
    });
  }
}

function rowHtml(f) {
  const corr = state.corrections[f.id] || { corrected: false, note: "", dateISO: "" };
  const checked = corr.corrected ? "checked" : "";
  const sev = (f.severity || "info").toLowerCase();

  const findingMain = escapeHtml(f.finding || "");
  const requirement = escapeHtml(f.requirement || "");
  const recommendation = escapeHtml(f.recommendation || "");
  const category = escapeHtml(f.category || "General");

  return `
    <tr>
      <td><span class="small">${escapeHtml(f.id)}</span></td>
      <td>
        <div class="finding-title">${category}</div>
        <div class="finding-sub">
          ${f.owner ? `Owner: ${escapeHtml(f.owner)}` : ""}
          ${f.dueDate ? `${f.owner ? " • " : ""}Due: ${escapeHtml(f.dueDate)}` : ""}
        </div>
      </td>
      <td>
        <div class="finding-title">${findingMain}</div>
        ${requirement ? `<div class="finding-sub">Requirement: ${requirement}</div>` : ""}
      </td>
      <td><span class="badge ${sev}">${escapeHtml(capitalize(sev))}</span></td>
      <td>
        <div class="finding-sub">${recommendation || `<span class="muted">—</span>`}</div>
      </td>
      <td>
        <div class="corr-box">
          <div class="corr-row">
            <label class="checkbox" title="Mark as corrected">
              <input id="cb-${cssId(f.id)}" type="checkbox" ${checked} />
              Corrected
            </label>
            <input id="date-${cssId(f.id)}" class="input input-sm" style="max-width:160px"
              type="date" value="${escapeAttr(corr.dateISO || "")}" />
          </div>
          <textarea id="note-${cssId(f.id)}" class="note" placeholder="Correction details / evidence (what changed, ticket link, proof, etc.)">${escapeHtml(corr.note || "")}</textarea>
        </div>
      </td>
    </tr>
  `;
}

function renderResults() {
  if (!state.findings.length) {
    els.resultsBox.innerHTML = `<div class="empty">No report loaded yet.</div>`;
    return;
  }

  const total = state.findings.length;
  const correctedFindings = state.findings.filter((f) => state.corrections[f.id]?.corrected);
  const openFindings = state.findings.filter((f) => !state.corrections[f.id]?.corrected);

  const corrected = correctedFindings.length;
  const pct = total ? Math.round((corrected / total) * 100) : 0;

  const topOpen = openFindings.slice(0, 8);
  const topDone = correctedFindings.slice(0, 8);

  els.resultsBox.innerHTML = `
    <div class="split">
      <div>
        <h3>Summary</h3>
        <div class="muted">
          Progress: <b>${pct}%</b> (${corrected} corrected / ${total} total).
          ${openFindings.length ? `Remaining: <b>${openFindings.length}</b>.` : `All findings corrected.`}
        </div>

        <h3 style="margin-top:14px;">Recently corrected (preview)</h3>
        ${
          topDone.length
            ? `<ul>${topDone.map((f) => `<li><b>${escapeHtml(f.id)}</b> — ${escapeHtml(f.finding)}</li>`).join("")}</ul>`
            : `<div class="muted">No corrected items yet.</div>`
        }
      </div>

      <div>
        <h3>Open findings (preview)</h3>
        ${
          topOpen.length
            ? `<ul>${topOpen.map((f) => `<li><b>${escapeHtml(f.id)}</b> — ${escapeHtml(f.finding)}</li>`).join("")}</ul>`
            : `<div class="muted">No open findings.</div>`
        }

        <h3 style="margin-top:14px;">What to deliver as “evidence”</h3>
        <ul>
          <li>Link to ticket / change request</li>
          <li>Before/after config or screenshot</li>
          <li>Policy/procedure update reference</li>
          <li>Access logs / audit trail snapshot</li>
        </ul>
      </div>
    </div>
  `;
}

function exportResultsCSV() {
  if (!state.findings.length) {
    alert("No report loaded.");
    return;
  }

  const headers = [
    "ID",
    "Category",
    "Requirement",
    "Finding",
    "Severity",
    "Recommendation",
    "Owner",
    "DueDate",
    "Corrected",
    "CorrectionDate",
    "CorrectionNote",
  ];

  const lines = [headers.join(",")];

  for (const f of state.findings) {
    const corr = state.corrections[f.id] || { corrected: false, note: "", dateISO: "" };

    const row = [
      f.id,
      f.category,
      f.requirement,
      f.finding,
      f.severity,
      f.recommendation,
      f.owner,
      f.dueDate,
      corr.corrected ? "YES" : "NO",
      corr.dateISO || "",
      corr.note || "",
    ].map(csvCell);

    lines.push(row.join(","));
  }

  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
  const name = `AMIGOS_Corrections_${new Date().toISOString().slice(0, 10)}.csv`;
  downloadBlob(blob, name);
}

function csvCell(v) {
  const s = String(v ?? "");
  // Escape quotes by doubling them and wrap in quotes if needed
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

function loadCorrections() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}
function saveCorrections(obj) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(obj || {}));
  } catch {
    // ignore
  }
}

function parseCsv(text) {
  // Minimal CSV parser (handles quotes)
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];

    if (ch === '"' && inQuotes && next === '"') {
      cell += '"';
      i++;
      continue;
    }
    if (ch === '"') {
      inQuotes = !inQuotes;
      continue;
    }
    if (ch === "," && !inQuotes) {
      row.push(cell);
      cell = "";
      continue;
    }
    if ((ch === "\n" || ch === "\r") && !inQuotes) {
      if (ch === "\r" && next === "\n") i++;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      continue;
    }
    cell += ch;
  }
  // last cell
  row.push(cell);
  rows.push(row);
  // remove trailing empty rows
  return rows.filter((r) => r.some((c) => String(c).trim() !== ""));
}

function capitalize(s) {
  if (!s) return "";
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
function escapeAttr(s) {
  return escapeHtml(s).replaceAll("`", "&#096;");
}
function cssId(s) {
  return String(s).replace(/[^a-zA-Z0-9_-]/g, "_");
}
