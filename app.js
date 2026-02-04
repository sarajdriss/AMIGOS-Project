/* global XLSX */

/**
 * AMIGOS — corrective actions tracker (client-only)
 * Viewer (default): shows CLOSED NCs only (read-only)
 * Admin: can load report, edit checklist, attach evidence, export, reset
 *
 * Admin access (ONLY):
 *  - User: SARA  | Password: AMIGOS 55
 *  - User: SARA1 | Password: AMIGOS 555
 *
 * Note: Client-only authentication (UI-level). Not a secure backend auth.
 */

const APP_VERSION = "v2.3.0";

/* ---------------- Direct source links (PDF) ---------------- */
const MOROCCO_LAW_PDF =
  "https://www.cour-constitutionnelle.ma/Documents/Lois/%D9%85%D8%AF%D9%88%D9%86%D8%A9%20%D8%A7%D9%84%D8%B4%D8%BA%D9%84.pdf";

const INDITEX_COC_PDF =
  "https://www.inditex.com/itxcomweb/api/media/8cd88d29-0571-43d5-a6c3-a6c34671e4c1/inditex_code_of_conduct_for_manufacturers_and_suppliers.pdf";

/* ---------------- Report source defaults ---------------- */
const DEFAULT_SHEET_URL =
  "https://docs.google.com/spreadsheets/d/1wpp99NP1l83r_hDi5klWy3NKmHn9n-mX/edit?usp=drive_link&ouid=109090514128015175660&rtpof=true&sd=true";

/* ---------------- Storage keys ---------------- */
const STORAGE_KEY = "amigos_nc_tracker_v2_progress"; // checklist progress + evidence
const DATA_KEY = "amigos_nc_tracker_v2_data"; // last loaded report items/meta (so Viewer can read)

/* ---------------- Session keys (admin mode) ---------------- */
const SESSION_ADMIN_KEY = "amigos_nc_tracker_admin";
const SESSION_USER_KEY = "amigos_nc_tracker_admin_user";

/* ---------------- Admin users (ONLY these two) ---------------- */
const ALLOWED_USERS = [
  { user: "SARA", pass: "AMIGOS 55" },
  { user: "SARA1", pass: "AMIGOS 555" },
];

/* ---------------- Evidence file limits (localStorage) ---------------- */
const MAX_EVIDENCE_FILES = 6; // per NC
const MAX_EVIDENCE_MB_EACH = 2.0; // per file (DataURL stored in localStorage)

/* ---------------- DOM elements ---------------- */
const els = {
  // Admin / login UI (from index.html)
  adminBtn: document.getElementById("adminBtn"),
  logoutBtn: document.getElementById("logoutBtn"),
  modeText: document.getElementById("modeText"),

  loginDialog: document.getElementById("loginDialog"),
  loginForm: document.getElementById("loginForm"),
  closeLoginDialog: document.getElementById("closeLoginDialog"),
  cancelLogin: document.getElementById("cancelLogin"),
  adminCode: document.getElementById("adminCode"),
  adminPass: document.getElementById("adminPass"),
  loginError: document.getElementById("loginError"),

  // Load sources
  sheetUrl: document.getElementById("sheetUrl"),
  btnLoadSheet: document.getElementById("btnLoadSheet"),
  fileInput: document.getElementById("fileInput"),

  // Exports / reset (admin only)
  btnExportCSV: document.getElementById("btnExportCSV"),
  btnExportJSON: document.getElementById("btnExportJSON"),
  btnReset: document.getElementById("btnReset"),

  // Options (admin only in the new index)
  toggleOnlyNC: document.getElementById("toggleOnlyNC"),
  toggleRequireEvidence: document.getElementById("toggleRequireEvidence"),

  reportMeta: document.getElementById("reportMeta"),

  // KPIs
  kpiTotal: document.getElementById("kpiTotal"),
  kpiNC: document.getElementById("kpiNC"),
  kpiClosed: document.getElementById("kpiClosed"),
  kpiOpen: document.getElementById("kpiOpen"),

  ringPct: document.getElementById("ringPct"),
  barFill: document.getElementById("barFill"),
  progressText: document.getElementById("progressText"),

  // Filters
  searchInput: document.getElementById("searchInput"),
  filterSeverity: document.getElementById("filterSeverity"),
  filterState: document.getElementById("filterState"),

  // Lists
  dashboardList: document.getElementById("dashboardList"),
  ncList: document.getElementById("ncList"),

  // Navigation
  navItems: Array.from(document.querySelectorAll(".nav-item")),
  views: Array.from(document.querySelectorAll(".view")),
};

/* ---------------- App state ---------------- */
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

/* ============================================================
   Admin mode helpers
   ============================================================ */

function isAdmin() {
  return sessionStorage.getItem(SESSION_ADMIN_KEY) === "1";
}

function setAdminMode(on, username = "") {
  sessionStorage.setItem(SESSION_ADMIN_KEY, on ? "1" : "0");
  if (on) sessionStorage.setItem(SESSION_USER_KEY, username);
  else sessionStorage.removeItem(SESSION_USER_KEY);

  updateUIForMode();
  render();
}

function updateUIForMode() {
  const admin = isAdmin();
  const username = sessionStorage.getItem(SESSION_USER_KEY) || "";

  // Buttons
  if (els.adminBtn) els.adminBtn.classList.toggle("hidden", admin);
  if (els.logoutBtn) els.logoutBtn.classList.toggle("hidden", !admin);

  // Toggle all adminOnly blocks
  document.querySelectorAll(".adminOnly").forEach((el) => {
    el.classList.toggle("hidden", !admin);
  });

  // Mode text (Arabic)
  if (els.modeText) {
    if (admin) {
      els.modeText.innerHTML =
        `أنت في <strong>وضع المسؤول</strong>${username ? ` (${escapeHtml(username)})` : ""}: يمكنك تحميل التقرير، تعديل البيانات، رفع الأدلة، وإغلاق حالات NC.`;
    } else {
      els.modeText.innerHTML =
        `أنت في <strong>وضع المشاهدة</strong>: يتم عرض <strong>عدم المطابقة المغلقة فقط</strong>. للتعديل أو رفع الأدلة يلزم دخول المسؤول.`;
    }
  }

  // Default options for viewer
  if (!admin) {
    if (els.toggleOnlyNC) els.toggleOnlyNC.checked = true;
    if (els.toggleRequireEvidence) els.toggleRequireEvidence.checked = true;
    if (els.filterState) els.filterState.value = "closed";
  }
}

function openLogin() {
  if (!els.loginDialog) return;

  els.loginError?.classList.add("hidden");
  if (els.adminCode) els.adminCode.value = "";
  if (els.adminPass) els.adminPass.value = "";

  try {
    els.loginDialog.showModal();
  } catch {
    // fallback
    els.loginDialog.setAttribute("open", "open");
  }

  setTimeout(() => els.adminCode?.focus(), 50);
}

function closeLogin() {
  if (!els.loginDialog) return;
  try {
    els.loginDialog.close();
  } catch {
    els.loginDialog.removeAttribute("open");
  }
}

function handleLoginSubmit(e) {
  e.preventDefault();

  const user = (els.adminCode?.value || "").trim();
  const pass = (els.adminPass?.value || "").trim();

  const ok = ALLOWED_USERS.some((u) => u.user === user && u.pass === pass);
  if (ok) {
    setAdminMode(true, user);
    closeLogin();
  } else {
    els.loginError?.classList.remove("hidden");
  }
}

/* ============================================================
   Init / navigation / events
   ============================================================ */

function init() {
  // Load last saved dataset for Viewer
  const saved = loadData();
  if (saved?.items?.length) {
    state.items = saved.items;
    state.loadedAt = saved.loadedAtISO ? new Date(saved.loadedAtISO) : null;
    state.source = saved.source || null;

    if (els.reportMeta) {
      const rows = state.items.length;
      els.reportMeta.textContent = `${state.reportName} • تم التحميل (محفوظ محلياً) • ${rows} صف`;
    }
  }

  if (els.sheetUrl) els.sheetUrl.value = DEFAULT_SHEET_URL;

  updateUIForMode();

  // Navigation
  els.navItems.forEach((btn) => {
    btn.addEventListener("click", () => {
      els.navItems.forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      const viewId = btn.dataset.view;
      els.views.forEach((v) => v.classList.toggle("hidden", v.id !== viewId));
    });
  });

  // Admin buttons
  els.adminBtn?.addEventListener("click", openLogin);
  els.logoutBtn?.addEventListener("click", () => setAdminMode(false));

  // Login dialog events
  els.loginForm?.addEventListener("submit", handleLoginSubmit);
  els.closeLoginDialog?.addEventListener("click", closeLogin);
  els.cancelLogin?.addEventListener("click", closeLogin);

  // Load from Google Sheet (ADMIN ONLY)
  els.btnLoadSheet?.addEventListener("click", async () => {
    if (!isAdmin()) return openLogin();

    const url = els.sheetUrl?.value.trim();
    if (!url) return;

    setEmpty("جاري تحميل التقرير من Google Sheets…");
    try {
      await loadFromGoogleSheetUrl(url);
    } catch (e) {
      console.error(e);
      setEmpty("تعذر التحميل من Google Sheets. قم برفع ملف XLSX/CSV بدلاً من ذلك (قد يكون الملف خاصاً/محجوباً).");
      if (els.reportMeta) {
        els.reportMeta.textContent = "فشل التحميل من Google Sheets (قد يكون خاصاً/محجوباً). قم برفع XLSX/CSV.";
      }
    }
  });

  // Upload file (ADMIN ONLY)
  els.fileInput?.addEventListener("change", async (e) => {
    if (!isAdmin()) {
      if (els.fileInput) els.fileInput.value = "";
      return openLogin();
    }

    const file = e.target.files?.[0];
    if (!file) return;

    setEmpty(`جاري تحميل الملف: ${escapeHtml(file.name)} …`);
    try {
      await loadFromLocalFile(file);
    } catch (err) {
      console.error(err);
      setEmpty("تعذر قراءة الملف. الرجاء رفع ملف XLSX/CSV صالح.");
    } finally {
      if (els.fileInput) els.fileInput.value = "";
    }
  });

  // Options
  els.toggleOnlyNC?.addEventListener("change", () => render());
  els.toggleRequireEvidence?.addEventListener("change", () => {
    // Important: keep checklist.evidence aligned when rule changes
    syncEvidenceChecklistForAll();
    saveProgress(state.progress);
    render();
  });

  // Filters
  els.searchInput?.addEventListener("input", () => render());
  els.filterSeverity?.addEventListener("change", () => render());
  els.filterState?.addEventListener("change", () => render());

  // Exports (ADMIN ONLY)
  els.btnExportCSV?.addEventListener("click", () => {
    if (!isAdmin()) return openLogin();
    exportCSV();
  });
  els.btnExportJSON?.addEventListener("click", () => {
    if (!isAdmin()) return openLogin();
    exportJSON();
  });

  // Reset (ADMIN ONLY)
  els.btnReset?.addEventListener("click", () => {
    if (!isAdmin()) return openLogin();
    if (!confirm("هل تريد إعادة ضبط الحالات المحفوظة على هذا المتصفح؟")) return;

    localStorage.removeItem(STORAGE_KEY);
    localStorage.removeItem(DATA_KEY);

    state.items = [];
    state.source = null;
    state.loadedAt = null;
    state.progress = {};

    if (els.reportMeta) {
      els.reportMeta.textContent = "تمت إعادة الضبط. الرجاء تحميل التقرير من جديد.";
    }

    render();
  });

  // Initial render
  render();
}

/* ============================================================
   Data loading + parsing
   ============================================================ */

/**
 * Item model (normalized)
 * id, category, finding, recommendation, severity, isNC,
 * reqMorocco, reqInditex, rawStatus
 */

function setEmpty(message) {
  const html = `<div class="empty">${escapeHtml(message)}</div>`;
  if (els.dashboardList) els.dashboardList.innerHTML = html;
  if (els.ncList) els.ncList.innerHTML = html;
}

async function loadFromGoogleSheetUrl(sheetUrl) {
  const sheetId = extractSheetId(sheetUrl);
  if (!sheetId) throw new Error("Could not extract sheetId from URL.");

  const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
  const res = await fetch(exportUrl, { cache: "no-store" });
  if (!res.ok) throw new Error(`Fetch failed: ${res.status}`);

  const ab = await res.arrayBuffer();
  const { items, meta } = parseXlsx(ab);

  state.items = items;
  state.loadedAt = new Date();
  state.source = meta;

  saveData({
    items: state.items,
    loadedAtISO: state.loadedAt.toISOString(),
    source: state.source,
    appVersion: APP_VERSION,
  });

  if (els.reportMeta) {
    els.reportMeta.textContent = `${state.reportName} • تم التحميل • ${state.items.length} صف`;
  }
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

    saveData({
      items: state.items,
      loadedAtISO: state.loadedAt.toISOString(),
      source: state.source,
      appVersion: APP_VERSION,
    });

    if (els.reportMeta) {
      els.reportMeta.textContent = `${state.reportName} • تم التحميل (CSV) • ${items.length} صف`;
    }
    render();
    return;
  }

  const ab = await file.arrayBuffer();
  const { items, meta } = parseXlsx(ab);

  state.items = items;
  state.loadedAt = new Date();
  state.source = { ...meta, name: file.name };

  saveData({
    items: state.items,
    loadedAtISO: state.loadedAt.toISOString(),
    source: state.source,
    appVersion: APP_VERSION,
  });

  if (els.reportMeta) {
    els.reportMeta.textContent = `${state.reportName} • تم التحميل (XLSX) • ${items.length} صف`;
  }
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
    const header = rows[0].map((h) => String(h ?? ""));
    rows = rows.slice(1).map((r) => {
      const obj = {};
      header.forEach((h, i) => (obj[h] = r[i]));
      return obj;
    });
  }

  const items = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || {};
    const keys = Object.keys(r);

    const anyVal = keys.some((k) => String(r[k] ?? "").trim() !== "");
    if (!anyVal) continue;

    const get = (candidates) => {
      const k = findKey(keys, candidates);
      return k ? String(r[k] ?? "").trim() : "";
    };

    const idRaw = get(["id", "ref", "#", "no", "n", "index", "finding id", "findingid"]);
    const category = get(["category", "section", "domain", "area", "topic"]);
    const finding = get(["finding", "non conformity", "nonconformity", "issue", "observation", "gap", "problem"]);
    const recommendation = get(["recommendation", "corrective action", "action", "remediation", "proposed action"]);
    const severity = normalizeSeverity(get(["severity", "risk", "priority", "criticality", "rating"]));

    const rawStatus = get(["status", "conformity", "compliance result", "result", "nc status", "state"]);
    const reqMorocco = get([
      "morocco law",
      "legal reference",
      "law reference",
      "moroccan law",
      "maroc law",
      "code du travail",
    ]);
    const reqInditex = get([
      "inditex",
      "inditex reference",
      "inditex requirement",
      "code of conduct",
      "ics",
      "social audit",
    ]);

    const id = idRaw || `F-${String(i + 1).padStart(3, "0")}`;
    const isNC = detectNonConformity({ rawStatus, finding, recommendation });

    items.push({
      id,
      category: category || "عام",
      finding: finding || summarizeFallback(r),
      recommendation,
      severity,
      rawStatus,
      reqMorocco,
      reqInditex,
      isNC,
    });

    // Ensure progress state exists (create for all items so closure state is stable)
    if (!state.progress[id]) {
      state.progress[id] = defaultProgressForItem();
    }
  }

  // Keep evidence checklist aligned if rule is ON
  syncEvidenceChecklistForAll();

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
      managementSignoff: false,
    },
    evidenceLink: "",
    evidenceNote: "",
    evidenceFiles: [], // [{name,type,size,dataUrl,uploadedAtISO}]
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
  const ncWords = [
    "non conform",
    "non-conform",
    "nonconform",
    "nc",
    "nok",
    "no",
    "fail",
    "not compliant",
    "non compliant",
    "غير مطابق",
    "عدم مطابقة",
  ];
  const okWords = ["conform", "ok", "yes", "pass", "compliant", "closed", "مطابق", "مغلق", "مغلقة"];

  if (s) {
    if (ncWords.some((w) => s === w || s.includes(w))) return true;
    if (okWords.some((w) => s === w || s.includes(w))) return false;
  }

  // Heuristic fallback (weak): if text clearly says non-compliant
  if (["non conform", "non-compliant", "not compliant", "non compliant", "violation", "غير مطابق"].some((w) => f.includes(w))) {
    return true;
  }

  return false;
}

/* ============================================================
   Rendering
   ============================================================ */

function render() {
  // Keep evidence checkbox aligned (important when data already has evidence)
  syncEvidenceChecklistForAll();

  renderKPIs();
  renderLists();
}

function getVisibleItems() {
  const q = (els.searchInput?.value || "").trim().toLowerCase();
  const sev = els.filterSeverity?.value || "all";
  const st = els.filterState?.value || "all";
  const onlyNC = !!els.toggleOnlyNC?.checked;

  let items = state.items.slice();

  // Viewer mode: CLOSED NCs only, always
  if (!isAdmin()) {
    items = items.filter((x) => x.isNC && getClosureState(x) === "closed");
  } else {
    if (onlyNC) items = items.filter((x) => x.isNC);
    if (st !== "all") items = items.filter((x) => getClosureState(x) === st);
  }

  if (q) {
    items = items.filter((x) => {
      const hay = `${x.id} ${x.category} ${x.finding} ${x.recommendation} ${x.reqMorocco} ${x.reqInditex}`.toLowerCase();
      return hay.includes(q);
    });
  }

  if (sev !== "all") items = items.filter((x) => (x.severity || "info") === sev);

  return items;
}

function renderKPIs() {
  const admin = isAdmin();

  const all = state.items.slice();
  const allNC = all.filter((x) => x.isNC);

  const baseItems = admin ? allNC : allNC.filter((x) => getClosureState(x) === "closed");

  const totalItems = admin ? all.length : baseItems.length;
  const ncCount = baseItems.length;

  const closed = baseItems.filter((x) => getClosureState(x) === "closed").length;
  const open = admin
    ? allNC.length - allNC.filter((x) => getClosureState(x) === "closed").length
    : 0;

  if (els.kpiTotal) els.kpiTotal.textContent = String(totalItems || 0);
  if (els.kpiNC) els.kpiNC.textContent = String(ncCount || 0);
  if (els.kpiClosed) els.kpiClosed.textContent = String(closed || 0);
  if (els.kpiOpen) els.kpiOpen.textContent = String(open || 0);

  const pct = baseItems.length ? Math.round((closed / baseItems.length) * 100) : 0;

  if (els.ringPct) els.ringPct.textContent = baseItems.length ? `${pct}%` : "—";
  const ring = document.querySelector(".ring");
  if (ring) ring.style.setProperty("--ring", `${pct}%`);

  if (els.barFill) els.barFill.style.width = `${pct}%`;

  if (els.progressText) {
    if (!admin) {
      els.progressText.textContent = baseItems.length
        ? `يتم عرض حالات NC المغلقة فقط (${baseItems.length}).`
        : "لا توجد حالات NC مغلقة للعرض حالياً.";
    } else {
      const closedAll = allNC.filter((x) => getClosureState(x) === "closed").length;
      const pctAll = allNC.length ? Math.round((closedAll / allNC.length) * 100) : 0;
      els.progressText.textContent = allNC.length
        ? `تم إغلاق ${closedAll} من أصل ${allNC.length} حالة NC (${pctAll}%).`
        : "لم يتم اكتشاف حالات NC بعد (تحقق من أعمدة الحالة/النتيجة في الملف).";
    }
  }
}

function renderLists() {
  if (!state.items.length) {
    setEmpty(
      isAdmin()
        ? "قم بتحميل تقرير لعرض عدم المطابقة."
        : "لا توجد بيانات محفوظة على هذا المتصفح. اطلب من المسؤول تحميل التقرير هنا لعرض حالات NC المغلقة."
    );
    return;
  }

  const visible = getVisibleItems();
  if (!visible.length) {
    const msg = isAdmin()
      ? "لا توجد عناصر مطابقة للفلاتر."
      : "لا توجد حالات NC مغلقة مطابقة للبحث/الفلاتر.";
    if (els.dashboardList) els.dashboardList.innerHTML = `<div class="empty">${escapeHtml(msg)}</div>`;
    if (els.ncList) els.ncList.innerHTML = `<div class="empty">${escapeHtml(msg)}</div>`;
    return;
  }

  if (els.dashboardList) els.dashboardList.innerHTML = visible.slice(0, 12).map(renderCard).join("");
  if (els.ncList) els.ncList.innerHTML = visible.map(renderCard).join("");

  bindCardEvents();
}

function severityLabel(sev) {
  const s = (sev || "info").toLowerCase();
  const map = {
    critical: "حرج",
    high: "عالي",
    medium: "متوسط",
    low: "منخفض",
    info: "معلومة",
  };
  return map[s] || "معلومة";
}

function renderEvidenceFilesList(p) {
  const files = Array.isArray(p.evidenceFiles) ? p.evidenceFiles : [];
  if (!files.length) return `<div class="hint">لا توجد ملفات دليل مرفوعة.</div>`;

  return files
    .map((f, idx) => {
      const name = escapeHtml(f.name || `file-${idx + 1}`);
      const size = typeof f.size === "number" ? `${Math.round(f.size / 1024)} KB` : "";
      const href = f.dataUrl ? escapeAttr(f.dataUrl) : "";
      return `
        <div class="row" style="justify-content:space-between; gap:8px; margin-top:6px;">
          <div class="muted">${name} ${size ? `— ${escapeHtml(size)}` : ""}</div>
          <div class="row" style="gap:8px;">
            ${href ? `<a class="btn btn-dark" href="${href}" download="${escapeAttr(f.name || "evidence")}">تحميل</a>` : ""}
            ${
              isAdmin()
                ? `<button class="btn btn-ghost-danger js-evRemove" type="button" data-ev-idx="${idx}">حذف</button>`
                : ""
            }
          </div>
        </div>
      `;
    })
    .join("");
}

function renderCard(item) {
  const p = state.progress[item.id] || defaultProgressForItem();
  const stKey = getClosureState(item);
  const sev = (item.severity || "info").toLowerCase();

  const admin = isAdmin();
  const dis = admin ? "" : "disabled";
  const roHint = admin ? "" : `<div class="hint">وضع المشاهدة: البيانات للعرض فقط.</div>`;

  const resultLabel = stKey === "closed" ? "مطابق (مغلقة)" : "غير مطابق (غير مغلقة)";
  const checklist = p.checklist;

  return `
    <article class="nc-card" data-id="${escapeAttr(item.id)}">
      <div class="nc-card-head">
        <div>
          <div class="nc-title">${escapeHtml(item.id)} — ${escapeHtml(item.category)}</div>
          <div class="nc-sub">${escapeHtml(item.finding)}</div>
        </div>

        <div class="pills">
          <span class="pill sev-${escapeAttr(sev)}">الشدّة: ${escapeHtml(severityLabel(sev))}</span>
          <span class="pill state-${stKey === "closed" ? "closed" : "open"}">النتيجة: ${escapeHtml(resultLabel)}</span>
          ${item.isNC ? `<span class="pill">NC</span>` : `<span class="pill">معلومة</span>`}
        </div>
      </div>

      <div class="nc-card-body">
        <div class="nc-grid">
          <div class="block">
            <div class="block-title">مراجع المتطلبات</div>

            <div class="hint">
              مصادر مباشرة:
              <a href="${escapeAttr(MOROCCO_LAW_PDF)}" target="_blank" rel="noopener">مدونة الشغل (65-99) PDF</a>
              •
              <a href="${escapeAttr(INDITEX_COC_PDF)}" target="_blank" rel="noopener">Inditex CoC PDF</a>
            </div>

            <div class="field">
              <label>مرجع القانون المغربي (حسب فريق الامتثال)</label>
              <input class="input js-moroccoRef" type="text"
                placeholder="مثال: مدونة الشغل — المادة/الفقرة…"
                value="${escapeAttr(item.reqMorocco || "")}" ${dis}>
            </div>

            <div class="field">
              <label>مرجع متطلبات Inditex</label>
              <input class="input js-inditexRef" type="text"
                placeholder="مثال: Inditex CoC — البند/الفقرة…"
                value="${escapeAttr(item.reqInditex || "")}" ${dis}>
            </div>

            <div class="field">
              <label>التوصية / الإجراء التصحيحي المطلوب</label>
              <textarea class="input js-reco" placeholder="اكتب الإجراء المطلوب…" ${dis}>${escapeHtml(item.recommendation || "")}</textarea>
              <div class="hint">يمكنك تعديل النص محلياً ليتوافق مع صياغة CAPA.</div>
            </div>

            ${roHint}
          </div>

          <div class="block">
            <div class="block-title">قائمة التحقق (شروط الإغلاق)</div>

            <div class="checklist">
              ${checkRow("containment", "تم تطبيق إجراء احتواء فوري (إيقاف الخطر/حل مؤقت)", checklist.containment, dis)}
              ${checkRow("rootCause", "تم إنجاز تحليل السبب الجذري", checklist.rootCause, dis)}
              ${checkRow("correctiveAction", "تم تنفيذ الإجراء التصحيحي (معالجة سبب NC)", checklist.correctiveAction, dis)}
              ${checkRow("preventiveAction", "تم تنفيذ الإجراء الوقائي (منع التكرار)", checklist.preventiveAction, dis)}
              ${checkRow("evidence", "تم إرفاق دليل (رابط/ملاحظة/ملف)", checklist.evidence, dis)}
              ${checkRow("verification", "تم إجراء التحقق الداخلي", checklist.verification, dis)}
              ${checkRow("managementSignoff", "تم اعتماد الإدارة / التوقيع", checklist.managementSignoff, dis)}
            </div>

            <div class="field">
              <label>رابط الدليل (تذكرة / Drive / صورة / مستند)</label>
              <input class="input js-evidenceLink" type="url" 
