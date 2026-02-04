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
const MAX_EVIDENCE_FILES = 6;       // per NC
const MAX_EVIDENCE_MB_EACH = 2.0;   // per file (DataURL stored in localStorage)

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

  // Default options for admin / viewer
  if (!admin) {
    // Viewer cannot change these; but keep safe defaults if present
    if (els.toggleOnlyNC) els.toggleOnlyNC.checked = true;
    if (els.toggleRequireEvidence) els.toggleRequireEvidence.checked = true;
    if (els.filterState) els.filterState.value = "closed";
  }
}

function openLogin() {
  if (!els.loginDialog) return;

  if (els.loginError) els.loginError.classList.add("hidden");
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

  // Apply mode on startup
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
      els.fileInput.value = "";
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
      els.fileInput.value = "";
    }
  });

  // Options
  els.toggleOnlyNC?.addEventListener("change", () => render());
  els.toggleRequireEvidence?.addEventListener("change", () => render());

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
  const res = await fetch(exportUrl);
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
  ];
  const okWords = ["conform", "ok", "yes", "pass", "compliant", "closed"];

  if (s) {
    if (ncWords.some((w) => s === w || s.includes(w))) return true;
    if (okWords.some((w) => s === w || s.includes(w))) return false;
  }

  // Heuristic fallback (weak): if text clearly says non-compliant
  if (["non conform", "non-compliant", "not compliant", "non compliant", "violation"].some((w) => f.includes(w))) {
    return true;
  }

  // Default: avoid false positives
  return false;
}

/* ============================================================
   Rendering
   ============================================================ */

function render() {
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
  // Admin shows overall stats; Viewer shows only visible (closed NCs)
  const admin = isAdmin();

  const all = state.items.slice();
  const allNC = all.filter((x) => x.isNC);

  let baseItems;
  if (admin) {
    baseItems = allNC;
  } else {
    baseItems = allNC.filter((x) => getClosureState(x) === "closed");
  }

  const totalItems = admin ? all.length : baseItems.length;
  const ncCount = baseItems.length;

  const closed = baseItems.filter((x) => getClosureState(x) === "closed").length;
  const open = admin ? (allNC.length - allNC.filter((x) => getClosureState(x) === "closed").length) : 0;

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
      els.progressText.textContent = allNC.length
        ? `تم إغلاق ${allNC.filter((x) => getClosureState(x) === "closed").length} من أصل ${allNC.length} حالة NC (${Math.round((allNC.filter((x) => getClosureState(x) === "closed").length / allNC.length) * 100)}%).`
        : "لم يتم اكتشاف حالات NC بعد (تحقق من أعمدة الحالة/النتيجة في الملف).";
    }
  }
}

function renderLists() {
  // If no data: Viewer should see guidance (admin load required on same browser)
  if (!state.items.length) {
    setEmpty(isAdmin()
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

  // Dashboard shows compact list
  if (els.dashboardList) els.dashboardList.innerHTML = visible.slice(0, 12).map(renderCard).join("");
  // NC view shows full list
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

function stateLabel(key) {
  const map = {
    open: "مفتوحة",
    progress: "قيد التنفيذ",
    ready: "جاهزة للتحقق",
    closed: "مغلقة",
  };
  return map[key] || key;
}

function renderEvidenceFilesList(p) {
  const files = Array.isArray(p.evidenceFiles) ? p.evidenceFiles : [];
  if (!files.length) return `<div class="hint">لا توجد ملفات دليل مرفوعة.</div>`;

  const items = files.map((f, idx) => {
    const name = escapeHtml(f.name || `file-${idx + 1}`);
    const size = typeof f.size === "number" ? `${Math.round(f.size / 1024)} KB` : "";
    const href = f.dataUrl ? escapeAttr(f.dataUrl) : "";
    return `
      <div class="row" style="justify-content:space-between; gap:8px; margin-top:6px;">
        <div class="muted">${name} ${size ? `— ${escapeHtml(size)}` : ""}</div>
        <div class="row" style="gap:8px;">
          ${href ? `<a class="btn btn-dark" href="${href}" download="${escapeAttr(f.name || "evidence")}">تحميل</a>` : ""}
          ${isAdmin() ? `<button class="btn btn-ghost-danger js-evRemove" type="button" data-ev-idx="${idx}">حذف</button>` : ""}
        </div>
      </div>
    `;
  }).join("");

  return items;
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
              <textarea class="input js-reco" placeholder="اكتب الإجراء المطلوب…"
                ${dis}>${escapeHtml(item.recommendation || "")}</textarea>
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
              <input class="input js-evidenceLink" type="url" placeholder="https://..."
                value="${escapeAttr(p.evidenceLink || "")}" ${dis}>
            </div>

            <div class="field">
              <label>ملاحظة دليل (إذا لا يوجد رابط)</label>
              <textarea class="input js-evidenceNote" placeholder="صف الدليل: ما الذي تغير؟ أين؟ من تحقق؟…"
                ${dis}>${escapeHtml(p.evidenceNote || "")}</textarea>
            </div>

            <div class="field">
              <label>رفع ملفات الدليل (اختياري / حسب السياسة)</label>
              <input class="input js-evidenceFiles" type="file" multiple accept="image/*,application/pdf" ${dis}>
              <div class="hint">
                حد مقترح: ${MAX_EVIDENCE_FILES} ملفات لكل NC، وحجم ${MAX_EVIDENCE_MB_EACH}MB لكل ملف (تخزين محلي).
              </div>
              ${isAdmin() ? `<div class="row" style="margin-top:8px;">
                  <button class="btn btn-ghost-danger js-evClear" type="button">مسح ملفات الدليل</button>
                </div>` : ""}
              <div class="card-lite" style="margin-top:10px;">
                <div class="card-lite-title">ملفات الدليل</div>
                ${renderEvidenceFilesList(p)}
              </div>
            </div>

            <div class="field">
              <label>المسؤول</label>
              <input class="input js-owner" type="text" placeholder="الاسم / القسم" value="${escapeAttr(p.owner || "")}" ${dis}>
            </div>

            <div class="row">
              <div class="field" style="flex:1">
                <label>تاريخ الاستحقاق</label>
                <input class="input js-dueDate" type="date" value="${escapeAttr(p.dueDate || "")}" ${dis}>
              </div>
              <div class="field" style="flex:1">
                <label>آخر تحديث</label>
                <input class="input" type="text" value="${escapeAttr(formatDateTime(p.updatedAtISO))}" disabled>
              </div>
            </div>

            <div class="field">
              <label>تعليق المراجع</label>
              <textarea class="input js-comment" placeholder="ملاحظات CAPA / رد المدقق / متابعة…"
                ${dis}>${escapeHtml(p.comment || "")}</textarea>
            </div>

            <div class="hint">
              تصبح الحالة <b>مغلقة</b> فقط عند استكمال عناصر قائمة التحقق الأساسية + التحقق + اعتماد الإدارة
              ${els.toggleRequireEvidence?.checked ? " + وجود دليل (رابط/ملاحظة/ملف)." : "."}
            </div>
          </div>
        </div>
      </div>
    </article>
  `;
}

function checkRow(key, label, checked, disAttr) {
  return `
    <label class="check">
      <input class="js-check" data-key="${escapeAttr(key)}" type="checkbox" ${checked ? "checked" : ""} ${disAttr}>
      <span>${escapeHtml(label)}</span>
    </label>
  `;
}

/* ============================================================
   Card events (ADMIN ONLY)
   ============================================================ */

function bindCardEvents() {
  if (!isAdmin()) return; // Viewer = read-only

  const cards = document.querySelectorAll(".nc-card");
  cards.forEach((card) => {
    const id = card.getAttribute("data-id");
    const item = state.items.find((x) => x.id === id);
    if (!item) return;

    const p = state.progress[id] || defaultProgressForItem();

    // checklist toggles
    card.querySelectorAll(".js-check").forEach((cb) => {
      cb.addEventListener("change", () => {
        const k = cb.dataset.key;
        p.checklist[k] = cb.checked;

        touchProgress(id, p);
        if (!saveProgress(state.progress)) {
          alert("تعذر حفظ التغييرات (قد تكون مساحة التخزين ممتلئة).");
        }
        render();
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

    // evidence file uploads
    const evidenceFiles = card.querySelector(".js-evidenceFiles");
    evidenceFiles?.addEventListener("change", async () => {
      const files = Array.from(evidenceFiles.files || []);
      evidenceFiles.value = "";
      if (!files.length) return;

      try {
        await addEvidenceFilesToProgress(p, files);
        autoSetEvidenceChecklist(p);
        touchProgress(id, p);

        if (!saveProgress(state.progress)) {
          alert("تعذر حفظ ملفات الدليل (قد تكون مساحة التخزين ممتلئة). جرّب ملفات أصغر أو استخدم رابط دليل.");
          // do not renderLists to avoid misleading; but keep in-memory
        } else {
          render();
        }
      } catch (err) {
        console.error(err);
        alert("تعذر معالجة ملفات الدليل.");
      }
    });

    // clear evidence files
    const evClear = card.querySelector(".js-evClear");
    evClear?.addEventListener("click", () => {
      if (!confirm("هل تريد مسح جميع ملفات الدليل لهذا الـ NC؟")) return;
      p.evidenceFiles = [];
      autoSetEvidenceChecklist(p);
      touchProgress(id, p);
      saveProgress(state.progress);
      render();
    });

    // remove single evidence file
    card.querySelectorAll(".js-evRemove").forEach((btn) => {
      btn.addEventListener("click", () => {
        const idx = Number(btn.getAttribute("data-ev-idx"));
        if (!Number.isFinite(idx)) return;
        p.evidenceFiles.splice(idx, 1);
        autoSetEvidenceChecklist(p);
        touchProgress(id, p);
        saveProgress(state.progress);
        render();
      });
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

    // Requirement refs & reco (editable)
    const moroccoRef = card.querySelector(".js-moroccoRef");
    moroccoRef?.addEventListener("input", () => {
      item.reqMorocco = moroccoRef.value;
      // persist report edits locally too
      saveData({
        items: state.items,
        loadedAtISO: state.loadedAt ? state.loadedAt.toISOString() : null,
        source: state.source,
      });
    });

    const inditexRef = card.querySelector(".js-inditexRef");
    inditexRef?.addEventListener("input", () => {
      item.reqInditex = inditexRef.value;
      saveData({
        items: state.items,
        loadedAtISO: state.loadedAt ? state.loadedAt.toISOString() : null,
        source: state.source,
      });
    });

    const reco = card.querySelector(".js-reco");
    reco?.addEventListener("input", () => {
      item.recommendation = reco.value;
      saveData({
        items: state.items,
        loadedAtISO: state.loadedAt ? state.loadedAt.toISOString() : null,
        source: state.source,
      });
    });
  });
}

async function addEvidenceFilesToProgress(p, files) {
  p.evidenceFiles = Array.isArray(p.evidenceFiles) ? p.evidenceFiles : [];

  for (const f of files) {
    if (p.evidenceFiles.length >= MAX_EVIDENCE_FILES) break;

    const mb = f.size / (1024 * 1024);
    if (mb > MAX_EVIDENCE_MB_EACH) {
      // skip large files
      continue;
    }

    const dataUrl = await fileToDataUrl(f);
    p.evidenceFiles.push({
      name: f.name,
      type: f.type,
      size: f.size,
      dataUrl,
      uploadedAtISO: new Date().toISOString(),
    });
  }
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(String(fr.result));
    fr.onerror = reject;
    fr.readAsDataURL(file);
  });
}

function autoSetEvidenceChecklist(p) {
  const requireEvidence = !!els.toggleRequireEvidence?.checked;
  if (!requireEvidence) return;
  p.checklist.evidence = hasEvidence(p);
}

function hasEvidence(p) {
  const linkOk = (p.evidenceLink || "").trim().length > 8;
  const noteOk = (p.evidenceNote || "").trim().length >= 20;
  const filesOk = Array.isArray(p.evidenceFiles) && p.evidenceFiles.length > 0;
  return linkOk || noteOk || filesOk;
}

/**
 * Closure states:
 * - open: nothing done
 * - progress: partial checklist
 * - ready: almost done
 * - closed: base checklist + verification + signoff (+ evidence if required)
 */
function getClosureState(item) {
  if (!item.isNC) return "closed"; // info items are not treated as NC

  const p = state.progress[item.id] || defaultProgressForItem();
  const c = p.checklist || {};

  const requireEvidence = !!els.toggleRequireEvidence?.checked;

  const requiredKeys = [
    "containment",
    "rootCause",
    "correctiveAction",
    "preventiveAction",
    "verification",
    "managementSignoff",
  ];
  const baseAllDone = requiredKeys.every((k) => !!c[k]);

  const evidenceOk = !requireEvidence || (c.evidence && hasEvidence(p));
  const allDone = baseAllDone && evidenceOk;

  const anyDone = Object.values(c).some(Boolean);

  if (!anyDone) return "open";
  if (allDone) return "closed";

  const missing = requiredKeys.filter((k) => !c[k]);
  if (missing.length === 1 && (missing[0] === "managementSignoff" || missing[0] === "verification")) return "ready";
  if (requireEvidence && !evidenceOk && baseAllDone) return "ready";

  return "progress";
}

function touchProgress(id, p) {
  p.updatedAtISO = new Date().toISOString();
  state.progress[id] = p;
}

/* ============================================================
   Exports (ADMIN ONLY)
   ============================================================ */

function exportCSV() {
  if (!state.items.length) return alert("لا يوجد تقرير محمّل.");

  const headers = [
    "ID",
    "IsNC",
    "Category",
    "Finding",
    "Severity",
    "MoroccoLawRef",
    "InditexRef",
    "Recommendation",
    "State",
    "Result",
    "Owner",
    "DueDate",
    "EvidenceLink",
    "EvidenceNote",
    "EvidenceFilesCount",
    "EvidenceFilesNames",
    "Containment",
    "RootCause",
    "CorrectiveAction",
    "PreventiveAction",
    "Evidence",
    "Verification",
    "ManagementSignoff",
    "UpdatedAt",
  ];

  const lines = [headers.join(",")];

  for (const it of state.items) {
    const p = state.progress[it.id] || defaultProgressForItem();
    const st = getClosureState(it);
    const result = it.isNC ? (st === "closed" ? "CONFORM" : "NON-CONFORM") : "CONFORM";

    const files = Array.isArray(p.evidenceFiles) ? p.evidenceFiles : [];
    const names = files.map((f) => f.name).filter(Boolean).join(" | ");

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
      String(files.length),
      names,
      bool(p.checklist.containment),
      bool(p.checklist.rootCause),
      bool(p.checklist.correctiveAction),
      bool(p.checklist.preventiveAction),
      bool(p.checklist.evidence),
      bool(p.checklist.verification),
      bool(p.checklist.managementSignoff),
      p.updatedAtISO || "",
    ].map(csvCell);

    lines.push(row.join(","));
  }

  downloadBlob(
    new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" }),
    `AMIGOS_NC_Corrections_${new Date().toISOString().slice(0, 10)}.csv`
  );
}

function exportJSON() {
  if (!state.items.length) return alert("لا يوجد تقرير محمّل.");

  const payload = {
    license: "AMIGOS — corrective actions tracker (client-only).",
    reportName: state.reportName,
    loadedAt: state.loadedAt?.toISOString() || null,
    source: state.source,
    requireEvidenceForClosure: !!els.toggleRequireEvidence?.checked,
    items: state.items,
    progress: state.progress,
  };

  downloadBlob(
    new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" }),
    `AMIGOS_NC_Corrections_${new Date().toISOString().slice(0, 10)}.json`
  );
}

/* ============================================================
   Storage (progress + last loaded data)
   ============================================================ */

function loadProgress() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function saveProgress(obj) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(obj || {}));
    return true;
  } catch {
    return false;
  }
}

function loadData() {
  try {
    const raw = localStorage.getItem(DATA_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function saveData(data) {
  try {
    localStorage.setItem(DATA_KEY, JSON.stringify(data || null));
    return true;
  } catch {
    // If too large, still keep progress; admin can rely on loading file again.
    return false;
  }
}

/* ============================================================
   Utilities
   ============================================================ */

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
  if (["1", "p1"].includes(s)) return "critical";
  if (["2", "p2"].includes(s)) return "high";
  if (["3", "p3"].includes(s)) return "medium";
  if (["4", "p4"].includes(s)) return "low";
  return "info";
}

function summarizeFallback(obj) {
  const pairs = Object.entries(obj)
    .map(([k, v]) => [String(k).trim(), String(v ?? "").trim()])
    .filter(([k, v]) => k && v);
  return pairs.slice(0, 3).map(([k, v]) => `${k}: ${v}`).join(" • ") || "—";
}

function findKey(keys, candidates) {
  const normalized = new Map(keys.map((k) => [norm(k), k]));
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
  return rows.filter((r) => r.some((c) => String(c).trim() !== ""));
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
function escapeAttr(s) {
  return escapeHtml(s).replaceAll("`", "&#096;");
}

function bool(x) {
  return x ? "YES" : "NO";
}

function formatDateTime(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleString();
}
