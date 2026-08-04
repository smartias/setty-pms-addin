// ─── CONFIG ───────────────────────────────────────────────────────────────────
const MSAL_CONFIG = {
  auth: {
    clientId: "8e5155fb-6221-4508-97ea-3661438c6688",
    authority: "https://login.microsoftonline.com/f374c024-71c2-48b6-8420-076fff97327c",
    redirectUri: "https://smartias.github.io/setty-pms-addin/word/taskpane.html",
  },
  cache: { cacheLocation: "localStorage" }
};
// Files.ReadWrite.All removed 2026-08-04: never admin-consented for this app, so requesting
// it at sign-in walled the whole add-in behind "Approval required" (user self-consent is off).
// The "save .docx to SharePoint" checkbox will 403 until it is re-wired to the narrow
// Sites.Selected scope (single-site grant on NYCProjects) instead of firm-wide "all files".
const GRAPH_SCOPES = ["User.Read", "Notes.ReadWrite"];
// Requested ON-DEMAND only (same pattern as the Outlook pane's shared-mail
// scope): kept out of the default sign-in so first-run consent is unchanged.
// The first click of "Send in Email" triggers a one-time per-user consent
// popup; afterwards the token is acquired silently like any other.
const MAIL_DRAFT_SCOPES = ["Mail.ReadWrite"];
const TEAMS_TEAM_ID = "a4c48361-7991-43db-af83-4c854918a760";
// SharePoint — same hardcoded drive ID the Outlook add-in uses (no admin consent needed).
const SP_DRIVE_ID = "b!ZARYqukTtE6K1Mpv9bngAehneskb-yNKopp1Ol1X1BBnJPKsNGM-TaGmbGiL3ZaU";
const SP_BASE_URL = "https://setty.sharepoint.com/sites/NYCProjects/Project%20Document%20Library";
const SUPABASE_URL  = "https://khxmgjilwhdguuepbhne.supabase.co";
const SUPABASE_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtoeG1namlsd2hkZ3V1ZXBiaG5lIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMwNjg2MDYsImV4cCI6MjA4ODY0NDYwNn0.vtHt2eydU2iQ426iYOzLrqpH2WLXdRnicq-3sNfoNq8";
const SB_HEADERS = { "apikey": SUPABASE_ANON, "Authorization": "Bearer " + SUPABASE_ANON };

// Shared suite sign-in (setty-auth.js, loaded cross-repo). Shim keeps the pane
// working if the script fails to load; popup flow because taskpanes can't
// full-page redirect. SB_HEADERS is read at call time, so mutating it covers
// every Supabase call below.
const _settyAuth = window.settyAuth || {
  init: async () => null, onChange() {}, token: () => SUPABASE_ANON,
  isSignedIn: () => false, mountPill() {}, signInPopup: async () => null,
};
const settyAuthReady = _settyAuth.init().catch(() => null).then(() => _syncSettyAuth());
function _syncSettyAuth() { SB_HEADERS.Authorization = "Bearer " + _settyAuth.token(); }
_settyAuth.onChange(_syncSettyAuth);
_settyAuth.mountPill({ label: "🔐 Sign in", onClick: () => _settyAuth.signInPopup() });
const PROJECTS_CACHE_KEY = "settyPmsWord:projectsCache";
const PROJECTS_CACHE_TTL_MS = 1000 * 60 * 60 * 6; // 6h
const LAST_ACCOUNT_STORAGE_KEY = "settyPmsWord:lastMsalAccountHomeId";
const TARGET_SECTION_NAME = "Documents"; // fixed section per product decision
const PMS_PROJECT_BASE_URL = "https://smartias.github.io/setty-pms/SettyPMS.html#project:";
// Setty document templates — same drive as everything else, so the existing
// Graph token covers them. Word templates only (.docx/.dotx); Excel ones in
// the folder are filtered out.
const TEMPLATES_PATH = "SAPX26XXX - NY 2026 Templates and Standards/01 📋 Project Management/Templates for MCP Connector";
const TEMPLATES_CACHE_KEY = "settyPmsWord:templatesCache";
const TEMPLATES_CACHE_TTL_MS = 1000 * 60 * 60 * 6; // 6h

// ─── DOCUMENT STATUS ─────────────────────────────────────────────────────────
const DOC_STATUS_OPTIONS = [
  { value: "draft", label: "Draft", dot: "#888",    hint: "Still editing" },
  { value: "final", label: "Final", dot: "#4ade80", hint: "PDF exported"  },
];
const DOC_DONE_KEY = "settyPms:doneEditing";

// ─── STATE ────────────────────────────────────────────────────────────────────
let msalApp = null;
let msalAccount = null;
let allProjects = [];
let allClients = []; // global directory — used by the POC picker
let selectedProject = null;
let docFilename = "";
let currentDocUrl = "";   // full URL of the open document ("" if unsaved); http(s) = cloud-hosted (AutoSaves)
let docFirstPageText = "";
let saveInFlight = false;
let docStatus = "draft"; // persisted in Office document settings
let dirtySinceFiled = false;     // true once the local copy is edited after filing — the filed SP copy is now stale
let isExporting = false;         // guard: ignore the add-in's own file/settings writes in the change handler
let changeHandlerRegistered = false;
let doneEditingList = []; // [{ name, email, ts }] — persisted in document settings
let draftSaved = false;     // true once a .docx draft has been uploaded to SharePoint
let draftFolderPath = "";   // drive-relative SP folder the draft was filed to
let draftBaseName = "";     // intended filename at last save (from Document Name field)
let draftFileName = "";     // actual filename SharePoint stored the draft under
let draftWebUrl = "";       // SharePoint webUrl of the filed .docx (for "open the copy")
// SharePoint folder picker state. spCurrentPath is drive-relative (e.g.
// "24-105 Acme HVAC/Documents"). Defaults to the project folder root when
// the SharePoint checkbox is first ticked.
let spCurrentPath = "";
let spProjectRootPath = "";
let _spFolders = [];   // folder names at the currently-browsed level (for the picker filter)
// Last successful PDF export, kept in memory so "Send in Email" can attach it
// without re-exporting. Session-only: the button hides on reload or when a new
// export starts, and reappears once that export succeeds.
let lastPdfBlob = null;
let lastPdfName = "";
let lastPdfWebUrl = "";
let emailInFlight = false;
// New-document-from-template state.
let _templates = [];        // [{ id, name }] from the templates folder
let newDocInFlight = false;

// ─── INIT ─────────────────────────────────────────────────────────────────────
// Surface any unhandled error directly into the pane so a startup crash never
// leaves the user staring at a blank black box. Without this, an early throw
// (before any view is shown) produces a silent failure with no UI feedback.
function showFatalError(label, err) {
  try {
    const msg = (err && (err.message || err.toString())) || "(no detail)";
    const stack = (err && err.stack) ? "\n\n" + err.stack : "";
    document.body.innerHTML =
      '<div style="padding:16px;font-family:Segoe UI,sans-serif;color:#f87171;background:#1a1a1a;height:100vh;font-size:12px;overflow:auto">'
      + '<div style="font-weight:700;margin-bottom:8px">⚠ Setty PMS Filer — startup error</div>'
      + '<div style="color:#fff;margin-bottom:8px">' + label + '</div>'
      + '<pre style="white-space:pre-wrap;color:#fbbf24;font-size:11px;margin:0">' + msg + stack + '</pre>'
      + '</div>';
  } catch {}
}
// Only trip the fatal-error panel for errors that look like they came from
// our own code, not from Office.js telemetry noise in a plain browser context.
function looksLikeOurError(err) {
  const text = String((err && (err.message || err.stack)) || err || "");
  if (/Office\.js|outside of Office client|telemetryservice|message channel closed/i.test(text)) return false;
  return true;
}
window.addEventListener("error", (e) => {
  if (looksLikeOurError(e.error || e.message)) showFatalError("Uncaught error", e.error || e.message);
});
window.addEventListener("unhandledrejection", (e) => {
  if (looksLikeOurError(e.reason)) showFatalError("Unhandled promise rejection", e.reason);
});

Office.onReady(async (info) => {
  // Host check removed — manifest already restricts this add-in to Word, and
  // the previous check could swap the body for an invisible message in some
  // hosts where info.host doesn't strict-equal Office.HostType.Word.
  console.log("[Setty PMS Filer] Office.onReady fired, info.host =", info && info.host);
  try {
    if (typeof msal === "undefined") {
      throw new Error("MSAL bundle did not load (../msal-browser.min.js). Check that the file is published next to /word/ on GitHub Pages.");
    }
    msalApp = new msal.PublicClientApplication(MSAL_CONFIG);
    await msalApp.initialize();
    const accounts = msalApp.getAllAccounts();
    if (accounts.length > 0) {
      const lastId = localStorage.getItem(LAST_ACCOUNT_STORAGE_KEY);
      msalAccount = accounts.find(a => a.homeAccountId === lastId) || accounts[0];
      msalApp.setActiveAccount(msalAccount);
      await onSignedIn();
    } else {
      showView("signInView");
    }
    setupListeners();
  } catch (e) {
    console.error("[Setty PMS Filer] startup failed:", e);
    showView("signInView");
    setStatus("signInStatus", "error", "Startup error: " + e.message);
  }
});

function setupListeners() {
  document.getElementById("signInBtn").onclick        = doSignIn;
  document.getElementById("signOutBtn").onclick       = doSignOut;
  document.getElementById("saveOneNoteBtn").onclick   = doSaveToOneNote;
  document.getElementById("saveSpDraftBtn").onclick   = doSaveDraft;
  document.getElementById("savePdfBtn").onclick       = doSavePdf;
  document.getElementById("sendEmailBtn").onclick     = doSendPdfEmail;
  document.getElementById("insertNameBtn").onclick    = () => doInsertField("name");
  document.getElementById("insertNumberBtn").onclick  = () => doInsertField("number");
  document.getElementById("insertClientBtn").onclick  = () => doInsertField("client");
  document.getElementById("insertPocToggleBtn").onclick = togglePocPicker;
  document.getElementById("searchInput").addEventListener("input", () => renderProjectList());
  document.getElementById("pocSearch").addEventListener("input", renderPocList);
  document.getElementById("pillChangeLink").onclick   = expandProjectPicker;
  document.getElementById("openPmsBtn").onclick       = openSelectedProjectInPms;
  document.getElementById("openSpFolderBtn").onclick  = openSelectedProjectSpFolder;
  document.getElementById("destChangeBtn").onclick    = () => toggleOpt("spFolderEdit");
  document.getElementById("toggleDoneBtn").onclick    = toggleCurrentUserDone;
  document.getElementById("newVersionBtn").onclick    = doStartNewVersion;
  document.getElementById("newDocToggleBtn").onclick  = toggleNewDocPicker;
  document.getElementById("createNewDocBtn").onclick  = doCreateFromTemplate;
  document.getElementById("templateSelect").onchange  = onTemplatePicked;
  document.getElementById("newDocNameInput").oninput  = () => { updateNewDocPreview(); updateNewDocButtons(); };
  document.getElementById("titleInput").addEventListener("input", updateSavePreview);
  const spSearch = document.getElementById("spFolderSearch");
  if (spSearch) spSearch.addEventListener("input", () => renderSpFolderRows(spSearch.value));
  document.getElementById("conflictReplaceBtn").onclick = () => _conflictResolver && _conflictResolver("replace");
  document.getElementById("conflictRenameBtn").onclick  = () => _conflictResolver && _conflictResolver("rename");
  document.getElementById("conflictCancelBtn").onclick  = () => _conflictResolver && _conflictResolver("cancel");
}

// ─── NAME-CONFLICT GUARD ───────────────────────────────────────────────────────
// True if a file already lives at targetPath/filename in the SP drive. A 404
// means the name is free; any error falls back to false so a check failure can
// never block a save.
async function spFileExists(targetPath, filename) {
  try {
    const token = await getToken();
    const safeName = filename.replace(/[\\/:*?"<>|]/g, "_");
    const url = "https://graph.microsoft.com/v1.0/drives/" + SP_DRIVE_ID
      + "/root:/" + encodeDrivePath(targetPath + "/" + safeName);
    const res = await fetch(url, { headers: { "Authorization": "Bearer " + token } });
    return res.ok;
  } catch {
    return false;
  }
}

// Inline prompt shown when a first-time save would collide with an existing
// file. Resolves to "replace" | "rename" | "cancel". Promise-based so doSaveDraft
// can simply `await` the user's choice.
let _conflictResolver = null;
function askNameConflict(filename) {
  return new Promise((resolve) => {
    const card = document.getElementById("conflictCard");
    const msg  = document.getElementById("conflictMsg");
    const renameBtn = document.getElementById("conflictRenameBtn");
    if (!card) { resolve("rename"); return; }   // no UI — preserve old behavior
    msg.innerHTML = `<b>${escapeHtml(filename)}</b> already exists in this folder. What would you like to do?`;
    if (renameBtn) renameBtn.textContent = `Save as “${filename.replace(/\.docx$/i, "")} (1)”`;
    card.style.display = "block";
    _conflictResolver = (choice) => {
      card.style.display = "none";
      _conflictResolver = null;
      resolve(choice);
    };
  });
}

// Collapse the search + list into the compact pill once a project is chosen.
function collapseProjectPickerToPill() {
  if (!selectedProject) return;
  document.getElementById("projectSearchWrap").style.display  = "none";
  document.getElementById("projectPill").style.display        = "flex";
  document.getElementById("pillProjectNumber").textContent    = selectedProject.projectNumber || "";
  document.getElementById("pillProjectName").textContent      = selectedProject.name || "";
  // Quick links
  const ql = document.getElementById("projectQuickLinks");
  if (ql) ql.style.display = "grid";
  updateQuickLinks();
  // Destination card — only when project has a SharePoint folder
  updateDestCard();
  // Title field for OneNote — always visible once a project is picked
  const tf = document.getElementById("titleFieldWrap");
  if (tf) tf.style.display = "block";
}

function expandProjectPicker() {
  document.getElementById("projectSearchWrap").style.display = "block";
  document.getElementById("projectPill").style.display       = "none";
  const ql = document.getElementById("projectQuickLinks");
  if (ql) ql.style.display = "none";
  const s = document.getElementById("searchInput");
  s.value = "";
  s.focus();
  renderProjectList();
}

function updateQuickLinks() {
  const pmsBtn = document.getElementById("openPmsBtn");
  const spBtn  = document.getElementById("openSpFolderBtn");
  if (pmsBtn) pmsBtn.disabled = !selectedProject;
  if (spBtn)  spBtn.disabled  = !(selectedProject?.projectFolderUrl);
}

function updateDestCard() {
  const card = document.getElementById("destCard");
  if (!card) return;
  // The destination card shows whenever a project is selected — saving is no
  // longer gated on the project having a pre-linked folder; the user can browse
  // the whole library and pick any folder (defaulting into the project's own).
  card.style.display = selectedProject ? "block" : "none";
  if (selectedProject) {
    const display = document.getElementById("destFolderDisplay");
    if (display) display.textContent = prettySpPath(spCurrentPath);
  }
  updateSavePreview();
  updateFiledCard();
}

// A drive-relative path ("a/b/c") shown as readable breadcrumbs; "" = library root.
function prettySpPath(path) {
  const parts = path ? path.split("/").filter(Boolean) : [];
  return parts.length ? parts.join(" / ") : "Library root";
}

// Live "this is what gets saved" line under the File name field — combines the
// typed name, the .docx extension, and the destination folder so there's no
// guessing what the save will produce or where it lands.
function updateSavePreview() {
  const el = document.getElementById("savePreview");
  if (!el) return;
  const base = (document.getElementById("titleInput")?.value || "").trim().replace(/\.(docx|pdf)$/i, "");
  if (!base || !selectedProject) { el.style.display = "none"; return; }
  el.style.display = "block";
  el.innerHTML = `Saves as <b>${escapeHtml(base)}.docx</b> in 📂 ${escapeHtml(prettySpPath(spCurrentPath))}`;
}

// Confirmation card shown once a .docx draft has been filed to SharePoint.
// Crucially it surfaces the "Open the SharePoint copy" action — the only way
// to get true AutoSave, since this open document is a detached local copy.
function updateFiledCard() {
  const card = document.getElementById("draftFiledCard");
  if (!card) return;
  if (!draftSaved) { card.style.display = "none"; return; }
  card.style.display = "block";
  const nameEl = document.getElementById("filedName");
  if (nameEl) {
    nameEl.textContent = (draftFileName || "(filed)") + "  ·  " + prettySpPath(draftFolderPath);
  }

  const headEl = document.getElementById("filedHead");
  const hintEl = document.getElementById("filedHint");
  const openBtn = document.getElementById("openSpCopyBtn");
  const inner   = document.getElementById("filedCardInner");

  // If the open document is itself cloud-hosted (http/https URL), it's the
  // SharePoint copy — edits AutoSave and there's nothing to reopen. Show a calm
  // confirmation and hide the open action, so people aren't told to "open the
  // SharePoint copy" when they're already in it.
  const inCloudCopy = isEditingFiledCloudCopy();
  if (inCloudCopy) {
    if (inner) inner.classList.remove("is-dirty");
    if (headEl) headEl.textContent = "✓ Filed — AutoSave is on";
    if (hintEl) hintEl.textContent = "You're editing the SharePoint copy. Changes save automatically.";
    if (openBtn) openBtn.style.display = "none";
    return;
  }

  // Detached local copy, two sub-states:
  //  • dirty — edited since filing, so the SharePoint copy is now stale (amber).
  //  • clean — just filed; frame it as one step from done and hand them off.
  if (inner) inner.classList.toggle("is-dirty", dirtySinceFiled);
  if (dirtySinceFiled) {
    if (headEl) headEl.textContent = "● Unsynced edits — changes since you filed";
    if (hintEl) hintEl.textContent = "This open file is a copy on your computer. Click “Update Draft” to push your edits to SharePoint, or switch to the SharePoint copy (it saves on its own).";
  } else {
    if (headEl) headEl.textContent = "✓ Filed to SharePoint — one step left";
    if (hintEl) hintEl.textContent = "You're editing a copy on your computer. Continue in the SharePoint copy so your edits save automatically.";
  }
  // Primary hand-off action — labeled to lead people to the synced copy, not
  // just "open". Only works when we know the SP URL (a copy reopened without one
  // hides the button; the user relies on Update Draft instead).
  if (openBtn) {
    openBtn.style.display = draftWebUrl ? "block" : "none";
    openBtn.textContent = "Continue in the SharePoint copy →";
    openBtn.onclick = (e) => { e.preventDefault(); openSharePointCopy(); };
    openBtn.removeAttribute("href");
  }
}

// Opens the filed SharePoint copy so further edits AutoSave. Opens the file's
// SharePoint URL directly: in the browser that's Word on the web (which
// AutoSaves), and if the tenant/library is set to open in the client app,
// SharePoint hands off to desktop Word automatically.
//
// We deliberately do NOT use the ms-word:ofe|u| protocol here. It throws
// "Office doesn't recognize the command it was given" on real SharePoint paths
// (the document library and project folders contain spaces and other characters
// the protocol parser mishandles regardless of encoding) — that error is the
// exact bug this replaces.
function openSharePointCopy() {
  if (!draftWebUrl) return;
  // The honest ceiling for the two-window confusion: Office.js can't close this
  // local copy for the user, so tell them it's safe to abandon once Word opens.
  setStatus("spStatus", "info", "Opening in Word… once it's up, you can close this window safely.");
  openExternalUrl(draftWebUrl);
}

// Opens a URL in the system browser. openBrowserWindow is the reliable path
// from inside an Office taskpane (WebView2 may swallow window.open).
function openExternalUrl(url) {
  if (!url) return;
  try {
    if (Office?.context?.ui?.openBrowserWindow) {
      Office.context.ui.openBrowserWindow(url);
      return;
    }
  } catch (e) {
    console.warn("openBrowserWindow failed, falling back to window.open:", e);
  }
  window.open(url, "_blank");
}

function openSelectedProjectInPms() {
  if (!selectedProject) return;
  window.open(PMS_PROJECT_BASE_URL + selectedProject.id, "_blank");
}

function openSelectedProjectSpFolder() {
  if (!selectedProject?.projectFolderUrl) return;
  window.open(selectedProject.projectFolderUrl, "_blank");
}

function toggleOpt(editId, focusId) {
  const el = document.getElementById(editId);
  el.classList.toggle("active");
  if (el.classList.contains("active") && focusId) {
    document.getElementById(focusId).focus();
  }
}

// Refresh SP picker for the currently-selected project (or empty it if none).
// Called whenever selectedProject changes.
async function refreshSpPickerForProject() {
  const folders = document.getElementById("spFolders");
  const crumbs  = document.getElementById("spBreadcrumbs");
  if (!selectedProject) {
    if (crumbs)  crumbs.textContent  = "Pick a project to enable folder browsing.";
    if (folders) folders.innerHTML   = "";
    updateDestCard();
    return;
  }
  // Default landing folder: the project's own folder when it has one, else the
  // library root. This is just the starting point — the breadcrumb root is the
  // library, so the user can browse to ANY folder from here.
  spProjectRootPath = spDrivePath(selectedProject.projectFolderUrl || "") || "";
  spCurrentPath     = spProjectRootPath;
  // Formal documents default into the project's "Project Management" subfolder
  // when one exists (names vary, e.g. "01 📋 Project Management" — match by
  // contains). Falls back to the project root when the probe finds nothing.
  if (spProjectRootPath && !(draftSaved && draftFolderPath)) {
    try {
      const token = await getToken();
      const kids = await fetchAllChildFolders(token, spProjectRootPath);
      const pm = kids.find(n => n.toLowerCase().includes("project management"));
      if (pm) spCurrentPath = spProjectRootPath + "/" + pm;
    } catch (e) {
      console.warn("PM-folder probe failed, defaulting to project root:", e.message);
    }
  }
  // Once a draft has been filed, default to the folder it went to so the next
  // save lands next to it.
  if (draftSaved && draftFolderPath) spCurrentPath = draftFolderPath;
  updateDestCard();
  await renderSpPicker();
}

async function renderSpPicker() {
  const crumbs = document.getElementById("spBreadcrumbs");
  const folders = document.getElementById("spFolders");
  // Breadcrumbs are anchored at the LIBRARY ROOT so every folder is reachable.
  const parts = spCurrentPath ? spCurrentPath.split("/").filter(Boolean) : [];
  const crumbHtml = ['<span data-depth="0">📁 Library root</span>']
    .concat(parts.map((p, i) => ` / <span data-depth="${i + 1}">${escapeHtml(p)}</span>`))
    .join("");
  crumbs.innerHTML = "Save to: " + crumbHtml;
  // Mirror the current folder into the destination card display.
  const destDisplay = document.getElementById("destFolderDisplay");
  if (destDisplay) destDisplay.textContent = parts.length ? parts.join(" / ") : "Library root";
  updateSavePreview();
  updateNewDocPreview();   // the new-from-template preview shows the same folder
  crumbs.querySelectorAll("span").forEach(s => {
    s.onclick = async () => {
      const depth = parseInt(s.getAttribute("data-depth"), 10);
      spCurrentPath = parts.slice(0, depth).join("/");   // depth 0 -> "" (library root)
      await renderSpPicker();
    };
  });
  const search = document.getElementById("spFolderSearch");
  if (search) { search.value = ""; search.style.display = "none"; }   // reset filter on every level change
  folders.innerHTML = '<div class="sp-loading">Loading folders…</div>';
  try {
    const token = await getToken();
    _spFolders = await fetchAllChildFolders(token, spCurrentPath);
    // Only surface the filter box when the list is long enough to need it.
    if (search) search.style.display = _spFolders.length > 8 ? "block" : "none";
    renderSpFolderRows("");
  } catch (e) {
    folders.innerHTML = '<div class="sp-empty">Error: ' + escapeHtml(e.message) + '</div>';
  }
}

// Fetches ALL immediate subfolders of a drive-relative path, following
// pagination so a level with >200 folders (e.g. the library root) isn't
// silently truncated. Capped at ~3000 to stay sane. Returns sorted names.
async function fetchAllChildFolders(token, path) {
  const driveBase = "https://graph.microsoft.com/v1.0/drives/" + SP_DRIVE_ID;
  let url = (path
    ? driveBase + "/root:/" + encodeDrivePath(path) + ":/children"
    : driveBase + "/root/children")
    + "?$select=name,folder&$top=200";
  const names = [];
  let guard = 0;
  while (url && guard < 15) {
    const res = await fetch(url, { headers: { "Authorization": "Bearer " + token } });
    if (!res.ok) throw new Error("list folders " + res.status);
    const data = await res.json();
    for (const it of (data.value || [])) if (it.folder) names.push(it.name);
    url = data["@odata.nextLink"] || null;
    guard++;
  }
  return names.sort((a, b) => a.localeCompare(b));
}

// Renders the (optionally filtered) folder rows from _spFolders into #spFolders.
// Separate from the fetch so typing in the filter re-renders without re-hitting
// Graph.
function renderSpFolderRows(filter) {
  const folders = document.getElementById("spFolders");
  if (!folders) return;
  if (!_spFolders.length) {
    folders.innerHTML = '<div class="sp-empty">No subfolders here — saves will land in this folder.</div>';
    return;
  }
  const q = (filter || "").trim().toLowerCase();
  const list = q ? _spFolders.filter(n => n.toLowerCase().includes(q)) : _spFolders;
  if (!list.length) {
    folders.innerHTML = '<div class="sp-empty">No folders match “' + escapeHtml(filter) + '”.</div>';
    return;
  }
  folders.innerHTML = list.map(name =>
    `<div class="sp-folder-row" data-name="${escapeHtml(name)}">📁 ${escapeHtml(name)}</div>`
  ).join("");
  folders.querySelectorAll(".sp-folder-row").forEach(row => {
    row.onclick = async () => {
      const name = row.getAttribute("data-name");
      spCurrentPath = spCurrentPath ? spCurrentPath + "/" + name : name;
      await renderSpPicker();
    };
  });
}

function showView(id) {
  document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
  document.getElementById(id).classList.add("active");
}

function setStatus(elId, kind, msg) {
  const el = document.getElementById(elId);
  if (!el) return;
  el.className = "status " + kind;
  el.textContent = msg;
}

// ─── AUTH ─────────────────────────────────────────────────────────────────────
async function doSignIn() {
  try {
    setStatus("signInStatus", "info", "Signing in…");
    const result = await msalApp.loginPopup({ scopes: GRAPH_SCOPES });
    msalAccount = result.account;
    msalApp.setActiveAccount(msalAccount);
    localStorage.setItem(LAST_ACCOUNT_STORAGE_KEY, msalAccount.homeAccountId);
    await onSignedIn();
  } catch (e) {
    setStatus("signInStatus", "error", "Sign-in failed: " + e.message);
  }
}

async function doSignOut() {
  try {
    await msalApp.logoutPopup({ account: msalAccount });
  } catch {}
  msalAccount = null;
  localStorage.removeItem(LAST_ACCOUNT_STORAGE_KEY);
  showView("signInView");
}

async function getToken() {
  const result = await msalApp.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: msalAccount });
  return result.accessToken;
}

// Mail token is separate from getToken(): Mail.ReadWrite is requested on
// demand, so the silent call fails the first time (no consent yet) — fall
// back to a consent popup once; every later call resolves silently.
async function getMailToken() {
  try {
    const r = await msalApp.acquireTokenSilent({ scopes: MAIL_DRAFT_SCOPES, account: msalAccount });
    return r.accessToken;
  } catch {
    const r = await msalApp.acquireTokenPopup({ scopes: MAIL_DRAFT_SCOPES, account: msalAccount });
    return r.accessToken;
  }
}

async function onSignedIn() {
  showView("mainView");
  await Promise.all([loadProjects(), loadDocumentContext(), loadDocStatus(), loadDoneEditing()]);
  loadSelectedProject();
  loadDraftSaved();
  renderProjectList();
  // A project tagged in a previous session wins over a fresh auto-suggest.
  if (selectedProject) {
    updateSaveButtons();
    refreshSpPickerForProject();
    collapseProjectPickerToPill();
  } else {
    applyAutoSuggest();
  }
  renderStatusBar();
  renderDoneEditing();
  registerChangeWatcher();
  fxWelcome(fxFirstName());   // header-logo wiggle + greeting toast
}

// Watches for edits to the open document so we can warn when a filed draft has
// drifted out of sync. The common API has no "content changed" event, so we use
// DocumentSelectionChanged as a proxy for "the user is working in the doc" —
// good enough for a STICKY dirty flag that's only ever cleared by a successful
// "Update Draft" (never auto-cleared on a guess).
function registerChangeWatcher() {
  if (changeHandlerRegistered) return;
  try {
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      onDocMaybeChanged,
      (r) => { if (r.status === Office.AsyncResultStatus.Succeeded) changeHandlerRegistered = true; }
    );
  } catch (e) {
    console.warn("Could not register change watcher:", e.message);
  }
}

function onDocMaybeChanged() {
  if (isExporting) return;                                   // our own file/settings writes — not a user edit
  if (dirtySinceFiled || !draftSaved) return;                // already flagged, or nothing filed yet to drift from
  if (isEditingFiledCloudCopy()) return;                     // cloud copy AutoSaves — never "unsynced"
  dirtySinceFiled = true;
  updateFiledCard();
  updateSaveButtons();   // reflect "● edits pending" on the update button
}

// True when the OPEN document is the same file the pane last filed — only then
// does "the cloud copy AutoSaves" apply. After "Start a new version" (or a
// rename-and-refile), the filed draft is a DIFFERENT file from the one open in
// Word, so edits here no longer land in the filed copy.
function isEditingFiledCloudCopy() {
  if (!/^https?:\/\//i.test(currentDocUrl || "")) return false;
  if (!draftFileName) return true;   // nothing filed to compare against — old behavior
  try {
    const urlName = decodeURIComponent(currentDocUrl).split(/[\\/]/).pop() || "";
    return urlName.toLowerCase() === draftFileName.toLowerCase();
  } catch {
    return true;
  }
}

// ─── PROJECTS ─────────────────────────────────────────────────────────────────
async function loadProjects() {
  await settyAuthReady;   // session header set before the first Supabase read
  // Cache-first hydrate
  try {
    const raw = localStorage.getItem(PROJECTS_CACHE_KEY);
    if (raw) {
      const cached = JSON.parse(raw);
      if (cached.savedAt && (Date.now() - cached.savedAt) < PROJECTS_CACHE_TTL_MS) {
        allProjects = cached.projects || [];
        allClients  = cached.clients  || [];
      }
    }
  } catch {}

  try {
    const [pRes, cRes] = await Promise.all([
      fetch(SUPABASE_URL + "/rest/v1/pms_projects?select=id,project,version", { headers: SB_HEADERS }),
      fetch(SUPABASE_URL + "/rest/v1/pms_clients?select=client", { headers: SB_HEADERS }),
    ]);
    if (pRes.ok) {
      const pRows = await pRes.json();
      if (pRows && pRows.length > 0) {
        allProjects = pRows.map(r => r.project).filter(p => p && !p.archived);
        if (cRes.ok) {
          const cRows = await cRes.json();
          allClients = (cRows || []).map(r => r.client).filter(Boolean);
        }
        try {
          localStorage.setItem(PROJECTS_CACHE_KEY, JSON.stringify({
            projects: allProjects, clients: allClients, savedAt: Date.now()
          }));
        } catch {}
        return;
      }
    }
  } catch (e) {
    console.warn("V2 load failed, trying legacy:", e.message);
  }
  // Legacy fallback
  try {
    const res = await fetch(SUPABASE_URL + "/rest/v1/pms_data?id=eq.singleton&select=projects,clients", { headers: SB_HEADERS });
    const rows = await res.json();
    if (rows?.[0]?.projects) {
      allProjects = rows[0].projects.filter(p => !p.archived);
      allClients  = rows[0].clients || [];
    }
  } catch (e) {
    console.error("Failed to load projects:", e);
  }
}

// ─── DOCUMENT SETTINGS PERSISTENCE ─────────────────────────────────────────────
// The tagged project and document status are stored in the document's own
// settings so they travel inside the .docx and are restored when the file is
// reopened. saveAsync must run before the add-in reads the file out for upload,
// otherwise the uploaded copy is missing the tags.
function persistDocSettings() {
  try {
    Office.context.document.settings.set("settyPms:projectId", selectedProject?.id || "");
    Office.context.document.settings.set("settyPms:docStatus", docStatus);
    Office.context.document.settings.set("settyPms:draftSaved", draftSaved);
    Office.context.document.settings.set("settyPms:draftFolderPath", draftFolderPath);
    Office.context.document.settings.set("settyPms:draftBaseName", draftBaseName);
    Office.context.document.settings.set("settyPms:draftFileName", draftFileName);
    Office.context.document.settings.set("settyPms:draftWebUrl", draftWebUrl);
  } catch (e) {
    console.warn("Could not stage document settings:", e.message);
  }
}

function flushSettings() {
  return new Promise((resolve) => {
    try {
      Office.context.document.settings.saveAsync(r => {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          console.warn("settings.saveAsync failed:", r.error?.message);
        }
        resolve();
      });
    } catch (e) {
      console.warn("settings.saveAsync threw:", e.message);
      resolve();
    }
  });
}

// Restore the project tagged on this document in an earlier session. Must run
// after loadProjects() so allProjects is populated.
function loadSelectedProject() {
  try {
    const id = Office.context.document.settings.get("settyPms:projectId");
    if (id) selectedProject = allProjects.find(p => p.id === id) || null;
  } catch (e) {
    console.warn("Could not read saved project:", e.message);
  }
}

// Restore whether a .docx draft has already been filed to SharePoint.
function loadDraftSaved() {
  try {
    draftSaved      = Office.context.document.settings.get("settyPms:draftSaved") === true;
    draftFolderPath = Office.context.document.settings.get("settyPms:draftFolderPath") || "";
    draftBaseName   = Office.context.document.settings.get("settyPms:draftBaseName") || "";
    draftFileName   = Office.context.document.settings.get("settyPms:draftFileName") || "";
    draftWebUrl     = Office.context.document.settings.get("settyPms:draftWebUrl") || "";
  } catch (e) {
    draftSaved = false;
    draftFolderPath = draftBaseName = draftFileName = draftWebUrl = "";
  }
}

// ─── DOCUMENT CONTEXT ─────────────────────────────────────────────────────────

// Office.js exposes the document URL only via this async, callback-based API —
// there is no synchronous `Office.context.document.url`. Returns "" for an
// unsaved document.
function getDocumentUrl() {
  return new Promise((resolve) => {
    Office.context.document.getFilePropertiesAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value.url || "");
      } else {
        console.warn("getFilePropertiesAsync failed:", result.error?.message);
        resolve("");
      }
    });
  });
}

// Default name for a document that has never been saved (no filename yet).
// Date-stamped so unsaved docs don't all collide on the same OneNote page
// title; the user can still type over it in the Document Name field.
function deriveUnsavedTitle() {
  const d = new Date();
  const stamp = d.getFullYear() + "-"
    + String(d.getMonth() + 1).padStart(2, "0") + "-"
    + String(d.getDate()).padStart(2, "0");
  return "Document " + stamp;
}

// Turn a document URL into a readable folder path. SharePoint-hosted docs are
// shown project-library-relative; anything else falls back to the decoded
// directory portion of the URL.
function prettyDocLocation(url) {
  const noFile = String(url).replace(/[\\/][^\\/]*$/, "");
  if (noFile.startsWith(SP_BASE_URL)) {
    const rel = decodeURIComponent(noFile.slice(SP_BASE_URL.length).replace(/^\/+/, ""));
    return rel || "(library root)";
  }
  try { return decodeURIComponent(noFile); } catch { return noFile; }
}

async function loadDocumentContext() {
  // Filename — resolved via the async Office API (see getDocumentUrl above).
  const docUrl = await getDocumentUrl();
  currentDocUrl = docUrl;
  docFilename = docUrl.split(/[\\/]/).pop() || "";
  document.getElementById("docFilename").textContent = docFilename || "(unsaved document)";
  // Show where the document currently lives so a reopened file's location is
  // visible at a glance.
  const locEl = document.getElementById("docLocation");
  if (locEl) {
    locEl.textContent = docUrl ? "📂 " + prettyDocLocation(docUrl) : "Not yet saved";
    locEl.style.display = "block";
  }
  // Pre-fill the Document Name field with the filename minus extension, or the
  // unsaved fallback. This value drives the saved filename and OneNote title.
  const defaultTitle = docFilename
    ? docFilename.replace(/\.[^.]+$/, "")
    : deriveUnsavedTitle();
  const titleInput = document.getElementById("titleInput");
  if (titleInput) titleInput.value = defaultTitle;

  // First-page text via Word.run — used by the auto-suggest scorer.
  // Capped at the first ~2000 chars; project numbers/names typically appear
  // on cover pages, headers, or in the first paragraph.
  try {
    await Word.run(async (ctx) => {
      const body = ctx.document.body;
      body.load("text");
      await ctx.sync();
      docFirstPageText = (body.text || "").slice(0, 2000);
    });
  } catch (e) {
    console.warn("Could not read document body:", e.message);
    docFirstPageText = "";
  }
}

// ─── AUTO-SUGGEST ─────────────────────────────────────────────────────────────
// TODO(user): Implement the scorer. Given a project, the document filename, and
// the first ~2000 chars of body text, return a numeric score where higher =
// more likely the right project. Return 0 to mean "no match — don't suggest".
//
// Available project fields: projectNumber, name, clientName (and others —
// inspect a sample with `console.log(allProjects[0])` if useful).
//
// Things to consider:
//   • Project number is the strongest signal (exact substring match in
//     filename or body should dominate). What if the number appears partial
//     (e.g. "24-105" matched against filename "24105_charter")?
//   • Project name match in filename is medium-strong; in body text, weaker
//     (lots of false positives from generic words).
//   • Client name match is the weakest — useful as a tiebreaker.
//   • Filename matches should outweigh body-text matches (filenames are
//     intentional; body text often mentions other projects in passing).
//
// Return 0 for "skip"; the UI auto-selects only if top score >= 5 (tune freely).
function scoreProjectMatch(project, filename, bodyText) {
  const fn = filename.toLowerCase();
  const tx = bodyText.toLowerCase();
  const num = (project.projectNumber || "").toLowerCase();
  const name = (project.name || "").toLowerCase();
  const client = (project.clientName || "").toLowerCase();
  let score = 0;
  if (num && fn.includes(num))                          score += 10;
  if (num && tx.includes(num))                          score += 5;
  if (name && name.length > 4 && fn.includes(name))     score += 6;
  if (name && name.length > 4 && tx.includes(name))     score += 2;
  if (client && client.length > 4 && fn.includes(client)) score += 3;
  return score;
}

function applyAutoSuggest() {
  if (!allProjects.length) return;
  const scored = allProjects
    .map(p => ({ p, score: scoreProjectMatch(p, docFilename, docFirstPageText) }))
    .filter(s => s.score > 0)
    .sort((a, b) => b.score - a.score);
  if (scored.length && scored[0].score >= 5) {
    selectedProject = scored[0].p;
    updateSaveButtons();
    refreshSpPickerForProject();
    collapseProjectPickerToPill();
    fxSuggestPill();   // celebrate the correct guess
  }
}

// ─── PROJECT LIST UI ──────────────────────────────────────────────────────────
function renderProjectList(suggestedIds = []) {
  const q = (document.getElementById("searchInput").value || "").trim().toLowerCase();
  const list = document.getElementById("projectList");
  const suggestedSet = new Set(suggestedIds);
  const filtered = allProjects.filter(p => {
    if (!q) return true;
    return (p.projectNumber || "").toLowerCase().includes(q)
        || (p.name || "").toLowerCase().includes(q)
        || (p.clientName || "").toLowerCase().includes(q);
  });
  // Suggested projects float to top, then alphabetical
  filtered.sort((a, b) => {
    const aS = suggestedSet.has(a.id) ? 0 : 1;
    const bS = suggestedSet.has(b.id) ? 0 : 1;
    if (aS !== bS) return aS - bS;
    return (a.projectNumber || "").localeCompare(b.projectNumber || "");
  });
  list.innerHTML = filtered.slice(0, 50).map(p => {
    const sel = selectedProject?.id === p.id ? " selected" : "";
    const sug = suggestedSet.has(p.id) ? " suggested" : "";
    const badge = suggestedSet.has(p.id) ? '<span class="suggested-badge">SUGGESTED</span>' : "";
    return `<div class="project-row${sel}${sug}" data-id="${p.id}">
      <div class="pn">${escapeHtml(p.projectNumber || "")}${badge}</div>
      <div class="nm">${escapeHtml(p.name || "")}</div>
      <div class="cl">${escapeHtml(p.clientName || "")}</div>
    </div>`;
  }).join("");
  list.querySelectorAll(".project-row").forEach(row => {
    row.onclick = async () => {
      const id = row.getAttribute("data-id");
      selectedProject = allProjects.find(p => p.id === id) || null;
      updateSaveButtons();
      refreshSpPickerForProject();
      collapseProjectPickerToPill();
      // Reset POC picker — its list depends on the selected project
      document.getElementById("pocPicker").style.display = "none";
      document.getElementById("pocSearch").value = "";
      setStatus("insertStatus", "info", "");
      document.getElementById("insertStatus").className = "status";
      // Close the folder picker if it was open for the previous project
      document.getElementById("spFolderEdit").classList.remove("active");
      // Tag the document with the chosen project so it survives a reopen.
      persistDocSettings();
      await flushSettings();
    };
  });
}

function updateSaveButtons() {
  const oneBtn   = document.getElementById("saveOneNoteBtn");
  const draftBtn = document.getElementById("saveSpDraftBtn");
  const pdfBtn   = document.getElementById("savePdfBtn");
  const projTag  = selectedProject ? ` → ${selectedProject.projectNumber || selectedProject.name}` : "";
  // Saving only needs a selected project now — the destination is the folder
  // chosen in the picker (defaults to the project folder, or the library root),
  // so a project without a pre-linked folder can still save.
  const spReady  = !!selectedProject;

  // Relabel via the .brand-label span only — writing textContent on the button
  // itself would wipe the icon tile inside it.
  oneBtn.disabled = !selectedProject || saveInFlight;
  document.getElementById("saveOneNoteLabel").textContent = "Save to OneNote" + projTag;

  // When the open document IS the filed SharePoint copy, an "update" would PUT
  // to the very file Word has locked — SharePoint rejects it with 423. AutoSave
  // already covers this case, so disable the button rather than offer a doomed
  // upload.
  const inCloudCopy = draftSaved && isEditingFiledCloudCopy();
  draftBtn.disabled = !spReady || saveInFlight || inCloudCopy;
  // Plain-language states. No project-tag suffix here — the labels are long and
  // the pill above already shows the project. The "●" on a dirty update ties to
  // the amber filed card's "● Unsynced edits" so they read as the same signal.
  document.getElementById("saveSpDraftLabel").textContent =
    inCloudCopy        ? "Filed — AutoSave is on"
    : !draftSaved      ? "File to SharePoint"
    : dirtySinceFiled  ? "Update filed copy ●"
    :                    "Update filed copy";
  draftBtn.style.display = "";

  pdfBtn.disabled = !spReady || saveInFlight;
  document.getElementById("savePdfLabel").textContent =
    "Export PDF to SharePoint" + (spReady ? projTag : "");

  const nvBtn = document.getElementById("newVersionBtn");
  if (nvBtn) nvBtn.disabled = !spReady || saveInFlight;

  updateNewDocButtons();

  const hasProject = !!selectedProject;
  document.getElementById("insertNameBtn").disabled      = !hasProject;
  document.getElementById("insertNumberBtn").disabled    = !hasProject;
  document.getElementById("insertClientBtn").disabled    = !hasProject;
  document.getElementById("insertPocToggleBtn").disabled = !hasProject;
}

// ─── SLICE DECODE ─────────────────────────────────────────────────────────────
// Office.js returns slice data as Base64 strings on Office on the web and some
// desktop builds. Passing those strings directly into new Blob() produces a
// text file, not binary — Word then rejects the upload as unreadable.
function decodeSliceData(data) {
  if (typeof data === "string") {
    const binary = atob(data);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return bytes;
  }
  return data instanceof Uint8Array ? data : new Uint8Array(data);
}

// ─── DOCX EXPORT ──────────────────────────────────────────────────────────────
// Office.FileType.Compressed returns the doc in its native format (a .docx is
// already a zip). Same slice-assembly pattern as the PDF export.
async function getDocumentAsDocxBlob() {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        return reject(new Error("getFileAsync (docx) failed: " + (result.error?.message || "unknown")));
      }
      const file = result.value;
      const sliceCount = file.sliceCount;
      const slices = new Array(sliceCount);
      let received = 0;
      for (let i = 0; i < sliceCount; i++) {
        const idx = i;
        file.getSliceAsync(idx, (sliceRes) => {
          if (sliceRes.status !== Office.AsyncResultStatus.Succeeded) {
            file.closeAsync();
            return reject(new Error("getSliceAsync (docx) failed at " + idx));
          }
          slices[idx] = decodeSliceData(sliceRes.value.data);
          received++;
          if (received === sliceCount) {
            file.closeAsync();
            resolve(new Blob(slices, { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
          }
        });
      }
    });
  });
}

// ─── SHAREPOINT UPLOAD ────────────────────────────────────────────────────────
function encodeDrivePath(path) {
  return String(path || "").split("/").filter(Boolean).map(p => encodeURIComponent(p)).join("/");
}
function spDrivePath(spFolderUrl) {
  const base = SP_BASE_URL + "/";
  if (!spFolderUrl || !spFolderUrl.startsWith(base)) return null;
  return decodeURIComponent(spFolderUrl.slice(base.length));
}
async function uploadFileToSharePoint(project, blob, filename, contentType, targetPathOverride, conflictBehavior) {
  // The destination is the user-chosen library-relative folder ("" = library
  // root). Fall back to the project's own folder only when no path is supplied.
  const targetPath = (targetPathOverride !== undefined && targetPathOverride !== null)
    ? targetPathOverride
    : (spDrivePath(project.projectFolderUrl) || "");
  const token = await getToken();
  const safeName = filename.replace(/[\\/:*?"<>|]/g, "_");
  // conflictBehavior defaults to "rename" (appends " (1)" / " (2)" so a fresh
  // upload never clobbers an unrelated file). The draft re-save passes
  // "replace" to update a previously-filed draft in place — SharePoint version
  // history still keeps every prior version.
  const url = "https://graph.microsoft.com/v1.0/drives/" + SP_DRIVE_ID
    + "/root:/" + encodeDrivePath(targetPath + "/" + safeName)
    + ":/content?@microsoft.graph.conflictBehavior=" + (conflictBehavior || "rename");
  const res = await fetch(url, {
    method: "PUT",
    headers: { "Authorization": "Bearer " + token, "Content-Type": contentType },
    body: blob,
  });
  if (!res.ok) {
    // 423 = the target file is open in a Word session (co-authoring lock).
    // The raw Graph JSON is useless to the user — say what to actually do.
    if (res.status === 423) {
      throw new Error("The SharePoint copy is open in another Word window or browser tab, so SharePoint has it locked. Close the other copy, wait a minute, and retry — or change the Document Name to file it as a fresh copy.");
    }
    const errText = await res.text().catch(() => "");
    throw new Error("SharePoint " + res.status + ": " + errText.slice(0, 300));
  }
  const item = await res.json();
  return { name: item.name, webUrl: item.webUrl };
}

function uploadDocxToSharePoint(project, blob, filename, targetPathOverride, conflictBehavior) {
  return uploadFileToSharePoint(
    project, blob, filename,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    targetPathOverride, conflictBehavior
  );
}

// ─── NEW DOCUMENT FROM TEMPLATE ───────────────────────────────────────────────
// Creates the .docx directly in the project folder and opens it, so the doc is
// born on SharePoint with AutoSave — the file-then-hand-off flow (and its
// two-copies confusion / 423 lock) never applies to documents started here.

async function loadTemplates(force) {
  const sel = document.getElementById("templateSelect");
  if (!sel) return;
  if (!force) {
    try {
      const raw = localStorage.getItem(TEMPLATES_CACHE_KEY);
      if (raw) {
        const cached = JSON.parse(raw);
        if (cached.savedAt && (Date.now() - cached.savedAt) < TEMPLATES_CACHE_TTL_MS
            && Array.isArray(cached.templates) && cached.templates.length) {
          _templates = cached.templates;
          renderTemplateSelect();
          return;
        }
      }
    } catch {}
  }
  try {
    const token = await getToken();
    const url = "https://graph.microsoft.com/v1.0/drives/" + SP_DRIVE_ID
      + "/root:/" + encodeDrivePath(TEMPLATES_PATH) + ":/children?$select=id,name,file&$top=200";
    const res = await fetch(url, { headers: { "Authorization": "Bearer " + token } });
    if (!res.ok) throw new Error("Template list failed (" + res.status + ")");
    const data = await res.json();
    _templates = (data.value || [])
      .filter(it => it.file && /\.(docx|dotx)$/i.test(it.name || ""))
      .map(it => ({ id: it.id, name: it.name }))
      .sort((a, b) => a.name.localeCompare(b.name));
    localStorage.setItem(TEMPLATES_CACHE_KEY, JSON.stringify({ savedAt: Date.now(), templates: _templates }));
    renderTemplateSelect();
  } catch (e) {
    sel.innerHTML = '<option value="">Couldn’t load templates</option>';
    setStatus("newDocStatus", "error", e.message);
  }
}

// Display/prefill name: drop the extension and the trailing "SA Template" /
// "TEMPLATE" noise the template files carry.
function cleanTemplateBase(name) {
  return String(name || "")
    .replace(/\.(docx|dotx)$/i, "")
    .replace(/\s*(SA\s+Template|Template)\s*$/i, "")
    .trim();
}

function renderTemplateSelect() {
  const sel = document.getElementById("templateSelect");
  if (!sel) return;
  sel.innerHTML = '<option value="">— pick a template —</option>'
    + _templates.map(t =>
        `<option value="${escapeHtml(t.id)}">${escapeHtml(cleanTemplateBase(t.name))}</option>`
      ).join("");
  updateNewDocButtons();
}

function onTemplatePicked() {
  const sel = document.getElementById("templateSelect");
  const t = _templates.find(x => x.id === sel.value);
  if (t) document.getElementById("newDocNameInput").value = cleanTemplateBase(t.name);
  updateNewDocPreview();
  updateNewDocButtons();
}

function updateNewDocPreview() {
  const el = document.getElementById("newDocPreview");
  if (!el) return;
  const name = (document.getElementById("newDocNameInput").value || "").trim();
  const destEl = document.getElementById("destFolderDisplay");
  const folder = (destEl && destEl.textContent) || "Library root";
  el.innerHTML = name
    ? `→ <b>${escapeHtml(name.replace(/\.docx$/i, ""))}.docx</b> in 📂 ${escapeHtml(folder)}`
    : "";
}

function updateNewDocButtons() {
  const toggle = document.getElementById("newDocToggleBtn");
  const create = document.getElementById("createNewDocBtn");
  if (toggle) toggle.disabled = !selectedProject || newDocInFlight;
  if (create) {
    const sel = document.getElementById("templateSelect");
    const name = (document.getElementById("newDocNameInput").value || "").trim();
    create.disabled = !selectedProject || newDocInFlight || !sel || !sel.value || !name;
  }
}

function toggleNewDocPicker() {
  const pick = document.getElementById("newDocPicker");
  if (!pick) return;
  const open = pick.style.display !== "none";
  pick.style.display = open ? "none" : "block";
  if (!open) {
    loadTemplates();
    updateNewDocPreview();
    updateNewDocButtons();
  }
}

// A .dotx is byte-identical to a .docx except the main-part content type inside
// [Content_Types].xml — patch that one string so Word opens the copy as a
// normal document instead of a template. Plain .docx templates pass through.
async function templateBufferToDocxBlob(buf, srcName) {
  const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  if (!/\.dotx$/i.test(srcName || "")) return new Blob([buf], { type: DOCX_MIME });
  if (typeof JSZip === "undefined") {
    throw new Error("JSZip bundle did not load (../jszip.min.js). Check that the file is published on GitHub Pages.");
  }
  const zip = await JSZip.loadAsync(buf);
  const ctFile = zip.file("[Content_Types].xml");
  if (!ctFile) throw new Error("Template is missing [Content_Types].xml — not a valid Office file.");
  const ct = await ctFile.async("string");
  zip.file("[Content_Types].xml",
    ct.replace(/wordprocessingml\.template\.main\+xml/g, "wordprocessingml.document.main+xml"));
  return zip.generateAsync({ type: "blob", mimeType: DOCX_MIME, compression: "DEFLATE" });
}

async function doCreateFromTemplate() {
  if (!selectedProject) { setStatus("newDocStatus", "error", "Pick a project first."); return; }
  if (newDocInFlight) return;
  const sel = document.getElementById("templateSelect");
  const tpl = _templates.find(x => x.id === (sel && sel.value));
  if (!tpl) { setStatus("newDocStatus", "error", "Pick a template first."); return; }
  const baseName = (document.getElementById("newDocNameInput").value || "").trim();
  if (!baseName) { setStatus("newDocStatus", "error", "Give the document a name."); return; }
  const uploadFilename = baseName.replace(/\.docx$/i, "") + ".docx";

  newDocInFlight = true;
  updateNewDocButtons();
  document.getElementById("newDocLink").innerHTML = "";
  try {
    // Same conflict etiquette as filing a draft: never silently clobber.
    let conflictBehavior = "rename";
    if (await spFileExists(spCurrentPath, uploadFilename)) {
      const choice = await askNameConflict(uploadFilename);
      if (choice === "cancel") { setStatus("newDocStatus", "info", "Cancelled."); return; }
      conflictBehavior = choice;
    }

    setStatus("newDocStatus", "info", "⏳ Fetching template…");
    const token = await getToken();
    // Re-read the item for a fresh pre-authenticated downloadUrl — the cached
    // template list can be hours old and downloadUrls expire quickly.
    const metaRes = await fetch(
      "https://graph.microsoft.com/v1.0/drives/" + SP_DRIVE_ID + "/items/" + tpl.id,
      { headers: { "Authorization": "Bearer " + token } }
    );
    if (!metaRes.ok) throw new Error("Couldn’t read the template (" + metaRes.status + "). Try reopening the pane.");
    const meta = await metaRes.json();
    const dlUrl = meta["@microsoft.graph.downloadUrl"];
    if (!dlUrl) throw new Error("Template has no download URL — is it still in the Templates folder?");
    const dlRes = await fetch(dlUrl);
    if (!dlRes.ok) throw new Error("Template download failed (" + dlRes.status + ")");
    const buf = await dlRes.arrayBuffer();
    const docxBlob = await templateBufferToDocxBlob(buf, tpl.name);

    setStatus("newDocStatus", "info", "⏳ Creating " + uploadFilename + " in SharePoint…");
    const item = await uploadDocxToSharePoint(selectedProject, docxBlob, uploadFilename, spCurrentPath, conflictBehavior);

    setStatus("newDocStatus", "success", "✓ " + item.name + " created — opening in Word. AutoSave is on from the start.");
    document.getElementById("newDocLink").innerHTML =
      `<a href="${item.webUrl}" target="_blank">📁 ${escapeHtml(item.name)}</a>`;
    fxFileDrop(document.getElementById("createNewDocBtn"), "📄");
    if (item.webUrl) openExternalUrl(item.webUrl);
  } catch (e) {
    setStatus("newDocStatus", "error", e.message);
  } finally {
    newDocInFlight = false;
    updateNewDocButtons();
  }
}

// ─── PDF EXPORT ───────────────────────────────────────────────────────────────
// Office.js getFileAsync streams the doc in slices; assemble into a Blob.
async function getDocumentAsPdfBlob() {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Pdf, { sliceSize: 65536 }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        return reject(new Error("getFileAsync failed: " + (result.error?.message || "unknown")));
      }
      const file = result.value;
      const sliceCount = file.sliceCount;
      const slices = new Array(sliceCount);
      let received = 0;
      for (let i = 0; i < sliceCount; i++) {
        const idx = i;
        file.getSliceAsync(idx, (sliceRes) => {
          if (sliceRes.status !== Office.AsyncResultStatus.Succeeded) {
            file.closeAsync();
            return reject(new Error("getSliceAsync failed at " + idx));
          }
          slices[idx] = decodeSliceData(sliceRes.value.data);
          received++;
          if (received === sliceCount) {
            file.closeAsync();
            const blob = new Blob(slices, { type: "application/pdf" });
            resolve(blob);
          }
        });
      }
    });
  });
}

// ─── ONENOTE POST ─────────────────────────────────────────────────────────────
async function findOrCreateSection(baseUrl, notebookId, sectionName, token) {
  const sectionsRes = await fetch(`https://graph.microsoft.com/v1.0/${baseUrl}/notebooks/${notebookId}/sections`, {
    headers: { "Authorization": "Bearer " + token }
  });
  if (!sectionsRes.ok) throw new Error("Sections GET " + sectionsRes.status);
  const data = await sectionsRes.json();
  let section = (data.value || []).find(s => s.displayName === sectionName);
  if (section) return section;
  // Create
  const createRes = await fetch(`https://graph.microsoft.com/v1.0/${baseUrl}/notebooks/${notebookId}/sections`, {
    method: "POST",
    headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ displayName: sectionName })
  });
  if (createRes.ok) return await createRes.json();
  // Race: re-fetch
  const refetch = await fetch(`https://graph.microsoft.com/v1.0/${baseUrl}/notebooks/${notebookId}/sections`, {
    headers: { "Authorization": "Bearer " + token }
  });
  const refetchData = await refetch.json();
  section = (refetchData.value || []).find(s => s.displayName === sectionName);
  if (!section) throw new Error("Could not find or create section " + sectionName);
  return section;
}

async function postPdfPrintoutToOneNote(project, pageTitle, pdfBlob) {
  const useTeams   = !!project.teamsOneNoteNotebookId;
  const notebookId = project.teamsOneNoteNotebookId || project.oneNoteNotebookId;
  if (!notebookId) throw new Error("Project has no OneNote notebook linked. Create one in the PMS first.");
  const baseUrl = useTeams ? `groups/${TEAMS_TEAM_ID}/onenote` : `me/onenote`;

  const token = await getToken();
  const section = await findOrCreateSection(baseUrl, notebookId, TARGET_SECTION_NAME, token);

  const safeTitle = pageTitle.replace(/[<&>]/g, c => ({ "<":"&lt;", ">":"&gt;", "&":"&amp;" }[c]));
  const safeProj  = (project.projectNumber || "").replace(/[<&>]/g, c => ({ "<":"&lt;", ">":"&gt;", "&":"&amp;" }[c]));
  const created   = new Date().toISOString();
  // The blob is a PDF, so name the attachment .pdf regardless of the source
  // .docx name — otherwise OneNote stored PDF bytes under a .docx name.
  const fileName  = (docFilename ? docFilename.replace(/\.[^.]+$/, "") : "document") + ".pdf";

  // OneNote multipart: HTML "Presentation" part + named PDF binary part.
  //   • <img data-render-src> renders the PDF's pages as images — the actual
  //     "printout" (one image per page) the user sees inline.
  //   • <object data-attachment> attaches the same part as a downloadable file.
  // Both point at the same multipart part ("filePart"). Previously only the
  // <object> was present, so OneNote showed a file icon, not a printout.
  // If the source .docx has been filed to SharePoint (draftWebUrl is set when
  // the user has run "Save to SharePoint" in this session), inject a clickable
  // link to the source document so the OneNote reader can jump back to the
  // editable Word file — not just the PDF printout.
  const sourceLinkHtml = draftWebUrl
    ? `<div style="font-family:sans-serif;font-size:12px;margin-bottom:14px;padding:6px 10px;background:#f0f4fa;border-left:3px solid #003865;border-radius:3px">` +
        `<a href="${escapeHtml(draftWebUrl)}" target="_blank" style="color:#003865;text-decoration:none;font-weight:600">` +
          `📄 Open source Word document →` +
        `</a>` +
      `</div>`
    : "";

  const presentation =
    `<!DOCTYPE html><html><head><title>${safeTitle}</title><meta name="created" content="${created}" /></head><body>` +
    `<div style="border-bottom:2px solid #003865;padding-bottom:8px;margin-bottom:16px;font-family:sans-serif">` +
    `<span style="background:#003865;color:#fff;padding:2px 8px;border-radius:3px;font-size:11px;margin-right:6px">${safeProj}</span>` +
    `<span style="font-size:11px;color:#666">${escapeHtml(fileName)}</span>` +
    `</div>` +
    sourceLinkHtml +
    `<img data-render-src="name:filePart" alt="${escapeHtml(pageTitle)}" width="800" />` +
    `<p style="font-family:sans-serif;font-size:11px;color:#666;margin-top:12px">📎 Attached file:</p>` +
    `<object data-attachment="${escapeHtml(fileName)}" data="name:filePart" type="application/pdf" />` +
    `</body></html>`;

  const form = new FormData();
  form.append("Presentation", new Blob([presentation], { type: "text/html" }));
  form.append("filePart", pdfBlob, fileName);

  const url = `https://graph.microsoft.com/v1.0/${baseUrl}/sections/${section.id}/pages`;
  const maxAttempts = 3;
  let lastErr;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const res = await fetch(url, {
      method: "POST",
      headers: { "Authorization": "Bearer " + token },
      body: form,
    });
    if (res.ok) {
      const page = await res.json();
      return { id: page.id, webUrl: page.links?.oneNoteWebUrl?.href || page.webUrl || "" };
    }
    if (res.status === 429 || res.status === 503) {
      const retryAfter = parseInt(res.headers.get("Retry-After") || "0", 10);
      const wait = retryAfter > 0 ? retryAfter * 1000 : Math.min(15000, 1000 * Math.pow(2, attempt));
      await new Promise(r => setTimeout(r, wait));
      lastErr = new Error("OneNote throttled (" + res.status + ")");
      continue;
    }
    const errText = await res.text().catch(() => "");
    throw new Error("OneNote " + res.status + ": " + errText.slice(0, 300));
  }
  throw lastErr || new Error("OneNote page creation failed");
}

// ─── INSERT FLOWS ─────────────────────────────────────────────────────────────
async function insertHtmlAtCursor(html) {
  await Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.insertHtml(html, Word.InsertLocation.replace);
    await ctx.sync();
  });
}

async function doInsertField(which) {
  if (!selectedProject) return;
  const fieldMap = {
    name:   { value: selectedProject.name,          label: "project name"   },
    number: { value: selectedProject.projectNumber, label: "project number" },
    client: { value: selectedProject.clientName,    label: "client name"    },
  };
  const f = fieldMap[which];
  const text = (f?.value || "").trim();
  if (!text) {
    setStatus("insertStatus", "error", `Project has no ${f?.label || which}.`);
    return;
  }
  try {
    await insertHtmlAtCursor(escapeHtml(text));
    setStatus("insertStatus", "success", `✓ Inserted ${f.label}`);
  } catch (e) {
    setStatus("insertStatus", "error", "Insert failed: " + e.message);
  }
}

function togglePocPicker() {
  const picker = document.getElementById("pocPicker");
  const open = picker.style.display !== "none";
  if (open) {
    picker.style.display = "none";
  } else {
    picker.style.display = "block";
    document.getElementById("pocSearch").value = "";
    renderPocList();
    document.getElementById("pocSearch").focus();
  }
}

// Build a unified POC list. Project people appear first (group: "On this
// project"), followed by the global directory minus anyone already in the
// project list (deduped by lowercase email).
function buildPocList() {
  const projectPeople = (selectedProject?.directory || []).map(d => ({
    name: d.name || "",
    title: d.title || "",
    email: d.email || "",
    phone: d.phone || "",
    company: d.company || "",
    group: "On this project",
  }));
  const projectEmails = new Set(
    projectPeople.map(p => (p.email || "").toLowerCase()).filter(Boolean)
  );
  const others = [];
  for (const client of allClients || []) {
    for (const c of (client.contacts || [])) {
      const emailLc = (c.email || "").toLowerCase();
      if (emailLc && projectEmails.has(emailLc)) continue;
      others.push({
        name: c.name || "",
        title: c.title || "",
        email: c.email || "",
        phone: c.phone || "",
        company: client.name || "",
        group: "Directory",
      });
    }
  }
  return [...projectPeople, ...others];
}

function renderPocList() {
  const q = (document.getElementById("pocSearch").value || "").trim().toLowerCase();
  const list = document.getElementById("pocList");
  const all = buildPocList();
  const filtered = all.filter(p => {
    if (!q) return true;
    return p.name.toLowerCase().includes(q)
        || p.company.toLowerCase().includes(q)
        || p.email.toLowerCase().includes(q)
        || p.title.toLowerCase().includes(q);
  });
  if (!filtered.length) {
    list.innerHTML = '<div class="poc-empty">No people match. Add contacts via the Outlook add-in or PMS Directory.</div>';
    return;
  }
  // Capped at 100 results — user should refine the search if they need more.
  const capped = filtered.slice(0, 100);
  let html = "";
  let lastGroup = "";
  capped.forEach((p, i) => {
    if (p.group !== lastGroup) {
      html += `<div class="poc-group-header">${escapeHtml(p.group)}</div>`;
      lastGroup = p.group;
    }
    const meta = [p.title, p.company].filter(Boolean).join(" · ");
    html += `<div class="poc-row" data-i="${i}">
      <div class="poc-name">${escapeHtml(p.name) || "<em>(no name)</em>"}</div>
      <div class="poc-meta">${escapeHtml(meta)}${p.email ? " · " + escapeHtml(p.email) : ""}</div>
    </div>`;
  });
  list.innerHTML = html;
  list.querySelectorAll(".poc-row").forEach(row => {
    row.onclick = () => {
      const idx = parseInt(row.getAttribute("data-i"), 10);
      doInsertPocBlock(capped[idx]);
    };
  });
}

async function doInsertPocBlock(person) {
  const lines = [person.name, person.title, person.company, person.email]
    .filter(Boolean)
    .map(escapeHtml);
  if (!lines.length) {
    setStatus("insertStatus", "error", "This contact has no fields to insert.");
    return;
  }
  const html = lines.join("<br>");
  try {
    await insertHtmlAtCursor(html);
    setStatus("insertStatus", "success", `✓ Inserted ${person.name || "contact"}`);
    document.getElementById("pocPicker").style.display = "none";
  } catch (e) {
    setStatus("insertStatus", "error", "Insert failed: " + e.message);
  }
}

// ─── DOCUMENT STATUS ─────────────────────────────────────────────────────────
async function loadDocStatus() {
  try {
    const saved = Office.context.document.settings.get("settyPms:docStatus");
    if (saved && DOC_STATUS_OPTIONS.find(o => o.value === saved)) docStatus = saved;
  } catch (e) {
    console.warn("Could not read doc status:", e.message);
  }
}

async function saveDocStatus(status) {
  docStatus = status;
  try {
    Office.context.document.settings.set("settyPms:docStatus", status);
    await new Promise((resolve, reject) => {
      Office.context.document.settings.saveAsync(r =>
        r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error?.message))
      );
    });
  } catch (e) {
    console.warn("Could not persist doc status:", e.message);
  }
  renderStatusBar();
  updateSaveButtons();
}

function renderStatusBar() {
  const opt  = DOC_STATUS_OPTIONS.find(o => o.value === docStatus) || DOC_STATUS_OPTIONS[0];
  const dot  = document.getElementById("statusDot");
  const lbl  = document.getElementById("statusLabelText");
  const hint = document.getElementById("statusHint");
  if (dot)  dot.style.background  = opt.dot;
  if (lbl)  lbl.textContent       = opt.label;
  if (hint) hint.textContent      = opt.hint;
  // "Start a new version" only makes sense once the document has gone Final
  // (PDF exported) — before that, plain editing and re-filing cover it.
  const nvWrap = document.getElementById("newVersionWrap");
  if (nvWrap) nvWrap.style.display = docStatus === "final" ? "block" : "none";
}

// ─── DONE-EDITING ─────────────────────────────────────────────────────────────
async function loadDoneEditing() {
  try {
    const raw = Office.context.document.settings.get(DOC_DONE_KEY);
    doneEditingList = (typeof raw === "string" ? JSON.parse(raw) : raw) || [];
  } catch {
    doneEditingList = [];
  }
}

async function saveDoneEditingList() {
  try {
    Office.context.document.settings.set(DOC_DONE_KEY, JSON.stringify(doneEditingList));
    await new Promise((resolve, reject) => {
      Office.context.document.settings.saveAsync(r =>
        r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error?.message))
      );
    });
  } catch (e) {
    console.warn("Could not persist done list:", e.message);
  }
}

async function toggleCurrentUserDone() {
  const email = (msalAccount?.username || "").toLowerCase();
  const name  = msalAccount?.name || email;
  const idx   = doneEditingList.findIndex(d => (d.email || "").toLowerCase() === email);
  if (idx >= 0) {
    doneEditingList.splice(idx, 1);
  } else {
    doneEditingList.push({ name, email, ts: Date.now() });
  }
  await saveDoneEditingList();
  renderDoneEditing();
}

function renderDoneEditing() {
  const section    = document.getElementById("doneEditingSection");
  const listEl     = document.getElementById("doneEditingList");
  const toggleBtn  = document.getElementById("toggleDoneBtn");
  if (!section) return;

  // Stays visible even when Final — the record of who signed off before the
  // document was printed is useful history. Only the "Mark me done" toggle is
  // hidden once final, since the editing phase is over.
  const isFinal = docStatus === "final";
  section.style.display = "block";
  if (toggleBtn) toggleBtn.style.display = isFinal ? "none" : "";

  const email   = (msalAccount?.username || "").toLowerCase();
  const iAmDone = doneEditingList.some(d => (d.email || "").toLowerCase() === email);
  if (toggleBtn && !isFinal) {
    toggleBtn.textContent = iAmDone ? "✓ I'm done" : "Mark me done";
    toggleBtn.classList.toggle("is-done", iAmDone);
  }
  if (!listEl) return;
  if (!doneEditingList.length) {
    listEl.innerHTML = '<div class="done-empty">No one has marked done yet.</div>';
    return;
  }
  listEl.innerHTML = doneEditingList.map(d =>
    `<div class="done-item">
      <span class="done-check">✓</span>
      <span class="done-name">${escapeHtml(d.name || d.email)}</span>
      <span class="done-time">${formatTimeAgo(d.ts)}</span>
    </div>`
  ).join("");
}

function formatTimeAgo(ts) {
  const m = Math.floor((Date.now() - ts) / 60000);
  if (m < 1)  return "just now";
  if (m < 60) return m + "m ago";
  const h = Math.floor(m / 60);
  if (h < 24) return h + "h ago";
  return Math.floor(h / 24) + "d ago";
}

// ─── SAVE FLOWS ───────────────────────────────────────────────────────────────
async function doSaveToOneNote() {
  if (!selectedProject) { setStatus("oneNoteStatus", "error", "Pick a project first."); return; }
  if (saveInFlight) return;
  saveInFlight = true;
  updateSaveButtons();
  document.getElementById("oneNoteLink").innerHTML = "";
  const titleInput = document.getElementById("titleInput");
  const pageTitle = (titleInput.value || "").trim() || docFilename.replace(/\.[^.]+$/, "");
  try {
    isExporting = true;
    setStatus("oneNoteStatus", "info", "⏳ Exporting document as PDF…");
    const pdfBlob = await getDocumentAsPdfBlob();
    setStatus("oneNoteStatus", "info", "⏳ Sending to OneNote…");
    const page = await postPdfPrintoutToOneNote(selectedProject, pageTitle, pdfBlob);
    setStatus("oneNoteStatus", "success", "✓ Saved to OneNote");
    fxPaperPlane(document.getElementById("saveOneNoteBtn"));
    if (page.webUrl) {
      document.getElementById("oneNoteLink").innerHTML =
        `<a href="${page.webUrl}" target="_blank">📓 Open in OneNote</a>`;
    }
  } catch (e) {
    setStatus("oneNoteStatus", "error", e.message);
  } finally {
    isExporting = false;
    saveInFlight = false;
    updateSaveButtons();
  }
}

// Saves the .docx to SharePoint — safe to use at any status, any time.
async function doSaveDraft() {
  if (!selectedProject) { setStatus("spStatus", "error", "Pick a project first."); return; }
  if (saveInFlight) return;
  saveInFlight = true;
  updateSaveButtons();
  document.getElementById("spLink").innerHTML = "";
  const titleInput = document.getElementById("titleInput");
  // The Document Name field drives the saved filename — fall back to the
  // original filename, then a generic name only if both are empty.
  const baseName = (titleInput.value || "").trim()
    || docFilename.replace(/\.[^.]+$/, "")
    || "Document";
  const uploadFilename = baseName.replace(/\.docx$/i, "") + ".docx";

  // "Overwrite unless renamed": if this document already has a filed draft and
  // neither the name nor the destination folder has changed, replace that file
  // in place. Otherwise upload a fresh, rename-on-conflict copy.
  const sameDraft = draftSaved
    && uploadFilename === draftBaseName
    && spCurrentPath  === draftFolderPath;

  // First-time (or renamed/moved) save: if the target name is already taken,
  // ask how to resolve it instead of silently appending " (1)".
  let conflictBehavior = "rename";
  if (!sameDraft && await spFileExists(spCurrentPath, uploadFilename)) {
    const choice = await askNameConflict(uploadFilename);
    if (choice === "cancel") {
      setStatus("spStatus", "info", "Save cancelled.");
      saveInFlight = false;
      updateSaveButtons();
      return;
    }
    conflictBehavior = choice;   // "replace" overwrites in place; "rename" keeps both
  }

  try {
    isExporting = true;   // our own settings/file writes below are not user edits
    setStatus("spStatus", "info", "⏳ Exporting .docx…");
    // Stage the project/status tags into the document before reading it out, so
    // the uploaded copy carries them when someone reopens it.
    persistDocSettings();
    await flushSettings();
    const docxBlob = await getDocumentAsDocxBlob();
    setStatus("spStatus", "info", sameDraft ? "⏳ Updating the SharePoint copy…" : "⏳ Uploading to SharePoint…");
    // A replace targets the file actually on SharePoint (draftFileName), which
    // can differ from the intended name if the first save hit a name clash.
    const item = sameDraft
      ? await uploadDocxToSharePoint(selectedProject, docxBlob, draftFileName, draftFolderPath, "replace")
      : await uploadDocxToSharePoint(selectedProject, docxBlob, uploadFilename, spCurrentPath, conflictBehavior);

    // Commit the filed-draft state — only now that the upload has succeeded.
    draftSaved      = true;
    draftBaseName   = uploadFilename;
    draftFileName   = item.name;
    draftFolderPath = spCurrentPath;
    draftWebUrl     = item.webUrl || "";
    dirtySinceFiled = false;   // the local copy now matches the filed copy again
    persistDocSettings();
    await flushSettings();

    setStatus("spStatus", "success", sameDraft ? "✓ SharePoint copy updated" : "✓ Draft filed to SharePoint");
    fxFileDrop(document.getElementById("saveSpDraftBtn"), "💾");
    document.getElementById("spLink").innerHTML =
      `<a href="${item.webUrl}" target="_blank">📁 ${escapeHtml(item.name)}</a>`;
  } catch (e) {
    // Upload failed — leave draftSaved untouched: a previously filed draft (if
    // any) still exists, and a first-time save simply stays un-filed.
    setStatus("spStatus", "error", e.message);
  } finally {
    isExporting = false;
    saveInFlight = false;
    updateSaveButtons();
    updateDestCard();
  }
}

// Exports a PDF and saves it to SharePoint. Auto-sets status to Final on success.
async function doSavePdf() {
  if (!selectedProject) { setStatus("spStatus", "error", "Pick a project first."); return; }
  if (saveInFlight) return;
  saveInFlight = true;
  updateSaveButtons();
  document.getElementById("spLink").innerHTML = "";
  hideSendEmail();   // stale PDF — the button reappears when this export succeeds
  const titleInput = document.getElementById("titleInput");
  // The Document Name field drives the saved filename.
  const baseName = (titleInput.value || "").trim()
    || docFilename.replace(/\.[^.]+$/, "")
    || "Document";
  const uploadFilename = baseName.replace(/\.[^.]+$/, "") + ".pdf";
  try {
    isExporting = true;
    setStatus("spStatus", "info", "⏳ Exporting PDF…");
    const pdfBlob = await getDocumentAsPdfBlob();
    setStatus("spStatus", "info", "⏳ Uploading PDF to SharePoint…");
    const pdfItem = await uploadFileToSharePoint(
      selectedProject, pdfBlob, uploadFilename, "application/pdf", spCurrentPath
    );
    // Mark as Final. The done-editing list is kept as a record of who signed
    // off before the document was printed. The project tag is staged too so the
    // open document keeps it (saveDocStatus flushes both to the file).
    persistDocSettings();
    await saveDocStatus("final");
    renderDoneEditing();
    setStatus("spStatus", "success", "✓ PDF saved to SharePoint — document marked Final");
    fxFileDrop(document.getElementById("savePdfBtn"), "📄");
    document.getElementById("spLink").innerHTML =
      `<a href="${pdfItem.webUrl}" target="_blank">📄 ${escapeHtml(pdfItem.name)}</a>`;
    // Keep the exported PDF around so "Send in Email" can attach it.
    lastPdfBlob = pdfBlob;
    lastPdfName = pdfItem.name;
    lastPdfWebUrl = pdfItem.webUrl;
    showSendEmail();
  } catch (e) {
    setStatus("spStatus", "error", e.message);
  } finally {
    isExporting = false;
    saveInFlight = false;
    updateSaveButtons();
  }
}

// ─── NEW VERSION ─────────────────────────────────────────────────────────────
// "Report" → "Report v2"; "Report v2" → "Report v3". The printed original
// counts as v1, so an unversioned name jumps straight to v2.
function bumpVersionName(base) {
  const m = base.match(/^(.*?)[ _-]*[vV](\d+)$/);
  if (m) return m[1].replace(/[ _.-]+$/, "") + " v" + (parseInt(m[2], 10) + 1);
  return base + " v2";
}

// After a PDF export marks the document Final, this starts the next round:
// files the current content as a fresh "<name> v#.docx" in the SAME folder,
// clears the sign-off list, and flips status back to Draft — both in this open
// document and inside the uploaded copy (settings are flushed before the file
// is read out, so the v# copy opens clean).
async function doStartNewVersion() {
  if (!selectedProject) { setStatus("newVersionStatus", "error", "Pick a project first."); return; }
  if (saveInFlight) return;
  saveInFlight = true;
  updateSaveButtons();
  const linkEl = document.getElementById("newVersionLink");
  if (linkEl) linkEl.innerHTML = "";

  const titleInput = document.getElementById("titleInput");
  const currentBase = (titleInput?.value || "").trim().replace(/\.(docx|pdf)$/i, "")
    || docFilename.replace(/\.[^.]+$/, "")
    || "Document";
  const newBase = bumpVersionName(currentBase);
  const uploadFilename = newBase + ".docx";
  // Same folder as the filed draft; a doc that was only ever PDF-exported falls
  // back to the current picker folder (where that PDF just went).
  const targetPath = (draftSaved && draftFolderPath) ? draftFolderPath : spCurrentPath;

  // Snapshot so a failed upload can put the sign-offs and status back.
  const prevStatus = docStatus;
  const prevDone   = doneEditingList.slice();
  const prevDraft  = { draftSaved, draftFolderPath, draftBaseName, draftFileName, draftWebUrl };
  try {
    isExporting = true;   // our own settings/file writes are not user edits
    setStatus("newVersionStatus", "info", "⏳ Preparing " + uploadFilename + "…");
    docStatus = "draft";
    doneEditingList = [];
    Office.context.document.settings.set(DOC_DONE_KEY, JSON.stringify(doneEditingList));
    persistDocSettings();
    await flushSettings();
    const docxBlob = await getDocumentAsDocxBlob();
    setStatus("newVersionStatus", "info", "⏳ Filing " + uploadFilename + " to SharePoint…");
    const item = await uploadDocxToSharePoint(selectedProject, docxBlob, uploadFilename, targetPath, "rename");

    // The new version is now the filed draft this pane tracks.
    draftSaved      = true;
    draftBaseName   = uploadFilename;
    draftFileName   = item.name;
    draftFolderPath = targetPath;
    draftWebUrl     = item.webUrl || "";
    dirtySinceFiled = false;
    spCurrentPath   = targetPath;
    persistDocSettings();
    await flushSettings();
    if (titleInput) titleInput.value = item.name.replace(/\.docx$/i, "");

    setStatus("newVersionStatus", "success",
      "✓ " + item.name + " created — sign-offs cleared, status back to Draft");
    fxFileDrop(document.getElementById("newVersionBtn"), "📄");
    if (linkEl && item.webUrl) {
      const openUrl = item.webUrl;
      linkEl.innerHTML = `<a href="#" id="openNewVersionLink">📝 Open ${escapeHtml(item.name)} to edit →</a>`;
      document.getElementById("openNewVersionLink").onclick = (e) => { e.preventDefault(); openExternalUrl(openUrl); };
    }
  } catch (e) {
    // The new version never landed — restore status and sign-offs untouched.
    docStatus = prevStatus;
    doneEditingList = prevDone;
    ({ draftSaved, draftFolderPath, draftBaseName, draftFileName, draftWebUrl } = prevDraft);
    Office.context.document.settings.set(DOC_DONE_KEY, JSON.stringify(doneEditingList));
    persistDocSettings();
    await flushSettings();
    setStatus("newVersionStatus", "error", e.message);
  } finally {
    isExporting = false;
    saveInFlight = false;
    renderStatusBar();
    renderDoneEditing();
    updateSaveButtons();
    updateFiledCard();
    updateDestCard();
  }
}

// ─── SEND IN EMAIL ────────────────────────────────────────────────────────────
// Creates an Outlook draft with the exported PDF attached, then opens it in
// Outlook on the web for the user to address and send. Nothing is sent from
// here — recipients and the Send button stay with the user.

// Graph caps JSON-inline attachments at ~3 MB; bigger PDFs go through an
// attachment upload session in chunks (size must be a multiple of 320 KiB).
const INLINE_ATTACH_MAX = 3 * 1024 * 1024;
const UPLOAD_CHUNK = 327680 * 10;   // 3.2 MB

function showSendEmail() {
  const btn = document.getElementById("sendEmailBtn");
  if (btn) btn.style.display = "flex";
}

function hideSendEmail() {
  const btn = document.getElementById("sendEmailBtn");
  if (btn) btn.style.display = "none";
  const status = document.getElementById("emailStatus");
  if (status) { status.className = "status"; status.textContent = ""; }
  const link = document.getElementById("emailLink");
  if (link) link.innerHTML = "";
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload  = () => resolve(String(fr.result).split(",", 2)[1] || "");
    fr.onerror = () => reject(new Error("Could not read the PDF for attaching"));
    fr.readAsDataURL(blob);
  });
}

async function doSendPdfEmail() {
  if (!lastPdfBlob) { setStatus("emailStatus", "error", "Export a PDF first."); return; }
  if (emailInFlight) return;
  emailInFlight = true;
  const btn = document.getElementById("sendEmailBtn");
  if (btn) btn.disabled = true;
  const linkEl = document.getElementById("emailLink");
  if (linkEl) linkEl.innerHTML = "";
  try {
    setStatus("emailStatus", "info", "⏳ Creating email draft…");
    const token = await getMailToken();
    // Footer only — the message itself is the user's to write. The SharePoint
    // link gives recipients (and the sender, later) the filed copy.
    const bodyHtml =
      `<br/><br/><div style="font-size:12px;color:#666">Attached: ${escapeHtml(lastPdfName)}` +
      (lastPdfWebUrl ? ` — <a href="${escapeHtml(lastPdfWebUrl)}">filed copy on SharePoint</a>` : "") +
      `</div>`;
    const draftPayload = {
      subject: lastPdfName.replace(/\.pdf$/i, ""),
      body: { contentType: "html", content: bodyHtml },
    };
    const smallEnough = lastPdfBlob.size <= INLINE_ATTACH_MAX;
    if (smallEnough) {
      draftPayload.attachments = [{
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: lastPdfName,
        contentType: "application/pdf",
        contentBytes: await blobToBase64(lastPdfBlob),
      }];
    }
    const res = await fetch("https://graph.microsoft.com/v1.0/me/messages", {
      method: "POST",
      headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify(draftPayload),
    });
    if (!res.ok) {
      const errText = await res.text().catch(() => "");
      throw new Error("Draft create failed (" + res.status + "): " + errText.slice(0, 300));
    }
    const draft = await res.json();
    if (!smallEnough) {
      setStatus("emailStatus", "info", "⏳ Attaching PDF (large file)…");
      await uploadLargeAttachment(token, draft.id, lastPdfBlob, lastPdfName);
    }
    setStatus("emailStatus", "success", "✓ Draft created — opening in Outlook…");
    fxPaperPlane(btn);
    // Fallback link in case the browser window is blocked or closed too soon.
    if (linkEl && draft.webLink) {
      linkEl.innerHTML = `<a href="${escapeHtml(draft.webLink)}" target="_blank">✉️ Open the draft</a>`;
    }
    openExternalUrl(draft.webLink);
  } catch (e) {
    setStatus("emailStatus", "error", e.message);
  } finally {
    emailInFlight = false;
    if (btn) btn.disabled = false;
  }
}

// Outlook large-attachment upload session. The uploadUrl is pre-authenticated,
// so the chunk PUTs carry no Authorization header.
async function uploadLargeAttachment(token, messageId, blob, name) {
  const sessRes = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments/createUploadSession`, {
    method: "POST",
    headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({
      AttachmentItem: { attachmentType: "file", name, size: blob.size }
    }),
  });
  if (!sessRes.ok) {
    const errText = await sessRes.text().catch(() => "");
    throw new Error("Attachment session failed (" + sessRes.status + "): " + errText.slice(0, 300));
  }
  const session = await sessRes.json();
  let offset = 0;
  while (offset < blob.size) {
    const end = Math.min(offset + UPLOAD_CHUNK, blob.size);
    const putRes = await fetch(session.uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Type": "application/octet-stream",
        "Content-Range": `bytes ${offset}-${end - 1}/${blob.size}`,
      },
      body: blob.slice(offset, end),
    });
    if (!putRes.ok) {
      const errText = await putRes.text().catch(() => "");
      throw new Error("Attachment upload failed (" + putRes.status + "): " + errText.slice(0, 300));
    }
    offset = end;
  }
}

// ─── UTIL ─────────────────────────────────────────────────────────────────────
function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, c => ({
    "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;"
  }[c]));
}

// ─── CELEBRATIONS ─────────────────────────────────────────────────────────────
// Fun, lightweight feedback for the add-in's key moments. Every effect is
// decorative — it never blocks the UI and always cleans up after itself.

// Master guard: false when the OS asks for reduced motion, so every effect
// below becomes a no-op for users who opted out of animation.
function fxOn() {
  return !(window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches);
}

// Shared full-pane overlay that hosts the floating effects.
function fxLayer() {
  let layer = document.getElementById("wordFx");
  if (!layer) {
    layer = document.createElement("div");
    layer.id = "wordFx";
    document.body.appendChild(layer);
  }
  return layer;
}

// Center point of an element, in viewport coordinates.
function fxCenterOf(el) {
  const r = el.getBoundingClientRect();
  return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
}

// 1 ─ Paper airplane: launches from `originEl` and flies off the pane.
//     Fired on a successful "Save to OneNote".
function fxPaperPlane(originEl) {
  if (!fxOn() || !originEl) return;
  const { x, y } = fxCenterOf(originEl);
  const plane = document.createElement("div");
  plane.className = "fx-plane";
  plane.style.left = x + "px";
  plane.style.top  = y + "px";
  fxLayer().appendChild(plane);
  setTimeout(() => plane.remove(), 1300);
}

// 2 ─ File drop: a page icon falls into a folder just above `originEl`.
//     Fired on a successful SharePoint save (PDF export or .docx draft).
function fxFileDrop(originEl, pageIcon) {
  if (!fxOn() || !originEl) return;
  const { x, y } = fxCenterOf(originEl);
  const drop = document.createElement("div");
  drop.className = "fx-drop";
  drop.style.left = x + "px";
  drop.style.top  = (y - 34) + "px";        // sit just above the clicked button
  drop.innerHTML =
    '<span class="fx-page">' + (pageIcon || "📄") + '</span>' +
    '<span class="fx-folder fx-gulp">📁</span>';
  fxLayer().appendChild(drop);
  setTimeout(() => drop.remove(), 1300);
}

// 3 ─ Auto-suggest celebration: pops + glows the project pill and pings a
//     few sparkles, so a correct guess feels intentional, not accidental.
function fxSuggestPill() {
  if (!fxOn()) return;
  const pill = document.getElementById("projectPill");
  if (!pill || pill.style.display === "none") return;
  pill.classList.remove("fx-pill-pop", "fx-pill-glow");
  void pill.offsetWidth;                     // restart the animations
  pill.classList.add("fx-pill-pop", "fx-pill-glow");
  setTimeout(() => pill.classList.remove("fx-pill-pop", "fx-pill-glow"), 1600);

  const { x } = fxCenterOf(pill);
  const top = pill.getBoundingClientRect().top;
  ["✨", "⭐", "✨"].forEach((s, i) => {
    const sp = document.createElement("div");
    sp.className = "fx-sparkle";
    sp.textContent = s;
    sp.style.left = (x + (i - 1) * 28) + "px";
    sp.style.top  = top + "px";
    sp.style.setProperty("--dx", ((i - 1) * 16) + "px");
    sp.style.setProperty("--dy", (-22 - i * 6) + "px");
    fxLayer().appendChild(sp);
    setTimeout(() => sp.remove(), 1100);
  });
}

// 4 ─ Welcome: wiggles the header logo and slides in a greeting toast.
function fxWelcome(firstName) {
  if (!fxOn()) return;
  const logo = document.querySelector(".header-logo");
  if (logo) {
    logo.classList.remove("fx-logo-wiggle");
    void logo.offsetWidth;
    logo.classList.add("fx-logo-wiggle");
  }
  const toast = document.createElement("div");
  toast.className = "fx-welcome";
  toast.textContent = pickGreeting(firstName);
  document.body.appendChild(toast);
  setTimeout(() => toast.remove(), 3500);
}

// Best-effort first name from the signed-in account's display name.
function fxFirstName() {
  const n = ((msalAccount && msalAccount.name) || "").trim();
  if (!n) return "";
  // Handle "Last, First" as well as "First Last".
  if (n.includes(",")) return n.split(",")[1].trim().split(/\s+/)[0];
  return n.split(/\s+/)[0];
}

// The greeting shown in the welcome toast. THIS IS YOURS TO WRITE — see chat.
// Each entry is a function: it receives the first name ("" if unknown) and
// returns the toast text.
const WELCOME_GREETINGS = [
  name => name ? `Welcome back, ${name}!` : "Welcome back!",
  name => name ? `Welcome to the wordy party, ${name}!` : "Welcome to the wordy party!",
  name => name ? `Word on the street: ${name}'s got docs to file.` : "Word on the street: docs to file.",
  name => name ? `${name}, let's give this doc a forever home.` : "Let's give this doc a forever home.",
  name => name ? `The doc whisperer returns. Hey, ${name}.` : "The doc whisperer returns.",
  name => name ? `Filing time, ${name} — make it look easy.` : "Filing time — make it look easy.",
  name => name ? `It won't file itself, ${name}. Lucky you've got us.` : "It won't file itself. Lucky you've got us.",
  name => name ? `Look who's being organized today — hi, ${name}.` : "Look who's being organized today.",
  name => name ? `Another masterpiece to file, ${name}?` : "Another masterpiece to file?",
  name => name ? `${name}, let's turn "where'd I save that" into "filed."` : `Let's turn "where'd I save that" into "filed."`,
  name => name ? `Your future self thanks you for filing this, ${name}.` : "Your future self thanks you for filing this.",
  name => name ? `Fresh doc, fresh start. Let's file it, ${name}.` : "Fresh doc, fresh start. Let's file it.",
  name => name ? `${name}, inbox-zero energy starts here.` : "Inbox-zero energy starts here.",
  name => name ? `Hey ${name} — let's make this doc easy to find later.` : "Let's make this doc easy to find later.",
];
function pickGreeting(firstName) {
  const pick = WELCOME_GREETINGS[Math.floor(Math.random() * WELCOME_GREETINGS.length)];
  return pick(firstName || "");
}
