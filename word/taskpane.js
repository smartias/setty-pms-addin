// ─── CONFIG ───────────────────────────────────────────────────────────────────
const MSAL_CONFIG = {
  auth: {
    clientId: "8e5155fb-6221-4508-97ea-3661438c6688",
    authority: "https://login.microsoftonline.com/f374c024-71c2-48b6-8420-076fff97327c",
    redirectUri: "https://smartias.github.io/setty-pms-addin/word/taskpane.html",
  },
  cache: { cacheLocation: "localStorage" }
};
const GRAPH_SCOPES = ["User.Read", "Notes.ReadWrite", "Files.ReadWrite.All"];
const TEAMS_TEAM_ID = "a4c48361-7991-43db-af83-4c854918a760";
// SharePoint — same hardcoded drive ID the Outlook add-in uses (no admin consent needed).
const SP_DRIVE_ID = "b!ZARYqukTtE6K1Mpv9bngAehneskb-yNKopp1Ol1X1BBnJPKsNGM-TaGmbGiL3ZaU";
const SP_BASE_URL = "https://setty.sharepoint.com/sites/NYCProjects/Project%20Document%20Library";
const SUPABASE_URL  = "https://khxmgjilwhdguuepbhne.supabase.co";
const SUPABASE_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtoeG1namlsd2hkZ3V1ZXBiaG5lIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMwNjg2MDYsImV4cCI6MjA4ODY0NDYwNn0.vtHt2eydU2iQ426iYOzLrqpH2WLXdRnicq-3sNfoNq8";
const SB_HEADERS = { "apikey": SUPABASE_ANON, "Authorization": "Bearer " + SUPABASE_ANON };
const PROJECTS_CACHE_KEY = "settyPmsWord:projectsCache";
const PROJECTS_CACHE_TTL_MS = 1000 * 60 * 60 * 6; // 6h
const LAST_ACCOUNT_STORAGE_KEY = "settyPmsWord:lastMsalAccountHomeId";
const TARGET_SECTION_NAME = "Documents"; // fixed section per product decision
const PMS_PROJECT_BASE_URL = "https://smartias.github.io/setty-pms/SettyPMS.html#project:";

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
  document.getElementById("titleInput").addEventListener("input", updateSavePreview);
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
  // The destination picker stays available even after a draft is filed — it
  // shows where the *next* save (Update Draft / PDF) will go.
  const hasFolder = !!(selectedProject?.projectFolderUrl);
  card.style.display = hasFolder ? "block" : "none";
  if (hasFolder) {
    const display = document.getElementById("destFolderDisplay");
    const rel = spCurrentPath && spProjectRootPath && spCurrentPath.startsWith(spProjectRootPath)
      ? spCurrentPath.slice(spProjectRootPath.length).replace(/^\/+/, "")
      : "";
    if (display) display.textContent = rel || "project root";
  }
  updateSavePreview();
  updateFiledCard();
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
  if (selectedProject.projectFolderUrl) {
    const rel = spCurrentPath && spProjectRootPath && spCurrentPath.startsWith(spProjectRootPath)
      ? spCurrentPath.slice(spProjectRootPath.length).replace(/^\/+/, "")
      : "";
    el.innerHTML = `Saves as <b>${escapeHtml(base)}.docx</b> in 📂 ${escapeHtml(rel || "project root")}`;
  } else {
    el.innerHTML = `Saves as <b>${escapeHtml(base)}.docx</b>`;
  }
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
    const folder = draftFolderPath && spProjectRootPath && draftFolderPath.startsWith(spProjectRootPath)
      ? draftFolderPath.slice(spProjectRootPath.length).replace(/^\/+/, "") || "project root"
      : (draftFolderPath || "project root");
    nameEl.textContent = (draftFileName || "(filed)") + "  ·  " + folder;
  }

  const headEl = document.getElementById("filedHead");
  const hintEl = document.getElementById("filedHint");
  const openBtn = document.getElementById("openSpCopyBtn");
  const inner   = document.getElementById("filedCardInner");

  // If the open document is itself cloud-hosted (http/https URL), it's the
  // SharePoint copy — edits AutoSave and there's nothing to reopen. Show a calm
  // confirmation and hide the open action, so people aren't told to "open the
  // SharePoint copy" when they're already in it.
  const inCloudCopy = /^https?:\/\//i.test(currentDocUrl || "");
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
  try {
    if (Office?.context?.ui?.openBrowserWindow) {
      Office.context.ui.openBrowserWindow(draftWebUrl);
      return;
    }
  } catch (e) {
    console.warn("openBrowserWindow failed, falling back to window.open:", e);
  }
  window.open(draftWebUrl, "_blank");
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
  if (!selectedProject.projectFolderUrl) {
    if (crumbs)  crumbs.textContent  = "This project has no SharePoint folder linked.";
    if (folders) folders.innerHTML   = '<div class="sp-empty">Create one in the PMS first to enable saving here.</div>';
    updateDestCard();
    return;
  }
  spProjectRootPath = spDrivePath(selectedProject.projectFolderUrl) || "";
  spCurrentPath     = spProjectRootPath;
  // Once a draft has been filed, exports default to the folder it went to so
  // the PDF lands next to the .docx.
  if (draftSaved && draftFolderPath) spCurrentPath = draftFolderPath;
  updateDestCard();
  await renderSpPicker();
}

async function renderSpPicker() {
  const crumbs = document.getElementById("spBreadcrumbs");
  const folders = document.getElementById("spFolders");
  // Breadcrumbs: project-root-relative
  const rel = spCurrentPath.startsWith(spProjectRootPath)
    ? spCurrentPath.slice(spProjectRootPath.length).replace(/^\/+/, "")
    : spCurrentPath;
  const parts = rel ? rel.split("/") : [];
  const crumbHtml = ['<span data-depth="0">📁 (project root)</span>']
    .concat(parts.map((p, i) => ` / <span data-depth="${i + 1}">${escapeHtml(p)}</span>`))
    .join("");
  crumbs.innerHTML = "Save to: " + crumbHtml;
  // Mirror the current folder into the destination card display.
  const destDisplay = document.getElementById("destFolderDisplay");
  if (destDisplay) destDisplay.textContent = parts.length ? parts.join(" / ") : "project root";
  updateSavePreview();
  crumbs.querySelectorAll("span").forEach(s => {
    s.onclick = async () => {
      const depth = parseInt(s.getAttribute("data-depth"), 10);
      spCurrentPath = depth === 0
        ? spProjectRootPath
        : spProjectRootPath + "/" + parts.slice(0, depth).join("/");
      await renderSpPicker();
    };
  });
  folders.innerHTML = '<div class="sp-loading">Loading folders…</div>';
  try {
    const token = await getToken();
    const url = "https://graph.microsoft.com/v1.0/drives/" + SP_DRIVE_ID
      + "/root:/" + encodeDrivePath(spCurrentPath)
      + ":/children?$select=name,folder&$filter=folder ne null&$top=200";
    const res = await fetch(url, { headers: { "Authorization": "Bearer " + token } });
    if (!res.ok) {
      folders.innerHTML = '<div class="sp-empty">Could not list folders (' + res.status + ')</div>';
      return;
    }
    const data = await res.json();
    const subs = (data.value || []).filter(it => it.folder);
    if (!subs.length) {
      folders.innerHTML = '<div class="sp-empty">No subfolders here. Click Send to save at this location.</div>';
      return;
    }
    folders.innerHTML = subs.map(f =>
      `<div class="sp-folder-row" data-name="${escapeHtml(f.name)}">📁 ${escapeHtml(f.name)}</div>`
    ).join("");
    folders.querySelectorAll(".sp-folder-row").forEach(row => {
      row.onclick = async () => {
        spCurrentPath = spCurrentPath + "/" + row.getAttribute("data-name");
        await renderSpPicker();
      };
    });
  } catch (e) {
    folders.innerHTML = '<div class="sp-empty">Error: ' + escapeHtml(e.message) + '</div>';
  }
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
  if (/^https?:\/\//i.test(currentDocUrl || "")) return;     // cloud copy AutoSaves — never "unsynced"
  dirtySinceFiled = true;
  updateFiledCard();
}

// ─── PROJECTS ─────────────────────────────────────────────────────────────────
async function loadProjects() {
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
  const spReady  = !!(selectedProject && selectedProject.projectFolderUrl);

  // Relabel via the .brand-label span only — writing textContent on the button
  // itself would wipe the icon tile inside it.
  oneBtn.disabled = !selectedProject || saveInFlight;
  document.getElementById("saveOneNoteLabel").textContent = "Save to OneNote" + projTag;

  draftBtn.disabled = !spReady || saveInFlight;
  // Re-saving is allowed: a filed draft can be updated in place, so the button
  // stays put and just relabels to make the repeat action obvious.
  document.getElementById("saveSpDraftLabel").textContent =
    (draftSaved ? "Update Draft" : "Save Draft") + (spReady ? projTag : "");
  draftBtn.style.display = "";

  pdfBtn.disabled = !spReady || saveInFlight;
  document.getElementById("savePdfLabel").textContent =
    "Export PDF to SharePoint" + (spReady ? projTag : "");

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
  if (!project.projectFolderUrl) {
    throw new Error("Project has no SharePoint folder linked. Create one in the PMS first.");
  }
  const targetPath = targetPathOverride || spDrivePath(project.projectFolderUrl);
  if (!targetPath) throw new Error("Project SharePoint folder URL is not in the expected library.");
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
  const presentation =
    `<!DOCTYPE html><html><head><title>${safeTitle}</title><meta name="created" content="${created}" /></head><body>` +
    `<div style="border-bottom:2px solid #003865;padding-bottom:8px;margin-bottom:16px;font-family:sans-serif">` +
    `<span style="background:#003865;color:#fff;padding:2px 8px;border-radius:3px;font-size:11px;margin-right:6px">${safeProj}</span>` +
    `<span style="font-size:11px;color:#666">${escapeHtml(fileName)}</span>` +
    `</div>` +
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
  if (!selectedProject.projectFolderUrl) {
    setStatus("spStatus", "error", "Project has no SharePoint folder linked. Create one in the PMS first.");
    return;
  }
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
  if (!selectedProject.projectFolderUrl) {
    setStatus("spStatus", "error", "Project has no SharePoint folder linked. Create one in the PMS first.");
    return;
  }
  if (saveInFlight) return;
  saveInFlight = true;
  updateSaveButtons();
  document.getElementById("spLink").innerHTML = "";
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
  } catch (e) {
    setStatus("spStatus", "error", e.message);
  } finally {
    isExporting = false;
    saveInFlight = false;
    updateSaveButtons();
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
