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

// ─── STATE ────────────────────────────────────────────────────────────────────
let msalApp = null;
let msalAccount = null;
let allProjects = [];
let selectedProject = null;
let docFilename = "";
let docFirstPageText = "";
let saveInFlight = false;
// SharePoint folder picker state. spCurrentPath is drive-relative (e.g.
// "24-105 Acme HVAC/Documents"). Defaults to the project folder root when
// the SharePoint checkbox is first ticked.
let spCurrentPath = "";
let spProjectRootPath = "";

// ─── INIT ─────────────────────────────────────────────────────────────────────
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Word) {
    document.body.innerHTML = '<p style="color:#f87171;padding:16px;font-family:sans-serif">This add-in only runs in Word.</p>';
    return;
  }
  try {
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
    showView("signInView");
    setStatus("signInStatus", "error", "Startup error: " + e.message);
  }
});

function setupListeners() {
  document.getElementById("signInBtn").onclick     = doSignIn;
  document.getElementById("signOutBtn").onclick    = doSignOut;
  document.getElementById("saveOneNoteBtn").onclick = doSaveToOneNote;
  document.getElementById("saveSpBtn").onclick      = doSaveToSharePoint;
  document.getElementById("searchInput").addEventListener("input", () => renderProjectList());
}

// Refresh SP picker for the currently-selected project (or empty it if none).
// Called whenever selectedProject changes.
async function refreshSpPickerForProject() {
  const folders = document.getElementById("spFolders");
  const crumbs = document.getElementById("spBreadcrumbs");
  if (!selectedProject) {
    crumbs.textContent = "Pick a project to enable folder browsing.";
    folders.innerHTML = "";
    return;
  }
  if (!selectedProject.projectFolderUrl) {
    crumbs.textContent = "This project has no SharePoint folder linked.";
    folders.innerHTML = '<div class="sp-empty">Create one in the PMS first to enable saving here.</div>';
    return;
  }
  spProjectRootPath = spDrivePath(selectedProject.projectFolderUrl) || "";
  spCurrentPath = spProjectRootPath;
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
  await Promise.all([loadProjects(), loadDocumentContext()]);
  renderProjectList();
  applyAutoSuggest();
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
      }
    }
  } catch {}

  try {
    const res = await fetch(SUPABASE_URL + "/rest/v1/pms_projects?select=id,project,version", { headers: SB_HEADERS });
    if (res.ok) {
      const rows = await res.json();
      if (rows && rows.length > 0) {
        allProjects = rows.map(r => r.project).filter(p => p && !p.archived);
        try {
          localStorage.setItem(PROJECTS_CACHE_KEY, JSON.stringify({ projects: allProjects, savedAt: Date.now() }));
        } catch {}
        return;
      }
    }
  } catch (e) {
    console.warn("V2 load failed, trying legacy:", e.message);
  }
  // Legacy fallback
  try {
    const res = await fetch(SUPABASE_URL + "/rest/v1/pms_data?id=eq.singleton&select=projects", { headers: SB_HEADERS });
    const rows = await res.json();
    if (rows?.[0]?.projects) {
      allProjects = rows[0].projects.filter(p => !p.archived);
    }
  } catch (e) {
    console.error("Failed to load projects:", e);
  }
}

// ─── DOCUMENT CONTEXT ─────────────────────────────────────────────────────────
async function loadDocumentContext() {
  // Filename
  docFilename = (Office.context.document.url || "")
    .split(/[\\/]/).pop() || "Untitled.docx";
  document.getElementById("docFilename").textContent = docFilename;
  // Pre-fill page title with filename minus extension
  const titleInput = document.getElementById("titleInput");
  titleInput.value = docFilename.replace(/\.[^.]+$/, "");

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
    renderProjectList(scored.map(s => s.p.id));
    updateSaveButtons();
    refreshSpPickerForProject();
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
    row.onclick = () => {
      const id = row.getAttribute("data-id");
      selectedProject = allProjects.find(p => p.id === id) || null;
      renderProjectList(suggestedIds);
      updateSaveButtons();
      refreshSpPickerForProject();
    };
  });
}

function updateSaveButtons() {
  const oneBtn = document.getElementById("saveOneNoteBtn");
  const spBtn  = document.getElementById("saveSpBtn");
  const projTag = selectedProject ? ` → ${selectedProject.projectNumber || selectedProject.name}` : "";
  oneBtn.disabled = !selectedProject || saveInFlight;
  oneBtn.textContent = "Save to OneNote" + projTag;
  // SharePoint requires a project AND a linked SharePoint folder
  const spReady = !!(selectedProject && selectedProject.projectFolderUrl);
  spBtn.disabled = !spReady || saveInFlight;
  spBtn.textContent = "Save to SharePoint" + (spReady ? projTag : "");
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
          slices[idx] = sliceRes.value.data;
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
async function uploadDocxToSharePoint(project, docxBlob, filename, targetPathOverride) {
  if (!project.projectFolderUrl) {
    throw new Error("Project has no SharePoint folder linked. Create one in the PMS first.");
  }
  const targetPath = targetPathOverride || spDrivePath(project.projectFolderUrl);
  if (!targetPath) throw new Error("Project SharePoint folder URL is not in the expected library.");
  const token = await getToken();
  const safeName = filename.replace(/[\\/:*?"<>|]/g, "_");
  // conflictBehavior=rename appends " (1)" / " (2)" so re-saves don't overwrite
  // prior versions. SharePoint version history would track replaces too, but
  // rename is more discoverable for non-power-users browsing the folder.
  const url = "https://graph.microsoft.com/v1.0/drives/" + SP_DRIVE_ID
    + "/root:/" + encodeDrivePath(targetPath + "/" + safeName)
    + ":/content?@microsoft.graph.conflictBehavior=rename";
  const res = await fetch(url, {
    method: "PUT",
    headers: {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    },
    body: docxBlob,
  });
  if (!res.ok) {
    const errText = await res.text().catch(() => "");
    throw new Error("SharePoint " + res.status + ": " + errText.slice(0, 300));
  }
  const item = await res.json();
  return { name: item.name, webUrl: item.webUrl };
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
          slices[idx] = sliceRes.value.data; // Uint8Array
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
  const fileName  = docFilename || "document.pdf";

  // OneNote multipart: HTML "Presentation" part + named PDF binary part.
  // The <object> with type="application/pdf" tells OneNote to render the PDF
  // as a printout (one image per page) inside the note.
  const presentation =
    `<!DOCTYPE html><html><head><title>${safeTitle}</title><meta name="created" content="${created}" /></head><body>` +
    `<div style="border-bottom:2px solid #003865;padding-bottom:8px;margin-bottom:16px;font-family:sans-serif">` +
    `<span style="background:#003865;color:#fff;padding:2px 8px;border-radius:3px;font-size:11px;margin-right:6px">${safeProj}</span>` +
    `<span style="font-size:11px;color:#666">${escapeHtml(fileName)}</span>` +
    `</div>` +
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
    setStatus("oneNoteStatus", "info", "⏳ Exporting document as PDF…");
    const pdfBlob = await getDocumentAsPdfBlob();
    setStatus("oneNoteStatus", "info", "⏳ Sending to OneNote…");
    const page = await postPdfPrintoutToOneNote(selectedProject, pageTitle, pdfBlob);
    setStatus("oneNoteStatus", "success", "✓ Saved to OneNote");
    if (page.webUrl) {
      document.getElementById("oneNoteLink").innerHTML =
        `<a href="${page.webUrl}" target="_blank">📓 Open in OneNote</a>`;
    }
  } catch (e) {
    setStatus("oneNoteStatus", "error", e.message);
  } finally {
    saveInFlight = false;
    updateSaveButtons();
  }
}

async function doSaveToSharePoint() {
  if (!selectedProject) { setStatus("spStatus", "error", "Pick a project first."); return; }
  if (!selectedProject.projectFolderUrl) {
    setStatus("spStatus", "error", "Project has no SharePoint folder linked. Create one in the PMS first.");
    return;
  }
  if (saveInFlight) return;
  saveInFlight = true;
  updateSaveButtons();
  document.getElementById("spLink").innerHTML = "";
  try {
    setStatus("spStatus", "info", "⏳ Exporting document…");
    const docxBlob = await getDocumentAsDocxBlob();
    setStatus("spStatus", "info", "⏳ Uploading to SharePoint…");
    const item = await uploadDocxToSharePoint(selectedProject, docxBlob, docFilename, spCurrentPath);
    setStatus("spStatus", "success", "✓ Saved to SharePoint");
    if (item.webUrl) {
      document.getElementById("spLink").innerHTML =
        `<a href="${item.webUrl}" target="_blank">📁 Open in SharePoint (${escapeHtml(item.name)})</a>`;
    }
  } catch (e) {
    setStatus("spStatus", "error", e.message);
  } finally {
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
