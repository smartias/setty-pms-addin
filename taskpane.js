// ─── CONFIG ───────────────────────────────────────────────────────────────────
const MSAL_CONFIG = {
  auth: {
    clientId: "c4739c11-e89b-4a04-9580-f2d886356301",
    authority: "https://login.microsoftonline.com/f374c024-71c2-48b6-8420-076fff97327c",
    redirectUri: "https://smartias.github.io/setty-pms-addin/taskpane.html",
  },
  cache: { cacheLocation: "sessionStorage" }
};
const GRAPH_SCOPES = [
  "User.Read",
  "Mail.Read",
  "Files.ReadWrite.All",
  "Sites.Read.All",
];
const SUPABASE_URL  = "https://khxmgjilwhdguuepbhne.supabase.co";
const SUPABASE_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtoeG1namlsd2hkZ3V1ZXBiaG5lIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMwNjg2MDYsImV4cCI6MjA4ODY0NDYwNn0.vtHt2eydU2iQ426iYOzLrqpH2WLXdRnicq-3sNfoNq8";
const SP_SITE      = "setty.sharepoint.com:/sites/NYCProjects:";
const SP_LIBRARY   = "Project Document Library";

// ─── STATE ────────────────────────────────────────────────────────────────────
let msalApp = null;
let msalAccount = null;
let allProjects = [];
let selectedProject = null;
let emailItem = null;
let emailBody = "";
let emailFrom = "";
let emailFromAddress = "";
let _spIds = null; // cached per session — avoids repeated site/drive lookups

// ─── INIT ─────────────────────────────────────────────────────────────────────
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) return;
  msalApp = new msal.PublicClientApplication(MSAL_CONFIG);
  await msalApp.initialize();

  // Try silent sign-in first
  const accounts = msalApp.getAllAccounts();
  if (accounts.length > 0) {
    msalAccount = accounts[0];
    await onSignedIn();
  } else {
    showView("signInView");
  }

  setupEventListeners();
  loadEmailContext();
});

function setupEventListeners() {
  document.getElementById("signInBtn").onclick     = doSignIn;
  document.getElementById("signOutBtn").onclick    = doSignOut;
  document.getElementById("saveSpBtn").onclick     = doSaveToSharePoint;
  document.getElementById("logNoteBtn").onclick    = () => showView("noteView");
  document.getElementById("logRfiBtn").onclick     = () => { prefillRfi(); showView("rfiView"); };
  document.getElementById("logSubBtn").onclick     = () => showView("subView");
  document.getElementById("extractContactBtn").onclick = doExtractContact;

  document.getElementById("noteBack").onclick    = () => showView("mainView");
  document.getElementById("rfiBack").onclick     = () => showView("mainView");
  document.getElementById("subBack").onclick     = () => showView("mainView");
  document.getElementById("contactBack").onclick = () => showView("mainView");

  document.getElementById("saveNoteBtn").onclick    = doSaveNote;
  document.getElementById("saveRfiBtn").onclick     = doSaveRfi;
  document.getElementById("saveSubBtn").onclick     = doSaveSub;
  document.getElementById("saveContactBtn").onclick = doSaveContact;

  // Project search
  const searchInput = document.getElementById("projectSearch");
  const dropdown    = document.getElementById("projectDropdown");
  searchInput.addEventListener("input", () => {
    const q = searchInput.value.trim().toLowerCase();
    if (!q) { dropdown.style.display = "none"; return; }
    const matches = allProjects.filter(p =>
      (p.name || "").toLowerCase().includes(q) ||
      (p.projectNumber || "").toLowerCase().includes(q)
    ).slice(0, 10);
    if (!matches.length) { dropdown.style.display = "none"; return; }
    dropdown.innerHTML = matches.map(p => `
      <div class="proj-option" data-id="${p.id}">
        <div class="proj-num">${p.projectNumber || ""}</div>
        <div class="proj-name">${p.name || ""}</div>
      </div>
    `).join("");
    dropdown.style.display = "block";
    dropdown.querySelectorAll(".proj-option").forEach(el => {
      el.onclick = () => {
        selectedProject = allProjects.find(p => p.id === el.dataset.id);
        searchInput.value = "";
        dropdown.style.display = "none";
        const badge = document.getElementById("selectedProjectBadge");
        badge.textContent = "✓ " + (selectedProject.projectNumber ? selectedProject.projectNumber + " — " : "") + selectedProject.name;
        badge.style.display = "block";
      };
    });
  });
  document.addEventListener("click", (e) => {
    if (!e.target.closest(".project-search-wrapper")) dropdown.style.display = "none";
  });
}

// ─── EMAIL CONTEXT ────────────────────────────────────────────────────────────
function loadEmailContext() {
  emailItem = Office.context.mailbox.item;
  if (!emailItem) return;

  document.getElementById("emailSubject").textContent = emailItem.subject || "(No subject)";
  const from = emailItem.from;
  emailFrom = from?.displayName || "";
  emailFromAddress = from?.emailAddress || "";
  const date = emailItem.dateTimeCreated ? new Date(emailItem.dateTimeCreated).toLocaleDateString("en-US", {
    month: "short", day: "numeric", year: "numeric"
  }) : "";
  document.getElementById("emailMeta").textContent =
    "From: " + (emailFrom || emailFromAddress) + (date ? "  ·  " + date : "");

  // Pre-fill note body
  document.getElementById("noteBody").value = emailItem.subject || "";

  // Pre-fill RFI from
  document.getElementById("rfiFrom").value = emailFrom;
  document.getElementById("subFrom").value = emailFrom;
}

// ─── AUTH ─────────────────────────────────────────────────────────────────────
async function doSignIn() {
  setStatus("signInStatus", "info", "⏳ Signing in…");
  try {
    const result = await msalApp.loginPopup({ scopes: GRAPH_SCOPES });
    msalAccount = result.account;
    await onSignedIn();
  } catch (e) {
    setStatus("signInStatus", "error", "✗ Sign-in failed: " + e.message);
  }
}

async function doSignOut() {
  await msalApp.logoutPopup({ account: msalAccount });
  msalAccount = null;
  selectedProject = null;
  allProjects = [];
  showView("signInView");
}

async function onSignedIn() {
  showView("mainView");
  await loadProjects();
}

async function getToken() {
  try {
    const r = await msalApp.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: msalAccount });
    return r.accessToken;
  } catch {
    const r = await msalApp.acquireTokenPopup({ scopes: GRAPH_SCOPES, account: msalAccount });
    return r.accessToken;
  }
}

// ─── SUPABASE ─────────────────────────────────────────────────────────────────
const SB_HEADERS = {
  "apikey": SUPABASE_ANON,
  "Authorization": "Bearer " + SUPABASE_ANON,
  "Content-Type": "application/json",
  "Prefer": "return=minimal",
};

async function loadProjects() {
  try {
    const res = await fetch(SUPABASE_URL + "/rest/v1/pms_data?id=eq.singleton&select=projects", {
      headers: SB_HEADERS,
    });
    const rows = await res.json();
    if (!rows || !rows[0]) return;
    allProjects = (rows[0].projects || []).filter(p => !p.archived);
  } catch (e) {
    console.error("Failed to load projects:", e);
  }
}

async function saveToSupabase(updatedProjects) {
  await fetch(SUPABASE_URL + "/rest/v1/pms_data?id=eq.singleton", {
    method: "PATCH",
    headers: SB_HEADERS,
    body: JSON.stringify({ projects: updatedProjects, updated_at: new Date().toISOString() }),
  });
}

function updateProjectInList(updatedProject) {
  allProjects = allProjects.map(p => p.id === updatedProject.id ? updatedProject : p);
}

// ─── GRAPH HELPERS ────────────────────────────────────────────────────────────
async function graphFetch(method, path, body, token) {
  const t = token || await getToken();
  const res = await fetch("https://graph.microsoft.com/v1.0" + path, {
    method,
    headers: {
      "Authorization": "Bearer " + t,
      "Content-Type": "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error("Graph " + res.status + ": " + err.slice(0, 200));
  }
  return res.status === 204 ? null : res.json();
}

async function resolveSpIds() {
  if (_spIds) return _spIds;
  const site = await graphFetch("GET", "/sites/" + SP_SITE);
  const siteId = site.id;
  const drives = await graphFetch("GET", "/sites/" + siteId + "/drives");
  const drive = (drives.value || []).find(d => d.name === SP_LIBRARY || d.name.toLowerCase() === SP_LIBRARY.toLowerCase());
  if (!drive) throw new Error("Document library not found: " + SP_LIBRARY);
  _spIds = { siteId, driveId: drive.id };
  return _spIds;
}

async function getEmailBodyHtml(token) {
  try {
    const msgId = Office.context.mailbox.item.itemId;
    const restId = Office.context.mailbox.convertToRestId(msgId, Office.MailboxEnums.RestVersion.v2_0);
    const data = await graphFetch("GET", "/me/messages/" + restId + "?$select=body", null, token);
    return data?.body?.content || "";
  } catch { return ""; }
}

// ─── SP / EMAIL HELPERS ──────────────────────────────────────────────────────

const SP_BASE_URL = "https://setty.sharepoint.com/sites/NYCProjects/Project%20Document%20Library";

// Create a folder idempotently (conflictBehavior:replace is a no-op on existing folders)
async function ensureSpFolder(driveId, token, parentPath, name) {
  try {
    await fetch("https://graph.microsoft.com/v1.0/drives/" + driveId + "/root:/" + encodeURIComponent(parentPath) + ":/children", {
      method: "POST",
      headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "replace" }),
    });
  } catch {}
  return parentPath + "/" + name;
}

// Extract the drive-relative path from a full SharePoint web URL
function spDrivePath(spFolderUrl) {
  const base = SP_BASE_URL + "/";
  if (!spFolderUrl || !spFolderUrl.startsWith(base)) return null;
  return decodeURIComponent(spFolderUrl.slice(base.length));
}

// Build the email HTML file content from the current emailItem
function buildEmailHtml(bodyHtml) {
  const from = emailItem.from;
  const header = `<div style="font-family:sans-serif;font-size:12px;padding:12px 16px;border-bottom:1px solid #ddd;margin-bottom:16px">
    <strong>Subject:</strong> ${emailItem.subject || ""}<br>
    <strong>From:</strong> ${from?.displayName || ""} &lt;${from?.emailAddress || ""}&gt;<br>
    <strong>Date:</strong> ${new Date(emailItem.dateTimeCreated).toLocaleString()}
  </div>`;
  return "<!DOCTYPE html><html><head><meta charset='utf-8'></head><body>" + header + (bodyHtml || "(No body)") + "</body></html>";
}

// Upload email.html + any attachments into targetPath. Returns attachment count.
async function uploadEmailAndAttachments(driveId, token, targetPath) {
  const bodyHtml = await getEmailBodyHtml(token);
  await fetch("https://graph.microsoft.com/v1.0/drives/" + driveId + "/root:/" + encodeURIComponent(targetPath) + "/email.html:/content", {
    method: "PUT",
    headers: { "Authorization": "Bearer " + token, "Content-Type": "text/html" },
    body: buildEmailHtml(bodyHtml),
  });

  if (!emailItem.hasAttachments) return 0;
  try {
    const restId = Office.context.mailbox.convertToRestId(emailItem.itemId, Office.MailboxEnums.RestVersion.v2_0);
    const attData = await graphFetch("GET", "/me/messages/" + restId + "/attachments", null, token);
    let count = 0;
    for (const att of (attData?.value || [])) {
      if (!att.contentBytes) continue;
      const binary = atob(att.contentBytes);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      const safeName = (att.name || "attachment").replace(/[\\/:*?"<>|]/g, "-").trim();
      await fetch("https://graph.microsoft.com/v1.0/drives/" + driveId + "/root:/" + encodeURIComponent(targetPath) + "/" + encodeURIComponent(safeName) + ":/content", {
        method: "PUT",
        headers: { "Authorization": "Bearer " + token, "Content-Type": att.contentType || "application/octet-stream" },
        body: bytes,
      });
      count++;
    }
    return count;
  } catch (e) {
    console.warn("Attachment upload failed:", e.message);
    return 0;
  }
}

// ─── SAVE TO SHAREPOINT ───────────────────────────────────────────────────────
async function doSaveToSharePoint() {
  if (!selectedProject) { setStatus("actionStatus", "error", "Select a project first."); return; }
  if (!selectedProject.projectFolderUrl) { setStatus("actionStatus", "error", "No SharePoint folder on this project. Create one in the PMS first."); return; }

  setStatus("actionStatus", "info", "⏳ Saving to SharePoint…");
  try {
    const token = await getToken();
    const { driveId } = await resolveSpIds();
    const projFolderName = decodeURIComponent(selectedProject.projectFolderUrl.split("/").pop());

    const d = new Date(emailItem.dateTimeCreated);
    const safeSubject = (emailItem.subject || "No Subject").replace(/[\\/:*?"<>|]/g, "-").replace(/\s+/g, " ").trim().slice(0, 70);
    const emailFolderName = d.getFullYear() + "_" + String(d.getMonth() + 1).padStart(2, "0") + "_" + String(d.getDate()).padStart(2, "0") + " " + safeSubject;

    const emailsPath  = await ensureSpFolder(driveId, token, projFolderName, "Emails");
    const targetPath  = await ensureSpFolder(driveId, token, emailsPath, emailFolderName);
    const attCount    = await uploadEmailAndAttachments(driveId, token, targetPath);

    const from = emailItem.from;
    const spFolderUrl = SP_BASE_URL + "/" + encodeURIComponent(projFolderName) + "/Emails/" + encodeURIComponent(emailFolderName);
    const msgId = Office.context.mailbox.convertToRestId(emailItem.itemId, Office.MailboxEnums.RestVersion.v2_0);
    const emailRecord = {
      id: uid(), msgId,
      subject: emailItem.subject || "",
      from: from?.displayName || "",
      fromAddress: from?.emailAddress || "",
      date: emailItem.dateTimeCreated,
      bodyText: "", spFolderUrl,
      savedAt: new Date().toISOString(),
    };
    const updatedProject = { ...selectedProject, emails: [...(selectedProject.emails || []), emailRecord] };
    updateProjectInList(updatedProject);
    selectedProject = updatedProject;
    await saveToSupabase(allProjects);

    const attMsg = attCount ? " + " + attCount + " attachment" + (attCount > 1 ? "s" : "") : "";
    setStatus("actionStatus", "success", "✓ Saved to SharePoint" + attMsg + " and project record.");
  } catch (e) {
    setStatus("actionStatus", "error", "✗ " + e.message);
  }
}

// ─── LOG NOTE ─────────────────────────────────────────────────────────────────
async function doSaveNote() {
  if (!selectedProject) { setStatus("noteStatus", "error", "No project selected."); return; }
  const category = document.getElementById("noteCategory").value;
  const body = document.getElementById("noteBody").value.trim();
  if (!body) { setStatus("noteStatus", "error", "Note body is empty."); return; }

  setStatus("noteStatus", "info", "⏳ Saving…");
  try {
    const note = {
      id: uid(),
      body,
      category,
      actionItem: false,
      author: msalAccount?.name || msalAccount?.username || "Unknown",
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      importedFromEmail: true,
    };
    const updated = { ...selectedProject, notes: [...(selectedProject.notes || []), note] };
    updateProjectInList(updated);
    selectedProject = updated;
    await saveToSupabase(allProjects);
    setStatus("noteStatus", "success", "✓ Note saved to " + selectedProject.name);
    document.getElementById("noteBody").value = "";
  } catch (e) {
    setStatus("noteStatus", "error", "✗ " + e.message);
  }
}

// ─── SHARED: file email+attachments into a project subfolder ─────────────────
async function uploadEmailUnderFolder(driveId, token, projFolderName, subfolder, recordFolderName) {
  const subPath    = await ensureSpFolder(driveId, token, projFolderName, subfolder);
  const recordPath = await ensureSpFolder(driveId, token, subPath, recordFolderName);
  await uploadEmailAndAttachments(driveId, token, recordPath);
  return SP_BASE_URL + "/" + encodeURIComponent(projFolderName) + "/" + encodeURIComponent(subfolder) + "/" + encodeURIComponent(recordFolderName);
}

// ─── RFI MODE TOGGLE ─────────────────────────────────────────────────────────
function setRfiMode(mode) {
  document.getElementById("rfiNewForm").style.display      = mode === "new"      ? "" : "none";
  document.getElementById("rfiExistingForm").style.display = mode === "existing" ? "" : "none";
  document.getElementById("rfiModeNew").className      = "btn mode-tab " + (mode === "new"      ? "btn-blue"  : "btn-ghost");
  document.getElementById("rfiModeExisting").className = "btn mode-tab " + (mode === "existing" ? "btn-blue"  : "btn-ghost");
}

function setSubMode(mode) {
  document.getElementById("subNewForm").style.display      = mode === "new"      ? "" : "none";
  document.getElementById("subExistingForm").style.display = mode === "existing" ? "" : "none";
  document.getElementById("subModeNew").className      = "btn mode-tab " + (mode === "new"      ? "btn-purple" : "btn-ghost");
  document.getElementById("subModeExisting").className = "btn mode-tab " + (mode === "existing" ? "btn-purple" : "btn-ghost");
}

function renderRfiPicker() {
  const sel  = document.getElementById("rfiExistingSelect");
  const rfis = selectedProject?.rfis || [];
  sel.innerHTML = rfis.length
    ? rfis.map(r => `<option value="${r.id}">${r.number}${r.title ? " — " + r.title.slice(0, 45) : ""}</option>`).join("")
    : '<option value="">No RFIs on this project</option>';
}

function renderSubPicker() {
  const sel  = document.getElementById("subExistingSelect");
  const subs = selectedProject?.submittals || [];
  sel.innerHTML = subs.length
    ? subs.map(s => `<option value="${s.id}">${s.number}${s.description ? " — " + s.description.slice(0, 45) : ""}</option>`).join("")
    : '<option value="">No submittals on this project</option>';
}

// ─── LOG RFI ──────────────────────────────────────────────────────────────────
function prefillRfi() {
  document.getElementById("rfiTitle").value = emailItem?.subject || "";
  setRfiMode("new");
  renderRfiPicker();
}

async function doSaveRfi() {
  if (!selectedProject) { setStatus("rfiStatus", "error", "No project selected."); return; }
  const title = document.getElementById("rfiTitle").value.trim();
  if (!title) { setStatus("rfiStatus", "error", "Title is required."); return; }

  setStatus("rfiStatus", "info", "⏳ Saving…");
  try {
    const existingRfis = selectedProject.rfis || [];
    const nextNum = "RFI-" + String(existingRfis.length + 1).padStart(3, "0");
    const received = new Date();

    let spFolderUrl = "";
    if (selectedProject.projectFolderUrl) {
      try {
        const token = await getToken();
        const { driveId } = await resolveSpIds();
        const projFolderName = decodeURIComponent(selectedProject.projectFolderUrl.split("/").pop());
        const safeName = (nextNum + " " + title).replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
        spFolderUrl = await uploadEmailUnderFolder(driveId, token, projFolderName, "RFIs", safeName);
      } catch (e) { console.warn("RFI SP upload failed:", e.message); }
    }

    const rfi = {
      id: uid(), number: nextNum, title,
      discipline: document.getElementById("rfiDiscipline").value,
      from: document.getElementById("rfiFrom").value.trim(),
      dateReceived: received.toISOString().slice(0, 10),
      dueDate: addBizDays(received, 5),
      status: "Open",
      notes: document.getElementById("rfiNotes").value.trim(),
      assignedTo: [], spFolderUrl,
      createdAt: new Date().toISOString(),
    };
    const updated = { ...selectedProject, rfis: [...existingRfis, rfi] };
    updateProjectInList(updated);
    selectedProject = updated;
    await saveToSupabase(allProjects);
    setStatus("rfiStatus", "success", "✓ " + nextNum + " logged" + (spFolderUrl ? " · filed to SharePoint" : ""));
    document.getElementById("rfiTitle").value = "";
    document.getElementById("rfiNotes").value = "";
  } catch (e) {
    setStatus("rfiStatus", "error", "✗ " + e.message);
  }
}

async function doFileToExistingRfi() {
  if (!selectedProject) { setStatus("rfiExistingStatus", "error", "Select a project first."); return; }
  const rfiId = document.getElementById("rfiExistingSelect").value;
  if (!rfiId) { setStatus("rfiExistingStatus", "error", "Select an RFI."); return; }

  const rfi = (selectedProject.rfis || []).find(r => r.id === rfiId);
  if (!rfi) { setStatus("rfiExistingStatus", "error", "RFI not found."); return; }

  setStatus("rfiExistingStatus", "info", "⏳ Filing email…");
  try {
    const token = await getToken();
    const { driveId } = await resolveSpIds();

    // Use the existing SP folder if set, otherwise create it under the project folder
    let targetPath = spDrivePath(rfi.spFolderUrl);
    if (!targetPath) {
      if (!selectedProject.projectFolderUrl) throw new Error("No SharePoint folder on this project. Create one in the PMS first.");
      const projFolderName = decodeURIComponent(selectedProject.projectFolderUrl.split("/").pop());
      const safeName = (rfi.number + " " + (rfi.title || "")).replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
      const rfisPath = await ensureSpFolder(driveId, token, projFolderName, "RFIs");
      targetPath = await ensureSpFolder(driveId, token, rfisPath, safeName);
    }

    const attCount = await uploadEmailAndAttachments(driveId, token, targetPath);

    // If this RFI didn't have a spFolderUrl yet, store it now
    if (!rfi.spFolderUrl) {
      const newUrl = SP_BASE_URL + "/" + targetPath.split("/").map(encodeURIComponent).join("/");
      const updatedRfis = (selectedProject.rfis || []).map(r => r.id === rfi.id ? { ...r, spFolderUrl: newUrl } : r);
      const updated = { ...selectedProject, rfis: updatedRfis };
      updateProjectInList(updated);
      selectedProject = updated;
      await saveToSupabase(allProjects);
    }

    const attMsg = attCount ? " + " + attCount + " attachment" + (attCount > 1 ? "s" : "") : "";
    setStatus("rfiExistingStatus", "success", "✓ Filed to " + rfi.number + attMsg);
  } catch (e) {
    setStatus("rfiExistingStatus", "error", "✗ " + e.message);
  }
}

// ─── LOG SUBMITTAL ────────────────────────────────────────────────────────────
function prefillSub() {
  setSubMode("new");
  renderSubPicker();
}

async function doSaveSub() {
  if (!selectedProject) { setStatus("subStatus", "error", "No project selected."); return; }
  const desc = document.getElementById("subDesc").value.trim();
  if (!desc) { setStatus("subStatus", "error", "Description is required."); return; }

  setStatus("subStatus", "info", "⏳ Saving…");
  try {
    const existing = selectedProject.submittals || [];
    const nextNum = "SUB-" + String(existing.length + 1).padStart(3, "0");
    const received = new Date();

    let spFolderUrl = "";
    if (selectedProject.projectFolderUrl) {
      try {
        const token = await getToken();
        const { driveId } = await resolveSpIds();
        const projFolderName = decodeURIComponent(selectedProject.projectFolderUrl.split("/").pop());
        const safeName = (nextNum + " " + desc).replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
        spFolderUrl = await uploadEmailUnderFolder(driveId, token, projFolderName, "Submittals", safeName);
      } catch (e) { console.warn("Submittal SP upload failed:", e.message); }
    }

    const sub = {
      id: uid(), number: nextNum,
      specSection: document.getElementById("subSpec").value.trim(),
      description: desc,
      discipline: document.getElementById("subDiscipline").value,
      from: document.getElementById("subFrom").value.trim(),
      dateReceived: received.toISOString().slice(0, 10),
      dueDate: addBizDays(received, 10),
      status: "Received",
      notes: document.getElementById("subNotes").value.trim(),
      assignedTo: [], spFolderUrl,
      createdAt: new Date().toISOString(),
    };
    const updated = { ...selectedProject, submittals: [...existing, sub] };
    updateProjectInList(updated);
    selectedProject = updated;
    await saveToSupabase(allProjects);
    setStatus("subStatus", "success", "✓ " + nextNum + " logged" + (spFolderUrl ? " · filed to SharePoint" : ""));
    document.getElementById("subDesc").value = "";
    document.getElementById("subSpec").value = "";
    document.getElementById("subNotes").value = "";
  } catch (e) {
    setStatus("subStatus", "error", "✗ " + e.message);
  }
}

async function doFileToExistingSub() {
  if (!selectedProject) { setStatus("subExistingStatus", "error", "Select a project first."); return; }
  const subId = document.getElementById("subExistingSelect").value;
  if (!subId) { setStatus("subExistingStatus", "error", "Select a submittal."); return; }

  const sub = (selectedProject.submittals || []).find(s => s.id === subId);
  if (!sub) { setStatus("subExistingStatus", "error", "Submittal not found."); return; }

  setStatus("subExistingStatus", "info", "⏳ Filing email…");
  try {
    const token = await getToken();
    const { driveId } = await resolveSpIds();

    let targetPath = spDrivePath(sub.spFolderUrl);
    if (!targetPath) {
      if (!selectedProject.projectFolderUrl) throw new Error("No SharePoint folder on this project. Create one in the PMS first.");
      const projFolderName = decodeURIComponent(selectedProject.projectFolderUrl.split("/").pop());
      const safeName = (sub.number + " " + (sub.description || "")).replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
      const subsPath = await ensureSpFolder(driveId, token, projFolderName, "Submittals");
      targetPath = await ensureSpFolder(driveId, token, subsPath, safeName);
    }

    const attCount = await uploadEmailAndAttachments(driveId, token, targetPath);

    if (!sub.spFolderUrl) {
      const newUrl = SP_BASE_URL + "/" + targetPath.split("/").map(encodeURIComponent).join("/");
      const updatedSubs = (selectedProject.submittals || []).map(s => s.id === sub.id ? { ...s, spFolderUrl: newUrl } : s);
      const updated = { ...selectedProject, submittals: updatedSubs };
      updateProjectInList(updated);
      selectedProject = updated;
      await saveToSupabase(allProjects);
    }

    const attMsg = attCount ? " + " + attCount + " attachment" + (attCount > 1 ? "s" : "") : "";
    setStatus("subExistingStatus", "success", "✓ Filed to " + sub.number + attMsg);
  } catch (e) {
    setStatus("subExistingStatus", "error", "✗ " + e.message);
  }
}

// ─── EXTRACT CONTACT ──────────────────────────────────────────────────────────
async function doExtractContact() {
  setStatus("actionStatus", "info", "⏳ Extracting contact…");
  const token = await getToken();
  const body = await getEmailBodyHtml(token);
  const contact = parseSignature(body, emailFrom, emailFromAddress);
  document.getElementById("contactName").value    = contact.name;
  document.getElementById("contactTitle").value   = contact.title;
  document.getElementById("contactCompany").value = contact.company;
  document.getElementById("contactEmail").value   = contact.email;
  document.getElementById("contactPhone").value   = contact.phone;
  setStatus("actionStatus", "info", "");
  showView("contactView");
}

function parseSignature(html, fromName, fromEmail) {
  // Strip HTML tags for text parsing
  const text = html.replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ").replace(/\s+/g, " ");

  // Phone: various formats
  const phoneMatch = text.match(/(?:Tel|Phone|Cell|Mobile|Direct|T|P|M|D)[:\.]?\s*([\+\d\s\(\)\-\.]{7,20})/i)
    || text.match(/\b(\+?1?[\s\-\.]?\(?\d{3}\)?[\s\-\.]?\d{3}[\s\-\.]?\d{4})\b/);
  const phone = phoneMatch ? phoneMatch[1].trim() : "";

  // Email — look for any email that isn't the from address
  const emailMatches = [...text.matchAll(/\b([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})\b/g)];
  const sigEmail = emailMatches.find(m => m[1].toLowerCase() !== (fromEmail || "").toLowerCase())?.[1] || fromEmail;

  // Title: look for common title patterns
  const titlePatterns = [
    /\b(Senior|Junior|Lead|Principal|Director|Manager|Engineer|Architect|Associate|Project|Design|MEP|VP|Vice President|President|Principal|Officer|Superintendent|Foreman|Coordinator|Analyst|Consultant|Officer|Executive)[^\n,|•\|]{0,50}/i,
    /(?:Title|Position|Role)[:\s]+([^\n,|•]{5,60})/i,
  ];
  let title = "";
  for (const pat of titlePatterns) {
    const m = text.match(pat);
    if (m) { title = m[0].trim().replace(/[|•].*$/, "").trim(); break; }
  }

  // Company: look for lines after signature break or before phone
  const companyPatterns = [
    /(?:Company|Firm|Organization|Corp|LLC|Inc|PLLC|PC|LLP|Associates|Group|Construction|Consulting)[^\n,|•]{0,60}/i,
    /(Setty|AECOM|WSP|Arup|Thornton|Langan|STV|Jacobs|Turner|Skanska|Suffolk|Structure Tone)[^\n,|•]{0,40}/i,
  ];
  let company = "";
  for (const pat of companyPatterns) {
    const m = text.match(pat);
    if (m) { company = m[0].trim().replace(/[|•].*$/, "").trim().slice(0, 60); break; }
  }

  return {
    name: fromName || "",
    title: title.slice(0, 80),
    company: company.slice(0, 80),
    email: sigEmail || "",
    phone: phone.replace(/\s+/g, " ").trim(),
  };
}

async function doSaveContact() {
  const name    = document.getElementById("contactName").value.trim();
  const title   = document.getElementById("contactTitle").value.trim();
  const company = document.getElementById("contactCompany").value.trim();
  const email   = document.getElementById("contactEmail").value.trim();
  const phone   = document.getElementById("contactPhone").value.trim();
  const type    = document.getElementById("contactType").value;
  const saveTo  = document.getElementById("contactSaveTo").value;

  if (!name && !email) { setStatus("contactStatus", "error", "Name or email required."); return; }
  setStatus("contactStatus", "info", "⏳ Saving…");

  try {
    const res = await fetch(SUPABASE_URL + "/rest/v1/pms_data?id=eq.singleton&select=clients,projects", {
      headers: SB_HEADERS,
    });
    const rows = await res.json();
    const patchBody = {};

    if (saveTo === "client") {
      const clients = rows[0]?.clients || [];
      const existing = clients.find(c => c.name && c.name.trim().toLowerCase() === (company || name).trim().toLowerCase());
      const contact = { id: uid(), name, title, email, phone, role: type };
      if (existing) {
        existing.contacts = [...(existing.contacts || []), contact];
      } else {
        clients.push({ id: uid(), name: company || name, type, contacts: [contact], address: "" });
      }
      patchBody.clients = clients;
    } else {
      if (!selectedProject) { setStatus("contactStatus", "error", "Select a project first."); return; }
      const poc = { id: uid(), name, title, email, phone, role: type };
      const projects = rows[0]?.projects || allProjects;
      const proj = projects.find(p => p.id === selectedProject.id);
      if (proj) {
        proj.projectContacts = proj.projectContacts || {};
        proj.projectContacts.pm = [...(proj.projectContacts.pm || []), poc];
      }
      patchBody.projects = projects;
    }

    await fetch(SUPABASE_URL + "/rest/v1/pms_data?id=eq.singleton", {
      method: "PATCH",
      headers: SB_HEADERS,
      body: JSON.stringify({ ...patchBody, updated_at: new Date().toISOString() }),
    });

    setStatus("contactStatus", "success", "✓ Contact saved.");
  } catch (e) {
    setStatus("contactStatus", "error", "✗ " + e.message);
  }
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function uid() {
  return "addin-" + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function addBizDays(date, days) {
  const d = new Date(date);
  let added = 0;
  while (added < days) {
    d.setDate(d.getDate() + 1);
    if (d.getDay() !== 0 && d.getDay() !== 6) added++;
  }
  return d.toISOString().slice(0, 10);
}

function showView(id) {
  ["signInView","mainView","noteView","rfiView","subView","contactView"].forEach(v => {
    const el = document.getElementById(v);
    if (el) el.classList.toggle("active", v === id);
  });
}

function setStatus(elId, type, msg) {
  const el = document.getElementById(elId);
  if (!el) return;
  el.className = "status-msg" + (msg ? " show " + type : "");
  el.textContent = msg;
}
