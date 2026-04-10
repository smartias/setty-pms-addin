// ─── CONFIG ───────────────────────────────────────────────────────────────────
const MSAL_CONFIG = {
  auth: {
    clientId: "8e5155fb-6221-4508-97ea-3661438c6688",
    authority: "https://login.microsoftonline.com/f374c024-71c2-48b6-8420-076fff97327c",
    redirectUri: "https://smartias.github.io/setty-pms-addin/taskpane.html",
  },
  // Persist login across taskpane reloads while users open different emails.
  cache: { cacheLocation: "localStorage" }
};
const GRAPH_SCOPES = [
  "User.Read",
  "Mail.Read",
  "Files.ReadWrite.All",
  "Calendars.ReadWrite.Shared",
  // Sites.Read.All removed — site and drive IDs are hardcoded below (no admin consent needed)
];
const SUPABASE_URL  = "https://khxmgjilwhdguuepbhne.supabase.co";
const SUPABASE_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtoeG1namlsd2hkZ3V1ZXBiaG5lIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMwNjg2MDYsImV4cCI6MjA4ODY0NDYwNn0.vtHt2eydU2iQ426iYOzLrqpH2WLXdRnicq-3sNfoNq8";
const PMS_PROJECT_BASE_URL = "https://smartias.github.io/setty-pms/SettyPMS.html#project:";
const PMS_DASHBOARD_URL = "https://smartias.github.io/setty-pms/SettyPMS.html#dashboard";
const SP_SITE      = "setty.sharepoint.com:/sites/NYCProjects:";
const SP_LIBRARY   = "Project Document Library";

// ─── STATE ────────────────────────────────────────────────────────────────────
let msalApp = null;
let msalAccount = null;
let allProjects = [];
let allClients = [];
let selectedProject = null;
let emailItem = null;
let emailBody = "";
let emailFrom = "";
let emailFromAddress = "";
let emailParticipants = []; // { label, displayName, emailAddress }
let lastAttachmentUploadStats = null;
// Hardcoded SharePoint IDs — eliminates Sites.Read.All (the only admin-consent scope).
// Retrieved once via https://setty.sharepoint.com/sites/NYCProjects/_api/v2.0/drives
const SP_SITE_ID_HARDCODED  = "setty.sharepoint.com,aa580464-13e9-4eb4-8ad4-ca6ff5b9e001,c97a67e8-fb1b-4a23-a29a-753a5d57d410";
const SP_DRIVE_ID_HARDCODED = "b!ZARYqukTtE6K1Mpv9bngAehneskb-yNKopp1Ol1X1BBnJPKsNGM-TaGmbGiL3ZaU";
let _spIds = { siteId: SP_SITE_ID_HARDCODED, driveId: SP_DRIVE_ID_HARDCODED };
const LAST_ACCOUNT_STORAGE_KEY = "settyPms:lastMsalAccountHomeId";
const EMAIL_PROJECT_MAP_STORAGE_KEY = "settyPms:emailProjectMap";

// ─── INIT ─────────────────────────────────────────────────────────────────────
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) {
    document.body.innerHTML = '<p style="color:#f87171;padding:16px;font-family:sans-serif;">This add-in only runs in Outlook.</p>';
    return;
  }
  try {
    msalApp = new msal.PublicClientApplication(MSAL_CONFIG);
    await msalApp.initialize();

    const accounts = msalApp.getAllAccounts();
    if (accounts.length > 0) {
      const lastAccountId = localStorage.getItem(LAST_ACCOUNT_STORAGE_KEY);
      msalAccount = accounts.find(a => a.homeAccountId === lastAccountId) || accounts[0];
      msalApp.setActiveAccount(msalAccount);
      await onSignedIn();
    } else {
      showView("signInView");
    }

    setupEventListeners();
    loadEmailContext();
  } catch (e) {
    // Show something rather than a black screen if init fails
    showView("signInView");
    setStatus("signInStatus", "error", "Startup error: " + e.message);
  }
});

function setupEventListeners() {
  document.getElementById("signInBtn").onclick     = doSignIn;
  document.getElementById("signOutBtn").onclick    = doSignOut;
  document.getElementById("saveSpBtn").onclick     = doSaveToSharePoint;
  document.getElementById("saveRecordBtn").onclick = doSaveToProjectRecordOnly;
  document.getElementById("logNoteBtn").onclick    = () => showView("noteView");
  document.getElementById("logRfiBtn").onclick       = () => { prefillRfi(); showView("rfiView"); };
  document.getElementById("logSubBtn").onclick       = () => { prefillSub(); showView("subView"); };
  document.getElementById("extractContactBtn").onclick = doExtractContact;

  document.getElementById("noteBack").onclick    = () => showView("mainView");
  document.getElementById("rfiBack").onclick     = () => showView("mainView");
  document.getElementById("subBack").onclick     = () => showView("mainView");
  document.getElementById("peopleBack").onclick  = () => showView("mainView");
  document.getElementById("contactBack").onclick = () => showView("mainView");
  document.getElementById("datesBack").onclick   = () => showView("mainView");

  document.getElementById("findDatesBtn").onclick    = showDatesView;
  document.getElementById("manualMilestoneBtn").onclick = showManualMilestoneForm;
  document.getElementById("addParticipantBtn").onclick = showPeopleView;
  document.getElementById("saveMilestoneBtn").onclick = doSaveMilestone;

  document.getElementById("saveNoteBtn").onclick    = doSaveNote;
  document.getElementById("saveRfiBtn").onclick     = doSaveRfi;
  document.getElementById("saveSubBtn").onclick     = doSaveSub;
  document.getElementById("saveContactBtn").onclick = doSaveContact;
  document.getElementById("openPmsBtn").onclick = openSelectedProjectInPms;
  document.getElementById("openSpFolderBtn").onclick = openSelectedProjectSpFolder;
  document.getElementById("openDashboardBtn").onclick = openPmsDashboard;

  // RFI mode toggles
  document.getElementById("rfiModeNew").onclick      = () => setRfiMode("new");
  document.getElementById("rfiModeExisting").onclick = () => setRfiMode("existing");
  document.getElementById("fileRfiBtn").onclick      = doFileToExistingRfi;

  // Submittal mode toggles
  document.getElementById("subModeNew").onclick      = () => setSubMode("new");
  document.getElementById("subModeExisting").onclick = () => setSubMode("existing");
  document.getElementById("fileSubBtn").onclick      = doFileToExistingSub;

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
        setSelectedProject(allProjects.find(p => p.id === el.dataset.id), true);
        searchInput.value = "";
        dropdown.style.display = "none";
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

  // Build participants list from From / To / CC
  const toList = emailItem.to || [];
  const ccList = emailItem.cc || [];
  emailParticipants = [
    { label: "From", displayName: emailFrom, emailAddress: emailFromAddress },
    ...toList.map(r => ({ label: "To", displayName: r.displayName || "", emailAddress: r.emailAddress || "" })),
    ...ccList.map(r => ({ label: "CC", displayName: r.displayName || "", emailAddress: r.emailAddress || "" })),
  ].filter(p => p.emailAddress);

  // Pre-fill note body
  document.getElementById("noteBody").value = emailItem.subject || "";

  // Pre-fill RFI from
  document.getElementById("rfiFrom").value = emailFrom;
  document.getElementById("subFrom").value = emailFrom;
  restoreProjectSelectionForCurrentEmail();
  refreshEmailSavedIndicator();
}

function getCurrentMessageRestId() {
  if (!emailItem?.itemId) return "";
  return Office.context.mailbox.convertToRestId(emailItem.itemId, Office.MailboxEnums.RestVersion.v2_0);
}

function findSavedEmailRecord(project, msgId) {
  if (!project || !msgId) return null;
  return (project.emails || []).find(e => e.msgId === msgId) || null;
}

function refreshEmailSavedIndicator() {
  const btnSharePoint = document.getElementById("saveSpBtn");
  const btnRecordOnly = document.getElementById("saveRecordBtn");
  if (!btnSharePoint || !btnRecordOnly) return;

  btnSharePoint.disabled = false;
  btnRecordOnly.disabled = false;

  if (!selectedProject || !emailItem?.itemId) return;
  const existing = findSavedEmailRecord(selectedProject, getCurrentMessageRestId());
  if (!existing) return;

  const savedDate = existing.savedAt ? new Date(existing.savedAt).toLocaleString("en-US") : "an earlier time";
  setStatus("actionStatus", "info", "This email was already saved to this project on " + savedDate + ".");
  btnSharePoint.disabled = true;
  btnRecordOnly.disabled = true;
}

// ─── AUTH ─────────────────────────────────────────────────────────────────────
async function doSignIn() {
  setStatus("signInStatus", "info", "⏳ Signing in…");
  try {
    const result = await msalApp.loginPopup({ scopes: GRAPH_SCOPES });
    msalAccount = result.account;
    msalApp.setActiveAccount(msalAccount);
    localStorage.setItem(LAST_ACCOUNT_STORAGE_KEY, msalAccount?.homeAccountId || "");
    await onSignedIn();
  } catch (e) {
    setStatus("signInStatus", "error", "✗ Sign-in failed: " + e.message);
  }
}

async function doSignOut() {
  await msalApp.logoutPopup({ account: msalAccount });
  msalAccount = null;
  msalApp.setActiveAccount(null);
  localStorage.removeItem(LAST_ACCOUNT_STORAGE_KEY);
  selectedProject = null;
  allProjects = [];
  allClients = [];
  showView("signInView");
  updateProjectQuickLinks();
}

async function onSignedIn() {
  showView("mainView");
  await loadProjects();
  restoreProjectSelectionForCurrentEmail();
  updateProjectQuickLinks();
}

async function getToken() {
  const account = msalAccount || msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
  if (!account) throw new Error("Not signed in");
  msalAccount = account;
  try {
    const r = await msalApp.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
    return r.accessToken;
  } catch {
    const r = await msalApp.acquireTokenPopup({ scopes: GRAPH_SCOPES, account });
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
    const res = await fetch(SUPABASE_URL + "/rest/v1/pms_data?id=eq.singleton&select=projects,clients", {
      headers: SB_HEADERS,
    });
    const rows = await res.json();
    if (!rows || !rows[0]) return;
    allProjects = (rows[0].projects || []).filter(p => !p.archived);
    allClients = rows[0].clients || [];
    renderCompanySuggestions();
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

function getEmailProjectMap() {
  try {
    return JSON.parse(localStorage.getItem(EMAIL_PROJECT_MAP_STORAGE_KEY) || "{}");
  } catch {
    return {};
  }
}

function setSelectedProject(project, persistForEmail = false) {
  selectedProject = project || null;
  const badge = document.getElementById("selectedProjectBadge");
  if (badge) {
    if (selectedProject) {
      badge.textContent = "✓ " + (selectedProject.projectNumber ? selectedProject.projectNumber + " — " : "") + selectedProject.name;
      badge.style.display = "block";
    } else {
      badge.style.display = "none";
      badge.textContent = "";
    }
  }
  if (persistForEmail && selectedProject) {
    const msgId = getCurrentMessageRestId();
    if (msgId) {
      const map = getEmailProjectMap();
      map[msgId] = selectedProject.id;
      localStorage.setItem(EMAIL_PROJECT_MAP_STORAGE_KEY, JSON.stringify(map));
    }
  }
  updateProjectQuickLinks();
  refreshEmailSavedIndicator();
}

function restoreProjectSelectionForCurrentEmail() {
  const msgId = getCurrentMessageRestId();
  if (!msgId || !allProjects.length) return;
  const map = getEmailProjectMap();
  const projectId = map[msgId];
  if (!projectId) return;
  const project = allProjects.find(p => p.id === projectId);
  if (project) setSelectedProject(project, false);
}

function renderCompanySuggestions() {
  const list = document.getElementById("companyList");
  if (!list) return;
  const companies = [...new Set((allClients || []).map(c => (c.name || "").trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  list.innerHTML = companies.map(name => `<option value="${escHtml(name)}"></option>`).join("");
}

function getClientByEmail(email) {
  if (!email) return null;
  const emailLc = email.toLowerCase();
  const domain = emailLc.includes("@") ? emailLc.split("@")[1] : "";
  return (allClients || []).find(c => {
    const contacts = c.contacts || [];
    if (contacts.some(ct => (ct.email || "").toLowerCase() === emailLc)) return true;
    return !!domain && contacts.some(ct => (ct.email || "").toLowerCase().endsWith("@" + domain));
  }) || null;
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
  // IDs are hardcoded — no Graph API call needed, no Sites.Read.All required.
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

function encodeDrivePath(path) {
  return String(path || "")
    .split("/")
    .filter(Boolean)
    .map(p => encodeURIComponent(p))
    .join("/");
}

// Create a folder idempotently (conflictBehavior:replace is a no-op on existing folders)
async function ensureSpFolder(driveId, token, parentPath, name) {
  try {
    await fetch("https://graph.microsoft.com/v1.0/drives/" + driveId + "/root:/" + encodeDrivePath(parentPath) + ":/children", {
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
  lastAttachmentUploadStats = { attempted: 0, uploaded: 0, failed: [] };
  const bodyHtml = await getEmailBodyHtml(token);
  await fetch("https://graph.microsoft.com/v1.0/drives/" + driveId + "/root:/" + encodeDrivePath(targetPath) + "/email.html:/content", {
    method: "PUT",
    headers: { "Authorization": "Bearer " + token, "Content-Type": "text/html" },
    body: buildEmailHtml(bodyHtml),
  });

  if (!emailItem.hasAttachments) return 0;
  try {
    let count = 0;
    // Prefer Outlook item APIs for attachment bytes; this is the most reliable in add-ins.
    const officeAtts = await getOfficeFileAttachments();
    if (officeAtts.length) {
      for (const att of officeAtts) {
        lastAttachmentUploadStats.attempted++;
        const uploaded = await uploadAttachmentToSharePoint(driveId, token, targetPath, att.name, att.contentType, att.bytes);
        if (uploaded) count++;
        else lastAttachmentUploadStats.failed.push(att.name || "attachment");
      }
      lastAttachmentUploadStats.uploaded = count;
      return count;
    }

    // Fallback to Graph attachment APIs when Office APIs are unavailable.
    const restId = Office.context.mailbox.convertToRestId(emailItem.itemId, Office.MailboxEnums.RestVersion.v2_0);
    const attData = await graphFetch("GET", "/me/messages/" + restId + "/attachments", null, token);
    for (const att of (attData?.value || [])) {
      if (att["@odata.type"] !== "#microsoft.graph.fileAttachment") continue;
      lastAttachmentUploadStats.attempted++;

      let bytes = null;
      if (att.contentBytes) {
        const binary = atob(att.contentBytes);
        bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      } else if (att.id) {
        const rawRes = await fetch(
          "https://graph.microsoft.com/v1.0/me/messages/" + restId + "/attachments/" + att.id + "/$value",
          { headers: { "Authorization": "Bearer " + token } }
        );
        if (!rawRes.ok) {
          console.warn("Attachment download failed:", att.name, rawRes.status);
          lastAttachmentUploadStats.failed.push((att.name || "attachment") + " (download " + rawRes.status + ")");
          continue;
        }
        bytes = new Uint8Array(await rawRes.arrayBuffer());
      }
      if (!bytes) continue;
      const uploaded = await uploadAttachmentToSharePoint(driveId, token, targetPath, att.name, att.contentType, bytes);
      if (uploaded) count++;
      else lastAttachmentUploadStats.failed.push(att.name || "attachment");
    }
    lastAttachmentUploadStats.uploaded = count;
    return count;
  } catch (e) {
    console.warn("Attachment upload failed:", e.message);
    lastAttachmentUploadStats.failed.push("Unhandled error: " + e.message);
    return 0;
  }
}

function toBytesFromBase64(base64) {
  const binary = atob(base64 || "");
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

async function getOfficeFileAttachments() {
  if (!emailItem?.getAttachmentsAsync || !emailItem?.getAttachmentContentAsync) return [];

  const atts = await new Promise((resolve, reject) => {
    emailItem.getAttachmentsAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value || []);
      else reject(new Error(res.error?.message || "getAttachmentsAsync failed"));
    });
  });

  const fileAtts = atts.filter(att => att.attachmentType === Office.MailboxEnums.AttachmentType.File);
  const out = [];
  for (const att of fileAtts) {
    const content = await new Promise((resolve, reject) => {
      emailItem.getAttachmentContentAsync(att.id, (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
        else reject(new Error(res.error?.message || "getAttachmentContentAsync failed"));
      });
    }).catch((e) => {
      console.warn("Office attachment content failed:", att.name, e.message);
      return null;
    });
    if (!content || content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) continue;
    out.push({
      name: att.name || "attachment",
      contentType: att.contentType || "application/octet-stream",
      bytes: toBytesFromBase64(content.content),
    });
  }
  return out;
}

async function uploadAttachmentToSharePoint(driveId, token, targetPath, name, contentType, bytes) {
  const safeName = (name || "attachment").replace(/[\\/:*?"<>|]/g, "-").trim() || "attachment";
  const uploadRes = await fetch("https://graph.microsoft.com/v1.0/drives/" + driveId + "/root:/" + encodeDrivePath(targetPath + "/" + safeName) + ":/content", {
    method: "PUT",
    headers: { "Authorization": "Bearer " + token, "Content-Type": contentType || "application/octet-stream" },
    body: bytes,
  });
  if (!uploadRes.ok) {
    console.warn("Attachment upload failed:", safeName, uploadRes.status);
    return false;
  }
  return true;
}

// ─── SAVE TO SHAREPOINT ───────────────────────────────────────────────────────
async function doSaveToSharePoint() {
  if (!selectedProject) { setStatus("actionStatus", "error", "Select a project first."); return; }
  if (!selectedProject.projectFolderUrl) { setStatus("actionStatus", "error", "No SharePoint folder on this project. Create one in the PMS first."); return; }
  if (findSavedEmailRecord(selectedProject, getCurrentMessageRestId())) {
    refreshEmailSavedIndicator();
    return;
  }

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
    if (emailItem.hasAttachments && attCount === 0) {
      const attempted = lastAttachmentUploadStats?.attempted || 0;
      const sample = (lastAttachmentUploadStats?.failed || []).slice(0, 2).join("; ");
      setStatus("actionStatus", "error", "Email saved, but 0/" + attempted + " attachments uploaded. " + (sample || "Open browser console for details."));
    } else {
      setStatus("actionStatus", "success", "✓ Saved to SharePoint" + attMsg + " and project record.");
    }
    refreshEmailSavedIndicator();
  } catch (e) {
    setStatus("actionStatus", "error", "✗ " + e.message);
  }
}

async function doSaveToProjectRecordOnly() {
  if (!selectedProject) { setStatus("actionStatus", "error", "Select a project first."); return; }
  if (emailItem?.hasAttachments) {
    setStatus("actionStatus", "error", "This email has attachments. Use 'Save to SharePoint + Project Record' instead.");
    return;
  }
  const msgId = getCurrentMessageRestId();
  if (findSavedEmailRecord(selectedProject, msgId)) {
    refreshEmailSavedIndicator();
    return;
  }

  setStatus("actionStatus", "info", "⏳ Saving to project record…");
  try {
    const from = emailItem.from;
    const emailRecord = {
      id: uid(), msgId,
      subject: emailItem.subject || "",
      from: from?.displayName || "",
      fromAddress: from?.emailAddress || "",
      date: emailItem.dateTimeCreated,
      bodyText: "",
      spFolderUrl: "",
      savedAt: new Date().toISOString(),
      savedToSharePoint: false,
    };
    const updatedProject = { ...selectedProject, emails: [...(selectedProject.emails || []), emailRecord] };
    updateProjectInList(updatedProject);
    selectedProject = updatedProject;
    await saveToSupabase(allProjects);
    setStatus("actionStatus", "success", "✓ Saved to project record (no SharePoint upload).");
    refreshEmailSavedIndicator();
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

// ─── CALENDAR HELPERS ─────────────────────────────────────────────────────────

let _nycCalendarId = null;

async function getNYCCalendarId() {
  if (_nycCalendarId) return _nycCalendarId;
  const cached = sessionStorage.getItem("setty_addin_cal_id");
  if (cached) { _nycCalendarId = cached; return _nycCalendarId; }
  try {
    const token = await getToken();
    const data  = await graphFetch("GET", "/me/calendars?$top=50", null, token);
    const nyc   = (data?.value || []).find(c =>
      c.name.toLowerCase().includes("nyc") || c.name.toLowerCase().includes("shared")
    );
    if (nyc) {
      _nycCalendarId = nyc.id;
      sessionStorage.setItem("setty_addin_cal_id", nyc.id);
    }
  } catch {}
  return _nycCalendarId || null;
}

async function createMilestoneCalendarEvent(milestone, project) {
  // All-day events need exclusive end = start + 1 day
  const endD = new Date(milestone.dueDate + "T12:00:00");
  endD.setDate(endD.getDate() + 1);
  const endStr = endD.getFullYear() + "-" + String(endD.getMonth()+1).padStart(2,"0") + "-" + String(endD.getDate()).padStart(2,"0");

  const prefix  = project.projectNumber ? "[" + project.projectNumber + "] " : "";
  const subject = prefix + project.name + " — " + milestone.name;

  const event = {
    subject,
    isAllDay: true,
    start: { dateTime: milestone.dueDate + "T00:00:00", timeZone: "Eastern Standard Time" },
    end:   { dateTime: endStr          + "T00:00:00", timeZone: "Eastern Standard Time" },
  };

  try {
    const token = await getToken();
    const calId = await getNYCCalendarId();
    const path  = calId ? "/me/calendars/" + calId + "/events" : "/me/events";
    const res   = await graphFetch("POST", path, event, token);
    return { success: true, eventId: res?.id, onShared: !!calId };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ─── DUE DATE EXTRACTOR ───────────────────────────────────────────────────────

function extractDueDates(text, emailReceivedDate) {
  const results = [];
  const seen    = new Set();
  const refDate = emailReceivedDate ? new Date(emailReceivedDate) : new Date();

  const MONTHS_LONG  = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const MONTHS_SHORT = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const DAYS         = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];

  function toISO(year, month1, day) {
    const y = year < 100 ? 2000 + year : year;
    return y + "-" + String(month1).padStart(2,"0") + "-" + String(day).padStart(2,"0");
  }

  function addResult(iso, display, idx) {
    if (seen.has(iso)) return;
    const d   = new Date(iso + "T12:00:00");
    const now = new Date(); now.setDate(now.getDate() - 30);
    const cap = new Date(); cap.setFullYear(cap.getFullYear() + 3);
    if (d < now || d > cap) return;

    const ctxStart = Math.max(0, idx - 120);
    const ctxEnd   = Math.min(text.length, idx + display.length + 80);
    let ctx = text.slice(ctxStart, ctxEnd).replace(/\s+/g, " ").trim();
    if (ctxStart > 0) ctx = "…" + ctx;
    if (ctxEnd < text.length) ctx += "…";

    const before     = text.slice(Math.max(0, idx - 150), idx).toLowerCase();
    const hasKeyword = /\b(due|deadline|by|no later than|nlt|ntp|submit|required|respond|return|need|complete|deliver|before|expected|must have)\b/.test(before);

    seen.add(iso);
    results.push({ iso, display, ctx, hasKeyword });
  }

  let m;
  // Long month name: "March 15, 2026" / "March 15th, 2026"
  const p1 = /\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})\b/gi;
  while ((m = p1.exec(text))) {
    const mo = MONTHS_LONG.findIndex(x => x.toLowerCase() === m[1].toLowerCase()) + 1;
    addResult(toISO(+m[3], mo, +m[2]), m[0], m.index);
  }
  // Short month name: "Mar 15, 2026" / "Mar. 15 2026"
  const p2 = /\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})\b/gi;
  while ((m = p2.exec(text))) {
    const mo = MONTHS_SHORT.findIndex(x => x.toLowerCase() === m[1].toLowerCase()) + 1;
    addResult(toISO(+m[3], mo, +m[2]), m[0], m.index);
  }
  // Day-first: "15 March 2026"
  const p3 = /\b(\d{1,2})(?:st|nd|rd|th)?\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})\b/gi;
  while ((m = p3.exec(text))) {
    const mo = MONTHS_LONG.findIndex(x => x.toLowerCase() === m[2].toLowerCase()) + 1;
    addResult(toISO(+m[3], mo, +m[1]), m[0], m.index);
  }
  // Slash notation: "3/15/2026" or "03/15/26"
  const p4 = /\b(\d{1,2})\/(\d{1,2})\/(\d{2,4})\b/g;
  while ((m = p4.exec(text))) {
    const mo = +m[1], dy = +m[2], yr = +m[3];
    if (mo >= 1 && mo <= 12 && dy >= 1 && dy <= 31)
      addResult(toISO(yr, mo, dy), m[0], m.index);
  }
  // ISO: "2026-03-15"
  const p5 = /\b(20\d{2})-(\d{2})-(\d{2})\b/g;
  while ((m = p5.exec(text))) addResult(m[0], m[0], m.index);

  // Relative weekday: "next Friday" / "this Thursday"
  const p6 = /\b(next|this)\s+(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b/gi;
  while ((m = p6.exec(text))) {
    const target = DAYS.findIndex(d => d.toLowerCase() === m[2].toLowerCase());
    const d      = new Date(refDate);
    let   delta  = target - d.getDay();
    if (m[1].toLowerCase() === "next" || delta <= 0) delta += 7;
    d.setDate(d.getDate() + delta);
    const iso = d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0");
    addResult(iso, m[0] + "  (" + iso + ")", m.index);
  }

  // Keyword hits first, then chronological
  return results.sort((a, b) => {
    if (a.hasKeyword !== b.hasKeyword) return a.hasKeyword ? -1 : 1;
    return a.iso.localeCompare(b.iso);
  });
}

function escHtml(s) {
  return (s || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

async function showDatesView() {
  showView("datesView");
  document.getElementById("milestoneForm").style.display = "none";
  const list = document.getElementById("datesList");
  list.innerHTML = '<p style="color:#64748b;font-size:12px;text-align:center;padding:16px 0;">⏳ Scanning email…</p>';

  try {
    const token = await getToken();
    const html  = await getEmailBodyHtml(token);
    const tmp   = document.createElement("div");
    tmp.innerHTML = html;
    const text = (tmp.innerText || tmp.textContent || "").replace(/\s+/g, " ");

    const dates = extractDueDates(text, emailItem?.dateTimeCreated);

    if (!dates.length) {
      list.innerHTML = '<p style="color:#64748b;font-size:12px;text-align:center;padding:20px 0;">No due dates found in this email.</p>';
      return;
    }

    list.innerHTML = dates.map((d, i) => `
      <div class="date-card">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:5px;">
          <span style="font-size:13px;font-weight:700;color:${d.hasKeyword ? "#60b4ff" : "#e2e8f0"};">${escHtml(d.display)}</span>
          ${d.hasKeyword ? '<span style="font-size:10px;background:#1e3a5f;color:#60b4ff;padding:1px 7px;border-radius:4px;flex-shrink:0;">deadline</span>' : ""}
        </div>
        <div style="font-size:11px;color:#64748b;line-height:1.5;margin-bottom:8px;font-style:italic;">${escHtml(d.ctx)}</div>
        <button class="btn btn-blue" style="padding:5px 12px;font-size:11px;margin-bottom:0;"
          onclick="prefillMilestone('${d.iso}')">➕ Use this date</button>
      </div>
    `).join("");
  } catch(e) {
    list.innerHTML = `<p style="color:#f87171;font-size:12px;">Error: ${escHtml(e.message)}</p>`;
  }
}

function prefillMilestone(iso) {
  document.getElementById("milestoneDate").value = iso;
  document.getElementById("milestoneName").value = (emailItem?.subject || "").slice(0, 80);
  document.getElementById("milestoneStatus").className = "status-msg";
  document.getElementById("milestoneStatus").textContent = "";
  const form = document.getElementById("milestoneForm");
  form.style.display = "block";
  form.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function showManualMilestoneForm() {
  showView("datesView");
  const list = document.getElementById("datesList");
  if (list) list.innerHTML = '<p style="color:#64748b;font-size:12px;text-align:center;padding:16px 0;">Manual mode: enter milestone details below.</p>';
  const defaultDate = new Date();
  const iso = defaultDate.getFullYear() + "-" + String(defaultDate.getMonth() + 1).padStart(2, "0") + "-" + String(defaultDate.getDate()).padStart(2, "0");
  prefillMilestone(iso);
}

async function doSaveMilestone() {
  const name    = document.getElementById("milestoneName").value.trim();
  const dueDate = document.getElementById("milestoneDate").value;
  if (!name)    { setStatus("milestoneStatus", "error", "Please enter a milestone name."); return; }
  if (!dueDate) { setStatus("milestoneStatus", "error", "Please select a date."); return; }
  if (!selectedProject) { setStatus("milestoneStatus", "error", "Select a project first (go back)."); return; }

  setStatus("milestoneStatus", "info", "⏳ Saving…");
  try {
    const res  = await fetch(SUPABASE_URL + "/rest/v1/pms_data?id=eq.singleton&select=projects", { headers: SB_HEADERS });
    const rows = await res.json();
    const projects = rows[0]?.projects || [];
    const proj = projects.find(p => p.id === selectedProject.id);
    if (!proj) { setStatus("milestoneStatus", "error", "Project not found in Supabase."); return; }

    // Create calendar event first so we can store its ID on the milestone
    const milestone = {
      id:          "addin-" + Date.now().toString(36) + Math.random().toString(36).slice(2,6),
      name,
      type:        "non-billable",
      phase:       "",
      dueDate,
      pctComplete: 0,
      fee:         0,
      notes:       "From email: " + (emailItem?.subject || ""),
      cancelled:   false,
    };

    setStatus("milestoneStatus", "info", "⏳ Syncing to calendar…");
    const calResult = await createMilestoneCalendarEvent(milestone, selectedProject);
    if (calResult.success) milestone.calendarEventId = calResult.eventId;

    proj.milestones = [...(proj.milestones || []), milestone];

    await fetch(SUPABASE_URL + "/rest/v1/pms_data?id=eq.singleton", {
      method:  "PATCH",
      headers: SB_HEADERS,
      body:    JSON.stringify({ projects, updated_at: new Date().toISOString() }),
    });

    const projLabel = (selectedProject.projectNumber ? selectedProject.projectNumber + " — " : "") + selectedProject.name;
    if (calResult.success) {
      const calLabel = calResult.onShared ? "NYC Shared Calendar" : "your personal calendar";
      setStatus("milestoneStatus", "success", "✓ Saved to " + projLabel + " · synced to " + calLabel);
    } else {
      setStatus("milestoneStatus", "success", "✓ Saved to " + projLabel + " (calendar sync failed: " + calResult.error + ")");
    }
    document.getElementById("milestoneForm").style.display = "none";
  } catch(e) {
    setStatus("milestoneStatus", "error", "✗ " + e.message);
  }
}

// ─── PEOPLE PICKER ────────────────────────────────────────────────────────────
function showPeopleView() {
  const list = document.getElementById("participantList");
  if (!emailParticipants.length) {
    list.innerHTML = '<p style="font-size:12px;color:#64748b;">No participants found.</p>';
  } else {
    const labelColor = { From: "#C00000", To: "#1d4ed8", CC: "#0f766e" };
    const labelBg    = { From: "#450a0a", To: "#1e3a5f", CC: "#134e4a" };
    list.innerHTML = emailParticipants.map((p, i) => `
      <div class="participant-row" data-idx="${i}">
        <div style="flex:1;min-width:0;">
          <div style="font-size:12px;font-weight:600;color:#e2e8f0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
            ${p.displayName || p.emailAddress}
          </div>
          <div style="font-size:11px;color:#64748b;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
            ${p.emailAddress}
          </div>
        </div>
        <span class="pill" style="background:${labelBg[p.label]||'#1e2235'};color:${labelColor[p.label]||'#94a3b8'};">
          ${p.label}
        </span>
      </div>
    `).join("");
    list.querySelectorAll(".participant-row").forEach(el => {
      el.onclick = () => prefillContactFromParticipant(emailParticipants[+el.dataset.idx]);
    });
  }
  showView("peopleView");
}

function prefillContactFromParticipant(p) {
  const matchedClient = getClientByEmail(p.emailAddress || "");
  document.getElementById("contactName").value    = p.displayName || "";
  document.getElementById("contactTitle").value   = "";
  document.getElementById("contactCompany").value = matchedClient?.name || "";
  document.getElementById("contactEmail").value   = p.emailAddress || "";
  document.getElementById("contactPhone").value   = "";
  setStatus("contactStatus", "", "");
  showView("contactView");
}

// ─── EXTRACT CONTACT ──────────────────────────────────────────────────────────
async function doExtractContact() {
  setStatus("actionStatus", "info", "⏳ Extracting contact…");
  const token = await getToken();
  const body = await getEmailBodyHtml(token);
  const contact = parseSignature(body, emailFrom, emailFromAddress);
  const matchedClient = getClientByEmail(contact.email || emailFromAddress);
  document.getElementById("contactName").value    = contact.name;
  document.getElementById("contactTitle").value   = contact.title;
  document.getElementById("contactCompany").value = matchedClient?.name || contact.company;
  document.getElementById("contactEmail").value   = contact.email;
  document.getElementById("contactPhone").value   = contact.phone;
  setStatus("actionStatus", "info", "");
  showView("contactView");
}

function projectPmsUrl(project) {
  if (!project) return "";
  if (project.pmsUrl) {
    // Normalize legacy links to the current hosted PMS path.
    return project.pmsUrl
      .replace("https://settypms.com/", "https://smartias.github.io/setty-pms/SettyPMS/");
  }
  if (project.slug) return PMS_PROJECT_BASE_URL + encodeURIComponent(project.slug);
  if (project.id) return PMS_PROJECT_BASE_URL + encodeURIComponent(project.id);
  return PMS_DASHBOARD_URL;
}

function updateProjectQuickLinks() {
  const pmsBtn = document.getElementById("openPmsBtn");
  const spBtn = document.getElementById("openSpFolderBtn");
  if (!pmsBtn || !spBtn) return;
  pmsBtn.disabled = !projectPmsUrl(selectedProject);
  spBtn.disabled = !selectedProject?.projectFolderUrl;
}

function openSelectedProjectInPms() {
  if (!selectedProject) { setStatus("actionStatus", "error", "Select a project first."); return; }
  const url = projectPmsUrl(selectedProject);
  if (!url) { setStatus("actionStatus", "error", "No PMS URL is available for this project."); return; }
  openExternalUrl(url);
}

function openSelectedProjectSpFolder() {
  if (!selectedProject) { setStatus("actionStatus", "error", "Select a project first."); return; }
  if (!selectedProject.projectFolderUrl) { setStatus("actionStatus", "error", "No SharePoint folder URL is set on this project."); return; }
  openExternalUrl(selectedProject.projectFolderUrl);
}

function openPmsDashboard() {
  openExternalUrl(PMS_DASHBOARD_URL);
}

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
  // Hide loading spinner on first real view
  const loading = document.getElementById("loadingView");
  if (loading) loading.style.display = "none";
  ["signInView","mainView","noteView","rfiView","subView","datesView","peopleView","contactView"].forEach(v => {
    const el = document.getElementById(v);
    if (el) el.classList.toggle("active", v === id);
  });
}

// Fallback: if Office.onReady never fires (browser preview / load failure),
// replace spinner with a plain message after 5 seconds
setTimeout(() => {
  const loading = document.getElementById("loadingView");
  if (loading && loading.style.display !== "none") {
    loading.innerHTML = '<p style="color:#94a3b8;font-size:12px;text-align:center;padding:0 16px;">Open this add-in from Outlook.<br/>To sideload, use the manifest.xml file.</p>';
  }
}, 5000);

function setStatus(elId, type, msg) {
  const el = document.getElementById(elId);
  if (!el) return;
  el.className = "status-msg" + (msg ? " show " + type : "");
  el.textContent = msg;
}
