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
  "Notes.ReadWrite",      // needed for OneNote page creation — no admin consent required
  // Sites.Read.All removed — site and drive IDs are hardcoded below (no admin consent needed)
];
const TEAMS_TEAM_ID = "a4c48361-7991-43db-af83-4c854918a760";
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
let currentItemKind = "message"; // message | appointment
let currentItemICalUId = "";    // iCalUId for appointments — same across all attendees' mailboxes
let lastAttachmentUploadStats = null;
let currentConversationId = "";
// Context generation: incremented every time loadItemContext fires (= every
// time the user clicks a different email/appointment). All in-flight async
// fetches capture the generation number at start and discard their results
// if the generation has advanced. Without this, slow Graph calls from email
// A complete after the user moved to email B and stamp module-level state
// (`currentItemICalUId`, `emailParticipants`) with values from the wrong item.
let itemContextGeneration = 0;
// Save in-flight flag — prevents double-clicks on save buttons from launching
// parallel save paths that race the version counter and produce phantom errors.
let saveInFlight = false;
// Hardcoded SharePoint IDs — eliminates Sites.Read.All (the only admin-consent scope).
// Retrieved once via https://setty.sharepoint.com/sites/NYCProjects/_api/v2.0/drives
const SP_SITE_ID_HARDCODED  = "setty.sharepoint.com,aa580464-13e9-4eb4-8ad4-ca6ff5b9e001,c97a67e8-fb1b-4a23-a29a-753a5d57d410";
const SP_DRIVE_ID_HARDCODED = "b!ZARYqukTtE6K1Mpv9bngAehneskb-yNKopp1Ol1X1BBnJPKsNGM-TaGmbGiL3ZaU";
let _spIds = { siteId: SP_SITE_ID_HARDCODED, driveId: SP_DRIVE_ID_HARDCODED };
const LAST_ACCOUNT_STORAGE_KEY = "settyPms:lastMsalAccountHomeId";
const EMAIL_PROJECT_MAP_STORAGE_KEY = "settyPms:emailProjectMap";
const EMAIL_CONVO_PROJECT_MAP_STORAGE_KEY = "settyPms:conversationProjectMap";
const EMAIL_THREAD_TAGS_TABLE = "pms_email_thread_tags";
const PROJECT_EMAILS_TABLE = "pms_project_emails";
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
    // When the task pane is pinned, Office swaps mailbox.item silently as the user
    // clicks different emails. ItemChanged fires each time — reload the pane context.
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      () => { showView("mainView"); loadItemContext(); }
    );
    loadItemContext();
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
  document.getElementById("newActionItemBtn").onclick = () => { prefillActionItem(); showView("actionItemView"); };
  document.getElementById("logRfiBtn").onclick       = () => { prefillRfi(); showView("rfiView"); };
  document.getElementById("logSubBtn").onclick       = () => { prefillSub(); showView("subView"); };
  document.getElementById("noteBack").onclick    = () => showView("mainView");
  document.getElementById("actionItemBack").onclick = () => showView("mainView");
  document.getElementById("rfiBack").onclick     = () => showView("mainView");
  document.getElementById("subBack").onclick     = () => showView("mainView");
  document.getElementById("peopleBack").onclick  = () => showView("mainView");
  // Contact form back returns to peopleView (the only entry point); peopleBack handles return-to-main
  document.getElementById("contactBack").onclick = () => showView("peopleView");
  document.getElementById("datesBack").onclick   = () => showView("mainView");
  // "More" expander — persist open/closed state across emails so power users
  // who expand it once don't have to keep doing so.
  const moreEl = document.getElementById("moreActions");
  if (moreEl) {
    if (localStorage.getItem("settyPms:moreExpanded") === "1") moreEl.open = true;
    moreEl.addEventListener("toggle", () => {
      localStorage.setItem("settyPms:moreExpanded", moreEl.open ? "1" : "0");
    });
  }
  document.getElementById("manualMilestoneBtn").onclick = showManualMilestoneForm;
  document.getElementById("addParticipantBtn").onclick = showPeopleView;
  document.getElementById("saveMilestoneBtn").onclick = doSaveMilestone;
  document.getElementById("saveNoteBtn").onclick    = doSaveNote;
  document.getElementById("saveActionItemBtn").onclick = doSaveActionItem;
  document.getElementById("saveRfiBtn").onclick     = doSaveRfi;
  document.getElementById("saveSubBtn").onclick     = doSaveSub;
  document.getElementById("saveContactBtn").onclick = doSaveContact;
  document.getElementById("openPmsBtn").onclick = openSelectedProjectInPms;
  document.getElementById("openSpFolderBtn").onclick = openSelectedProjectSpFolder;
  document.getElementById("openDashboardBtn").onclick = openPmsDashboard;
  document.getElementById("clearProjectTagBtn").onclick = () => { void clearProjectTagForCurrentEmail(); };
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
// ─── OUTLOOK ITEM CONTEXT (MAIL + CALENDAR) ─────────────────────────────────
function dedupeParticipants(participants) {
  const seen = new Set();
  return (participants || []).filter(p => {
    const email = (p.emailAddress || "").trim().toLowerCase();
    if (!email || seen.has(email)) return false;
    seen.add(email);
    return true;
  });
}
function getAppointmentDateLabel(item) {
  const start = item?.start;
  if (!start) return "";
  const d = new Date(start);
  if (Number.isNaN(d.getTime())) return "";
  return d.toLocaleString("en-US", {
    month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit"
  });
}
function buildMeetingNoteBody(item) {
  // In compose mode (organizer editing their own meeting), properties like subject,
  // start, attendees are async objects — read them defensively and skip if not plain values.
  const isComposeMode = typeof item?.subject !== "string";
  const lines = [];
  lines.push(isComposeMode ? "" : (item?.subject || "(No subject)"));
  lines.push("");
  lines.push("Meeting details:");
  if (!isComposeMode) {
    if (item?.location && typeof item.location === "string") lines.push("Location: " + item.location);
    if (item?.start && !(item.start?.getAsync)) lines.push("Start: " + new Date(item.start).toLocaleString("en-US"));
    if (item?.end   && !(item.end?.getAsync))   lines.push("End: "   + new Date(item.end).toLocaleString("en-US"));
    const organizer = item?.organizer;
    if (organizer?.displayName || organizer?.emailAddress)
      lines.push("Organizer: " + (organizer.displayName || organizer.emailAddress));
    const attendees = dedupeParticipants([
      ...(Array.isArray(item?.requiredAttendees) ? item.requiredAttendees : []).map(a => ({ displayName: a.displayName || "", emailAddress: a.emailAddress || "" })),
      ...(Array.isArray(item?.optionalAttendees) ? item.optionalAttendees : []).map(a => ({ displayName: a.displayName || "", emailAddress: a.emailAddress || "" })),
    ]);
    if (attendees.length) {
      lines.push("");
      lines.push("Attendees:");
      attendees.forEach(a => lines.push("- " + (a.displayName || a.emailAddress) + (a.emailAddress ? " <" + a.emailAddress + ">" : "")));
    }
  }
  lines.push("");
  lines.push("Summary:");
  lines.push("");
  lines.push("Action items:");
  return lines.join("\n");
}

function loadItemContext() {
  // Bump the generation. All async work below captures `myGen` at start and
  // bails out before writing module state if the generation has advanced
  // (= user clicked a different email mid-fetch).
  itemContextGeneration++;
  const myGen = itemContextGeneration;
  emailItem = Office.context.mailbox.item;
  currentConversationId = "";
  currentItemICalUId = "";
  emailParticipants = [];
  // Per-item ✓ "added this session" marks reset when item changes
  _sessionSavedContactEmails.clear();
  if (!emailItem) return;
  // For appointments, fetch the iCalUId in the background — it's the same across
  // all attendees' mailboxes so we can use it to match notes saved by anyone on the team.
  if (emailItem.itemType === Office.MailboxEnums.ItemType.Appointment) {
    void (async () => {
      try {
        const restId = Office.context.mailbox.convertToRestId(emailItem.itemId, Office.MailboxEnums.RestVersion.v2_0);
        const ev = await graphFetch("GET", `/me/events/${restId}?$select=iCalUId`);
        // Stale-result guard — discard if user has moved to another item
        if (myGen !== itemContextGeneration) return;
        currentItemICalUId = ev?.iCalUId || "";
        refreshOneNoteLinkBanner();
        refreshCalendarStatus();
        // Re-attempt restoration if no project was found on the first pass.
        // The first pass (in loadItemContext) might have run before iCalUId
        // was available — without this re-fire, an appointment opened on a
        // device that's never tagged it would never auto-restore the tag.
        if (!selectedProject && currentItemICalUId) {
          await restoreProjectSelectionForCurrentEmail();
        }
      } catch { /* non-fatal */ }
    })();
  }
  currentItemKind = emailItem.itemType === Office.MailboxEnums.ItemType.Appointment ? "appointment" : "message";
  if (currentItemKind === "appointment") {
    // Restore a previously-saved project association for this appointment.
    // (Same mechanism as emails — keyed on the REST item ID in localStorage.)
    setSelectedProject(null, false);
    void restoreProjectSelectionForCurrentEmail();

    // Detect compose vs read mode — subject is a plain string in read mode,
    // an async Subject object in compose mode (when the user is the organizer editing their own meeting).
    const isComposeMode = typeof emailItem.subject !== "string";

    // ── Subject ──────────────────────────────────────────────────────────────
    if (isComposeMode) {
      document.getElementById("emailSubject").textContent = "(Loading…)";
      emailItem.subject.getAsync(r => {
        if (myGen !== itemContextGeneration) return; // user moved on
        if (r.status === Office.AsyncResultStatus.Succeeded)
          document.getElementById("emailSubject").textContent = r.value || "(No subject)";
      });
    } else {
      document.getElementById("emailSubject").textContent = emailItem.subject || "(No subject)";
    }

    // ── Organizer ────────────────────────────────────────────────────────────
    // In compose mode item.organizer doesn't exist — the signed-in user IS the organizer.
    if (isComposeMode) {
      emailFrom = msalAccount?.name || "";
      emailFromAddress = msalAccount?.username || "";
    } else {
      const organizer = emailItem.organizer;
      emailFrom = organizer?.displayName || "";
      emailFromAddress = organizer?.emailAddress || "";
    }

    // ── Date display ─────────────────────────────────────────────────────────
    if (isComposeMode) {
      document.getElementById("emailMeta").textContent = "Organizer: " + (emailFrom || "(You)");
      emailItem.start.getAsync(r => {
        if (myGen !== itemContextGeneration) return;
        if (r.status === Office.AsyncResultStatus.Succeeded) {
          const d = new Date(r.value);
          const dateFmt = isNaN(d) ? "" : d.toLocaleString("en-US", { month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit" });
          if (dateFmt) document.getElementById("emailMeta").textContent += "  ·  " + dateFmt;
        }
      });
    } else {
      const date = getAppointmentDateLabel(emailItem);
      document.getElementById("emailMeta").textContent =
        "Organizer: " + (emailFrom || emailFromAddress || "(Unknown)") + (date ? "  ·  " + date : "");
    }

    // ── Participants ─────────────────────────────────────────────────────────
    emailParticipants = dedupeParticipants([
      { label: "Organizer", displayName: emailFrom, emailAddress: emailFromAddress },
    ]);
    if (isComposeMode) {
      // Compose mode — attendees are async Recipients objects. Without
      // generation guarding, a user clicking "Log as Note" within ~200ms of
      // opening a compose-mode appointment would save the note before
      // attendees finished loading. Now: each callback bails if the item
      // changed; loadAtts dispatches both required+optional in parallel and
      // doesn't update emailParticipants from a stale generation.
      const loadAtts = (getter, label) => getter.getAsync(r => {
        if (myGen !== itemContextGeneration) return;
        if (r.status === Office.AsyncResultStatus.Succeeded) {
          emailParticipants = dedupeParticipants([
            ...emailParticipants,
            ...(r.value || []).map(a => ({ label, displayName: a.displayName || "", emailAddress: a.emailAddress || "" })),
          ]);
        }
      });
      if (emailItem.requiredAttendees?.getAsync) loadAtts(emailItem.requiredAttendees, "Required");
      if (emailItem.optionalAttendees?.getAsync) loadAtts(emailItem.optionalAttendees, "Optional");
    } else {
      // Read mode — attendees are plain arrays
      emailParticipants = dedupeParticipants([
        ...emailParticipants,
        ...(emailItem.requiredAttendees || []).map(r => ({ label: "Required", displayName: r.displayName || "", emailAddress: r.emailAddress || "" })),
        ...(emailItem.optionalAttendees || []).map(r => ({ label: "Optional", displayName: r.displayName || "", emailAddress: r.emailAddress || "" })),
      ]);
    }

    document.getElementById("noteCategory").value = "Client Meeting";
    document.getElementById("noteBody").value = buildMeetingNoteBody(emailItem);
    document.getElementById("saveSpBtn").disabled = true;
    document.getElementById("saveRecordBtn").disabled = true;
    document.getElementById("logNoteBtn").disabled = true;
    document.getElementById("newActionItemBtn").disabled = true;
    document.getElementById("logRfiBtn").disabled = true;
    document.getElementById("logSubBtn").disabled = true;
    document.getElementById("manualMilestoneBtn").disabled = true;
    // Status depends on whether this event was already logged; refreshCalendarStatus()
    // is also called from setSelectedProject() so it re-evaluates once the project restores.
    refreshCalendarStatus();
  } else {
    document.getElementById("saveSpBtn").disabled = false;
    document.getElementById("saveRecordBtn").disabled = false;
    document.getElementById("logNoteBtn").disabled = false;
    document.getElementById("newActionItemBtn").disabled = false;
    document.getElementById("logRfiBtn").disabled = false;
    document.getElementById("logSubBtn").disabled = false;
    document.getElementById("manualMilestoneBtn").disabled = false;
    setStatus("actionStatus", "", "");
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
    emailParticipants = dedupeParticipants([
      { label: "From", displayName: emailFrom, emailAddress: emailFromAddress },
      ...toList.map(r => ({ label: "To", displayName: r.displayName || "", emailAddress: r.emailAddress || "" })),
      ...ccList.map(r => ({ label: "CC", displayName: r.displayName || "", emailAddress: r.emailAddress || "" })),
    ]);
    // Pre-fill note body
    document.getElementById("noteBody").value = emailItem.subject || "";
    // Pre-fill RFI from
    document.getElementById("rfiFrom").value = emailFrom;
    document.getElementById("subFrom").value = emailFrom;
    // Clear any previously selected project immediately — restoreProjectSelection
    // will re-populate it if this email/conversation has a saved tag.
    setSelectedProject(null, false);
    void restoreProjectSelectionForCurrentEmail();
    refreshEmailSavedIndicator();
  }
}
function getCurrentMessageRestId() {
  if (!emailItem?.itemId) return "";
  return Office.context.mailbox.convertToRestId(emailItem.itemId, Office.MailboxEnums.RestVersion.v2_0);
}
function getCurrentMessageRecordId() {
  // Prefer internetMessageId (shared across recipients) for cross-mailbox matching.
  // Keep REST/item IDs as fallbacks so existing records created before this change still resolve.
  return emailItem?.internetMessageId || getCurrentMessageRestId() || emailItem?.itemId || "";
}
function getCurrentMessageIdCandidates() {
  return [...new Set([
    emailItem?.internetMessageId || "",
    getCurrentMessageRestId(),
    emailItem?.itemId || "",
  ].filter(Boolean))];
}
function getCurrentSharedMessageId() {
  return emailItem?.internetMessageId || "";
}
async function getCurrentConversationId() {
  if (currentConversationId) return currentConversationId;
  try {
    const restId = getCurrentMessageRestId();
    if (!restId) return "";
    const data = await graphFetch("GET", "/me/messages/" + restId + "?$select=conversationId", null);
    currentConversationId = data?.conversationId || "";
    return currentConversationId;
  } catch {
    return "";
  }
}

// Shared key for the current item — used by the cross-device tag system in
// pms_email_thread_tags so a project tag set on one device shows up on others.
//   - Emails:      conversationId (Graph; shared across all recipients of a thread)
//   - Appointments: iCalUId         (Graph; shared across all attendees of an event)
//
// This was the missing piece for cross-device persistence on calendar events:
// previously appointments fell through to getCurrentConversationId() which
// returned "" (because the /me/messages endpoint 404s for appointments), so
// no shared tag was ever written or read for appointments. Now the iCalUId
// acts as the cross-device key, reusing the existing tag-table infrastructure.
async function getCurrentSharedKey() {
  if (currentItemKind === "appointment") {
    if (currentItemICalUId) return currentItemICalUId;
    // iCalUId not yet loaded — fetch synchronously now
    if (emailItem?.itemId) {
      try {
        const restId = Office.context.mailbox.convertToRestId(emailItem.itemId, Office.MailboxEnums.RestVersion.v2_0);
        const ev = await graphFetch("GET", "/me/events/" + restId + "?$select=iCalUId", null);
        currentItemICalUId = ev?.iCalUId || "";
        return currentItemICalUId;
      } catch {
        return "";
      }
    }
    return "";
  }
  // Default (emails, drafts, etc.) — use conversationId
  return await getCurrentConversationId();
}
function findSavedEmailRecord(project, msgId) {
  if (!project || !msgId) return null;
  const candidateIds = getCurrentMessageIdCandidates();
  return (project.emails || []).find(e => candidateIds.includes(e.msgId) || e.msgId === msgId) || null;
}
function getLoggedEmailArtifactLabels(project) {
  if (!project || !emailItem?.itemId) return [];
  const labels = [];
  const sourceItemId = emailItem.itemId;
  const sourceMessageIds = getCurrentMessageIdCandidates();

  const notes = project.notes || [];
  const hasActionItem = notes.some(n =>
    (n?.sourceItemId === sourceItemId || sourceMessageIds.includes(n?.sourceMessageId))
    && (n?.actionItem || n?.category === "Action Item")
  );
  const hasNote = notes.some(n =>
    (n?.sourceItemId === sourceItemId || sourceMessageIds.includes(n?.sourceMessageId))
    && !(n?.actionItem || n?.category === "Action Item")
  );
  const hasMilestone = (project.milestones || []).some(m =>
    m?.sourceItemId === sourceItemId || sourceMessageIds.includes(m?.sourceMessageId)
  );

  if (hasNote) labels.push("note");
  if (hasActionItem) labels.push("action item");
  if (hasMilestone) labels.push("milestone");
  return labels;
}
function refreshEmailSavedIndicator() {
  const btnSharePoint = document.getElementById("saveSpBtn");
  const btnRecordOnly = document.getElementById("saveRecordBtn");
  if (!btnSharePoint || !btnRecordOnly) return;

  // Reset to default state first
  btnSharePoint.disabled = false;
  btnRecordOnly.disabled = false;
  btnSharePoint.textContent = "📁 Save to SharePoint";
  btnRecordOnly.textContent = "🗂️ Save to Project";

  if (!selectedProject || !emailItem?.itemId) return;
  const existing = findSavedEmailRecord(selectedProject, getCurrentMessageRecordId());
  if (!existing) {
    applyEmailFlowEmphasis();
    return;
  }

  const savedDate = existing.savedAt
    ? new Date(existing.savedAt).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })
    : "a prior session";
  const loggedLabels = getLoggedEmailArtifactLabels(selectedProject);
  const loggedSuffix = loggedLabels.length ? " Logged as " + loggedLabels.join(", ") + "." : "";

  if (emailItem?.hasAttachments) {
    // Allow re-run to catch attachments, but make it clear it was filed
    setStatus("actionStatus", "success", "✓ Filed to project on " + savedDate + "." + loggedSuffix + " Re-run to sync any new attachments.");
    btnSharePoint.textContent = "↺ Re-sync attachments";
    btnRecordOnly.disabled = true;
    btnRecordOnly.textContent = "✓ In project record";
  } else {
    setStatus("actionStatus", "success", "✓ Filed to project on " + savedDate + "." + loggedSuffix);
    btnSharePoint.disabled = true;
    btnSharePoint.textContent = "✓ Filed to SharePoint";
    btnRecordOnly.disabled = true;
    btnRecordOnly.textContent = "✓ In project record";
  }
  applyPipelineUiRules();
  applyEmailFlowEmphasis();
}

// Visual emphasis between the twin save buttons.
// Both remain clickable; this only nudges which one looks like the "expected" choice
// based on the email's shape (attachments vs. body-only).
//
// TODO (UX policy — see request below): implement the emphasis rules.
//   - Inputs available: emailItem?.hasAttachments (bool), selectedProject (may be null)
//   - DOM hooks: btnSp, btnRecord (buttons), capSp, capRecord (caption spans)
//   - Toggle .btn-deemph on the off-flow button to soften it without disabling it.
//   - Optionally rewrite the caption text to reflect the recommended path.
function applyEmailFlowEmphasis() {
  const btnSp = document.getElementById("saveSpBtn");
  const btnRecord = document.getElementById("saveRecordBtn");
  const capSp = document.getElementById("saveSpCaption");
  const capRecord = document.getElementById("saveRecordCaption");
  if (!btnSp || !btnRecord) return;

  // Reset emphasis and captions each call so previous state doesn't leak
  btnSp.classList.remove("btn-deemph");
  btnRecord.classList.remove("btn-deemph");
  if (capSp) capSp.textContent = "Email + attachments → SharePoint";
  if (capRecord) capRecord.textContent = "Copy of email → project record";

  // No project picked yet — keep both neutral; we don't know what's relevant.
  if (!selectedProject) return;

  if (emailItem?.hasAttachments) {
    btnRecord.classList.add("btn-deemph");
    if (capSp) capSp.textContent = "Recommended — has attachments";
  } else {
    btnSp.classList.add("btn-deemph");
    if (capRecord) capRecord.textContent = "Recommended — no attachments";
  }
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
  await restoreProjectSelectionForCurrentEmail();
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
// Quick-Win #6: localStorage cache for the projects/clients list. Pane opens
// instantly with last-known data; freshness fetch runs in the background and
// re-renders if anything changed. Without this, the pane shows a blank state
// while the V2 fetch round-trips Supabase (~300-800ms cold, longer on slow
// VPN). With cache, perceived open time drops to ~50ms.
const PROJECTS_CACHE_KEY = "settyPms:addinProjectsCacheV2";
const PROJECTS_CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 24h hard limit; revalidate every open

function loadProjectsCache() {
  try {
    const raw = localStorage.getItem(PROJECTS_CACHE_KEY);
    if (!raw) return null;
    const cached = JSON.parse(raw);
    if (!cached || !Array.isArray(cached.projects)) return null;
    if (!cached.savedAt || (Date.now() - cached.savedAt) > PROJECTS_CACHE_TTL_MS) return null;
    return cached;
  } catch {
    return null;
  }
}

function saveProjectsCache(projects, clients, versionMap) {
  try {
    const payload = {
      projects,
      clients: clients || [],
      versionMap: versionMap || {},
      savedAt: Date.now(),
    };
    localStorage.setItem(PROJECTS_CACHE_KEY, JSON.stringify(payload));
  } catch (e) {
    // QuotaExceeded is the most likely failure; silently drop. Cache is an
    // optimization, not a correctness requirement.
    console.warn("Projects cache save failed (will work without cache):", e.message);
  }
}

async function loadProjects() {
  // Hydrate from cache instantly (if available) so the pane is responsive
  // even before the fresh fetch returns. The cache holds the *projects array*
  // (post-archived-filter) and the version map; we'll overwrite both when
  // the fresh fetch completes.
  const cached = loadProjectsCache();
  let renderedFromCache = false;
  if (cached) {
    allProjects = cached.projects;
    allClients = cached.clients;
    for (const [id, ver] of Object.entries(cached.versionMap || {})) {
      _projectVersionCache.set(id, ver);
    }
    renderCompanySuggestions();
    renderedFromCache = true;
  }

  // Prefer V2 (per-project rows). Falls back to legacy pms_data if V2 tables
  // don't exist yet or are empty (pre-migration). Once PMS migrates, V2 is
  // authoritative and the legacy row becomes a static safety net.
  try {
    const [pRes, cRes] = await Promise.all([
      fetch(SUPABASE_URL + "/rest/v1/pms_projects?select=id,project,version", { headers: SB_HEADERS }),
      fetch(SUPABASE_URL + "/rest/v1/pms_clients?select=client", { headers: SB_HEADERS }),
    ]);
    if (pRes.ok && cRes.ok) {
      const pRows = await pRes.json();
      const cRows = await cRes.json();
      if (pRows && pRows.length > 0) {
        // V2 path
        allProjects = pRows.map(r => r.project).filter(p => p && !p.archived);
        const versionMap = {};
        for (const r of pRows) {
          _projectVersionCache.set(r.id, r.version);
          versionMap[r.id] = r.version;
        }
        allClients = (cRows || []).map(r => r.client).filter(Boolean);
        renderCompanySuggestions();
        // Refresh the cache with the latest data
        saveProjectsCache(allProjects, allClients, versionMap);
        return;
      }
    }
  } catch (e) {
    console.warn("V2 loadProjects failed, falling back to legacy:", e.message);
    // If we have cached data and the network fetch failed, leave the cache
    // populated and surface a soft warning rather than blowing away the UI.
    if (renderedFromCache) {
      console.info("Working from cached projects (offline or transient error). Saves will revalidate.");
      return;
    }
  }
  // Legacy fallback
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

// ─── V2 SAVE GUARD (per-project rows + optimistic concurrency) ───────────────
// Uses pms_projects (one row per project) instead of pms_data.projects[]. The
// row carries a `version` int incremented on every UPDATE; saves do
// PATCH ... WHERE id=? AND version=? — if 0 rows match, someone else saved
// first and we throw a structured ConflictError that callers can surface.
//
// Falls back gracefully:
//   - If pms_projects doesn't exist or the row doesn't exist (pre-migration),
//     uses the legacy whole-array PATCH path.
//   - If the GET-fresh fails (offline), uses the in-memory cache.

class AddinConflictError extends Error {
  constructor(message, projectId, cloudRow) {
    super(message);
    this.name = "AddinConflictError";
    this.projectId = projectId;
    this.cloudRow = cloudRow;
  }
}

// Per-project version cache so we know what version we last loaded.
// Populated lazily on first fetch.
const _projectVersionCache = new Map();

async function fetchFreshProjectV2(projectId) {
  const url = SUPABASE_URL + "/rest/v1/pms_projects?id=eq." + encodeURIComponent(projectId) + "&select=project,version";
  const res = await fetch(url, { headers: SB_HEADERS });
  if (!res.ok) throw new Error("pms_projects GET HTTP " + res.status);
  const rows = await res.json();
  if (!rows || rows.length === 0) return null; // not migrated yet
  _projectVersionCache.set(projectId, rows[0].version);
  return { project: rows[0].project, version: rows[0].version };
}

async function saveProjectRowV2(project, expectedVersion) {
  const url = SUPABASE_URL + "/rest/v1/pms_projects?id=eq." + encodeURIComponent(project.id) +
              "&version=eq." + expectedVersion;
  const res = await fetch(url, {
    method: "PATCH",
    headers: { ...SB_HEADERS, "Prefer": "return=representation" },
    body: JSON.stringify({
      project,
      version: expectedVersion + 1,
      updated_at: new Date().toISOString(),
    }),
  });
  if (!res.ok) throw new Error("pms_projects PATCH HTTP " + res.status);
  const result = await res.json();
  if (!result || result.length === 0) {
    // version mismatch — re-fetch to give caller something to merge
    const fresh = await fetchFreshProjectV2(project.id);
    throw new AddinConflictError(
      "Project " + project.id + " was modified by someone else (cloud v" +
      (fresh?.version ?? "?") + ", you had v" + expectedVersion + ")",
      project.id, fresh
    );
  }
  _projectVersionCache.set(project.id, result[0].version);
  return result[0].version;
}

// Pre-migration fallback: legacy whole-array PATCH against pms_data.
async function legacyApplyLocalChangeAndSave(projectId, mutateProject) {
  let freshProjects;
  try {
    const res = await fetch(SUPABASE_URL + "/rest/v1/pms_data?id=eq.singleton&select=projects", {
      headers: SB_HEADERS,
    });
    if (!res.ok) throw new Error("HTTP " + res.status);
    const rows = await res.json();
    freshProjects = (rows?.[0]?.projects) || [];
  } catch (e) {
    console.warn("legacyApplyLocalChangeAndSave: re-fetch failed, using cached allProjects:", e.message);
    freshProjects = allProjects;
  }
  const idx = freshProjects.findIndex(p => p.id === projectId);
  if (idx < 0) throw new Error("Project no longer exists in PMS.");
  const mutated = mutateProject(freshProjects[idx]);
  if (!mutated || !mutated.id) throw new Error("mutator returned invalid project");
  await saveToSupabase(freshProjects.map((p, i) => i === idx ? mutated : p));
  allProjects = allProjects.map(p => p.id === projectId ? mutated : p);
  if (selectedProject && selectedProject.id === projectId) selectedProject = mutated;
  return mutated;
}

// Cached migration-status flag. Once we know V2 is canonical, we never fall
// back to the legacy path — falling back would write a stale projects-array
// snapshot to pms_data, creating a divergent shadow copy that nothing reads
// but might mislead future debugging.
let _migrationKnownComplete = false;

async function _checkAddinMigrationStatus() {
  if (_migrationKnownComplete) return true;
  try {
    const res = await fetch(SUPABASE_URL + "/rest/v1/pms_meta?id=eq.migration_status&select=data", { headers: SB_HEADERS });
    if (!res.ok) return false;
    const rows = await res.json();
    if (rows?.[0]?.data?.v1_complete) {
      _migrationKnownComplete = true;
      return true;
    }
    return false;
  } catch {
    return false;
  }
}

// Main entry point — used by all save callsites in the add-in.
async function applyLocalChangeAndSave(projectId, mutateProject) {
  if (!projectId) throw new Error("applyLocalChangeAndSave: missing projectId");

  // Try V2 path first
  let fresh;
  let v2FetchFailed = false;
  try {
    fresh = await fetchFreshProjectV2(projectId);
  } catch (e) {
    v2FetchFailed = true;
    console.warn("V2 fetch failed:", e.message);
  }

  if (fresh) {
    // V2 happy path
    const mutated = mutateProject(fresh.project);
    if (!mutated || !mutated.id) throw new Error("mutator returned invalid project");
    try {
      await saveProjectRowV2(mutated, fresh.version);
    } catch (e) {
      if (e instanceof AddinConflictError) {
        throw new Error("⚠ Save conflict: " + e.message + ". Refresh the add-in pane and try again.");
      }
      throw e;
    }
    allProjects = allProjects.map(p => p.id === projectId ? mutated : p);
    if (selectedProject && selectedProject.id === projectId) selectedProject = mutated;
    return mutated;
  }

  // No V2 row found OR V2 fetch errored. Before falling back to legacy, check
  // whether migration has already happened — if it has, the legacy path would
  // write a stale shadow copy that no one reads. In that case we surface a
  // clear error rather than silently writing to a dead-end table.
  const migrationDone = await _checkAddinMigrationStatus();
  if (migrationDone && !fresh) {
    // V2 migration is complete but this project doesn't have a V2 row. Either
    // the project was added in legacy and never migrated (unlikely), or this
    // is a brand-new add via the add-in. Treat as INSERT.
    try {
      const mutated = mutateProject({ id: projectId });
      if (!mutated?.id) throw new Error("mutator returned invalid project");
      const res = await fetch(SUPABASE_URL + "/rest/v1/pms_projects", {
        method: "POST",
        headers: SB_HEADERS,
        body: JSON.stringify({ id: projectId, project: mutated, version: 1, updated_at: new Date().toISOString() }),
      });
      if (!res.ok) throw new Error("pms_projects POST HTTP " + res.status);
      _projectVersionCache.set(projectId, 1);
      allProjects = allProjects.map(p => p.id === projectId ? mutated : p);
      if (selectedProject && selectedProject.id === projectId) selectedProject = mutated;
      return mutated;
    } catch (insertErr) {
      throw new Error("Could not save: V2 row missing for project " + projectId + " and INSERT failed: " + insertErr.message);
    }
  }
  if (migrationDone && v2FetchFailed) {
    // Migration is done but our V2 fetch failed transiently. Don't fall back
    // to legacy — surface the error so the user retries instead of writing
    // to a dead-end table.
    throw new Error("Cloud temporarily unreachable. Wait a few seconds and try again. (V2 fetch failed; not falling back to legacy because the data layer has migrated.)");
  }

  // Pre-migration: legacy path is still authoritative
  return legacyApplyLocalChangeAndSave(projectId, mutateProject);
}
async function saveProjectEmailRow(projectId, emailRecord, savedToSharePoint) {
  if (!projectId || !emailRecord?.msgId) return;
  const row = {
    record_id: emailRecord.id,
    project_id: projectId,
    msg_id: emailRecord.msgId,
    conversation_id: currentConversationId || null,
    subject: emailRecord.subject || "",
    from_name: emailRecord.from || "",
    from_address: emailRecord.fromAddress || "",
    email_date: emailRecord.date || null,
    saved_at: emailRecord.savedAt || new Date().toISOString(),
    sp_folder_url: emailRecord.spFolderUrl || "",
    saved_to_sharepoint: !!savedToSharePoint,
  };
  // Throws on failure (caller catches and surfaces a warning). Previously this
  // silently console.warn'd and returned, so users had no idea their email
  // was missing from the search index. The email is still in the project
  // record (saved by applyLocalChangeAndSave above), just not indexed for
  // PMS-side search until a re-save.
  const res = await fetch(SUPABASE_URL + "/rest/v1/" + PROJECT_EMAILS_TABLE, {
    method: "POST",
    headers: { ...SB_HEADERS, Prefer: "return=minimal" },
    body: JSON.stringify(row),
  });
  if (!res.ok) {
    const errText = await res.text();
    throw new Error("pms_project_emails POST HTTP " + res.status + ": " + errText.slice(0, 150));
  }
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
function getConversationProjectMap() {
  try {
    return JSON.parse(localStorage.getItem(EMAIL_CONVO_PROJECT_MAP_STORAGE_KEY) || "{}");
  } catch {
    return {};
  }
}
function saveConversationProjectMap(map) {
  localStorage.setItem(EMAIL_CONVO_PROJECT_MAP_STORAGE_KEY, JSON.stringify(map || {}));
}
async function saveSharedConversationProjectTag(conversationId, projectId) {
  if (!conversationId || !projectId) return;
  const payload = {
    conversation_id: conversationId,
    project_id: projectId,
    tagged_by: msalAccount?.username || msalAccount?.name || "unknown",
    updated_at: new Date().toISOString(),
  };
  const url = SUPABASE_URL + "/rest/v1/" + EMAIL_THREAD_TAGS_TABLE + "?on_conflict=conversation_id";
  try {
    const res = await fetch(url, {
      method: "POST",
      headers: { ...SB_HEADERS, "Prefer": "resolution=merge-duplicates,return=minimal" },
      body: JSON.stringify(payload),
    });
    if (!res.ok) {
      const errText = await res.text();
      console.warn("Shared conversation tag save failed:", res.status, errText);
    }
  } catch (e) {
    console.warn("Shared conversation tag save failed:", e);
  }
}
async function clearSharedConversationProjectTag(conversationId) {
  if (!conversationId) return;
  const url = SUPABASE_URL + "/rest/v1/" + EMAIL_THREAD_TAGS_TABLE + "?conversation_id=eq." + encodeURIComponent(conversationId);
  try {
    const res = await fetch(url, { method: "DELETE", headers: { ...SB_HEADERS, Prefer: "return=minimal" } });
    if (!res.ok) {
      const errText = await res.text();
      console.warn("Shared conversation tag clear failed:", res.status, errText);
    }
  } catch (e) {
    console.warn("Shared conversation tag clear failed:", e);
  }
}
async function getSharedConversationProjectId(conversationId) {
  if (!conversationId) return "";
  const url =
    SUPABASE_URL +
    "/rest/v1/" +
    EMAIL_THREAD_TAGS_TABLE +
    "?conversation_id=eq." + encodeURIComponent(conversationId) +
    "&select=project_id&limit=1";
  try {
    const res = await fetch(url, { headers: SB_HEADERS });
    if (!res.ok) return "";
    const rows = await res.json();
    return rows?.[0]?.project_id || "";
  } catch {
    return "";
  }
}

// Returns full thread-tag info including who tagged it. Used by
// restoreProjectSelectionForCurrentEmail to show an attribution banner
// when a colleague tagged this thread (so the user knows their tag
// inheritance came from someone else, not their own past click).
async function getSharedConversationTag(conversationId) {
  if (!conversationId) return null;
  const url =
    SUPABASE_URL +
    "/rest/v1/" +
    EMAIL_THREAD_TAGS_TABLE +
    "?conversation_id=eq." + encodeURIComponent(conversationId) +
    "&select=project_id,tagged_by,updated_at&limit=1";
  try {
    const res = await fetch(url, { headers: SB_HEADERS });
    if (!res.ok) return null;
    const rows = await res.json();
    return rows?.[0] || null;
  } catch {
    return null;
  }
}
function isProjectAwarded(project) {
  if (!project) return true;
  const explicit = [project.awarded, project.isAwarded, project.is_awarded].find(v => typeof v === "boolean");
  if (typeof explicit === "boolean") return explicit;

  const statusText = [
    project.awardStatus,
    project.projectStatus,
    project.status,
    project.stage,
    project.phase,
    project.lifecycleStatus,
    project.bidStatus,
  ].find(v => typeof v === "string" && v.trim());

  if (!statusText) return true;
  const normalized = statusText.trim().toLowerCase();
  if (/(awarded|active|construction|in progress|won)/.test(normalized)) return true;
  if (/(pipeline|not awarded|proposal|bidding|bid|pursuit|precon|opportunity|lead)/.test(normalized)) return false;
  return true;
}

function applyPipelineUiRules() {
  const isPipeline = !!selectedProject && !isProjectAwarded(selectedProject);
  const hint = document.getElementById("projectPipelineHint");
  if (hint) {
    if (isPipeline) {
      hint.textContent = "Pipeline project (not awarded yet): post-award actions are disabled.";
      hint.style.display = "block";
    } else {
      hint.textContent = "";
      hint.style.display = "none";
    }
  }
  if (!isPipeline) return;

  const keepEnabled = new Set([
    "saveRecordBtn",
    "openPmsBtn",
    "openDashboardBtn",
    "addParticipantBtn",
  ]);
  const actionButtons = [
    "saveSpBtn",
    "saveRecordBtn",
    "openPmsBtn",
    "openSpFolderBtn",
    "openDashboardBtn",
    "logNoteBtn",
    "newActionItemBtn",
    "logRfiBtn",
    "logSubBtn",
    "manualMilestoneBtn",
    "addParticipantBtn",
  ];
  actionButtons.forEach(id => {
    const btn = document.getElementById(id);
    if (!btn) return;
    btn.disabled = !keepEnabled.has(id);
  });
}

function refreshOneNoteLinkBanner() {
  const banner = document.getElementById("oneNoteLinkBanner");
  const anchor = document.getElementById("oneNoteLinkAnchor");
  if (!banner || !anchor) return;
  if (!selectedProject) { banner.style.display = "none"; return; }

  const itemId    = emailItem?.itemId || "";
  const icalUId   = currentItemICalUId || "";
  const notes     = selectedProject.notes || [];

  // Match on personal itemId first (instant), then fall back to iCalUId
  // which is shared across all attendees of the same meeting.
  const note = notes.find(n =>
    n.oneNoteUrl && (
      (itemId  && n.sourceItemId      === itemId)  ||
      (icalUId && n.sourceCalendarUId === icalUId)
    )
  ) || null;

  if (note?.oneNoteUrl) {
    anchor.href = note.oneNoteUrl;
    anchor.textContent = (note.category || "Client Meeting") + " notes — Open in OneNote";
    banner.style.display = "block";
  } else {
    banner.style.display = "none";
  }
}

function getProjectTeamMembers(project) {
  if (!project) return [];
  const rawLists = [
    project.projectTeam,
    project.teamMembers,
    project.team,
    project.internalTeam,
    project.staff,
    project.assignedTeam,
  ].filter(Array.isArray);

  const picked = rawLists.flat().map(member => {
    if (typeof member === "string") return member.trim();
    if (!member || typeof member !== "object") return "";
    return (member.name || member.displayName || member.fullName || member.userName || member.email || "").trim();
  }).filter(Boolean);

  return [...new Set(picked)].sort((a, b) => a.localeCompare(b));
}

function refreshActionItemOwnerOptions() {
  const ownerSelect = document.getElementById("actionItemOwner");
  if (!ownerSelect) return;
  const teamMembers = getProjectTeamMembers(selectedProject);
  const previous = ownerSelect.value || "";
  ownerSelect.innerHTML = '<option value="">— Select team member —</option>'
    + teamMembers.map(name => `<option value="${escHtml(name)}">${escHtml(name)}</option>`).join("");
  if (previous && teamMembers.includes(previous)) ownerSelect.value = previous;
}
async function clearProjectTagForCurrentEmail() {
  const msgId = getCurrentMessageRestId();
  if (msgId) {
    const map = getEmailProjectMap();
    if (map[msgId]) {
      delete map[msgId];
      localStorage.setItem(EMAIL_PROJECT_MAP_STORAGE_KEY, JSON.stringify(map));
    }
  }
  // Use shared key (handles both emails and appointments) so the cross-device
  // tag also gets cleared, not just the email-conversation one.
  const sharedKey = await getCurrentSharedKey();
  if (sharedKey) {
    const convoMap = getConversationProjectMap();
    if (convoMap[sharedKey]) {
      delete convoMap[sharedKey];
      saveConversationProjectMap(convoMap);
    }
    await clearSharedConversationProjectTag(sharedKey);
  }
  setSelectedProject(null, false);
  setStatus("actionStatus", "info", "Project tag cleared. Search and select the correct project.");
}

function setSelectedProject(project, persistForEmail = false) {
  selectedProject = project || null;
  const badge = document.getElementById("selectedProjectBadge");
  const badgeText = document.getElementById("selectedProjectBadgeText");
  const clearBtn = document.getElementById("clearProjectTagBtn");
  if (badge) {
    if (selectedProject) {
      const pipelineTag = isProjectAwarded(selectedProject) ? "" : " (Pipeline)";
      if (badgeText) {
        badgeText.textContent = "✓ " + (selectedProject.projectNumber ? selectedProject.projectNumber + " — " : "") + selectedProject.name + pipelineTag;
      } else {
        badge.textContent = "✓ " + (selectedProject.projectNumber ? selectedProject.projectNumber + " — " : "") + selectedProject.name + pipelineTag;
      }
      badge.style.display = "flex";
      if (clearBtn) clearBtn.style.display = "inline";
    } else {
      badge.style.display = "none";
      if (badgeText) badgeText.textContent = "";
      if (clearBtn) clearBtn.style.display = "none";
      if (!badgeText) badge.textContent = "";
    }
  }
  if (persistForEmail && selectedProject) {
    const msgId = getCurrentMessageRestId();
    if (msgId) {
      const map = getEmailProjectMap();
      map[msgId] = selectedProject.id;
      localStorage.setItem(EMAIL_PROJECT_MAP_STORAGE_KEY, JSON.stringify(map));
    }
    void (async () => {
      // Use the shared key (iCalUId for appointments, conversationId for emails)
      // so the tag is restorable from any device, by any attendee/recipient.
      // Previously this only handled emails — appointments silently skipped
      // the cloud tag write, meaning calendar-event project tags didn't sync.
      const sharedKey = await getCurrentSharedKey();
      if (!sharedKey) return;
      const convoMap = getConversationProjectMap();
      convoMap[sharedKey] = selectedProject.id;
      saveConversationProjectMap(convoMap);
      await saveSharedConversationProjectTag(sharedKey, selectedProject.id);
    })();
  }
  updateProjectQuickLinks();
  refreshActionItemOwnerOptions();
  refreshEmailSavedIndicator();
  refreshOneNoteLinkBanner();
  refreshCalendarStatus();
  applyPipelineUiRules();
  renderProjectSuggestions();
  void renderDateSuggestions();
}
// Refreshes the "Calendar event detected" / "Already logged" status message.
// Must be called AFTER selectedProject and currentItemICalUId are both resolved.
function refreshCalendarStatus() {
  if (currentItemKind !== "appointment") return;
  const itemId  = emailItem?.itemId || "";
  const icalUId = currentItemICalUId || "";
  const notes   = selectedProject?.notes || [];
  const logged  = notes.find(n =>
    (itemId  && n.sourceItemId      === itemId)  ||
    (icalUId && n.sourceCalendarUId === icalUId)
  );
  const logNoteBtn = document.getElementById("logNoteBtn");
  if (logged) {
    const loggedLabel = logged.category ? logged.category + " note" : "Note";
    setStatus("actionStatus", "success", "✓ " + loggedLabel + " already logged for this event." + (logged.oneNoteUrl ? " 📓" : ""));
    if (logNoteBtn) logNoteBtn.disabled = true;
  } else {
    setStatus("actionStatus", "info",
      "Calendar event detected: use 'Log as Note' for meetings/site visits and 'Add Participant to Contacts' for attendees.");
    if (logNoteBtn) logNoteBtn.disabled = false;
  }
}
async function restoreProjectSelectionForCurrentEmail() {
  const msgId = getCurrentMessageRestId();
  if (!allProjects.length) return;
  let projectId = "";
  let taggedByOther = null; // colleague who tagged this thread, if any
  let restoredVia = "";     // which path actually found the project (for logging)
  if (msgId) {
    const map = getEmailProjectMap();
    projectId = map[msgId] || "";
    if (projectId) restoredVia = "localStorage-msgId";
  }
  if (!projectId) {
    // Use shared key — iCalUId for appointments, conversationId for emails.
    // Awaits the iCalUId Graph fetch internally for appointments.
    const sharedKey = await getCurrentSharedKey();
    if (sharedKey) {
      const convoMap = getConversationProjectMap();
      projectId = convoMap[sharedKey] || "";
      if (projectId) restoredVia = "localStorage-sharedKey";
      if (!projectId) {
        const tag = await getSharedConversationTag(sharedKey);
        projectId = tag?.project_id || "";
        if (projectId) {
          restoredVia = "cloud-sharedKey";
          convoMap[sharedKey] = projectId;
          saveConversationProjectMap(convoMap);
          const myUsername = (msalAccount?.username || "").toLowerCase();
          const taggedBy = (tag?.tagged_by || "").toLowerCase();
          if (taggedBy && taggedBy !== myUsername) {
            taggedByOther = tag.tagged_by;
          }
        }
      }
    }
  }

  // FALLBACK — scan all projects for a note matching this item.
  // The localStorage maps and shared tag table can miss for notes saved before
  // recent fixes (or from a different itemId context — e.g., organizer view
  // vs attendee view, calendar-folder shifts, recurring meeting occurrences).
  // The note itself is the source of truth: if the note exists in a project
  // with sourceItemId or sourceCalendarUId matching the current item, that
  // project IS the right answer regardless of any external mapping.
  if (!projectId) {
    const itemId  = emailItem?.itemId || "";
    const icalUId = currentItemICalUId || "";
    if (itemId || icalUId) {
      for (const p of allProjects) {
        const notes = p.notes || [];
        const matchingNote = notes.find(n =>
          (itemId  && n.sourceItemId      === itemId)  ||
          (icalUId && n.sourceCalendarUId === icalUId)
        );
        if (matchingNote) {
          projectId = p.id;
          restoredVia = "note-scan-" + (matchingNote.sourceItemId === itemId ? "itemId" : "icalUId");
          // Backfill the localStorage map AND the cloud shared tag so future
          // lookups hit the fast path. Self-healing for legacy notes.
          if (msgId) {
            const map = getEmailProjectMap();
            map[msgId] = projectId;
            localStorage.setItem(EMAIL_PROJECT_MAP_STORAGE_KEY, JSON.stringify(map));
          }
          const sharedKey = currentItemICalUId || (await getCurrentSharedKey());
          if (sharedKey) {
            const convoMap = getConversationProjectMap();
            convoMap[sharedKey] = projectId;
            saveConversationProjectMap(convoMap);
            // Fire-and-forget cloud upsert
            void saveSharedConversationProjectTag(sharedKey, projectId);
          }
          break;
        }
      }
    }
  }

  if (!projectId) {
    // No tag exists for this thread — surface ranked suggestions instead so
    // users don't always have to type/search. The chip area is hidden by
    // setSelectedProject the moment they pick one.
    renderProjectSuggestions();
    return;
  }
  const project = allProjects.find(p => p.id === projectId);
  if (project) {
    console.info("[restore] Restored project via", restoredVia, ":", project.projectNumber || project.name);
    setSelectedProject(project, false);
    if (taggedByOther) {
      const projLabel = (project.projectNumber ? project.projectNumber + " — " : "") + project.name;
      setStatus("actionStatus", "info", "ℹ Auto-tagged to " + projLabel + " by " + taggedByOther + ". If wrong, click ✕ on the project chip to clear and pick another.");
    }
  }
}

// ─── PROJECT SUGGESTION ──────────────────────────────────────────────────────
// Ranks projects by likelihood of being "the project this email is about" using
// signals from subject, sender domain, and project number. We *suggest* (never
// auto-apply) — accuracy matters more than saved clicks given multiple active
// projects per client. Tier-1 (existing thread tag) is handled separately by
// restoreProjectSelectionForCurrentEmail.
//
// TUNE: weights below were chosen so that "subject contains a unique acronym"
// or "subject contains 1 distinctive name token" easily clears the threshold,
// but "sender domain matches client" alone does not. Adjust after real-world
// testing if you find the chip suggesting too eagerly or not enough.
const SUGGESTION_WEIGHTS = {
  projectNumberInSubject: 10,
  perNameTokenInSubject:   2,
  acronymInSubject:        3,
  senderDomainMatchClient: 1,
};
const SUGGESTION_MIN_SCORE = 2;
const SUGGESTION_MAX_RESULTS = 3;
// Words to ignore when tokenizing project names and subjects — too generic to
// signal anything ("Project Renovation" matching "Renovation Project" should
// not count as a hit).
const SUGGESTION_STOPWORDS = new Set([
  "the","and","of","for","at","to","a","an","or","by","on","in","with",
  "re","fwd","fw","project","renovation","reno","new","update","updated",
  "phase","building","bldg","floor","fl","st","ave","road","rd",
]);
function suggestionTokenize(s) {
  return (s || "")
    .toLowerCase()
    .replace(/[^\w\s]/g, " ")
    .split(/\s+/)
    .filter(t => t && t.length >= 2 && !SUGGESTION_STOPWORDS.has(t));
}
// Build the set of acronyms that COULD appear in a subject for a given project
// name. For "Queens College ADA": {qc, ca, qca}. We generate every contiguous
// 2+ word slice's initials so "QC" matches even when the project has more words
// after "Queens College".
function suggestionAcronyms(name) {
  const words = (name || "").split(/\s+/).filter(Boolean);
  const out = new Set();
  for (let i = 0; i < words.length; i++) {
    for (let j = i + 2; j <= words.length; j++) {
      const acro = words.slice(i, j).map(w => w.charAt(0).toLowerCase()).join("");
      if (acro.length >= 2 && acro.length <= 6 && /^[a-z]+$/.test(acro)) out.add(acro);
    }
  }
  return out;
}
function suggestProjects(subject, senderEmail) {
  const subj = (subject || "").toLowerCase();
  if (!subj && !senderEmail) return [];
  const subjTokens = new Set(suggestionTokenize(subj));
  // Project number heuristic — most Setty project numbers are 5 digits but we
  // accept 4-6 to be tolerant of legacy/special projects.
  const numMatches = [...subj.matchAll(/\b(\d{4,6})\b/g)].map(m => m[1]);
  const senderDomain = (senderEmail || "").toLowerCase().split("@")[1] || "";
  const senderClient = senderDomain ? getClientByEmail(senderEmail) : null;

  const scored = [];
  for (const p of (allProjects || [])) {
    if (!p || p.archived) continue;
    if (!p.name && !p.projectNumber) continue;
    let score = 0;
    const reasons = [];

    if (p.projectNumber && numMatches.includes(String(p.projectNumber))) {
      score += SUGGESTION_WEIGHTS.projectNumberInSubject;
      reasons.push("project # in subject");
    }

    const projTokens = suggestionTokenize(p.name || "");
    const tokenHits = projTokens.filter(t => subjTokens.has(t));
    if (tokenHits.length) {
      score += tokenHits.length * SUGGESTION_WEIGHTS.perNameTokenInSubject;
      reasons.push(tokenHits.length + " name word" + (tokenHits.length > 1 ? "s" : "") + " match");
    }

    const acros = suggestionAcronyms(p.name || "");
    const acroHit = [...acros].some(a => subjTokens.has(a));
    if (acroHit) {
      score += SUGGESTION_WEIGHTS.acronymInSubject;
      reasons.push("acronym match");
    }

    if (senderClient && p) {
      const projClient = (p.prime || p.clientName || "").toLowerCase().trim();
      if (projClient && projClient === (senderClient.name || "").toLowerCase().trim()) {
        score += SUGGESTION_WEIGHTS.senderDomainMatchClient;
        reasons.push("sender's company");
      }
    }

    if (score >= SUGGESTION_MIN_SCORE) {
      scored.push({ project: p, score, reasons });
    }
  }

  scored.sort((a, b) => b.score - a.score || (a.project.name || "").localeCompare(b.project.name || ""));
  return scored.slice(0, SUGGESTION_MAX_RESULTS);
}
function renderProjectSuggestions() {
  const block = document.getElementById("suggestionBlock");
  const chips = document.getElementById("suggestionChips");
  const labelText = document.getElementById("suggestionLabelText");
  if (!block || !chips) return;
  if (selectedProject) { block.style.display = "none"; chips.innerHTML = ""; return; }

  const subject = (typeof emailItem?.subject === "string") ? emailItem.subject : "";
  const results = suggestProjects(subject, emailFromAddress);
  if (!results.length) { block.style.display = "none"; chips.innerHTML = ""; return; }

  if (labelText) labelText.textContent = results.length === 1 ? "Suggested project" : "Possible projects";
  chips.innerHTML = results.map((r, i) => `
    <button type="button" class="suggestion-chip" data-id="${escHtml(r.project.id)}">
      <div class="sc-num">${escHtml(r.project.projectNumber || "")}</div>
      <div class="sc-name">${escHtml(r.project.name || "")}</div>
      <div class="sc-reason">${escHtml(r.reasons.join(" · "))}</div>
    </button>
  `).join("");
  chips.querySelectorAll(".suggestion-chip").forEach(el => {
    el.onclick = () => {
      const proj = allProjects.find(p => p.id === el.dataset.id);
      if (proj) setSelectedProject(proj, true);
    };
  });
  block.style.display = "block";
}

// Cached per-item body fetch — avoids re-hitting Graph each time
// renderDateSuggestions runs (e.g., on project re-selection within the same
// email). Keyed on itemContextGeneration so it auto-invalidates when the
// user opens a different email.
let _dateSuggestBodyCache = { gen: -1, text: "" };

async function renderDateSuggestions() {
  const block = document.getElementById("dateSuggestionBlock");
  const chips = document.getElementById("dateSuggestionChips");
  const labelText = document.getElementById("dateSuggestionLabelText");
  if (!block || !chips) return;

  // Hide chips when there's no project to attach a milestone to.
  // Also skip in compose-mode appointments and when no email is loaded.
  if (!selectedProject || !emailItem || currentItemKind === "appointment") {
    block.style.display = "none"; chips.innerHTML = ""; return;
  }

  const myGen = itemContextGeneration;
  let text = "";
  if (_dateSuggestBodyCache.gen === myGen) {
    text = _dateSuggestBodyCache.text;
  } else {
    try {
      const token = await getToken();
      const html  = await getEmailBodyHtml(token);
      if (myGen !== itemContextGeneration) return; // user moved on
      const tmp = document.createElement("div");
      tmp.innerHTML = html || "";
      text = (tmp.innerText || tmp.textContent || "").replace(/\s+/g, " ");
      _dateSuggestBodyCache = { gen: myGen, text };
    } catch { return; /* non-fatal — chips just won't show */ }
  }

  const dates = extractDueDates(text, emailItem?.dateTimeCreated).slice(0, 3);
  if (!dates.length) { block.style.display = "none"; chips.innerHTML = ""; return; }

  if (labelText) labelText.textContent = dates.length === 1 ? "Possible date" : "Possible dates";
  chips.innerHTML = dates.map(d => {
    const friendly = new Date(d.iso + "T12:00:00").toLocaleDateString("en-US", {
      weekday: "short", month: "short", day: "numeric", year: "numeric"
    });
    const ctxShort = (d.ctx || "").length > 110 ? d.ctx.slice(0, 110) + "…" : d.ctx;
    return `
      <button type="button" class="suggestion-chip" data-iso="${escHtml(d.iso)}">
        <div class="sc-name">${escHtml(friendly)}${d.hasKeyword ? ' <span class="pill" style="background:#1e3a5f;color:#60b4ff;">deadline</span>' : ""}</div>
        <div class="sc-reason">${escHtml(ctxShort)}</div>
      </button>
    `;
  }).join("");
  chips.querySelectorAll(".suggestion-chip").forEach(el => {
    el.onclick = () => openMilestoneFormFromChip(el.dataset.iso);
  });
  block.style.display = "block";
}

function openMilestoneFormFromChip(iso) {
  showView("datesView");
  const list = document.getElementById("datesList");
  if (list) list.innerHTML = "";
  prefillMilestone(iso);
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

// Phase 3: compress HTML email body to base64 deflate before storing in the
// project record. Identical implementation to PMS so records are interchangeable.
// Logs a clear warning when compression fails non-trivially (i.e., we had real
// HTML to compress but couldn't), so failures don't silently produce empty
// bodyHtmlCompressed fields the user later sees as "Live-fetched only".
function compressHtmlAddin(html) {
  if (!html) return "";
  if (typeof pako === "undefined") {
    console.warn("compressHtmlAddin: pako library not loaded — email will save with empty bodyHtmlCompressed and require live-fetch on view");
    return "";
  }
  try {
    const deflated = pako.deflate(html, { level: 6 });
    let binary = "";
    const CHUNK = 0x8000;
    for (let i = 0; i < deflated.length; i += CHUNK) {
      binary += String.fromCharCode.apply(null, deflated.subarray(i, i + CHUNK));
    }
    return btoa(binary);
  } catch (e) {
    console.warn("compressHtmlAddin failed for", html.length, "chars:", e);
    return "";
  }
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
  // Kick off the email.html upload in parallel with the attachment loop —
  // they don't depend on each other, so why serialize them.
  const emailHtmlPromise = fetch(
    "https://graph.microsoft.com/v1.0/drives/" + driveId + "/root:/" + encodeDrivePath(targetPath) + "/email.html:/content",
    {
      method: "PUT",
      headers: { "Authorization": "Bearer " + token, "Content-Type": "text/html" },
      body: buildEmailHtml(bodyHtml),
    }
  );

  // Quick-Win #3: parallelize attachment uploads with bounded concurrency.
  // Graph throttles aggressively above ~5 parallel writes per session; 3 is
  // a safe ceiling that gives meaningful speedup (~3x for emails with 5+
  // attachments) without triggering 429 backoff.
  const ATTACHMENT_CONCURRENCY = 3;
  async function uploadInBatches(items, doUpload) {
    const failures = [];
    let succeeded = 0;
    for (let i = 0; i < items.length; i += ATTACHMENT_CONCURRENCY) {
      const batch = items.slice(i, i + ATTACHMENT_CONCURRENCY);
      const results = await Promise.allSettled(batch.map(doUpload));
      results.forEach((r, idx) => {
        const item = batch[idx];
        if (r.status === "fulfilled" && r.value) succeeded++;
        else failures.push((item.name || "attachment") + (r.status === "rejected" ? " (" + (r.reason?.message || "error").slice(0, 60) + ")" : ""));
      });
    }
    return { succeeded, failures };
  }

  try {
    let count = 0;
    // Prefer Outlook item APIs for attachment bytes; this is the most reliable in add-ins.
    const officeAtts = await getOfficeFileAttachments();
    if (officeAtts.length) {
      lastAttachmentUploadStats.attempted = officeAtts.length;
      const { succeeded, failures } = await uploadInBatches(
        officeAtts,
        att => uploadAttachmentToSharePoint(driveId, token, targetPath, att.name, att.contentType, att.bytes)
      );
      count = succeeded;
      lastAttachmentUploadStats.failed.push(...failures);
      lastAttachmentUploadStats.uploaded = count;
      // Make sure email.html upload completed before returning
      await emailHtmlPromise;
      return count;
    }
    // Fallback to Graph attachment APIs when Office APIs are unavailable.
    // Download bytes in parallel too — for large emails with many attachments,
    // this is where most of the time was being spent.
    const restId = Office.context.mailbox.convertToRestId(emailItem.itemId, Office.MailboxEnums.RestVersion.v2_0);
    const attData = await graphFetch("GET", "/me/messages/" + restId + "/attachments", null, token);
    const fileAtts = (attData?.value || []).filter(att => att["@odata.type"] === "#microsoft.graph.fileAttachment");
    lastAttachmentUploadStats.attempted = fileAtts.length;

    const { succeeded, failures } = await uploadInBatches(fileAtts, async (att) => {
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
          throw new Error("download " + rawRes.status);
        }
        bytes = new Uint8Array(await rawRes.arrayBuffer());
      }
      if (!bytes) return false;
      return uploadAttachmentToSharePoint(driveId, token, targetPath, att.name, att.contentType, bytes);
    });
    count = succeeded;
    lastAttachmentUploadStats.failed.push(...failures);
    lastAttachmentUploadStats.uploaded = count;
    await emailHtmlPromise;
    return count;
  } catch (e) {
    console.warn("Attachment upload failed:", e.message);
    lastAttachmentUploadStats.failed.push("Unhandled error: " + e.message);
    try { await emailHtmlPromise; } catch {}
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
// ─── CONCURRENT-SAVE GUARD ────────────────────────────────────────────────────
// Wraps a save function so a second click during the first call's flight is
// ignored rather than launching a parallel save that races the version
// counter and produces phantom errors. Disables the supplied buttons during
// the flight; restores them when finished (success or error). Uses the
// `saveInFlight` module flag as a process-wide lock — only one save of any
// type can run at a time, since they all go through applyLocalChangeAndSave
// and would otherwise race on selectedProject's version.
async function withSaveGuard(name, fn, buttonIds = []) {
  if (saveInFlight) {
    setStatus("actionStatus", "info", "⏳ Another save is in progress; please wait.");
    return;
  }
  saveInFlight = true;
  const buttons = buttonIds.map(id => document.getElementById(id)).filter(Boolean);
  const wasDisabled = buttons.map(b => b.disabled);
  buttons.forEach(b => { b.disabled = true; });
  try {
    return await fn();
  } finally {
    saveInFlight = false;
    buttons.forEach((b, i) => { if (!wasDisabled[i]) b.disabled = false; });
  }
}

// ─── SAVE TO SHAREPOINT ───────────────────────────────────────────────────────
async function doSaveToSharePoint() {
  return withSaveGuard("save-sp", _doSaveToSharePoint, ["saveSpBtn", "saveRecordBtn"]);
}
async function _doSaveToSharePoint() {
  if (!selectedProject) { setStatus("actionStatus", "error", "Select a project first."); return; }
  if (!selectedProject.projectFolderUrl) { setStatus("actionStatus", "error", "No SharePoint folder on this project. Create one in the PMS first."); return; }
const currentMsgId = getCurrentMessageRecordId();
const existingRecord = findSavedEmailRecord(selectedProject, currentMsgId);
if (existingRecord) {
  refreshEmailSavedIndicator();
  return;
}
  setStatus("actionStatus", "info", existingRecord ? "⏳ Re-saving to SharePoint (retrying attachments)…" : "⏳ Saving to SharePoint…");
  try {
    const token = await getToken();
    const { driveId } = await resolveSpIds();
    // Phase 3: fetch body HTML once up front so we can both upload to SharePoint
    // AND store the compressed version on the project record.
    // Track body-fetch failure separately so we can surface it in the success
    // message — previously a silent "" fallback meant the user thought everything
    // worked but the email record had no readable body.
    const bodyHtml = await getEmailBodyHtml(token);
    const bodyFetchFailed = !bodyHtml || bodyHtml.length === 0;
    const compressedBody = bodyHtml ? compressHtmlAddin(bodyHtml) : "";
    const projFolderName = decodeURIComponent(selectedProject.projectFolderUrl.split("/").pop());
    const d = new Date(emailItem.dateTimeCreated);
    const safeSubject = (emailItem.subject || "No Subject").replace(/[\\/:*?"<>|]/g, "-").replace(/\s+/g, " ").trim().slice(0, 70);
    const emailFolderName = d.getFullYear() + "_" + String(d.getMonth() + 1).padStart(2, "0") + "_" + String(d.getDate()).padStart(2, "0") + " " + safeSubject;
    const emailsPath  = await ensureSpFolder(driveId, token, projFolderName, "Emails");
    const targetPath  = await ensureSpFolder(driveId, token, emailsPath, emailFolderName);
    await writeSpMetadataSidecar(driveId, token, targetPath, buildAddinMetadata(selectedProject, "correspondence"));
    const attCount    = await uploadEmailAndAttachments(driveId, token, targetPath);
    const from = emailItem.from;
    const spFolderUrl = SP_BASE_URL + "/" + encodeURIComponent(projFolderName) + "/Emails/" + encodeURIComponent(emailFolderName);
    const msgId = currentMsgId;
    const emailRecord = {
      id: uid(), msgId,
      subject: emailItem.subject || "",
      from: from?.displayName || "",
      fromAddress: from?.emailAddress || "",
      date: emailItem.dateTimeCreated,
      bodyText: "",
      bodyHtmlCompressed: compressedBody,
      bodyHtmlSize: bodyHtml.length,
      spFolderUrl, links: [],
      savedAt: new Date().toISOString(),
    };
    // Re-fetch latest projects, then append email to the FRESH copy of this project.
    // Prevents the add-in from overwriting concurrent PMS edits made during this session.
    await applyLocalChangeAndSave(selectedProject.id, fresh => ({
      ...fresh,
      emails: [...(fresh.emails || []), emailRecord],
    }));
    let indexSaveFailed = false;
    try {
      await saveProjectEmailRow(selectedProject.id, emailRecord, true);
    } catch (idxErr) {
      console.warn("saveProjectEmailRow failed:", idxErr);
      indexSaveFailed = true;
    }
    const attMsg = attCount ? " + " + attCount + " attachment" + (attCount > 1 ? "s" : "") : "";
    const attempted = lastAttachmentUploadStats?.attempted || 0;

    // Compose status message accounting for ALL partial-failure modes (Issues #6, #7, #8):
    //   - body fetch failed → email.html was uploaded with "(No body)"
    //   - some attachments failed → user sees "X / Y uploaded"
    //   - all attachments failed → error
    //   - search-index write failed → email won't appear in PMS email-search results
    const warnings = [];
    if (bodyFetchFailed) warnings.push("⚠ Email body could not be retrieved from Outlook — saved record will show '(No body)' until you re-save when the email is reachable.");
    if (attempted > 0 && attCount > 0 && attCount < attempted) {
      const failedNames = (lastAttachmentUploadStats?.failed || []).slice(0, 2).join("; ");
      warnings.push("⚠ Only " + attCount + "/" + attempted + " attachments uploaded" + (failedNames ? " (failed: " + failedNames + ")" : "") + ".");
    }
    if (indexSaveFailed) warnings.push("⚠ Email saved to project, but search-index write failed — it may not appear in PMS email searches until you resave or PMS is reloaded.");

    if (attempted > 0 && attCount === 0) {
      const sample = (lastAttachmentUploadStats?.failed || []).slice(0, 2).join("; ");
      setStatus("actionStatus", "error", "Email saved, but 0/" + attempted + " attachments uploaded. " + (sample || "Open browser console for details.") + (warnings.length ? " " + warnings.join(" ") : ""));
    } else if (warnings.length > 0) {
      setStatus("actionStatus", "info", "✓ Saved to SharePoint" + attMsg + " and project record. " + warnings.join(" "));
    } else if (attempted === 0) {
      setStatus("actionStatus", "info", "Email saved to SharePoint, but no attachments were detected by Outlook/Graph for this message.");
    } else {
      setStatus("actionStatus", "success", "✓ Saved to SharePoint" + attMsg + " and project record.");
    }
    refreshEmailSavedIndicator();
  } catch (e) {
    setStatus("actionStatus", "error", "✗ " + e.message);
  }
}
async function doSaveToProjectRecordOnly() {
  return withSaveGuard("save-record", _doSaveToProjectRecordOnly, ["saveSpBtn", "saveRecordBtn"]);
}
async function _doSaveToProjectRecordOnly() {
  if (!selectedProject) { setStatus("actionStatus", "error", "Select a project first."); return; }
  if (emailItem?.hasAttachments) {
    setStatus("actionStatus", "error", "This email has attachments. Use 'Save to SharePoint + Project Record' instead.");
    return;
  }
  const msgId = getCurrentMessageRecordId();
  if (findSavedEmailRecord(selectedProject, msgId)) {
    refreshEmailSavedIndicator();
    return;
  }
  setStatus("actionStatus", "info", "⏳ Saving to project record…");
  try {
    // Phase 3: capture and compress body so PMS can render it without a Graph round-trip.
    const token = await getToken();
    const bodyHtml = await getEmailBodyHtml(token);
    const bodyFetchFailed = !bodyHtml || bodyHtml.length === 0;
    const compressedBody = bodyHtml ? compressHtmlAddin(bodyHtml) : "";
    const from = emailItem.from;
    const emailRecord = {
      id: uid(), msgId,
      subject: emailItem.subject || "",
      from: from?.displayName || "",
      fromAddress: from?.emailAddress || "",
      date: emailItem.dateTimeCreated,
      bodyText: "",
      bodyHtmlCompressed: compressedBody,
      bodyHtmlSize: bodyHtml.length,
      spFolderUrl: "", links: [],
      savedAt: new Date().toISOString(),
      savedToSharePoint: false,
    };
    await applyLocalChangeAndSave(selectedProject.id, fresh => ({
      ...fresh,
      emails: [...(fresh.emails || []), emailRecord],
    }));
    let indexSaveFailed = false;
    try {
      await saveProjectEmailRow(selectedProject.id, emailRecord, false);
    } catch (idxErr) {
      console.warn("saveProjectEmailRow failed:", idxErr);
      indexSaveFailed = true;
    }
    const warnings = [];
    if (bodyFetchFailed) warnings.push("⚠ Email body could not be retrieved from Outlook.");
    if (indexSaveFailed) warnings.push("⚠ Search-index write failed — may not appear in PMS email searches until next resave.");
    if (warnings.length > 0) {
      setStatus("actionStatus", "info", "✓ Saved to project record. " + warnings.join(" "));
    } else {
      setStatus("actionStatus", "success", "✓ Saved to project record (no SharePoint upload).");
    }
    refreshEmailSavedIndicator();
  } catch (e) {
    setStatus("actionStatus", "error", "✗ " + e.message);
  }
}
// ─── LOG NOTE ─────────────────────────────────────────────────────────────────
function buildAddinMeetingPageHtml(title, category, dateStr, participants, body) {
  const th = "padding:6px 12px;font-weight:bold;background:#f0f0f0;text-align:left;width:130px";
  const td = "padding:6px 12px";
  const dateFmt = dateStr ? new Date(dateStr).toLocaleDateString("en-US", {
    weekday: "long", month: "long", day: "numeric", year: "numeric"
  }) : "";
  const attendeeStr = (participants || [])
    .map(p => (p.displayName || p.emailAddress) + (p.label ? " (" + p.label + ")" : ""))
    .join(", ");
  return "<h1>" + title + "</h1>"
    + "<table style='border-collapse:collapse;width:100%;font-size:13px;margin-bottom:16px'>"
    + (dateFmt    ? "<tr><td style='" + th + "'>Date</td><td style='" + td + "'>" + dateFmt + "</td></tr>" : "")
    + "<tr><td style='" + th + "'>Type</td><td style='" + td + "'>" + category + "</td></tr>"
    + (attendeeStr ? "<tr><td style='" + th + "'>Attendees</td><td style='" + td + "'>" + attendeeStr + "</td></tr>" : "")
    + (body        ? "<tr><td style='" + th + "'>Notes</td><td style='" + td + "'><pre style='font-family:inherit;white-space:pre-wrap'>" + body + "</pre></td></tr>" : "")
    + "</table>"
    + "<h2>Discussion</h2><p>&nbsp;</p>"
    + "<h2>Decisions</h2><p>&nbsp;</p>"
    + "<h2>Action Items</h2>"
    + "<table style='border-collapse:collapse;width:100%'>"
    + "<tr style='background:#f0f0f0'><th style='" + td + ";text-align:left'>Item</th><th style='" + td + ";text-align:left'>Owner</th><th style='" + td + ";text-align:left'>Due</th></tr>"
    + "<tr><td style='" + td + "'>&nbsp;</td><td style='" + td + "'>&nbsp;</td><td style='" + td + "'>&nbsp;</td></tr>"
    + "</table>";
}

// HTML-escape OneNote title / metadata text to prevent breakage on `<`, `&`, etc.
function escapeOneNoteTextAddin(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

async function createAddinOneNotePage(project, title, body, category, dateStr) {
  const useTeams   = !!project.teamsOneNoteNotebookId;
  const notebookId = project.teamsOneNoteNotebookId || project.oneNoteNotebookId;
  // Route to the correct OneNote namespace based on where the notebook lives.
  // Teams notebooks were created under /groups/{teamId}/onenote — the same ID returns
  // 404 if looked up via /me/onenote because they are different namespaces.
  // Notes.ReadWrite (no .All) is sufficient for the groups endpoint.
  const baseUrl = useTeams
    ? `/groups/${TEAMS_TEAM_ID}/onenote`
    : `/me/onenote`;
  const sectionName = {
    "Client Meeting":       "Client Meetings",
    "Internal Meeting":     "Internal Meetings",
    "Meeting":              "Meetings",
    "Site Visit":           "Site Visits",
    "Client Communication": "Client Communications",
    "Decision":             "Decisions",
    "Issue":                "Issues",
    "Action Item":          "Action Items",
    "Internal":             "Internal Notes",
    "General":              "General Notes",
  }[category] || "General Notes";

  // Race-safe section lookup-or-create: catch 409 if two saves try to create
  // the same section simultaneously, then re-fetch to find the winner.
  const sectionsResp = await graphFetch("GET", `${baseUrl}/notebooks/${notebookId}/sections`);
  let section = (sectionsResp?.value || []).find(s => s.displayName === sectionName);
  if (!section) {
    try {
      section = await graphFetch("POST", `${baseUrl}/notebooks/${notebookId}/sections`, { displayName: sectionName });
    } catch (e) {
      // Another concurrent save may have created the section first — that's fine.
      if (!(e.message || "").match(/409|nameconflict|already exists/i)) {
        // Don't rethrow; fall through to re-fetch
      }
    }
    if (!section?.id) {
      const refetch = await graphFetch("GET", `${baseUrl}/notebooks/${notebookId}/sections`);
      section = (refetch?.value || []).find(s => s.displayName === sectionName);
    }
  }
  if (!section?.id) throw new Error("Could not find or create OneNote section: " + sectionName);

  // Metadata badge header (matches SettyPMS style so pages look consistent).
  // Project number / category escaped — handles `&`, `<`, `>` in unusual project names.
  const safeProjNum = escapeOneNoteTextAddin(project.projectNumber || "");
  const safeCategory = escapeOneNoteTextAddin(category || "");
  const badge = [
    project.projectNumber && `<span style="background:#003865;color:#fff;padding:2px 8px;border-radius:3px;font-size:11px;margin-right:6px">${safeProjNum}</span>`,
    category              && `<span style="background:#e8edf2;color:#003865;padding:2px 8px;border-radius:3px;font-size:11px">${safeCategory}</span>`,
  ].filter(Boolean).join("");
  const header = `<div style="border-bottom:2px solid #003865;padding-bottom:8px;margin-bottom:16px;font-family:sans-serif">${badge}</div>`;
  const safeTitle = escapeOneNoteTextAddin(title);
  const pageHtml = `<!DOCTYPE html><html><head><title>${safeTitle}</title><meta name="created" content="${dateStr || new Date().toISOString()}" /></head><body>${header}${buildAddinMeetingPageHtml(title, category, dateStr, emailParticipants, body)}</body></html>`;

  // POST page with retry on 429/503 — Graph throttles OneNote aggressively
  // and the previous code would just fail on the first throttle response,
  // which manifested as random "OneNote 429" errors during heavy save bursts.
  const token = await getToken();
  const url = `https://graph.microsoft.com/v1.0/${baseUrl}/sections/${section.id}/pages`;
  const maxAttempts = 3;
  let lastErr;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const res = await fetch(url, {
      method: "POST",
      headers: { "Authorization": "Bearer " + token, "Content-Type": "text/html" },
      body: pageHtml,
    });
    if (res.ok) {
      const page = await res.json();
      return { id: page.id, webUrl: page.links?.oneNoteWebUrl?.href || page.webUrl || "" };
    }
    if (res.status === 429 || res.status === 503) {
      const retryAfter = parseInt(res.headers.get("Retry-After") || "0", 10);
      const wait = retryAfter > 0 ? retryAfter * 1000 : Math.min(15000, 1000 * Math.pow(2, attempt));
      console.warn("[OneNote-addin] " + res.status + " — retrying in " + wait + "ms (attempt " + attempt + "/" + maxAttempts + ")");
      await new Promise(r => setTimeout(r, wait));
      lastErr = new Error("OneNote throttled (" + res.status + ")");
      continue;
    }
    const errText = await res.text().catch(() => "");
    throw new Error("OneNote " + res.status + ": " + errText.slice(0, 200));
  }
  throw lastErr || new Error("OneNote page creation failed after " + maxAttempts + " attempts");
}

async function doSaveNote() {
  if (!selectedProject) { setStatus("noteStatus", "error", "No project selected."); return; }
  if (saveInFlight) { setStatus("noteStatus", "info", "⏳ Another save is in progress; please wait."); return; }
  const category = document.getElementById("noteCategory").value;
  const body = document.getElementById("noteBody").value.trim();
  if (!body) { setStatus("noteStatus", "error", "Note body is empty."); return; }

  // Disable immediately so a slow OneNote round-trip can't trigger a double-save.
  const saveNoteBtn = document.getElementById("saveNoteBtn");
  if (saveNoteBtn) saveNoteBtn.disabled = true;
  saveInFlight = true;

  // Create a OneNote page for every logged note when a notebook is linked
  let oneNoteUrl = "";
  let oneNoteErr = "";
  const notebookId = selectedProject.teamsOneNoteNotebookId || selectedProject.oneNoteNotebookId || "";
  if (!notebookId) {
    oneNoteErr = "No OneNote notebook linked to this project — create one in the PMS first.";
  } else {
    setStatus("noteStatus", "info", "⏳ Creating OneNote page…");
    try {
        // Use the subject element text as fallback — already resolved even in compose mode
        const subjectEl = document.getElementById("emailSubject").textContent;
        const resolvedSubject = (subjectEl && subjectEl !== "(Loading…)") ? subjectEl : null;
        const title = (typeof emailItem?.subject === "string" ? emailItem.subject : resolvedSubject)
          || body.split("\n")[0].slice(0, 80) || category;

        // In compose mode emailItem.start is an async Time object — fall back to now
        const apptStart = emailItem?.start && !emailItem.start?.getAsync ? emailItem.start : null;
        const dateStr = currentItemKind === "appointment"
          ? new Date(apptStart || Date.now()).toISOString()
          : new Date(emailItem?.dateTimeCreated || Date.now()).toISOString();

        const page = await createAddinOneNotePage(selectedProject, title, body, category, dateStr);
        oneNoteUrl = page.webUrl || "";
      } catch (e) {
        oneNoteErr = e.message;
      }
    }

  setStatus("noteStatus", "info", "⏳ Saving…");
  try {
    const note = {
      id: uid(), body, category, actionItem: false,
      author: msalAccount?.name || msalAccount?.username || "Unknown",
      createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
      importedFromEmail: true, links: [],
      // sourceItemId — matches for the person who saved the note (mailbox-specific).
      // sourceCalendarUId — matches for ALL attendees of the same meeting (shared iCal standard ID).
      ...(emailItem?.itemId ? { sourceItemId: emailItem.itemId } : {}),
      ...(getCurrentSharedMessageId() ? { sourceMessageId: getCurrentSharedMessageId() } : {}),
      ...(currentItemICalUId ? { sourceCalendarUId: currentItemICalUId } : {}),
      ...(oneNoteUrl ? { oneNoteUrl } : {}),
    };
    await applyLocalChangeAndSave(selectedProject.id, fresh => ({
      ...fresh,
      notes: [...(fresh.notes || []), note],
    }));
    // Persist the appointment → project mapping so it auto-restores on next open.
    setSelectedProject(selectedProject, true);
    const linkEl = document.getElementById("noteOneNoteLink");
    if (oneNoteUrl) {
      setStatus("noteStatus", "success", "✓ Note saved · OneNote page created");
      if (linkEl) linkEl.innerHTML = `<a href="${oneNoteUrl}" target="_blank" style="font-size:12px">📓 Open in OneNote</a>`;
    } else if (oneNoteErr) {
      setStatus("noteStatus", "error", "Note saved, but OneNote failed: " + oneNoteErr);
      if (linkEl) linkEl.innerHTML = "";
    } else {
      setStatus("noteStatus", "success", "✓ Note saved");
      if (linkEl) linkEl.innerHTML = "";
    }
    document.getElementById("noteBody").value = "";
    // saveNoteBtn stays disabled — note is saved, re-clicking would double-create the OneNote page.
    refreshOneNoteLinkBanner();
    refreshCalendarStatus();
    // After a successful save, return the user to the main view. The OneNote
    // link banner is now visible there (refreshOneNoteLinkBanner just ran),
    // showing the linked project + 📓 link. Without this nav, the user is
    // stuck on the note-edit view and has to manually click back to see the
    // result of their save.
    if (oneNoteUrl || !oneNoteErr) {
      // Brief delay so the user can see the success status flash before the
      // view changes — feels like confirmation, not abrupt.
      setTimeout(() => showView("mainView"), 700);
    }
  } catch (e) {
    setStatus("noteStatus", "error", "✗ " + e.message);
    // Re-enable the button so the user can retry after fixing the error.
    if (saveNoteBtn) saveNoteBtn.disabled = false;
  } finally {
    saveInFlight = false;
  }
}

function prefillActionItem() {
  const body = document.getElementById("actionItemBody");
  const ownerSelect = document.getElementById("actionItemOwner");
  const dueDate = document.getElementById("actionItemDueDate");
  const teamMembers = getProjectTeamMembers(selectedProject);
  if (body) body.value = (emailItem?.subject || "").trim();
  if (ownerSelect) {
    refreshActionItemOwnerOptions();
    const defaultOwner = [msalAccount?.name, msalAccount?.username, emailFrom]
      .map(v => (v || "").trim())
      .find(v => v && teamMembers.includes(v)) || "";
    ownerSelect.value = defaultOwner;
  }
  if (dueDate) dueDate.value = addBizDays(new Date(), 5);
  setStatus("actionItemStatus", "", "");
}

async function doSaveActionItem() {
  if (!selectedProject) { setStatus("actionItemStatus", "error", "No project selected."); return; }
  if (saveInFlight) { setStatus("actionItemStatus", "info", "⏳ Another save is in progress; please wait."); return; }
  const teamMembers = getProjectTeamMembers(selectedProject);
  const body = document.getElementById("actionItemBody").value.trim();
  const owner = document.getElementById("actionItemOwner").value.trim();
  const dueDate = document.getElementById("actionItemDueDate").value;
  if (!teamMembers.length) {
    setStatus("actionItemStatus", "error", "No project team members found. Add a team in PMS first.");
    return;
  }
  if (!body) { setStatus("actionItemStatus", "error", "Action item is required."); return; }
  if (!owner) { setStatus("actionItemStatus", "error", "Owner is required."); return; }
  if (!teamMembers.includes(owner)) { setStatus("actionItemStatus", "error", "Please select a valid team member."); return; }
  if (!dueDate) { setStatus("actionItemStatus", "error", "Due date is required."); return; }

  const saveBtn = document.getElementById("saveActionItemBtn");
  if (saveBtn) saveBtn.disabled = true;
  saveInFlight = true;
  setStatus("actionItemStatus", "info", "⏳ Saving…");

  try {
    const createdAt = new Date().toISOString();
    const actionNoteBody = `${body}\n\nOwner: ${owner}\nDue: ${dueDate}`;
    const note = {
      id: uid(),
      body: actionNoteBody,
      category: "Action Item",
      actionItem: true,
      owner,
      dueDate,
      status: "Open",
      author: msalAccount?.name || msalAccount?.username || "Unknown",
      createdAt,
      updatedAt: createdAt,
      importedFromEmail: true,
      links: [],
      ...(emailItem?.itemId ? { sourceItemId: emailItem.itemId } : {}),
      ...(getCurrentSharedMessageId() ? { sourceMessageId: getCurrentSharedMessageId() } : {}),
      ...(currentItemICalUId ? { sourceCalendarUId: currentItemICalUId } : {}),
    };
    await applyLocalChangeAndSave(selectedProject.id, fresh => ({
      ...fresh,
      notes: [...(fresh.notes || []), note],
    }));
    setSelectedProject(selectedProject, true);
    setStatus("actionItemStatus", "success", "✓ Action item saved");
    document.getElementById("actionItemBody").value = "";
    document.getElementById("actionItemOwner").value = "";
    document.getElementById("actionItemDueDate").value = "";
  } catch (e) {
    setStatus("actionItemStatus", "error", "✗ " + e.message);
    if (saveBtn) saveBtn.disabled = false;
  } finally {
    saveInFlight = false;
  }
}
// ─── SHARED: file email+attachments into a project subfolder ─────────────────
async function uploadEmailUnderFolder(driveId, token, projFolderName, subfolder, recordFolderName, metadata = null) {
  const subPath    = await ensureSpFolder(driveId, token, projFolderName, subfolder);
  const recordPath = await ensureSpFolder(driveId, token, subPath, recordFolderName);
  await uploadEmailAndAttachments(driveId, token, recordPath);
  if (metadata) await writeSpMetadataSidecar(driveId, token, recordPath, metadata);
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
  if (saveInFlight) { setStatus("rfiStatus", "info", "⏳ Another save is in progress; please wait."); return; }
  const title = document.getElementById("rfiTitle").value.trim();
  if (!title) { setStatus("rfiStatus", "error", "Title is required."); return; }
  saveInFlight = true;
  setStatus("rfiStatus", "info", "⏳ Saving…");
  try {
    // Re-fetch fresh project data so the RFI number reflects what's actually in
    // the cloud — not the add-in's possibly-stale cache. Prevents two users from
    // independently picking the same RFI number when both edit at the same time.
    // Reads from pms_projects (V2) — pms_data is frozen at migration time and
    // doesn't include projects created post-migration.
    let freshProject = selectedProject;
    try {
      const res = await fetch(
        SUPABASE_URL + "/rest/v1/pms_projects?id=eq." + encodeURIComponent(selectedProject.id) + "&select=project",
        { headers: SB_HEADERS }
      );
      if (res.ok) {
        const rows = await res.json();
        if (rows?.[0]?.project) freshProject = rows[0].project;
      }
    } catch { /* fall back to cache; logged in applyLocalChangeAndSave too */ }

    const existingRfis = freshProject.rfis || [];
    const nextNum = "RFI-" + String(existingRfis.length + 1).padStart(3, "0");
    const received = new Date();
    let spFolderUrl = "";
    if (freshProject.projectFolderUrl) {
      try {
        const token = await getToken();
        const { driveId } = await resolveSpIds();
        const projFolderName = decodeURIComponent(freshProject.projectFolderUrl.split("/").pop());
        const safeName = (nextNum + " " + title).replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
        spFolderUrl = await uploadEmailUnderFolder(driveId, token, projFolderName, "RFIs", safeName, buildAddinMetadata(freshProject, "rfi"));
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
      assignedTo: [], spFolderUrl, links: [],
      createdAt: new Date().toISOString(),
    };
    await applyLocalChangeAndSave(selectedProject.id, fresh => ({
      ...fresh,
      rfis: [...(fresh.rfis || []), rfi],
    }));
    setStatus("rfiStatus", "success", "✓ " + nextNum + " logged" + (spFolderUrl ? " · filed to SharePoint" : ""));
    document.getElementById("rfiTitle").value = "";
    document.getElementById("rfiNotes").value = "";
  } catch (e) {
    setStatus("rfiStatus", "error", "✗ " + e.message);
  } finally {
    saveInFlight = false;
  }
}
async function doFileToExistingRfi() {
  if (!selectedProject) { setStatus("rfiExistingStatus", "error", "Select a project first."); return; }
  const rfiId = document.getElementById("rfiExistingSelect").value;
  if (!rfiId) { setStatus("rfiExistingStatus", "error", "Select an RFI."); return; }
  if (saveInFlight) { setStatus("rfiExistingStatus", "info", "⏳ Another save is in progress; please wait."); return; }
  const rfi = (selectedProject.rfis || []).find(r => r.id === rfiId);
  if (!rfi) { setStatus("rfiExistingStatus", "error", "RFI not found."); return; }
  saveInFlight = true;
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
      await applyLocalChangeAndSave(selectedProject.id, fresh => ({
        ...fresh,
        rfis: (fresh.rfis || []).map(r => r.id === rfi.id ? { ...r, spFolderUrl: newUrl } : r),
      }));
    }
    const attMsg = attCount ? " + " + attCount + " attachment" + (attCount > 1 ? "s" : "") : "";
    setStatus("rfiExistingStatus", "success", "✓ Filed to " + rfi.number + attMsg);
  } catch (e) {
    setStatus("rfiExistingStatus", "error", "✗ " + e.message);
  } finally {
    saveInFlight = false;
  }
}
// ─── LOG SUBMITTAL ────────────────────────────────────────────────────────────
function prefillSub() {
  setSubMode("new");
  renderSubPicker();
}
async function doSaveSub() {
  if (!selectedProject) { setStatus("subStatus", "error", "No project selected."); return; }
  if (saveInFlight) { setStatus("subStatus", "info", "⏳ Another save is in progress; please wait."); return; }
  const desc = document.getElementById("subDesc").value.trim();
  if (!desc) { setStatus("subStatus", "error", "Description is required."); return; }
  saveInFlight = true;
  setStatus("subStatus", "info", "⏳ Saving…");
  try {
    // Re-fetch so submittal numbering reflects current cloud state.
    // Reads from pms_projects (V2) — pms_data is frozen at migration time.
    let freshProject = selectedProject;
    try {
      const res = await fetch(
        SUPABASE_URL + "/rest/v1/pms_projects?id=eq." + encodeURIComponent(selectedProject.id) + "&select=project",
        { headers: SB_HEADERS }
      );
      if (res.ok) {
        const rows = await res.json();
        if (rows?.[0]?.project) freshProject = rows[0].project;
      }
    } catch { /* fall back to cache */ }

    const existing = freshProject.submittals || [];
    const nextNum = "SUB-" + String(existing.length + 1).padStart(3, "0");
    const received = new Date();
    let spFolderUrl = "";
    if (freshProject.projectFolderUrl) {
      try {
        const token = await getToken();
        const { driveId } = await resolveSpIds();
        const projFolderName = decodeURIComponent(freshProject.projectFolderUrl.split("/").pop());
        const safeName = (nextNum + " " + desc).replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
        spFolderUrl = await uploadEmailUnderFolder(driveId, token, projFolderName, "Submittals", safeName, buildAddinMetadata(freshProject, "submittal"));
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
      assignedTo: [], spFolderUrl, links: [],
      createdAt: new Date().toISOString(),
    };
    await applyLocalChangeAndSave(selectedProject.id, fresh => ({
      ...fresh,
      submittals: [...(fresh.submittals || []), sub],
    }));
    setStatus("subStatus", "success", "✓ " + nextNum + " logged" + (spFolderUrl ? " · filed to SharePoint" : ""));
    document.getElementById("subDesc").value = "";
    document.getElementById("subSpec").value = "";
    document.getElementById("subNotes").value = "";
  } catch (e) {
    setStatus("subStatus", "error", "✗ " + e.message);
  } finally {
    saveInFlight = false;
  }
}
async function doFileToExistingSub() {
  if (!selectedProject) { setStatus("subExistingStatus", "error", "Select a project first."); return; }
  if (saveInFlight) { setStatus("subExistingStatus", "info", "⏳ Another save is in progress; please wait."); return; }
  const subId = document.getElementById("subExistingSelect").value;
  if (!subId) { setStatus("subExistingStatus", "error", "Select a submittal."); return; }
  const sub = (selectedProject.submittals || []).find(s => s.id === subId);
  if (!sub) { setStatus("subExistingStatus", "error", "Submittal not found."); return; }
  saveInFlight = true;
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
      await applyLocalChangeAndSave(selectedProject.id, fresh => ({
        ...fresh,
        submittals: (fresh.submittals || []).map(s => s.id === sub.id ? { ...s, spFolderUrl: newUrl } : s),
      }));
    }
    const attMsg = attCount ? " + " + attCount + " attachment" + (attCount > 1 ? "s" : "") : "";
    setStatus("subExistingStatus", "success", "✓ Filed to " + sub.number + attMsg);
  } catch (e) {
    setStatus("subExistingStatus", "error", "✗ " + e.message);
  } finally {
    saveInFlight = false;
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
// Strip Outlook reply chains so we don't surface dates from older messages
// quoted inside a forward. First marker past the first ~50 chars wins —
// the threshold prevents truncating an email whose first line happens to
// start with "On Monday, …".
function trimToCurrentMessage(text) {
  if (!text) return "";
  const markers = [
    /\bFrom:\s+\S[\s\S]{0,200}?\bSent:/i,    // Outlook header block
    /\bOn\s+[\s\S]{1,120}?wrote:/i,           // "On Tue, May 6 ... wrote:"
    /-{3,}\s*Original Message\s*-{3,}/i,
    /_{20,}/,
  ];
  let cutoff = text.length;
  for (const re of markers) {
    const m = re.exec(text);
    if (m && m.index > 50 && m.index < cutoff) cutoff = m.index;
  }
  return text.slice(0, cutoff);
}

function extractDueDates(rawText, emailReceivedDate) {
  const text = trimToCurrentMessage(rawText);
  const results = [];
  const seen    = new Set();
  const refDate = emailReceivedDate ? new Date(emailReceivedDate) : new Date();
  const MONTHS_LONG  = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const MONTHS_SHORT = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const DAYS         = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
  function isValidDateParts(year, month1, day) {
    const d = new Date(year, month1 - 1, day);
    return d.getFullYear() === year && d.getMonth() === (month1 - 1) && d.getDate() === day;
  }
  function toISO(year, month1, day) {
    const y = year < 100 ? 2000 + year : year;
    if (!isValidDateParts(y, month1, day)) return "";
    return y + "-" + String(month1).padStart(2,"0") + "-" + String(day).padStart(2,"0");
  }
  function resolveYearlessMonthDay(month1, day) {
    const refYear = refDate.getFullYear();
    const candidates = [refYear, refYear + 1].map(y => toISO(y, month1, day)).filter(Boolean);
    if (!candidates.length) return "";
    const refMid = new Date(refDate.getFullYear(), refDate.getMonth(), refDate.getDate(), 12, 0, 0, 0);
    const best = candidates.find(iso => {
      const d = new Date(iso + "T12:00:00");
      return d >= refMid;
    });
    return best || candidates[0];
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
    if (mo >= 1 && mo <= 12 && dy >= 1 && dy <= 31) {
      const iso = toISO(yr, mo, dy);
      if (iso) addResult(iso, m[0], m.index);
    }
  }
  // Slash notation without year: "4/22"
  const p4b = /\b(\d{1,2})\/(\d{1,2})(?!\/)\b/g;
  while ((m = p4b.exec(text))) {
    const mo = +m[1], dy = +m[2];
    if (mo >= 1 && mo <= 12 && dy >= 1 && dy <= 31) {
      const iso = resolveYearlessMonthDay(mo, dy);
      if (iso) addResult(iso, m[0] + "  (" + iso + ")", m.index);
    }
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
  // Weekday + ordinal day: "Tuesday the 29th" / "Tue the 29th" / "Tuesday 29th"
  const p7 = /\b(?:Mon(?:day)?|Tue(?:s(?:day)?)?|Wed(?:nesday)?|Thu(?:rs(?:day)?)?|Fri(?:day)?|Sat(?:urday)?|Sun(?:day)?)\s+(?:the\s+)?(\d{1,2})(?:st|nd|rd|th)\b/gi;
  while ((m = p7.exec(text))) {
    const day = +m[1];
    if (day < 1 || day > 31) continue;
    let found = "";
    for (let i = 0; i <= 12; i++) {
      const y = refDate.getFullYear() + Math.floor((refDate.getMonth() + i) / 12);
      const mo = ((refDate.getMonth() + i) % 12) + 1;
      const iso = toISO(y, mo, day);
      if (!iso) continue;
      const d = new Date(iso + "T12:00:00");
      const refMid = new Date(refDate.getFullYear(), refDate.getMonth(), refDate.getDate(), 12, 0, 0, 0);
      if (d >= refMid) { found = iso; break; }
    }
    if (found) addResult(found, m[0] + "  (" + found + ")", m.index);
  }
  // Bare weekday: "Friday", "by Tue", "ready Wed" — resolve to next future occurrence.
  // Skips matches already covered by p6 (next/this) or p7 (weekday + ordinal day),
  // and skips trailing "wrote:" footers (still possible if trimToCurrentMessage missed one).
  const p8 = /\b(Mon(?:day)?|Tue(?:s(?:day)?)?|Wed(?:nesday)?|Thu(?:rs(?:day)?)?|Fri(?:day)?|Sat(?:urday)?|Sun(?:day)?)\b/gi;
  while ((m = p8.exec(text))) {
    const before = text.slice(Math.max(0, m.index - 10), m.index).toLowerCase();
    if (/\b(next|this|last|on)\s+$/.test(before)) continue;
    const after = text.slice(m.index + m[0].length, m.index + m[0].length + 60);
    if (/^\s+(?:the\s+)?\d{1,2}(?:st|nd|rd|th)\b/.test(after)) continue; // "Tuesday the 29th"
    if (/wrote:/i.test(after) && /^,?\s+\w{3,}\s+\d{1,2}/.test(after)) continue; // "Mon, May 6, 2025 ... wrote:"
    const word   = m[1];
    const target = DAYS.findIndex(d => d.toLowerCase().startsWith(word.toLowerCase().slice(0, 3)));
    if (target < 0) continue;
    const d = new Date(refDate);
    let delta = target - d.getDay();
    if (delta <= 0) delta += 7;
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
  if (saveInFlight) { setStatus("milestoneStatus", "info", "⏳ Another save is in progress; please wait."); return; }
  saveInFlight = true;
  setStatus("milestoneStatus", "info", "⏳ Saving…");
  try {
    // Build the milestone first; sync calendar; then save via V2 path.
    // Previously this function PATCHed pms_data.projects directly, which
    // post-migration is a dead-end table — every milestone created here was
    // silently lost. Now uses applyLocalChangeAndSave like every other save.
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
      ...(emailItem?.itemId ? { sourceItemId: emailItem.itemId } : {}),
      ...(getCurrentSharedMessageId() ? { sourceMessageId: getCurrentSharedMessageId() } : {}),
    };
    setStatus("milestoneStatus", "info", "⏳ Syncing to calendar…");
    const calResult = await createMilestoneCalendarEvent(milestone, selectedProject);
    if (calResult.success) milestone.calendarEventId = calResult.eventId;

    setStatus("milestoneStatus", "info", "⏳ Saving to project…");
    await applyLocalChangeAndSave(selectedProject.id, fresh => ({
      ...fresh,
      milestones: [...(fresh.milestones || []), milestone],
    }));

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
  } finally {
    saveInFlight = false;
  }
}
// ─── PEOPLE PICKER ────────────────────────────────────────────────────────────
// Tracks emails saved as contacts during the current pane session — used to
// mark them with a ✓ when the user returns to the participant list after
// saving, so they can immediately move on to the next person without losing
// their place. Cleared per-email in loadItemContext.
const _sessionSavedContactEmails = new Set();
function showPeopleView() {
  const list = document.getElementById("participantList");
  if (!emailParticipants.length) {
    list.innerHTML = '<p style="font-size:12px;color:var(--text-soft);">No participants found.</p>';
  } else {
    const labelColor = { From: "#c50f1f", To: "#0f6cbd", CC: "#0e6d5c", Required: "#0f6cbd", Optional: "#616161", Organizer: "#c50f1f" };
    const labelBg    = { From: "#fde7e9", To: "#eaf3fb", CC: "#e0f5f0", Required: "#eaf3fb", Optional: "#f3f2f1", Organizer: "#fde7e9" };
    list.innerHTML = emailParticipants.map((p, i) => {
      const emailKey = (p.emailAddress || "").toLowerCase();
      const alreadyAdded = emailKey && _sessionSavedContactEmails.has(emailKey);
      return `
      <div class="participant-row${alreadyAdded ? ' added' : ''}" data-idx="${i}">
        <div style="flex:1;min-width:0;">
          <div style="font-size:13px;font-weight:600;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
            ${escHtml(p.displayName || p.emailAddress)}
          </div>
          <div style="font-size:11px;color:var(--text-soft);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
            ${escHtml(p.emailAddress || "")}
          </div>
        </div>
        ${alreadyAdded ? '<span class="pill added">✓ Added</span>' : ''}
        <span class="pill" style="background:${labelBg[p.label]||'var(--surface-2)'};color:${labelColor[p.label]||'var(--text-soft)'};">
          ${escHtml(p.label || "")}
        </span>
      </div>`;
    }).join("");
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
  applyPipelineUiRules();
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
function normalizeEmail(v) {
  return (v || "").trim().toLowerCase();
}
function contactExistsInList(list, email, name) {
  const targetEmail = normalizeEmail(email);
  const targetName = (name || "").trim().toLowerCase();
  return (list || []).some(c => {
    const cEmail = normalizeEmail(c.email);
    const cName = (c.name || "").trim().toLowerCase();
    if (targetEmail && cEmail && cEmail === targetEmail) return true;
    return !targetEmail && !!targetName && cName === targetName;
  });
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
  if (saveInFlight) { setStatus("contactStatus", "info", "⏳ Another save is in progress; please wait."); return; }
  saveInFlight = true;
  setStatus("contactStatus", "info", "⏳ Saving…");
  try {
    if (saveTo === "client") {
      // V2: write to pms_clients (per-client rows). Previously this PATCHed
      // pms_data.clients (legacy singleton blob), which post-migration is
      // never re-read by PMS — so the contact silently disappeared. Now we
      // upsert the client row directly.
      const targetCompany = (company || name).trim();
      const contact = { id: uid(), name, title, email, phone, role: type };
      // Find existing client by exact (case-insensitive) name match
      const queryUrl = SUPABASE_URL + "/rest/v1/pms_clients?select=id,client,version";
      const all = await fetch(queryUrl, { headers: SB_HEADERS });
      if (!all.ok) throw new Error("pms_clients GET HTTP " + all.status);
      const rows = await all.json();
      const existing = (rows || []).find(r => r.client?.name && r.client.name.trim().toLowerCase() === targetCompany.toLowerCase());
      if (existing) {
        const ec = existing.client;
        ec.contacts = ec.contacts || [];
        if (contactExistsInList(ec.contacts, email, name)) {
          setStatus("contactStatus", "info", "Contact already exists for this client. No duplicate was added.");
          return;
        }
        ec.contacts = [...ec.contacts, contact];
        // Optimistic-version PATCH
        const patchUrl = SUPABASE_URL + "/rest/v1/pms_clients?id=eq." + encodeURIComponent(existing.id) +
                         "&version=eq." + existing.version;
        const res = await fetch(patchUrl, {
          method: "PATCH",
          headers: { ...SB_HEADERS, "Prefer": "return=representation" },
          body: JSON.stringify({ client: ec, version: existing.version + 1, updated_at: new Date().toISOString() }),
        });
        if (!res.ok) throw new Error("pms_clients PATCH HTTP " + res.status);
        const result = await res.json();
        if (!result || result.length === 0) throw new Error("Client modified by someone else. Retry from Outlook.");
      } else {
        // New client — INSERT
        const newClient = { id: uid(), name: targetCompany, type, contacts: [contact], address: "" };
        const res = await fetch(SUPABASE_URL + "/rest/v1/pms_clients", {
          method: "POST",
          headers: SB_HEADERS,
          body: JSON.stringify({ id: newClient.id, client: newClient, version: 1, updated_at: new Date().toISOString() }),
        });
        if (!res.ok) throw new Error("pms_clients POST HTTP " + res.status);
        // Update in-memory cache so subsequent UI references see the new client
        allClients = [...(allClients || []), newClient];
        renderCompanySuggestions();
      }
    } else {
      if (!selectedProject) { setStatus("contactStatus", "error", "Select a project first."); return; }
      const poc = { id: uid(), name, title, email, phone, role: type };
      // V2: per-project save with version check via applyLocalChangeAndSave.
      // Already routes through pms_projects with optimistic concurrency.
      await applyLocalChangeAndSave(selectedProject.id, fresh => {
        const projectContacts = { ...(fresh.projectContacts || {}) };
        const pm = projectContacts.pm || [];
        if (contactExistsInList(pm, email, name)) {
          // Throw to abort the save and tell user it's a no-op
          throw new Error("__DUP__");
        }
        projectContacts.pm = [...pm, poc];
        return { ...fresh, projectContacts };
      });
    }
    // Mark this email as added-this-session so the participant list shows ✓
    // when we return there. Then bounce straight back to the list — the user
    // is almost always working through several participants in sequence.
    const savedEmailKey = (email || "").toLowerCase();
    if (savedEmailKey) _sessionSavedContactEmails.add(savedEmailKey);
    const destLabel = saveTo === "client"
      ? (company || name)
      : ((selectedProject?.projectNumber ? selectedProject.projectNumber + " — " : "") + (selectedProject?.name || "project POC"));
    setStatus("actionStatus", "success", "✓ Saved " + (name || email) + " to " + destLabel + ".");
    setStatus("contactStatus", "", "");
    showPeopleView();
    return;
  } catch (e) {
    if (e.message === "__DUP__") {
      // Treat dup as a benign success for the "add another" flow — mark them
      // as ✓ and return to the list rather than stranding the user on the form.
      const savedEmailKey = (email || "").toLowerCase();
      if (savedEmailKey) _sessionSavedContactEmails.add(savedEmailKey);
      setStatus("actionStatus", "info", "Already in this project's POC list — no duplicate added.");
      setStatus("contactStatus", "", "");
      showPeopleView();
      return;
    }
    setStatus("contactStatus", "error", "✗ " + e.message);
  } finally {
    saveInFlight = false;
  }
}
// ─── PMS METADATA HELPERS (mirrors SettyPMS.html metadata schema) ─────────────
function buildAddinMetadata(project, docType) {
  return {
    projectNumber: project.projectNumber || "",
    projectName:   project.name          || "",
    client:        project.prime || project.clientName || "",
    docType,
    date:          new Date().toISOString().slice(0, 10),
    createdBy:     msalAccount?.name || msalAccount?.username || "Unknown",
    source:        "outlook-addin",
    phase:         "",
    tags:          [],
    _schema:       "pms-v1",
  };
}
// Non-fatal — sidecar failure never blocks the primary save
async function writeSpMetadataSidecar(driveId, token, folderPath, metadata) {
  try {
    await fetch("https://graph.microsoft.com/v1.0/drives/" + driveId + "/root:/" + encodeDrivePath(folderPath + "/_pms-metadata.json") + ":/content", {
      method: "PUT",
      headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify(metadata, null, 2),
    });
  } catch (e) { console.warn("PMS metadata sidecar:", e.message); }
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
  ["signInView","mainView","noteView","actionItemView","rfiView","subView","datesView","peopleView","contactView"].forEach(v => {
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
