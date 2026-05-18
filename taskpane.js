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
// Scopes requested ON-DEMAND only — kept out of the default sign-in flow so
// users don't see a longer consent dialog at first run. First use of the
// associated feature triggers a one-time per-user consent popup; afterwards
// the token is cached silently like any other Graph token.
const CHANNEL_MESSAGE_SCOPES = ["ChannelMessage.Send"];
const TEAMS_TEAM_ID   = "a4c48361-7991-43db-af83-4c854918a760";
const TEAMS_TENANT_ID = "f374c024-71c2-48b6-8420-076fff97327c";
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
const TEAMS_SENT_MAP_STORAGE_KEY = "settyPms:teamsChannelSentMap";
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
    applyComposeModeUiGuard();
    // When the task pane is pinned, Office swaps mailbox.item silently as the user
    // clicks different emails. ItemChanged fires each time — reload the pane context.
    // The compose-mode guard re-runs here too because Reply/Forward swaps the item
    // from a Read-mode message to an Edit-mode draft.
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      () => { showView("mainView"); applyComposeModeUiGuard(); loadItemContext(); }
    );
    loadItemContext();
  } catch (e) {
    // Show something rather than a black screen if init fails
    showView("signInView");
    setStatus("signInStatus", "error", "Startup error: " + e.message);
  }
});
// In Compose mode (the user hit Reply / Forward / New) the email isn't sent
// yet, so the file-to-SharePoint / log-as-note / log-as-RFI / etc. flows
// don't apply. Hide everything except the project picker (still useful as
// context) and the Quick Text + Templates section (the only Compose-relevant
// feature). Detection: in Compose mode `mailbox.item.to` is a Recipients
// object with setAsync(); in Read mode it's an array of EmailAddressDetails.
function isComposeMode() {
  try {
    const item = Office.context.mailbox && Office.context.mailbox.item;
    return !!(item && item.to && typeof item.to.setAsync === "function");
  } catch { return false; }
}
// Hide the "Log as RFI" + "Log as Submittal" buttons on the main view when the
// selected project's status is not "In Construction Administration" — RFIs and
// submittals are only a CA-phase activity, so showing the entry points during
// other phases (Proposal, In Progress, Completed, etc.) just clutters the UI
// and invites mis-filing.
//
// Note: this only hides the buttons that LAUNCH those flows. The flows
// themselves (rfiView, subView) remain in the DOM and continue to work if
// reached — guarding the entry points is sufficient for the user-facing goal.
function applyConstructionAdminGuard() {
  const inCA = selectedProject?.status === "In Construction Administration";
  // If no project is selected, show the buttons by default (so users see the
  // moreActions row's full menu after picking a project). Without this, an
  // empty selection would also hide them, which is more aggressive than asked.
  const shouldHide = !!selectedProject && !inCA;
  const ids = ["logRfiBtn", "logSubBtn"];
  for (const id of ids) {
    const el = document.getElementById(id);
    if (el) el.style.display = shouldHide ? "none" : "";
  }
}

function applyComposeModeUiGuard() {
  const compose = isComposeMode();
  document.body.classList.toggle("compose-mode", compose);
  // Hide in Compose: there's no sent email yet to file or log.
  const hiddenInCompose = [
    "saveSpBtn", "saveRecordBtn", "saveConfirmation",
    "logNoteBtn", "sendToTeamsBtn", "newActionItemBtn",
    "moreActions", "oneNoteLinkBanner",
    "dateSuggestionBlock",
  ];
  // Hide in Read: templates only make sense while composing a reply.
  const hiddenInRead = ["quickTemplatesSection"];
  for (const id of hiddenInCompose) {
    const el = document.getElementById(id);
    if (el) el.style.display = compose ? "none" : "";
  }
  for (const id of hiddenInRead) {
    const el = document.getElementById(id);
    if (el) el.style.display = compose ? "" : "none";
  }
  // Save-row + caption sit in divs without ids — find via class.
  const saveRow = document.querySelector(".save-row");
  const saveRowCaption = document.querySelector(".save-row-caption");
  if (saveRow)        saveRow.style.display        = compose ? "none" : "";
  if (saveRowCaption) saveRowCaption.style.display = compose ? "none" : "";
}
function setupEventListeners() {
  document.getElementById("signInBtn").onclick     = doSignIn;
  document.getElementById("signOutBtn").onclick    = doSignOut;
  document.getElementById("saveSpBtn").onclick     = doSaveToSharePoint;
  document.getElementById("saveRecordBtn").onclick = doSaveToProjectRecordOnly;
  // 5-click easter egg on the SETTY PMS logo — reveals the cornerstone card.
  // Counter resets after 3 seconds idle so a curious user has time to discover
  // the pattern but doesn't accidentally trigger it across casual clicks.
  // NOTE: there are TWO `.header-logo` elements (one in signInView, one in
  // mainView), so bind to both via querySelectorAll. Counter is shared across
  // the two so a user who clicks 3x while signed-out and 2x after signing in
  // still gets the reveal.
  let _logoClickCount = 0;
  let _logoClickTimer = null;
  document.querySelectorAll(".header-logo").forEach(logoEl => {
    logoEl.title = "v" + (window.__appVersion || "");
    logoEl.onclick = () => {
      _logoClickCount++;
      clearTimeout(_logoClickTimer);
      if (_logoClickCount >= 5) {
        _logoClickCount = 0;
        const overlay = document.getElementById("creditsOverlay");
        if (overlay) overlay.classList.add("show");
        loadConfetti().then(ok => {
          if (!ok || typeof confetti !== "function") return;
          confetti({ particleCount: 40, spread: 60, origin: { y: 0.5 }, scalar: 0.7, ...(getSeasonalConfettiOpts() || {}) });
        });
        return;
      }
      _logoClickTimer = setTimeout(() => { _logoClickCount = 0; }, 3000);
    };
  });
  const credits = document.getElementById("creditsOverlay");
  if (credits) credits.onclick = () => credits.classList.remove("show");
  // Hint banner link — single entry point for "open project in PMS" so the
  // URL/permissions logic stays in one place (openSelectedProjectInPms).
  const spHintLink = document.getElementById("spFolderHintLink");
  if (spHintLink) spHintLink.onclick = (e) => { e.preventDefault(); openSelectedProjectInPms(); };
  document.getElementById("logNoteBtn").onclick    = () => showView("noteView");
  document.getElementById("sendToTeamsBtn").onclick = sendToTeamsChannel;
  document.getElementById("newActionItemBtn").onclick = () => { prefillActionItem(); showView("actionItemView"); };
  document.getElementById("logRfiBtn").onclick       = () => { prefillRfi(); showView("rfiView"); };
  document.getElementById("logSubBtn").onclick       = () => { prefillSub(); showView("subView"); };
  document.getElementById("noteBack").onclick    = () => showView("mainView");
  document.getElementById("actionItemBack").onclick = () => showView("mainView");
  document.getElementById("rfiBack").onclick     = () => showView("mainView");
  document.getElementById("subBack").onclick     = () => showView("mainView");
  // Response/review views (RFI Log Response, Submittal Log Review)
  const rfiRespBack = document.getElementById("rfiResponseBack");
  if (rfiRespBack) rfiRespBack.onclick = () => showView("mainView");
  const subRevBack = document.getElementById("subReviewBack");
  if (subRevBack) subRevBack.onclick = () => showView("mainView");
  const submitRfiResp = document.getElementById("submitRfiResponseBtn");
  if (submitRfiResp) submitRfiResp.onclick = submitRfiResponse;
  const submitSubRev = document.getElementById("submitSubReviewBtn");
  if (submitSubRev) submitSubRev.onclick = submitSubReview;
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
  // Quick text + templates — clipboard-based, no Outlook compose API needed.
  // Event delegation on the parent: click any .btn-quick → copy its template.
  document.querySelectorAll(".btn-quick").forEach(btn => {
    btn.onclick = () => copyTemplateToClipboard(btn.getAttribute("data-tpl"), btn);
  });
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
  // Compose-mode short-circuit: there's no sent email to read, and the item
  // properties have different shapes (e.g. subject is a Subject object with
  // getAsync, not a string). Skip the entire context load — the templates
  // section doesn't need email metadata.
  if (isComposeMode()) {
    emailItem = null;
    return;
  }
  // Bump the generation. All async work below captures `myGen` at start and
  // bails out before writing module state if the generation has advanced
  // (= user clicked a different email mid-fetch).
  itemContextGeneration++;
  const myGen = itemContextGeneration;
  emailItem = Office.context.mailbox.item;
  refreshTeamsBtn();
  currentConversationId = "";
  currentItemICalUId = "";
  emailParticipants = [];
  // Per-item ✓ "added this session" marks reset when item changes
  _sessionSavedContactEmails.clear();
  // Drop the per-item Graph caches (email body, etc.) so a different item
  // can't accidentally serve cached body HTML from the previous one. Cheap.
  if (typeof clearEmailBodyCache === "function") clearEmailBodyCache();
  // Custom SharePoint folder name is per-email; clear when switching emails so
  // last email's chosen name doesn't accidentally get applied to a new save.
  _customSpFolderName = "";
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
    try { refreshLoggedArtifactChips(); } catch {}
    maybeShowAecQuip();
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
  // Surface RFIs/Submittals here too so they show up in the "also logged as"
  // line of the saved-email indicator. The dedicated chip row above the save
  // buttons (refreshLoggedArtifactChips) gives the prominent affordance the
  // user expects for these specifically.
  const rfiMatch = (project.rfis || []).some(r =>
    r?.sourceItemId === sourceItemId || sourceMessageIds.includes(r?.sourceMessageId)
  );
  const subMatch = (project.submittals || []).some(s =>
    s?.sourceItemId === sourceItemId || sourceMessageIds.includes(s?.sourceMessageId)
  );
  if (rfiMatch) labels.push("RFI");
  if (subMatch) labels.push("Submittal");
  return labels;
}

// Returns RFI/Submittal records whose sourceItemId or sourceMessageId matches
// the currently-open mailbox item. Used by refreshLoggedArtifactChips to
// surface "Logged as RFI-XXX on Date" chips on the main view.
//
// `status` is included so the chip renderer can decide whether to show a
// "Log Response" / "Log Review" call-to-action (visible only while open).
function getLoggedRfiSubArtifacts(project) {
  if (!project || !emailItem?.itemId) return [];
  const sourceItemId = emailItem.itemId;
  const sourceMessageIds = getCurrentMessageIdCandidates();
  const matches = [];
  for (const r of (project.rfis || [])) {
    if (r?.sourceItemId === sourceItemId || (r?.sourceMessageId && sourceMessageIds.includes(r.sourceMessageId))) {
      matches.push({
        kind: "rfi",
        id: r.id,
        number: r.number || "RFI",
        title: r.title || "",
        date: r.createdAt || r.dateReceived || null,
        spFolderUrl: r.spFolderUrl || "",
        status: r.status || "Open",
      });
    }
  }
  for (const s of (project.submittals || [])) {
    if (s?.sourceItemId === sourceItemId || (s?.sourceMessageId && sourceMessageIds.includes(s.sourceMessageId))) {
      matches.push({
        kind: "sub",
        id: s.id,
        number: s.number || "SUB",
        title: s.description || "",
        date: s.createdAt || s.dateReceived || null,
        spFolderUrl: s.spFolderUrl || "",
        status: s.status || "Received",
      });
    }
  }
  return matches;
}

// "Open"-ish RFI statuses where a response action is still meaningful. We
// stop showing the Log Response button once the RFI is Responded/Closed/Void.
const RFI_OPEN_STATUSES = new Set(["Open", "Pending Sub Response"]);
// Same notion for submittals: while still under review (not yet returned).
const SUB_OPEN_STATUSES = new Set(["Received", "In Review", "Pending Sub Response"]);

// Render the "Logged as RFI-XXX / SUB-XXX" chips on the main view. Chips link
// to the SharePoint folder when one is stored. Open RFIs/Submittals also get
// a "Log Response" / "Log Review" button rendered below the chip so the user
// can finish the workflow in one click. Hidden when nothing applies.
function refreshLoggedArtifactChips() {
  const container = document.getElementById("loggedAsArtifactChips");
  if (!container) return;
  if (!selectedProject || !emailItem?.itemId) {
    container.innerHTML = "";
    container.style.display = "none";
    return;
  }
  const artifacts = getLoggedRfiSubArtifacts(selectedProject);
  if (artifacts.length === 0) {
    container.innerHTML = "";
    container.style.display = "none";
    return;
  }
  // Switch container to a column layout so each artifact's chip + (optional)
  // action button stack vertically.
  container.style.display = "flex";
  container.style.flexDirection = "column";
  container.style.gap = "6px";
  // Build per-artifact rows. Each row is chip + optional action button.
  // Action buttons are bound via event delegation below (one listener on the
  // container handles all clicks via data-action attributes).
  const rowsHtml = artifacts.map(a => {
    const dateStr = a.date
      ? new Date(a.date).toLocaleDateString("en-US", { month: "short", day: "numeric" })
      : "earlier";
    const icon = a.kind === "rfi" ? "🔵" : "📋";
    const bg = a.kind === "rfi" ? "#dbeafe" : "#ede9fe";
    const border = a.kind === "rfi" ? "#93c5fd" : "#c4b5fd";
    const color = a.kind === "rfi" ? "#1e3a8a" : "#5b21b6";
    const statusLabel = a.status && a.status !== "Open" && a.status !== "Received"
      ? ` · ${a.status}`
      : "";
    const label = `${icon} Logged as ${a.number}${a.title ? " — " + a.title.slice(0, 40) : ""} on ${dateStr}${statusLabel}`;
    const safeLabel = label.replace(/</g, "&lt;").replace(/>/g, "&gt;");
    let chipHtml;
    if (a.spFolderUrl) {
      const safeUrl = a.spFolderUrl.replace(/"/g, "&quot;");
      chipHtml = `<a href="${safeUrl}" target="_blank" rel="noreferrer" style="display:inline-flex;align-items:center;padding:4px 10px;border-radius:12px;background:${bg};border:1px solid ${border};color:${color};font-size:11px;font-weight:600;text-decoration:none;white-space:nowrap;">${safeLabel}</a>`;
    } else {
      chipHtml = `<span style="display:inline-flex;align-items:center;padding:4px 10px;border-radius:12px;background:${bg};border:1px solid ${border};color:${color};font-size:11px;font-weight:600;white-space:nowrap;">${safeLabel}</span>`;
    }
    // Action button — only for open artifacts.
    let actionHtml = "";
    const isOpen = a.kind === "rfi"
      ? RFI_OPEN_STATUSES.has(a.status)
      : SUB_OPEN_STATUSES.has(a.status);
    if (isOpen) {
      const action = a.kind === "rfi" ? "log-rfi-response" : "log-sub-review";
      const label  = a.kind === "rfi" ? "📤 Log Response"  : "📤 Log Review";
      const btnColor = a.kind === "rfi" ? "#1e40af" : "#6d28d9";
      actionHtml = `<button data-action="${action}" data-id="${a.id.replace(/"/g, "&quot;")}" style="align-self:flex-start;background:${btnColor};color:white;border:none;border-radius:6px;padding:5px 12px;font-size:11px;font-weight:600;cursor:pointer;margin-left:6px;">${label}</button>`;
    }
    return `<div style="display:flex;align-items:center;flex-wrap:wrap;gap:6px;">${chipHtml}${actionHtml}</div>`;
  });
  container.innerHTML = rowsHtml.join("");
  // Event delegation: one click listener handles all action buttons.
  // Re-binding every refresh is fine since innerHTML wiped previous handlers.
  container.onclick = (ev) => {
    const btn = ev.target.closest("button[data-action]");
    if (!btn) return;
    const action = btn.getAttribute("data-action");
    const id = btn.getAttribute("data-id");
    if (action === "log-rfi-response") openRfiResponseView(id);
    else if (action === "log-sub-review") openSubReviewView(id);
  };
}
function refreshEmailSavedIndicator(animate = false) {
  const btnSharePoint = document.getElementById("saveSpBtn");
  const btnRecordOnly = document.getElementById("saveRecordBtn");
  const saveRow = document.querySelector(".save-row");
  const saveCapRow = document.querySelector(".save-row-caption");
  const confirmation = document.getElementById("saveConfirmation");
  if (!btnSharePoint || !btnRecordOnly) return;

  // Default — buttons visible, confirmation hidden. Reset button text/state.
  if (saveRow) saveRow.style.display = "";
  if (saveCapRow) saveCapRow.style.display = "";
  if (confirmation) confirmation.style.display = "none";
  btnSharePoint.disabled = false;
  btnRecordOnly.disabled = false;
  const attCount = emailFileAttachmentCount();
  btnSharePoint.textContent = (attCount && attCount > 0)
    ? `📁 Save to SharePoint · ${attCount} file${attCount > 1 ? "s" : ""}`
    : "📁 Save to SharePoint";
  btnRecordOnly.textContent = "🗂️ Save to Project";

  if (!selectedProject || !emailItem?.itemId) {
    applyEmailFlowEmphasis();
    return;
  }
  const existing = findSavedEmailRecord(selectedProject, getCurrentMessageRecordId());
  if (!existing) {
    applyEmailFlowEmphasis();
    return;
  }

  // Saved → collapse the save row into a single big-check confirmation card.
  if (saveRow) saveRow.style.display = "none";
  if (saveCapRow) saveCapRow.style.display = "none";

  const savedDate = existing.savedAt
    ? new Date(existing.savedAt).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })
    : "earlier";
  const loggedLabels = getLoggedEmailArtifactLabels(selectedProject);
  const wasFiledToSharePoint = !!existing.spFolderUrl;

  const primary = wasFiledToSharePoint
    ? "Saved to SharePoint + project record"
    : "Saved to project record";
  const secondaryParts = [`Filed ${savedDate}`];
  if (wasFiledToSharePoint && attCount && attCount > 0) {
    secondaryParts.push(`${attCount} file${attCount > 1 ? "s" : ""}`);
  } else if (!wasFiledToSharePoint && attCount && attCount > 0) {
    secondaryParts.push(`${attCount} attachment${attCount > 1 ? "s" : ""} not filed`);
  }
  if (loggedLabels.length) {
    secondaryParts.push(`also logged as ${loggedLabels.join(", ")}`);
  }
  // Append the once-per-day greeting if pending. Cleared after one read so it
  // shows immediately after the save and never again on subsequent re-opens.
  if (_pendingDayGreeting) {
    secondaryParts.push(_pendingDayGreeting);
    _pendingDayGreeting = "";
  }
  // Same pattern for content/age-aware quips — fires once after the save.
  if (_pendingContentQuip) {
    secondaryParts.push(_pendingContentQuip);
    _pendingContentQuip = "";
  }

  if (confirmation) {
    const primaryEl = confirmation.querySelector(".sc-primary");
    const secondaryEl = confirmation.querySelector(".sc-secondary");
    const linkEl = confirmation.querySelector("#scSharePointLink");
    if (primaryEl) primaryEl.textContent = primary;
    if (secondaryEl) secondaryEl.textContent = secondaryParts.join(" · ");
    // SharePoint folder link — surfaces only when the email was actually filed
    // there. openExternalUrl handles Outlook's pop-out semantics; a bare
    // target="_blank" works in Outlook web but not always in desktop.
    if (linkEl) {
      if (wasFiledToSharePoint && existing.spFolderUrl) {
        linkEl.style.display = "inline-flex";
        linkEl.onclick = (e) => { e.preventDefault(); openExternalUrl(existing.spFolderUrl); };
      } else {
        linkEl.style.display = "none";
        linkEl.onclick = null;
      }
    }
    confirmation.style.display = "flex";
    // Animation only on a fresh save click — silent on email reopen so the card
    // feels like a stable "saved" state, not a celebration that happens twice.
    confirmation.classList.remove("entering");
    if (animate) {
      void confirmation.offsetWidth;
      confirmation.classList.add("entering");
    }
  }

  // Clear the transient status banner — the confirmation card now carries the
  // saved-state message, so showing both would be redundant.
  setStatus("actionStatus", "", "");
  applyPipelineUiRules();
}

// Single source of truth for "does this email have file attachments?".
// Returns true / false / null (unknown). null happens in edge cases — e.g.
// compose mode, or clients where neither signal is populated — and callers
// should treat it as "don't bias the UI either way".
//
// `emailItem.hasAttachments` was the original signal but it's not a documented
// Office.js property; it's undefined on some clients, which falsely read as
// "no attachments" and pushed users toward the wrong save button.
function emailLikelyHasAttachments() {
  const item = emailItem;
  if (!item) return null;
  if (item.hasAttachments === true) return true;
  if (Array.isArray(item.attachments)) {
    return item.attachments.some(a =>
      a.attachmentType === Office.MailboxEnums.AttachmentType.File && !a.isInline
    );
  }
  return null;
}

// Concrete file count for the SharePoint button label. Returns a number when
// we can read the attachments array, or null when unknown (Office.js timing
// edge cases). null callers should fall back to a generic label rather than
// showing "0 files", which would mislead users about what's being saved.
function emailFileAttachmentCount() {
  const item = emailItem;
  if (!item) return null;
  if (Array.isArray(item.attachments)) {
    return item.attachments.filter(a =>
      a.attachmentType === Office.MailboxEnums.AttachmentType.File && !a.isInline
    ).length;
  }
  return null;
}

// Custom folder name override for the next SharePoint save. Cleared after the
// save runs or when the user moves to a different email. Most users never
// touch this — the ✏ rename link in the SharePoint caption is the only entry
// point, so the default subject-based naming stays the path of least resistance.
let _customSpFolderName = "";
// Inline-editor state. window.prompt is blocked in Office add-ins, so the
// rename UI is an inline input that appears inside the pane.
let _renamingSpFolder = false;

function _getDefaultSpFolderSubject() {
  return (emailItem?.subject || "No Subject")
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 70);
}

// Opens the inline rename editor by flipping state and re-rendering. The
// editor itself lives in applyEmailFlowEmphasis since it shares the caption
// slot with the static text.
function openSpFolderRenameEditor() {
  _renamingSpFolder = true;
  applyEmailFlowEmphasis();
  // Focus the input on next paint so cursor is ready for typing.
  setTimeout(() => {
    const input = document.getElementById("saveSpRenameInput");
    if (input) { input.focus(); input.select(); }
  }, 0);
}
function commitSpFolderRename(value) {
  _customSpFolderName = (value || "").trim();
  _renamingSpFolder = false;
  applyEmailFlowEmphasis();
}
function cancelSpFolderRename() {
  _renamingSpFolder = false;
  applyEmailFlowEmphasis();
}

// Visibility + emphasis for the twin save buttons.
// SharePoint save is FOR attachments — when there are none, that path doesn't
// apply, so the button (and its caption) are hidden entirely and the layout
// collapses to a single column. When attachments exist, the Project Record
// button is dimmed to nudge toward SharePoint, which writes to BOTH places.
function applyEmailFlowEmphasis() {
  const btnSp = document.getElementById("saveSpBtn");
  const btnRecord = document.getElementById("saveRecordBtn");
  const capSp = document.getElementById("saveSpCaption");
  const capRecord = document.getElementById("saveRecordCaption");
  const row = document.querySelector(".save-row");
  const capRow = document.querySelector(".save-row-caption");
  if (!btnSp || !btnRecord) return;

  // Reset emphasis, visibility, and captions each call so previous state doesn't leak.
  btnSp.classList.remove("btn-deemph");
  btnRecord.classList.remove("btn-deemph");
  btnSp.style.display = "";
  if (capSp) capSp.style.display = "";
  if (capSp) {
    // Caption shows either the default path description, the chosen custom
    // folder name, or an inline editor. Office.js blocks window.prompt(), so
    // renaming must happen via embedded DOM controls, not a native modal.
    capSp.textContent = ""; // reset

    if (_renamingSpFolder) {
      // Inline editor: input + Save/Cancel. Date prefix is appended automatically
      // at save-time, so we only ask for the trailing portion.
      const label = document.createElement("span");
      label.textContent = "YYYY_MM_DD ";
      label.style.cssText = "color:var(--muted);font-size:11px;";
      capSp.appendChild(label);

      const input = document.createElement("input");
      input.type = "text";
      input.id = "saveSpRenameInput";
      input.value = _customSpFolderName || _getDefaultSpFolderSubject();
      input.maxLength = 70;
      input.style.cssText = "width:55%;font-size:11px;padding:2px 4px;border:1px solid var(--primary);border-radius:3px;";
      capSp.appendChild(input);

      const saveBtn = document.createElement("button");
      saveBtn.type = "button";
      saveBtn.textContent = "Save";
      saveBtn.style.cssText = "margin-left:6px;font-size:11px;padding:2px 6px;border:none;background:var(--primary);color:#fff;border-radius:3px;cursor:pointer;";
      saveBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        commitSpFolderRename(input.value);
      });
      capSp.appendChild(saveBtn);

      const cancelBtn = document.createElement("button");
      cancelBtn.type = "button";
      cancelBtn.textContent = "Cancel";
      cancelBtn.style.cssText = "margin-left:4px;font-size:11px;padding:2px 6px;border:1px solid #ccc;background:#fff;color:#555;border-radius:3px;cursor:pointer;";
      cancelBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        cancelSpFolderRename();
      });
      capSp.appendChild(cancelBtn);

      input.addEventListener("keydown", e => {
        if (e.key === "Enter") {
          e.preventDefault();
          commitSpFolderRename(input.value);
        } else if (e.key === "Escape") {
          e.preventDefault();
          cancelSpFolderRename();
        }
      });

      if (_customSpFolderName) {
        const clearBtn = document.createElement("button");
        clearBtn.type = "button";
        clearBtn.textContent = "Clear";
        clearBtn.title = "Revert to email subject as the folder name";
        clearBtn.style.cssText = "margin-left:4px;font-size:11px;padding:2px 6px;border:1px solid #ccc;background:#fff;color:#a00;border-radius:3px;cursor:pointer;";
        clearBtn.addEventListener("click", e => {
          e.preventDefault();
          e.stopPropagation();
          commitSpFolderRename("");
        });
        capSp.appendChild(clearBtn);
      }
    } else {
      if (_customSpFolderName) {
        const prefix = document.createTextNode("Folder: ");
        const strong = document.createElement("strong");
        strong.style.color = "var(--text)";
        strong.textContent = "YYYY_MM_DD " + _customSpFolderName;
        capSp.appendChild(prefix);
        capSp.appendChild(strong);
      } else {
        capSp.appendChild(document.createTextNode("Email + attachments → SharePoint + record"));
      }
      const renameBtn = document.createElement("button");
      renameBtn.type = "button";
      renameBtn.id = "saveSpRenameLink";
      renameBtn.textContent = _customSpFolderName ? "✏ change" : "✏ rename";
      renameBtn.title = "Set a custom folder name (the date prefix is added automatically)";
      renameBtn.style.cssText = "margin-left:8px;color:var(--primary);background:transparent;border:none;padding:0;font:inherit;font-size:11px;font-weight:600;cursor:pointer;text-decoration:underline;";
      renameBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        openSpFolderRenameEditor();
      });
      capSp.appendChild(renameBtn);
    }
  }
  if (capRecord) capRecord.textContent = "Email body → project record only";
  if (row) row.style.gridTemplateColumns = "1fr 1fr";
  if (capRow) capRow.style.gridTemplateColumns = "1fr 1fr";

  // No project picked yet — keep both visible/neutral; we don't know what's relevant.
  if (!selectedProject) return;

  const hasAtt = emailLikelyHasAttachments();
  if (hasAtt === false) {
    // No attachments → SharePoint isn't a meaningful path. Remove it entirely
    // so there's no wrong button to click. Project Record takes full width.
    btnSp.style.display = "none";
    if (capSp) capSp.style.display = "none";
    if (row) row.style.gridTemplateColumns = "1fr";
    if (capRow) capRow.style.gridTemplateColumns = "1fr";
  } else if (hasAtt === true) {
    // Attachments exist → SharePoint is the recommended path (also writes to record).
    btnRecord.classList.add("btn-deemph");
  }
  // hasAtt === null → keep both buttons visible and neutral; we don't know enough to bias.
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
  // Surface any saves that didn't complete on a previous session.
  try { showPendingFilingBanner(); } catch (e) { console.warn("[filing-queue] banner render failed:", e.message); }
}
async function getToken(forceRefresh = false) {
  const account = msalAccount || msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
  if (!account) throw new Error("Not signed in");
  msalAccount = account;
  try {
    // forceRefresh bypasses MSAL's cached AT — used after a 401 to recover
    // from server-side token revocation or clock-skew issues that the local
    // cache thinks are still valid.
    const r = await msalApp.acquireTokenSilent({ scopes: GRAPH_SCOPES, account, forceRefresh });
    return r.accessToken;
  } catch {
    const r = await msalApp.acquireTokenPopup({ scopes: GRAPH_SCOPES, account });
    return r.accessToken;
  }
}

// On-demand token for posting channel messages. Kept separate from getToken()
// so ChannelMessage.Send isn't bundled into the default sign-in scopes —
// users see the consent prompt only when they actually click Send-to-Teams.
async function getChannelMessageToken() {
  const account = msalAccount || msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
  if (!account) throw new Error("Not signed in");
  try {
    const r = await msalApp.acquireTokenSilent({ scopes: CHANNEL_MESSAGE_SCOPES, account });
    return r.accessToken;
  } catch {
    // First call → consent popup. Subsequent silent calls succeed because
    // the consent is cached on the user's MSAL account.
    const r = await msalApp.acquireTokenPopup({ scopes: CHANNEL_MESSAGE_SCOPES, account });
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
// localStorage cache for the projects/clients picker. Two changes vs prior
// "v2" format:
//   1. Strip projects + clients down to the fields the pane actually reads
//      (picker, status banners, OneNote/Teams display). Nested arrays
//      (rfis/submittals/changeOrders/milestones/emails/notes/links) are NOT
//      cached — they're not used in the pane, and any save flow re-fetches
//      the full project from Supabase via fetchFreshProjectV2 anyway.
//   2. pako-compress before storing. Even after stripping, a busy firm with
//      many clients/contacts can produce 100 KB+ of cache; compression
//      drops it to ~15 KB and pushes the quota ceiling out of reach.
// Cache key bumped to v3 so any stale uncompressed v2 entries are ignored
// (and overwritten on first save).
const PROJECTS_CACHE_KEY = "settyPms:addinProjectsCacheV3";
const PROJECTS_CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 24h hard limit; revalidate every open

// Cache-strip strategy: keep all top-level fields (the pane reads project.emails,
// .notes, .milestones, .rfis, .submittals for picker, banner, duplicate-check),
// but drop the LARGE content fields within each nested record. The biggest
// offenders are email bodies (bodyHtmlCompressed: ~5-30 KB per email) — those
// alone account for most localStorage growth. Note bodies, RFI/submittal long
// text, and link arrays are also stripped. None of these are needed by anything
// the add-in reads from cache: bodies are re-fetched from Supabase on display.
//
// Result: a project with 20 saved emails goes from ~250 KB → ~5 KB in cache,
// before compression. After pako, typically <1 KB per project.
function _stripProjectForCache(p) {
  if (!p) return p;
  const out = { ...p };
  if (Array.isArray(p.emails)) {
    out.emails = p.emails.map(e => {
      const { bodyHtmlCompressed, bodyHtml, bodyText, bodyHtmlSize, links, ...rest } = e || {};
      return rest;
    });
  }
  if (Array.isArray(p.notes)) {
    out.notes = p.notes.map(n => {
      const { body, bodyHtml, links, ...rest } = n || {};
      return rest;
    });
  }
  if (Array.isArray(p.rfis)) {
    out.rfis = p.rfis.map(r => {
      const { notes, links, ...rest } = r || {};
      return rest;
    });
  }
  if (Array.isArray(p.submittals)) {
    out.submittals = p.submittals.map(s => {
      const { notes, links, ...rest } = s || {};
      return rest;
    });
  }
  if (Array.isArray(p.changeOrders)) {
    out.changeOrders = p.changeOrders.map(co => {
      const { notes, links, ...rest } = co || {};
      return rest;
    });
  }
  return out;
}

function loadProjectsCache() {
  try {
    const raw = localStorage.getItem(PROJECTS_CACHE_KEY);
    if (!raw) return null;
    // New v3 format: base64-deflate of JSON. If pako isn't loaded yet, fall
    // back gracefully (caller treats null as cache miss).
    if (typeof pako === "undefined") return null;
    let parsed;
    try {
      const binary = atob(raw);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      const json = pako.inflate(bytes, { to: "string" });
      parsed = JSON.parse(json);
    } catch {
      // Malformed entry (could be a stale uncompressed v2 blob that was
      // overwritten under the same key by an older client). Toss it.
      localStorage.removeItem(PROJECTS_CACHE_KEY);
      return null;
    }
    if (!parsed || !Array.isArray(parsed.projects)) return null;
    if (!parsed.savedAt || (Date.now() - parsed.savedAt) > PROJECTS_CACHE_TTL_MS) return null;
    return parsed;
  } catch {
    return null;
  }
}

function saveProjectsCache(projects, clients, versionMap) {
  try {
    // Surgical strip: keep all top-level fields and nested arrays, drop only
    // large content fields within each nested record. See _stripProjectForCache
    // for details. Clients aren't stripped (already small).
    const stripped = {
      projects: (projects || []).map(_stripProjectForCache),
      clients:  clients || [],
      versionMap: versionMap || {},
      savedAt: Date.now(),
    };
    if (typeof pako === "undefined") {
      // No compressor available — at least don't write the unstripped legacy
      // shape. We could fall back to plain JSON.stringify(stripped) but if
      // the user's localStorage is already near the cap, even that may fail.
      // Better to skip caching this open and pick it up next open when pako
      // (which loads via <script defer>) has arrived.
      console.info("[cache] pako not yet loaded — skipping cache save this cycle");
      return;
    }
    const json = JSON.stringify(stripped);
    const deflated = pako.deflate(json, { level: 6 });
    // Uint8Array → binary string → base64. Batch fromCharCode to avoid stack
    // overflow on large inputs (same pattern used by compressHtmlAddin).
    let binary = "";
    const CHUNK = 0x8000;
    for (let i = 0; i < deflated.length; i += CHUNK) {
      binary += String.fromCharCode.apply(null, deflated.subarray(i, i + CHUNK));
    }
    localStorage.setItem(PROJECTS_CACHE_KEY, btoa(binary));
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
      fetchWithRetry(SUPABASE_URL + "/rest/v1/pms_projects?select=id,project,version", { headers: SB_HEADERS }, { label: "sb loadProjects v2 projects" }),
      fetchWithRetry(SUPABASE_URL + "/rest/v1/pms_clients?select=client", { headers: SB_HEADERS }, { label: "sb loadProjects v2 clients" }),
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
    const res = await fetchWithRetry(SUPABASE_URL + "/rest/v1/pms_data?id=eq.singleton&select=projects,clients", {
      headers: SB_HEADERS,
    }, { label: "sb loadProjects legacy" });
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
  await fetchWithRetry(SUPABASE_URL + "/rest/v1/pms_data?id=eq.singleton", {
    method: "PATCH",
    headers: SB_HEADERS,
    body: JSON.stringify({ projects: updatedProjects, updated_at: new Date().toISOString() }),
  }, { label: "sb saveToSupabase" });
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
  const res = await fetchWithRetry(url, { headers: SB_HEADERS }, { label: "sb fetchFreshProject" });
  if (!res.ok) throw new Error("pms_projects GET HTTP " + res.status);
  const rows = await res.json();
  if (!rows || rows.length === 0) return null; // not migrated yet
  _projectVersionCache.set(projectId, rows[0].version);
  return { project: rows[0].project, version: rows[0].version };
}

async function saveProjectRowV2(project, expectedVersion) {
  const url = SUPABASE_URL + "/rest/v1/pms_projects?id=eq." + encodeURIComponent(project.id) +
              "&version=eq." + expectedVersion;
  const res = await fetchWithRetry(url, {
    method: "PATCH",
    headers: { ...SB_HEADERS, "Prefer": "return=representation" },
    body: JSON.stringify({
      project,
      version: expectedVersion + 1,
      updated_at: new Date().toISOString(),
    }),
  }, { label: "sb saveProjectRow" });
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
    const res = await fetchWithRetry(SUPABASE_URL + "/rest/v1/pms_data?id=eq.singleton&select=projects", {
      headers: SB_HEADERS,
    }, { label: "sb legacyApply read" });
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
    const res = await fetchWithRetry(SUPABASE_URL + "/rest/v1/pms_meta?id=eq.migration_status&select=data", { headers: SB_HEADERS }, { label: "sb migration status", maxAttempts: 2 });
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
      const res = await fetchWithRetry(SUPABASE_URL + "/rest/v1/pms_projects", {
        method: "POST",
        headers: SB_HEADERS,
        body: JSON.stringify({ id: projectId, project: mutated, version: 1, updated_at: new Date().toISOString() }),
      }, { label: "sb pms_projects insert" });
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
  const res = await fetchWithRetry(SUPABASE_URL + "/rest/v1/" + PROJECT_EMAILS_TABLE, {
    method: "POST",
    headers: { ...SB_HEADERS, Prefer: "return=minimal" },
    body: JSON.stringify(row),
  }, { label: "sb project_emails insert" });
  if (!res.ok) {
    const errText = await res.text();
    throw new Error("pms_project_emails POST HTTP " + res.status + ": " + errText.slice(0, 150));
  }
}
function updateProjectInList(updatedProject) {
  allProjects = allProjects.map(p => p.id === updatedProject.id ? updatedProject : p);
}

// ─── FILING INTEGRITY: AUDIT LOG ─────────────────────────────────────────────
// Every filing operation (Save SP, Log RFI, Log Submittal, Log Note, …) writes
// a row to pms_filing_log so we have a permanent record of WHAT was filed,
// WHERE, and WHETHER VERIFICATION SUCCEEDED. The log is the foundation for
// (a) the reconcile sweep that PMS runs on open and (b) the resume-from-crash
// path in the failed-upload queue.
//
// Logging is fire-and-forget: a logging failure must NEVER block a save. We
// always console.warn but never throw to the caller.
const FILING_LOG_TABLE = "pms_filing_log";
const CLIENT_VERSION_STRING = (typeof window !== "undefined" && window.__appVersion) ? String(window.__appVersion) : "addin";

function _getCurrentUserEmail() {
  try { return msalAccount?.username || ""; } catch { return ""; }
}

// Write a single audit-log row. `record` shape:
//   {
//     project_id:     "<uuid>",                              // required
//     msg_id:         "<itemId>" | null,
//     operation:      "email-sp" | "rfi-new" | …,            // required
//     sp_folder_url:  "<url>" | null,
//     files:          [{ name, size, sha256, contentType, verified }],
//     email_subject:  "<subject>",
//     status:         "success" | "verified" | "failed" | "partial" | "queued" | "retrying",  // required
//     error:          "<message>" | null,
//     retried:        <int>,
//   }
async function logFilingOp(record) {
  if (!record || !record.project_id || !record.operation || !record.status) {
    console.warn("[filing-log] dropped malformed record:", record);
    return null;
  }
  const row = {
    project_id:     record.project_id,
    msg_id:         record.msg_id || null,
    operation:      record.operation,
    sp_folder_url:  record.sp_folder_url || null,
    files:          record.files || null,
    email_subject:  record.email_subject || null,
    status:         record.status,
    error:          record.error ? String(record.error).slice(0, 1000) : null,
    user_email:     _getCurrentUserEmail(),
    client_version: CLIENT_VERSION_STRING,
    retried:        record.retried || 0,
  };
  try {
    const res = await fetchWithRetry(SUPABASE_URL + "/rest/v1/" + FILING_LOG_TABLE, {
      method: "POST",
      headers: { ...SB_HEADERS, Prefer: "return=minimal" },
      body: JSON.stringify(row),
    }, { label: "sb filing-log insert", maxAttempts: 2 });
    if (!res.ok) {
      const txt = await res.text().catch(() => "");
      console.warn("[filing-log] insert failed:", res.status, txt.slice(0, 200));
      return null;
    }
    return true;
  } catch (e) {
    console.warn("[filing-log] insert threw (non-fatal):", e.message);
    return null;
  }
}

// Convenience: GET recent filing-log rows for a project. Used by the reconcile
// sweep in PMS (not the add-in), but kept here so the row shape stays in sync.
async function fetchRecentFilingLog(projectId, sinceIso) {
  try {
    const sinceClause = sinceIso ? "&created_at=gte." + encodeURIComponent(sinceIso) : "";
    const url = SUPABASE_URL + "/rest/v1/" + FILING_LOG_TABLE +
      "?project_id=eq." + encodeURIComponent(projectId) +
      "&order=created_at.desc" +
      sinceClause +
      "&limit=200";
    const res = await fetchWithRetry(url, { headers: SB_HEADERS }, { label: "sb filing-log fetch", maxAttempts: 2 });
    if (!res.ok) return [];
    return await res.json();
  } catch (e) {
    console.warn("[filing-log] fetch failed:", e.message);
    return [];
  }
}

// ─── FILING INTEGRITY: CRASH-RECOVERY QUEUE ─────────────────────────────────
// localStorage map of saves currently in flight. We can't rely on the audit
// log alone to detect interrupted saves — if the browser/tab crashes between
// the PUT and the log-row insert, no row ever lands. So at the start of every
// save flow we add an entry to this map; on success we delete it. On taskpane
// open, any leftover entries indicate "a save was interrupted — retry it."
//
// Entry shape: { queueId, project_id, project_name, msg_id, operation,
//                email_subject, started_at, attempts }
// Map shape: { [queueId]: entry }
const FILING_QUEUE_KEY = "settyPms:filingQueue";

function _readFilingQueue() {
  try { return JSON.parse(localStorage.getItem(FILING_QUEUE_KEY) || "{}"); } catch { return {}; }
}
function _writeFilingQueue(q) {
  try { localStorage.setItem(FILING_QUEUE_KEY, JSON.stringify(q)); } catch (e) {
    console.warn("[filing-queue] write failed:", e.message);
  }
}

function enqueueFilingIntent(entry) {
  if (!entry || !entry.project_id || !entry.operation) return null;
  const queueId = entry.queueId || (Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 7));
  const q = _readFilingQueue();
  q[queueId] = {
    queueId,
    project_id:    entry.project_id,
    project_name:  entry.project_name || "",
    msg_id:        entry.msg_id || null,
    operation:     entry.operation,
    email_subject: entry.email_subject || "",
    started_at:    new Date().toISOString(),
    attempts:      (q[queueId]?.attempts || 0) + 1,
  };
  _writeFilingQueue(q);
  return queueId;
}

function dequeueFilingIntent(queueId) {
  if (!queueId) return;
  const q = _readFilingQueue();
  if (q[queueId]) {
    delete q[queueId];
    _writeFilingQueue(q);
  }
}

// Returns array of entries that have been pending too long ("orphaned"). The
// definition of "too long" is a generous 60s — anything older than that on
// taskpane open is almost certainly a crash from a previous session, not an
// in-flight save from the current one.
function getOrphanedFilingIntents() {
  const q = _readFilingQueue();
  const now = Date.now();
  const out = [];
  for (const id in q) {
    const entry = q[id];
    const age = now - new Date(entry.started_at).getTime();
    if (age > 60 * 1000) out.push(entry);
  }
  return out;
}

// Best-effort: prune entries older than 7 days. They're stale enough that the
// user has presumably moved on; we don't want to nag forever.
function pruneAncientFilingIntents() {
  const q = _readFilingQueue();
  const cutoff = Date.now() - 7 * 24 * 60 * 60 * 1000;
  let changed = false;
  for (const id in q) {
    if (new Date(q[id].started_at).getTime() < cutoff) {
      delete q[id];
      changed = true;
    }
  }
  if (changed) _writeFilingQueue(q);
}

// Banner UI: show pending entries on the main view with a Dismiss button.
// We deliberately don't auto-resume — the user has to consciously open the
// email and click Save again. That avoids surprising re-uploads and keeps
// the recovery path simple (no Graph-only fetch + reconstruction logic).
function showPendingFilingBanner() {
  pruneAncientFilingIntents();
  const pending = getOrphanedFilingIntents();
  const banner = document.getElementById("filingPendingBanner");
  if (!banner) {
    // First-time render — inject the banner into the DOM
    const mainView = document.getElementById("mainView");
    if (!mainView) return;
    const el = document.createElement("div");
    el.id = "filingPendingBanner";
    el.style.cssText = "display:none;background:#fef3c7;border:1px solid #f59e0b;color:#78350f;padding:8px 12px;margin:8px 12px;border-radius:6px;font-size:12px;";
    mainView.insertBefore(el, mainView.firstChild);
  }
  const b = document.getElementById("filingPendingBanner");
  if (pending.length === 0) { b.style.display = "none"; return; }
  const lines = pending.slice(0, 5).map(e =>
    `<div style="margin:4px 0">⚠ <strong>${e.operation}</strong> · ${(e.email_subject || "(no subject)").replace(/</g, "&lt;").slice(0, 60)} — interrupted ${_relativeTime(e.started_at)}</div>`
  );
  const moreNote = pending.length > 5 ? `<div style="margin-top:4px;opacity:0.7">…and ${pending.length - 5} more</div>` : "";
  b.innerHTML = `
    <div style="font-weight:600;margin-bottom:4px">${pending.length} previous save${pending.length === 1 ? "" : "s"} did not complete</div>
    ${lines.join("")}
    ${moreNote}
    <div style="margin-top:6px;display:flex;gap:8px">
      <button id="filingBannerDismiss" style="font-size:11px;padding:3px 9px;border:1px solid #d97706;background:#fff;border-radius:4px;cursor:pointer">Dismiss all</button>
      <span style="font-size:11px;opacity:0.7;align-self:center">Open the email in Outlook and click Save again to retry.</span>
    </div>
  `;
  b.style.display = "";
  const dismissBtn = document.getElementById("filingBannerDismiss");
  if (dismissBtn) {
    dismissBtn.onclick = () => {
      _writeFilingQueue({});
      b.style.display = "none";
    };
  }
}

function _relativeTime(iso) {
  const ms = Date.now() - new Date(iso).getTime();
  if (ms < 60 * 1000) return "just now";
  if (ms < 60 * 60 * 1000) return Math.floor(ms / 60000) + "m ago";
  if (ms < 24 * 60 * 60 * 1000) return Math.floor(ms / 3600000) + "h ago";
  return Math.floor(ms / 86400000) + "d ago";
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
    "sendToTeamsBtn",
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
  // Hide phase-inappropriate logging buttons (Log as RFI / Submittal) when
  // the project isn't in CA. Runs every time the selection changes so a
  // status update in PMS naturally flows through on the next selection.
  applyConstructionAdminGuard();
  // Re-render the "Logged as RFI/Sub" chip row — different project may have
  // different artifacts sourced from the same email.
  try { refreshLoggedArtifactChips(); } catch (e) { console.warn("[chips] refresh failed:", e.message); }
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
  senderDomainMatchClient: 4,
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

// Per-email date dismissals — when the user clicks × on a date chip we
// remember it locally so it doesn't reappear on reload. Stored as
// { [emailId]: ["2026-05-18", ...] } in localStorage. Per-device by design;
// dismissals are UI preference, not project data.
const DISMISSED_DATES_KEY = "setty_pms_dismissed_dates_v1";
function _loadDismissedDatesMap() {
  try {
    const raw = localStorage.getItem(DISMISSED_DATES_KEY);
    return raw ? (JSON.parse(raw) || {}) : {};
  } catch { return {}; }
}
function getDismissedDatesForCurrentEmail() {
  const id = getCurrentMessageRecordId();
  if (!id) return new Set();
  const map = _loadDismissedDatesMap();
  return new Set(map[id] || []);
}
function dismissDateForCurrentEmail(iso) {
  const id = getCurrentMessageRecordId();
  if (!id || !iso) return;
  const map = _loadDismissedDatesMap();
  const list = new Set(map[id] || []);
  list.add(iso);
  map[id] = Array.from(list);
  try { localStorage.setItem(DISMISSED_DATES_KEY, JSON.stringify(map)); } catch {}
}

async function renderDateSuggestions() {
  const block = document.getElementById("dateSuggestionBlock");
  const chips = document.getElementById("dateSuggestionChips");
  const labelText = document.getElementById("dateSuggestionLabelText");
  if (!block || !chips) return;

  // Hide chips when there's no project to attach a milestone to.
  // Also skip in compose-mode appointments and when no email is loaded.
  // Internal emails (from @setty.com) never carry client-facing milestone dates.
  if (!selectedProject || !emailItem || currentItemKind === "appointment") {
    block.style.display = "none"; chips.innerHTML = ""; return;
  }
  if ((emailFromAddress || "").toLowerCase().endsWith("@setty.com")) {
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

  // Filter out dates that already correspond to a milestone created from this
  // email (acted-on dismissal, durable across devices) AND dates the user has
  // explicitly dismissed via × (UI preference, per-device).
  const itemId   = emailItem?.itemId || "";
  const sharedId = getCurrentSharedMessageId() || "";
  const usedDates = new Set(
    (selectedProject.milestones || [])
      .filter(m => (itemId && m.sourceItemId === itemId) || (sharedId && m.sourceMessageId === sharedId))
      .map(m => m.dueDate)
      .filter(Boolean)
  );
  const dismissedDates = getDismissedDatesForCurrentEmail();
  const dates = extractDueDates(text, emailItem?.dateTimeCreated)
    .filter(d => !usedDates.has(d.iso) && !dismissedDates.has(d.iso))
    .slice(0, 3);
  if (!dates.length) { block.style.display = "none"; chips.innerHTML = ""; return; }

  if (labelText) labelText.textContent = dates.length === 1 ? "Possible milestone date" : "Possible milestone dates";
  chips.innerHTML = dates.map(d => {
    const friendly = new Date(d.iso + "T12:00:00").toLocaleDateString("en-US", {
      weekday: "short", month: "short", day: "numeric", year: "numeric"
    });
    return `
      <div class="date-chip-row">
        <button type="button" class="suggestion-chip date-chip" data-iso="${escHtml(d.iso)}">
          <span class="sc-date">${escHtml(friendly)}${d.hasKeyword ? ' <span class="pill" style="background:var(--primary-soft);color:var(--primary-hov);border:1px solid #b7daf2;">deadline</span>' : ""}</span>
          <span class="sc-cta">+ Create →</span>
        </button>
        <button type="button" class="chip-dismiss" data-iso="${escHtml(d.iso)}" title="Not a deadline — dismiss" aria-label="Dismiss this date">×</button>
      </div>
    `;
  }).join("");
  chips.querySelectorAll(".suggestion-chip.date-chip").forEach(el => {
    el.onclick = () => openMilestoneFormFromChip(el.dataset.iso);
  });
  chips.querySelectorAll(".chip-dismiss").forEach(el => {
    el.onclick = (e) => {
      e.stopPropagation();
      dismissDateForCurrentEmail(el.dataset.iso);
      void renderDateSuggestions();
    };
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
// ─── TRANSIENT-FAILURE RETRY HELPER ──────────────────────────────────────────
// Single shared exponential-backoff wrapper. Retries network failures, 429,
// 503, 504. Honors Retry-After if Graph or Supabase provides it. Capped at
// 15s per wait so the taskpane never appears hung.
async function fetchWithRetry(url, init, opts = {}) {
  const maxAttempts = opts.maxAttempts ?? 3;
  const label = opts.label || "fetch";
  let lastErr;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const res = await fetch(url, init);
      if (res.status === 429 || res.status === 503 || res.status === 504) {
        if (attempt === maxAttempts) return res;
        const retryAfter = parseInt(res.headers.get("Retry-After") || "0", 10);
        const wait = retryAfter > 0
          ? Math.min(15000, retryAfter * 1000)
          : Math.min(15000, 500 * Math.pow(2, attempt));
        console.warn(`[retry:${label}] ${res.status} — waiting ${wait}ms (attempt ${attempt}/${maxAttempts})`);
        await new Promise(r => setTimeout(r, wait));
        continue;
      }
      return res;
    } catch (e) {
      lastErr = e;
      if (attempt === maxAttempts) throw e;
      const wait = Math.min(15000, 500 * Math.pow(2, attempt));
      console.warn(`[retry:${label}] network error — waiting ${wait}ms (attempt ${attempt}/${maxAttempts}): ${e.message}`);
      await new Promise(r => setTimeout(r, wait));
    }
  }
  throw lastErr || new Error(`${label}: exhausted retries`);
}

// ─── GRAPH HELPERS ────────────────────────────────────────────────────────────
async function graphFetch(method, path, body, token) {
  const t = token || await getToken();
  const makeReq = (tok) => fetchWithRetry("https://graph.microsoft.com/v1.0" + path, {
    method,
    headers: {
      "Authorization": "Bearer " + tok,
      "Content-Type": "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  }, { label: "graph " + method });

  let res = await makeReq(t);
  // 401 → force-refresh token + retry once. Long-running saves (slow VPN,
  // big attachments) can outlive MSAL's ~45 min token TTL — without this
  // refresh, the save fails with a useless "Graph 401" near the end.
  if (res.status === 401) {
    try {
      const fresh = await getToken(true);
      res = await makeReq(fresh);
    } catch { /* fall through to error path below */ }
  }
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
// Per-itemId cache for getEmailBodyHtml. Same email saved twice (e.g., retry
// after a transient error) skips the second Graph round-trip — big win on
// retry paths and double-action flows (Save SP then Log as Note in succession).
// Cleared on item switch (loadItemContext).
const _emailBodyCache = new Map();
function clearEmailBodyCache() { _emailBodyCache.clear(); }
async function getEmailBodyHtml(token) {
  try {
    const msgId = Office.context.mailbox.item.itemId;
    if (_emailBodyCache.has(msgId)) return _emailBodyCache.get(msgId);
    const restId = Office.context.mailbox.convertToRestId(msgId, Office.MailboxEnums.RestVersion.v2_0);
    const data = await graphFetch("GET", "/me/messages/" + restId + "?$select=body", null, token);
    const body = data?.body?.content || "";
    if (body) _emailBodyCache.set(msgId, body);
    return body;
  } catch { return ""; }
}

// In-memory idempotency cache for OneNote page creation (5-minute TTL).
// Prevents duplicate pages when a 5xx response masks a server-side success
// and the retry loop reposts. Keyed by url + html-prefix hash.
const _addinOneNoteCache = new Map();
const ADDIN_ONENOTE_DEDUP_TTL_MS = 5 * 60 * 1000;
function _hashOneNoteReq(url, html) {
  let h = 2166136261;
  const s = url + "|" + html.slice(0, 2000);
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return (h >>> 0).toString(36);
}

// Phase 3: compress HTML email body to base64 deflate before storing in the
// project record. Identical implementation to PMS so records are interchangeable.
// Logs a clear warning when compression fails non-trivially (i.e., we had real
// HTML to compress but couldn't), so failures don't silently produce empty
// bodyHtmlCompressed fields the user later sees as "Live-fetched only".
// ─── WEEKLY SAVE STREAK ──────────────────────────────────────────────────────
// Variable rewards on round numbers — most saves get the standard "✓ Filed"
// confirmation; only milestone counts (10/25/50/100 per week) trigger the
// celebration. Per-device counter via localStorage; cross-device sync would
// need a Supabase write per save and isn't worth it for a fun nudge.
const SAVE_STREAK_KEY = "setty_pms_save_streak_v1";

// ─── AEC FLAVOR PACK ────────────────────────────────────────────────────────
// Personality strings for the add-in. Goal: rare, varied, AEC-aware.
// Frequent repetition kills the charm, so quips rotate randomly and ambient
// ones fire at low probability (~8%). Plain "⏳ Saving…" is in the saving pool
// so ~1 in 8 saves gets the boring version, which keeps the rest feeling like
// nice surprises rather than wallpaper.
const SAVING_QUIPS = [
  "⏳ Saving…",
  "🏗️ Notarizing the email…",
  "📐 Calibrating documentation gravity…",
  "📁 Filing with extreme precision…",
  "🎯 Threading into the project record…",
  "📋 Asking the project record very politely…",
  "🔧 Tightening a metaphorical bolt…",
  "🗂️ Cross-referencing the master plan…",
];
const EMAIL_OPEN_QUIPS = [
  "*googly eyes* this one looks important",
  "*adjusts hard hat*",
  "*applies clipboard authority*",
  "*measures email* yep, that's definitely an email",
  "*examines metadata thoughtfully*",
  "*target acquired*",
  "*cross-references the master plan*",
  "*tightens a metaphorical bolt*",
  "*nods sagely*",
  "*consults the AIA standard for this exact moment*",
  "*sniffs for change orders*",
];
const MILESTONE_QUIPS_10 = [
  "🎉 10 saved! That's a respectable start.",
  "🎉 10 emails — foundation of project memory laid.",
  "🎉 10 down — documenting like a court reporter on caffeine.",
  "🎉 10 saved! Future-You will send a thank-you note.",
  "🎉 10 emails — measure twice, file once. You're doing both.",
];
const MILESTONE_QUIPS_25 = [
  "🔥 25 this week — strong rhythm!",
  "🔥 25 saved — architects are jealous of your filing game.",
  "🔥 25 emails — the project record gods are pleased.",
  "🔥 25 down — your project history is becoming legendary.",
  "🔥 25 emails — fixing project memory one save at a time.",
];
const MILESTONE_QUIPS_50 = [
  "🚀 50 emails — on a roll!",
  "🚀 50 saved! At this rate you'll need a bigger SharePoint folder.",
  "🚀 50 — basically the project's official scribe at this point.",
  "🚀 50 — *applies extra clipboard authority*",
  "🚀 50 emails! Documentation icon status: confirmed.",
];
const MILESTONE_QUIPS_100 = [
  "🏆 100 emails — legendary week!",
  "🏆 100 saved — you're now the project archivist. Update LinkedIn.",
  "🏆 100 emails — Setty docs hall of fame.",
  "🏆 100 — you've crossed from 'PM' to 'librarian'.",
  "🏆 100 — enough record to write an entire AIA standard.",
];
const STREAK_THRESHOLDS = [
  { count: 10,  pool: MILESTONE_QUIPS_10 },
  { count: 25,  pool: MILESTONE_QUIPS_25 },
  { count: 50,  pool: MILESTONE_QUIPS_50 },
  { count: 100, pool: MILESTONE_QUIPS_100 },
];

// Pep quips fired when a NEW milestone is saved (from a date chip or "New
// Milestone" form). Different vibe than the email-save quips — these are about
// commitment to a future date, so the encouragement leans toward "you got this".
const NEW_MILESTONE_QUIPS = [
  "📌 Locked in — keeping it on point!",
  "🎯 Don't forget about this one!",
  "📅 Milestone pinned — Future-You is grateful.",
  "📐 New dot on the project timeline.",
  "🗓️ Marked. Move with confidence.",
  "⏰ This one's on the radar now.",
  "✅ Date noted, project record updated.",
  "🔔 Milestone in the books.",
];

// Combinatorial silly-word generator (Claude Code style). With ~17 verbs ×
// ~10 nouns × 9 emojis = 1,530+ unique combos, daily users will rarely see
// the same one twice — much fresher than a curated pool alone.
const SILLY_VERBS = [
  "Frobnicating", "Hatching", "Pondering", "Cogitating", "Wrangling",
  "Marinating", "Conjuring", "Reticulating", "Burnishing", "Spelunking",
  "Quibbling", "Caboodling", "Bamboozling", "Filibustering", "Spellbinding",
  "Hornswoggling", "Yarning",
];
const SILLY_NOUNS = [
  "the email", "the metadata", "the project entropy", "the docu-mojo",
  "the timestamps", "the bytes", "the cosmic file order",
  "the project chunkings", "the file vibes", "the AEC ether",
];
const SILLY_EMOJI = ["🛠️", "🪄", "🌀", "🎩", "🧙", "🦴", "🐌", "🪅", "✨"];

function generateSillySavingMessage() {
  const e = SILLY_EMOJI[Math.floor(Math.random() * SILLY_EMOJI.length)];
  const v = SILLY_VERBS[Math.floor(Math.random() * SILLY_VERBS.length)];
  const n = SILLY_NOUNS[Math.floor(Math.random() * SILLY_NOUNS.length)];
  return `${e} ${v} ${n}…`;
}

function pickQuip(pool) {
  return pool[Math.floor(Math.random() * pool.length)];
}

// Lazy-load canvas-confetti on first celebration. Saves the ~5KB download
// (and parse cost) on every pane open for users who never hit a milestone.
// Browser caches the script after first load, so subsequent celebrations
// are instant. Returns a Promise<boolean> — true when ready, false on load error.
let _confettiLoadPromise = null;
function loadConfetti() {
  if (typeof confetti === "function") return Promise.resolve(true);
  if (_confettiLoadPromise) return _confettiLoadPromise;
  _confettiLoadPromise = new Promise(resolve => {
    const s = document.createElement("script");
    s.src = "https://cdn.jsdelivr.net/npm/canvas-confetti@1.9.3/dist/confetti.browser.min.js";
    s.onload = () => resolve(true);
    s.onerror = () => { _confettiLoadPromise = null; resolve(false); }; // allow retry next time
    document.head.appendChild(s);
  });
  return _confettiLoadPromise;
}

// Saving message picker — ~30% chance of a freshly-generated silly combo,
// otherwise pulls from the curated pool. The dual layer keeps things from
// getting stale even for users who save dozens of emails a day.
function pickSavingMessage() {
  if (Math.random() < 0.3) return generateSillySavingMessage();
  return pickQuip(SAVING_QUIPS);
}

// Time-of-day greeting — first save of the day gets a personalized line
// appended to the post-save card. Each slot has multiple variants so daily
// users see something different. Lunch and afternoon slots include break
// reminders, since those are the hours people actually skip breaks.
const TIME_GREETINGS = {
  morning: [
    "☕ Morning — first file of the day. Strong start.",
    "☕ Filing before 10am? Disciplined.",
    "🌅 Bright and early. Project record blessed.",
  ],
  lateBreakfast: [
    "🥐 Late breakfast filing. Solid.",
    "🥐 Pre-lunch productivity. Building momentum.",
    "🍵 Mid-morning groove. Nice pace.",
  ],
  lunch: [
    "🥪 Lunchtime filing — but eat something too, ok?",
    "🥪 Filing while you eat? AEC heroics. Don't forget the food.",
    "🥗 Lunch-hour documentation. Hydrate too.",
    "🍴 Filing through lunch? At least step away from the screen for 5.",
  ],
  afternoon: [
    "📊 Mid-afternoon focus. Respect.",
    "🧘 Afternoon files going in — also: stand up, stretch, water?",
    "📊 3pm momentum. You've earned a 5-min break soon.",
    "👀 Eyes off the screen for a sec? Then back to it.",
    "☕ Mid-afternoon — second coffee window is officially open.",
  ],
  evening: [
    "🌅 Evening filing — wrapping up clean.",
    "🌅 End-of-day cleanup. Tomorrow-You says thanks.",
    "🌇 Closing the loop on today. Nice.",
  ],
  lateEvening: [
    "🌙 9-to-9 day? Thanks for the dedication.",
    "🌙 Past 7pm? Make sure dinner happened.",
    "🌃 Evening shift respect.",
  ],
  lateNight: [
    "🦉 Filing past 10pm — admirable. Sleep is also good.",
    "🌙 Late shift respect. Set a hard stop?",
    "🦉 The owl hours. Don't let this become a habit.",
  ],
};
function timeOfDayGreeting() {
  const h = new Date().getHours();
  let pool;
  if (h >= 5  && h < 10) pool = TIME_GREETINGS.morning;
  else if (h >= 10 && h < 12) pool = TIME_GREETINGS.lateBreakfast;
  else if (h >= 12 && h < 14) pool = TIME_GREETINGS.lunch;
  else if (h >= 14 && h < 17) pool = TIME_GREETINGS.afternoon;
  else if (h >= 17 && h < 19) pool = TIME_GREETINGS.evening;
  else if (h >= 19 && h < 22) pool = TIME_GREETINGS.lateEvening;
  else if (h >= 22)           pool = TIME_GREETINGS.lateNight;
  else pool = ["🌌 Filing at " + h + ":00? You're a different kind of person. Respect."];
  return pickQuip(pool);
}

const LAST_SAVE_DATE_KEY = "setty_pms_last_save_date_v1";
let _pendingDayGreeting = "";
let _pendingContentQuip = "";

// Content-aware quips — fire on save when the subject contains specific
// keywords. ~30% chance per match so they stay surprising. Active voice
// makes the tool read like it's commenting on the email, not just labeling it.
const CONTENT_AWARE_QUIPS = [
  { pattern: /\basap\b/i,                              quip: "👀 the magic word, archived" },
  { pattern: /\bdeadline\b/i,                          quip: "🎯 deadline noted" },
  { pattern: /\bthank ?you\b|\bthanks\b/i,             quip: "📩 a thank-you, catalogued" },
  { pattern: /\burgent\b|\btime[- ]sensitive\b/i,      quip: "🚨 urgency, archived" },
  { pattern: /\bapprov(ed|al)\b/i,                     quip: "✅ approval, preserved" },
  { pattern: /\brfi\b|\brfi[- ]?\d+\b/i,               quip: "🔵 RFI logged in the official record" },
  { pattern: /\bsubmittal\b/i,                         quip: "📋 submittal, tracked" },
  { pattern: /\bchange order\b|\bco[- ]?\d{2,}\b/i,    quip: "💰 change order, noted" },
  { pattern: /\bdelay(ed)?\b|\bbehind schedule\b/i,    quip: "⏳ delay, on the record" },
  { pattern: /\bmeeting\b/i,                           quip: "🪑 meeting evidence captured" },
  { pattern: /\binvoice\b|\bpayment\b/i,               quip: "💵 financial trail extended" },
  { pattern: /\bsigned?\b|\bsignature\b/i,             quip: "✒️ signed, sealed, filed" },
];
function detectContentQuip(text) {
  if (!text) return "";
  for (const { pattern, quip } of CONTENT_AWARE_QUIPS) {
    if (pattern.test(text)) {
      // 30% probability per match — rare enough to feel discovered, frequent
      // enough that users with relevant emails actually see them.
      return Math.random() < 0.3 ? quip : "";
    }
  }
  return "";
}

// Age-based quip — saving an email older than ~6 months is rare enough to
// always fire when it happens. Plays into the project archivist vibe.
function detectAgeQuip() {
  if (!emailItem?.dateTimeCreated) return "";
  const ageDays = (Date.now() - new Date(emailItem.dateTimeCreated).getTime()) / 86400000;
  if (ageDays >= 180) return "🗿 ancient artifact filed for posterity";
  return "";
}

// Seasonal confetti modifiers — applied on top of the base confetti config
// during celebrations. Snowflakes in December, hearts on Valentine's,
// patriotic on July 4, etc. Returns null on ordinary days so the standard
// confetti palette runs. confetti.shapeFromText was added in v1.6+.
function getSeasonalConfettiOpts() {
  if (typeof confetti !== "function") return null;
  const d = new Date();
  const m = d.getMonth(); // 0-indexed
  const day = d.getDate();
  if (m === 11) {
    // December — snowflakes all month for that holiday vibe.
    return {
      colors: ["#ffffff", "#e3f2fd", "#bbdefb", "#90caf9"],
      shapes: [confetti.shapeFromText({ text: "❄️", scalar: 2 })],
      scalar: 1.6,
    };
  }
  if (m === 0 && day === 1) {
    // New Year's Day — gold/silver/pink/cyan party palette.
    return { colors: ["#ffd700", "#c0c0c0", "#ff6b9d", "#6bd4ff"] };
  }
  if (m === 1 && day === 14) {
    // Valentine's Day — hearts.
    return {
      colors: ["#ff3a7a", "#ff6b9d", "#ffb3c6", "#ffffff"],
      shapes: [confetti.shapeFromText({ text: "❤️", scalar: 2 })],
      scalar: 2,
    };
  }
  if (m === 6 && day === 4) {
    // July 4 — red/white/blue stars.
    return {
      colors: ["#bf0a30", "#ffffff", "#002868"],
      shapes: ["star", "circle"],
    };
  }
  if (m === 9 && day === 31) {
    // Halloween — orange & black.
    return { colors: ["#ff7518", "#000000", "#8a2be2", "#ffd700"] };
  }
  return null;
}
function consumeFirstSaveOfDay() {
  try {
    const today = new Date().toISOString().slice(0, 10);
    const last = localStorage.getItem(LAST_SAVE_DATE_KEY);
    if (last !== today) {
      localStorage.setItem(LAST_SAVE_DATE_KEY, today);
      return true;
    }
  } catch {}
  return false;
}

// Rare ambient observation under the email preview. Low probability (~8%) so
// it lands like a wink, not wallpaper. Skipped for calendar appointments
// (currentItemKind === "appointment") since the vibe doesn't fit.
function maybeShowAecQuip() {
  const line = document.getElementById("aecQuipLine");
  if (!line) return;
  if (!emailItem || currentItemKind === "appointment") {
    line.style.display = "none";
    return;
  }
  if (Math.random() < 0.08) {
    line.textContent = pickQuip(EMAIL_OPEN_QUIPS);
    line.style.display = "block";
  } else {
    line.style.display = "none";
  }
}

function _weekStartISO() {
  // Monday-anchored ISO week. Anyone working a Sun-Sat week feels off-by-one
  // for one day; switch the offset math if the firm prefers Sunday start.
  const d = new Date();
  const day = d.getDay(); // 0=Sun..6=Sat
  const offset = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + offset);
  return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0");
}

// First-ever save gets a one-time over-the-top welcome. Fires exactly once per
// device (localStorage flag), independent of the weekly streak counter. Marks
// the *transition* from "haven't done this yet" → "now I'm a person who does this."
const FIRST_SAVE_KEY = "setty_pms_first_save_done_v1";
function _isFirstSaveEver() {
  try { return !localStorage.getItem(FIRST_SAVE_KEY); } catch { return false; }
}
function _markFirstSaveDone() {
  try { localStorage.setItem(FIRST_SAVE_KEY, "1"); } catch {}
}

function recordSaveAndCelebrate() {
  // Check first-save BEFORE bumping the weekly counter so the welcome fires
  // even if this is also weekly-count #1.
  const isFirstEver = _isFirstSaveEver();
  if (isFirstEver) _markFirstSaveDone();

  const weekStart = _weekStartISO();
  let state = { weekStart, count: 0 };
  try {
    const raw = localStorage.getItem(SAVE_STREAK_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && parsed.weekStart === weekStart) state = parsed;
    }
  } catch { /* corrupt storage — start fresh */ }
  state.count = (state.count || 0) + 1;
  try { localStorage.setItem(SAVE_STREAK_KEY, JSON.stringify(state)); } catch {}

  // First save of the day — append a small greeting to the post-save card.
  // Read by refreshEmailSavedIndicator and cleared after one use.
  if (consumeFirstSaveOfDay()) _pendingDayGreeting = timeOfDayGreeting();

  // Content-aware + age-aware quips. Age wins when present — it's a much
  // rarer event (saving a 6+ month old email) and deserves the spotlight.
  const ageQuip = detectAgeQuip();
  const contentQuip = detectContentQuip(emailItem?.subject || "");
  _pendingContentQuip = ageQuip || contentQuip;

  if (isFirstEver) {
    triggerFirstSaveCelebration();
    return; // skip weekly thresholds; first-save is the headline.
  }
  const hit = STREAK_THRESHOLDS.find(t => t.count === state.count);
  if (hit) triggerCelebration(pickQuip(hit.pool));
}

// Fireworks-style burst pattern — two columns of bursts firing alternately for
// 2.5s, then a finale. The mix of star and circle shapes plus larger scalar
// makes it visually distinct from the standard milestone celebration.
function triggerFirstSaveCelebration() {
  const message = "🎊 First save! Welcome to the project record club 🎊";
  const toast = document.getElementById("celebrationToast");
  if (toast) {
    toast.textContent = message;
    toast.classList.remove("show", "first-save");
    void toast.offsetWidth;
    toast.classList.add("show", "first-save");
    clearTimeout(triggerCelebration._t);
    triggerCelebration._t = setTimeout(() => {
      toast.classList.remove("show", "first-save");
    }, 5000);
  }
  // Lazy-load confetti on first celebration. The fireworks pattern starts
  // once it resolves; toast/animation above is unaffected.
  loadConfetti().then(ok => {
    if (!ok || typeof confetti !== "function") return;
    const duration = 2500;
    const animationEnd = Date.now() + duration;
    const seasonal = getSeasonalConfettiOpts() || {};
    const defaults = { startVelocity: 32, spread: 360, ticks: 80, zIndex: 9999, scalar: 1.1, ...seasonal };
    const rand = (min, max) => Math.random() * (max - min) + min;
    const interval = setInterval(() => {
      const timeLeft = animationEnd - Date.now();
      if (timeLeft <= 0) { clearInterval(interval); return; }
      const particleCount = 60 * (timeLeft / duration);
      confetti({ ...defaults, particleCount, origin: { x: rand(0.1, 0.3), y: rand(0, 0.4) } });
      confetti({ ...defaults, particleCount, origin: { x: rand(0.7, 0.9), y: rand(0, 0.4) } });
    }, 200);
    // Big finale a beat after the rolling bursts end — mixed shapes for variety.
    setTimeout(() => {
      confetti({ particleCount: 160, spread: 110, startVelocity: 50, origin: { y: 0.55 }, scalar: 1.3, shapes: ["star", "circle"], ...seasonal });
    }, duration + 80);
  });
}

function triggerCelebration(message) {
  const toast = document.getElementById("celebrationToast");
  if (toast) {
    toast.textContent = message;
    // Force reflow so the .show transition fires even if a previous toast
    // is still in-flight (e.g., rapid back-to-back milestone hits).
    toast.classList.remove("show");
    void toast.offsetWidth;
    toast.classList.add("show");
    clearTimeout(triggerCelebration._t);
    triggerCelebration._t = setTimeout(() => {
      toast.classList.remove("show");
    }, 2800);
  }
  // Lazy-load confetti on first celebration. Toast above is sync; confetti
  // arrives a beat later — actually feels more dramatic.
  loadConfetti().then(ok => {
    if (!ok || typeof confetti !== "function") return;
    // Two bursts a beat apart — feels more alive than one big shot.
    // Seasonal opts layer on top (snowflakes in Dec, hearts on Valentine's, etc.)
    const seasonal = getSeasonalConfettiOpts() || {};
    confetti({ particleCount: 80, spread: 75, startVelocity: 35, origin: { y: 0.55 }, scalar: 0.85, ...seasonal });
    setTimeout(() => confetti({ particleCount: 50, spread: 110, startVelocity: 28, origin: { y: 0.45 }, scalar: 0.7, ...seasonal }), 220);
  });
}

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
// `item` is the Office mailbox-item snapshot captured by the caller. Falls
// back to the module global when undefined for legacy callers.
function buildEmailHtml(bodyHtml, item) {
  const src = item || emailItem;
  const esc = (s) => String(s || "").replace(/[&<>"']/g, c =>
    ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;" })[c]);
  const from = src?.from;
  const header = `<div style="font-family:sans-serif;font-size:12px;padding:12px 16px;border-bottom:1px solid #ddd;margin-bottom:16px">
    <strong>Subject:</strong> ${esc(src?.subject)}<br>
    <strong>From:</strong> ${esc(from?.displayName)} &lt;${esc(from?.emailAddress)}&gt;<br>
    <strong>Date:</strong> ${esc(new Date(src?.dateTimeCreated).toLocaleString())}
  </div>`;
  return "<!DOCTYPE html><html><head><meta charset='utf-8'></head><body>" + header + (bodyHtml || "<p style='color:#666;font-style:italic;padding:8px 12px;background:#f5f5f5;border-left:3px solid #ccc;'>No body content &mdash; this email may be a system notification, share invite, or attachment-only message.</p>") + "</body></html>";
}
// Upload email.html + any attachments into targetPath. Returns attachment count.
//
// `itemSnapshot` MUST be the Office mailbox item captured by the caller before
// any awaits. Without it, this function would read the module-level `emailItem`
// global which Office may have silently swapped (pinned-pane item switch),
// causing attachment bytes from a *different* email to be uploaded under the
// current email's folder. That's the worst kind of corruption — looks right,
// is wrong.
async function uploadEmailAndAttachments(driveId, token, targetPath, itemSnapshot) {
  const item = itemSnapshot || emailItem; // fallback for legacy callers
  // uploadedFiles[]: per-file metadata for the audit log. Each successful upload
  // appends { name, size, sha256, contentType, verified }. Phase 2 will populate
  // sha256 + verified once read-back lands.
  lastAttachmentUploadStats = { attempted: 0, uploaded: 0, failed: [], uploadedFiles: [] };
  const bodyHtml = await getEmailBodyHtml(token);
  // Kick off the email.html upload in parallel with the attachment loop —
  // they don't depend on each other, so why serialize them.
  const emailHtmlPromise = fetch(
    "https://graph.microsoft.com/v1.0/drives/" + driveId + "/root:/" + encodeDrivePath(targetPath) + "/email.html:/content",
    {
      method: "PUT",
      headers: { "Authorization": "Bearer " + token, "Content-Type": "text/html" },
      body: buildEmailHtml(bodyHtml, item),
    }
  );

  // Quick-Win #3: parallelize attachment uploads with bounded concurrency.
  const ATTACHMENT_CONCURRENCY = 3;
  async function uploadInBatches(items, doUpload) {
    const failures = [];
    let succeeded = 0;
    for (let i = 0; i < items.length; i += ATTACHMENT_CONCURRENCY) {
      const batch = items.slice(i, i + ATTACHMENT_CONCURRENCY);
      const results = await Promise.allSettled(batch.map(doUpload));
      results.forEach((r, idx) => {
        const it = batch[idx];
        if (r.status === "fulfilled" && r.value) succeeded++;
        else failures.push((it.name || "attachment") + (r.status === "rejected" ? " (" + (r.reason?.message || "error").slice(0, 60) + ")" : ""));
      });
    }
    return { succeeded, failures };
  }

  // De-collision: within a batch, if two attachments would produce the same
  // sanitized filename (or collide with the reserved "email.html"), suffix
  // the later ones with " (2)", " (3)", etc. Without this, a second image.png
  // would silently overwrite the first.
  function uniquifyNames(names) {
    const taken = new Set(["email.html"]);
    return names.map(rawName => {
      const safe = (rawName || "attachment").replace(/[\\/:*?"<>|]/g, "-").trim() || "attachment";
      if (!taken.has(safe)) { taken.add(safe); return safe; }
      const dot = safe.lastIndexOf(".");
      const stem = dot > 0 ? safe.slice(0, dot) : safe;
      const ext  = dot > 0 ? safe.slice(dot) : "";
      for (let n = 2; n < 1000; n++) {
        const candidate = `${stem} (${n})${ext}`;
        if (!taken.has(candidate)) { taken.add(candidate); return candidate; }
      }
      return safe + "-" + Math.random().toString(36).slice(2, 7); // last resort
    });
  }

  try {
    let count = 0;
    // Prefer Outlook item APIs for attachment bytes; this is the most reliable in add-ins.
    // Critical: pass `item` (the snapshot) — getOfficeFileAttachments would otherwise
    // re-read the module-level emailItem and could pick up a different email's attachments.
    const officeAtts = await getOfficeFileAttachments(item);
    if (officeAtts.length) {
      lastAttachmentUploadStats.attempted = officeAtts.length;
      const uniqueNames = uniquifyNames(officeAtts.map(a => a.name));
      const { succeeded, failures } = await uploadInBatches(
        officeAtts,
        async (att) => {
          const finalName = uniqueNames[officeAtts.indexOf(att)];
          // uploadAttachmentToSharePoint records verified metadata into
          // lastAttachmentUploadStats.uploadedFiles itself (Phase 2). No extra
          // push needed here.
          return uploadAttachmentToSharePoint(driveId, token, targetPath, finalName, att.contentType, att.bytes);
        }
      );
      count = succeeded;
      lastAttachmentUploadStats.failed.push(...failures);
      lastAttachmentUploadStats.uploaded = count;
      // Make sure email.html upload completed before returning
      await emailHtmlPromise;
      return count;
    }
    // Fallback to Graph attachment APIs when Office APIs are unavailable.
    const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
    const attData = await graphFetch("GET", "/me/messages/" + restId + "/attachments", null, token);
    const fileAtts = (attData?.value || []).filter(att => att["@odata.type"] === "#microsoft.graph.fileAttachment");
    lastAttachmentUploadStats.attempted = fileAtts.length;
    const uniqueGraphNames = uniquifyNames(fileAtts.map(a => a.name));

    const { succeeded, failures } = await uploadInBatches(fileAtts, async (att) => {
      let bytes = null;
      if (att.contentBytes) {
        const binary = atob(att.contentBytes);
        bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      } else if (att.id) {
        const rawRes = await fetchWithRetry(
          "https://graph.microsoft.com/v1.0/me/messages/" + restId + "/attachments/" + att.id + "/$value",
          { headers: { "Authorization": "Bearer " + token } },
          { label: "graph attachment $value" }
        );
        if (!rawRes.ok) {
          throw new Error("download " + rawRes.status);
        }
        bytes = new Uint8Array(await rawRes.arrayBuffer());
      }
      if (!bytes) return false;
      // NOTE: Graph's fileAttachment.size also reports MIME-encoded size, not
      // raw decoded bytes — same trap as Office.js. We log a warning if the
      // numbers look wildly off (could indicate genuine truncation on the
      // $value GET path) but don't reject the upload. Bytes from atob() of a
      // complete contentBytes string are by definition complete.
      if (typeof att.size === "number" && att.size > 0 && bytes.length > att.size + 16) {
        // Decoded bytes shouldn't ever be LARGER than the reported (encoded) size.
        console.warn(`[attachment] suspicious size for ${att.name}: decoded=${bytes.length} graph=${att.size} — uploading anyway`);
      }
      const safeName = uniqueGraphNames[fileAtts.indexOf(att)];
      // uploadAttachmentToSharePoint records verified metadata itself (Phase 2).
      return uploadAttachmentToSharePoint(driveId, token, targetPath, safeName, att.contentType, bytes);
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
// `item` MUST be passed in by the caller (a snapshot captured before any
// awaits). Reading the module-level emailItem inside this loop was a
// real-world corruption risk: Office swaps mailbox.item silently when the
// user clicks a different email, and a getAttachmentContentAsync call
// against the new item with an attachment ID from the old item produces
// either an error or — worse — bytes that *look* valid but are wrong.
async function getOfficeFileAttachments(item) {
  item = item || emailItem;
  if (!item?.getAttachmentsAsync || !item?.getAttachmentContentAsync) return [];
  const atts = await new Promise((resolve, reject) => {
    item.getAttachmentsAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value || []);
      else reject(new Error(res.error?.message || "getAttachmentsAsync failed"));
    });
  });
  const fileAtts = atts.filter(att => att.attachmentType === Office.MailboxEnums.AttachmentType.File);
  const out = [];
  for (const att of fileAtts) {
    const content = await new Promise((resolve, reject) => {
      // CRITICAL: use the captured `item`, not the live module global.
      item.getAttachmentContentAsync(att.id, (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
        else reject(new Error(res.error?.message || "getAttachmentContentAsync failed"));
      });
    }).catch((e) => {
      console.warn("Office attachment content failed:", att.name, e.message);
      return null;
    });
    if (!content || content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) continue;
    const bytes = toBytesFromBase64(content.content);
    // NOTE: Office.js's att.size is NOT the raw byte count — it's the
    // MIME-encoded size including base64 transport overhead (~+33%) and headers.
    // So we can't validate decoded bytes against att.size here. The Graph
    // fallback path (the other branch in uploadEmailAndAttachments) DOES get
    // raw byte size from Graph and validates there; that's the right place.
    // For Office.js, getAttachmentContentAsync arrives as one complete base64
    // payload — there's no streaming, so truncation is essentially impossible
    // and additional client-side validation would only produce false negatives.
    out.push({
      name: att.name || "attachment",
      contentType: att.contentType || "application/octet-stream",
      bytes,
    });
  }
  return out;
}
// Graph's simple PUT to /content is capped at 4 MiB. Above that, Graph
// returns 413 (or in some corner cases, just hangs) and the file fails to
// upload. We auto-dispatch to the upload-session API for anything > threshold.
const SP_SIMPLE_UPLOAD_MAX = 4 * 1024 * 1024;          // 4 MiB
const SP_UPLOAD_CHUNK_SIZE = 5 * 320 * 1024;           // 1.6 MiB — multiple of 320 KiB (Graph requirement)

// ─── INTEGRITY HASHING ───────────────────────────────────────────────────────
// SHA-256 of the local bytes via Web Crypto. Used by the read-back verification
// step to confirm that what SharePoint stored matches what we sent. Returns
// lowercase hex string. Cost on a 5MB file: ~30ms in Chrome.
async function sha256Hex(bytes) {
  if (!bytes || bytes.length === 0) return "";
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  const view = new Uint8Array(digest);
  let hex = "";
  for (let i = 0; i < view.length; i++) hex += view[i].toString(16).padStart(2, "0");
  return hex;
}

// Microsoft Graph's QuickXorHash — a custom XOR-with-rotation digest used by
// OneDrive/SharePoint. We only need this when sha256Hash isn't present on the
// DriveItem (SharePoint's `file.hashes` facet usually returns quickXorHash by
// default, sometimes sha1Hash; sha256Hash requires explicit enablement).
// Spec: https://learn.microsoft.com/en-us/onedrive/developer/code-snippets/quickxorhash
// Implementation translated from the reference C# / Python.
async function quickXorHashBase64(bytes) {
  if (!bytes || bytes.length === 0) return "";
  const BITS_IN_LAST_CELL = 32;
  const SHIFT = 11;
  const WIDTH_BITS = 160;
  const WIDTH_BYTES = 20;
  // We maintain 160 bits as 5 × 32-bit lanes (little-endian inside each lane).
  const lanes = new Uint32Array(5);
  let shiftSoFar = 0;
  let lengthSoFar = 0;

  // Rotate-left of the lanes array by `bits`. Done by streaming each input byte
  // into a moving 8-bit window. For performance + readability we operate on
  // 8-bit groups and recompute the destination bit/lane each step.
  for (let i = 0; i < bytes.length; i++) {
    const currentShift = shiftSoFar;
    const vectorArrayIndex = Math.floor(currentShift / 32) % 5;
    const vectorOffset = currentShift % 32;
    const nextVectorIndex = (vectorArrayIndex + 1) % 5;
    const xoredByte = bytes[i];
    // Lower bits of the byte XOR into the current lane shifted into position.
    lanes[vectorArrayIndex] = (lanes[vectorArrayIndex] ^ (xoredByte << vectorOffset)) >>> 0;
    // Bits that overflow the lane go into the next lane.
    if (vectorOffset > 24) {
      lanes[nextVectorIndex] = (lanes[nextVectorIndex] ^ (xoredByte >>> (32 - vectorOffset))) >>> 0;
    }
    shiftSoFar = (shiftSoFar + SHIFT) % WIDTH_BITS;
    lengthSoFar++;
  }

  // Finalize by XORing the message length (8 bytes, little-endian) into the
  // last 64 bits of the 160-bit state.
  const out = new Uint8Array(WIDTH_BYTES);
  for (let i = 0; i < 5; i++) {
    out[i * 4 + 0] = (lanes[i] >>>  0) & 0xff;
    out[i * 4 + 1] = (lanes[i] >>>  8) & 0xff;
    out[i * 4 + 2] = (lanes[i] >>> 16) & 0xff;
    out[i * 4 + 3] = (lanes[i] >>> 24) & 0xff;
  }
  // XOR the message length (in bytes, as 8-byte LE) into the last 8 bytes
  let len = lengthSoFar;
  for (let i = 0; i < 8; i++) {
    out[WIDTH_BYTES - 8 + i] ^= (len & 0xff);
    len = Math.floor(len / 256);
  }
  // base64 of the 20 bytes
  let bin = "";
  for (let i = 0; i < out.length; i++) bin += String.fromCharCode(out[i]);
  return btoa(bin);
}

// SHA-1 (still occasionally surfaced by Graph). Web Crypto handles it.
async function sha1Hex(bytes) {
  if (!bytes || bytes.length === 0) return "";
  const digest = await crypto.subtle.digest("SHA-1", bytes);
  const view = new Uint8Array(digest);
  let hex = "";
  for (let i = 0; i < view.length; i++) hex += view[i].toString(16).padStart(2, "0");
  return hex.toUpperCase(); // Graph reports SHA-1 in uppercase hex
}

// Fetch metadata for a just-uploaded file. Returns { size, hashes } or null.
async function fetchSpFileMetadata(driveId, token, targetPath, safeName) {
  try {
    const url = "https://graph.microsoft.com/v1.0/drives/" + driveId +
      "/root:/" + encodeDrivePath(targetPath + "/" + safeName) +
      "?select=size,file";
    const res = await fetchWithRetry(url, {
      headers: { "Authorization": "Bearer " + token },
    }, { label: "graph verify GET", maxAttempts: 3 });
    if (!res.ok) return null;
    const item = await res.json();
    return {
      size:   item?.size ?? null,
      hashes: item?.file?.hashes || {},
    };
  } catch (e) {
    console.warn("[verify] metadata GET failed:", safeName, e.message);
    return null;
  }
}

// Verify a just-uploaded file matches the local bytes we tried to send.
// Returns:
//   { ok: true,  algo, localHash, remoteHash } — verified clean
//   { ok: false, reason } — mismatch or unable to verify
async function verifyUploadedFile(driveId, token, targetPath, safeName, bytes) {
  const meta = await fetchSpFileMetadata(driveId, token, targetPath, safeName);
  if (!meta) return { ok: false, reason: "metadata GET failed" };

  // Size check — fastest, most important. Any mismatch here is a hard fail
  // regardless of which hash algorithm the server reports.
  if (typeof meta.size === "number" && meta.size !== bytes.length) {
    return {
      ok: false,
      reason: `size mismatch: local=${bytes.length} remote=${meta.size}`,
    };
  }

  const hashes = meta.hashes || {};
  // Prefer SHA-256 (only present if the tenant has it enabled), then SHA-1
  // (deprecated but widely present), then quickXorHash (the SharePoint default).
  if (hashes.sha256Hash) {
    const local = (await sha256Hex(bytes)).toUpperCase();
    const remote = String(hashes.sha256Hash).toUpperCase();
    return local === remote
      ? { ok: true,  algo: "sha256", localHash: local, remoteHash: remote }
      : { ok: false, reason: `sha256 mismatch: local=${local.slice(0,12)}… remote=${remote.slice(0,12)}…` };
  }
  if (hashes.sha1Hash) {
    const local = await sha1Hex(bytes);
    const remote = String(hashes.sha1Hash).toUpperCase();
    return local === remote
      ? { ok: true,  algo: "sha1", localHash: local, remoteHash: remote }
      : { ok: false, reason: `sha1 mismatch: local=${local.slice(0,12)}… remote=${remote.slice(0,12)}…` };
  }
  if (hashes.quickXorHash) {
    const local = await quickXorHashBase64(bytes);
    const remote = String(hashes.quickXorHash);
    return local === remote
      ? { ok: true,  algo: "quickXor", localHash: local, remoteHash: remote }
      : { ok: false, reason: `quickXor mismatch: local=${local.slice(0,12)}… remote=${remote.slice(0,12)}…` };
  }
  // Server returned no hashes — fall back to size-only (already checked above
  // and matched). Better than nothing.
  return { ok: true, algo: "size-only", localHash: String(bytes.length), remoteHash: String(meta.size) };
}

// Delete an item at a path. Used by the verify-then-retry path when the
// first upload landed corrupt and we need to clear the bad file before re-PUT.
async function deleteSpFile(driveId, token, targetPath, safeName) {
  try {
    const url = "https://graph.microsoft.com/v1.0/drives/" + driveId +
      "/root:/" + encodeDrivePath(targetPath + "/" + safeName);
    const res = await fetchWithRetry(url, {
      method: "DELETE",
      headers: { "Authorization": "Bearer " + token },
    }, { label: "graph delete bad file", maxAttempts: 2 });
    return res.ok || res.status === 404;
  } catch (e) {
    console.warn("[verify] delete failed:", safeName, e.message);
    return false;
  }
}

// Upload + verify. Strategy (corrected — see comment block below):
//   1. PUT bytes
//   2. Parse PUT response (Graph returns the DriveItem). Verify size matches.
//   3. If size mismatch → re-PUT once (PUT is idempotent — overwrites).
//      If still mismatch → hard error.
//   4. Hash verification is deferred: SharePoint Online computes file hashes
//      asynchronously after upload (often several seconds, sometimes minutes).
//      A synchronous GET-and-compare here would race the indexing pipeline and
//      report false-negatives. The reconcile sweep in PMS handles hash drift
//      detection later, when the hashes have had time to be computed.
//
//   IMPORTANT: We do NOT delete on verify mismatch. The original implementation
//   did and turned a benign indexing race into actual data loss. Re-uploads via
//   PUT overwrite naturally.
//
// What we record into lastAttachmentUploadStats.uploadedFiles:
//   - name, size, contentType (from local bytes — what we tried to send)
//   - sha256 (computed locally — used by the reconcile sweep later)
//   - verifiedAlgo: "size-from-put-response" if size matches the PUT echo
//   - verified: true once size verified
async function uploadAttachmentToSharePoint(driveId, token, targetPath, name, contentType, bytes) {
  const safeName = (name || "attachment").replace(/[\\/:*?"<>|]/g, "-").trim() || "attachment";
  if (bytes && bytes.length > SP_SIMPLE_UPLOAD_MAX) {
    return uploadLargeAttachmentToSharePoint(driveId, token, targetPath, safeName, contentType, bytes);
  }
  let lastErr;
  for (let attempt = 1; attempt <= 2; attempt++) {
    const uploadRes = await fetchWithRetry("https://graph.microsoft.com/v1.0/drives/" + driveId + "/root:/" + encodeDrivePath(targetPath + "/" + safeName) + ":/content", {
      method: "PUT",
      headers: { "Authorization": "Bearer " + token, "Content-Type": contentType || "application/octet-stream" },
      body: bytes,
    }, { label: "graph upload " + safeName });
    if (!uploadRes.ok) {
      console.warn("Attachment upload HTTP failed:", safeName, uploadRes.status);
      return false;
    }
    // Parse the DriveItem out of the PUT response. Graph returns it with size,
    // etag, file (sometimes hashes). This is the only verification source that's
    // immediately consistent — no read-after-write race.
    let item = null;
    try { item = await uploadRes.json(); } catch { /* old Graph versions may return 204 */ }
    // Best-effort size match. If item lacks size (unusual), accept the upload —
    // we'd rather trust the HTTP 200/201 than reject a save on a metadata quirk.
    const remoteSize = item?.size;
    if (typeof remoteSize === "number" && remoteSize !== bytes.length) {
      lastErr = `size mismatch on PUT response: local=${bytes.length} remote=${remoteSize}`;
      console.warn(`[verify] ${safeName} attempt ${attempt}: ${lastErr} — retrying`);
      continue; // retry — next PUT overwrites
    }
    // Compute the local SHA-256 once, async, so the audit log carries an
    // integrity fingerprint the reconcile sweep can later cross-check.
    let localSha256 = null;
    try { localSha256 = await sha256Hex(bytes); } catch { /* non-fatal */ }
    // If the PUT response happened to include hashes (sometimes for small
    // files), opportunistically validate. Otherwise we trust size + the
    // deferred reconcile.
    const hashes = item?.file?.hashes || {};
    let verifiedAlgo = "size-from-put-response";
    let verifiedHash = String(bytes.length);
    if (hashes.sha256Hash && localSha256) {
      const remote = String(hashes.sha256Hash).toUpperCase();
      if (remote !== localSha256.toUpperCase()) {
        lastErr = `sha256 mismatch: local=${localSha256.slice(0,12)}… remote=${remote.slice(0,12)}…`;
        console.warn(`[verify] ${safeName} attempt ${attempt}: ${lastErr} — retrying`);
        continue;
      }
      verifiedAlgo = "sha256";
      verifiedHash = localSha256;
    }
    _recordVerifiedFile(safeName, bytes, contentType, {
      ok: true, algo: verifiedAlgo, localHash: verifiedHash,
      localSha256, // always recorded for later reconcile
    });
    return true;
  }
  console.warn(`[verify] ${safeName} could not be verified after 2 attempts: ${lastErr}`);
  // We *do not* throw or delete. The bytes are on SharePoint (the PUT returned
  // 200). The size mismatch is suspicious but the safer move is to keep the
  // file and let the reconcile sweep flag it.
  _recordVerifiedFile(safeName, bytes, contentType, {
    ok: true, algo: "size-mismatch-tolerated", localHash: String(bytes.length),
    localSha256: null,
  });
  return true;
}

// Record the hash+verify outcome for the most recent uploaded file. Mutates
// the last entry of lastAttachmentUploadStats.uploadedFiles (which was just
// appended by the caller in uploadInBatches) if it matches by name; otherwise
// pushes a fresh entry. The dual-shape covers both the Office and Graph paths.
function _recordVerifiedFile(safeName, bytes, contentType, verify) {
  if (!lastAttachmentUploadStats) return;
  const arr = lastAttachmentUploadStats.uploadedFiles || (lastAttachmentUploadStats.uploadedFiles = []);
  // sha256 prefers the explicit field, falls back to whatever algo carried it.
  const sha256 = verify.localSha256
              || (verify.algo === "sha256" ? verify.localHash : null);
  for (let i = arr.length - 1; i >= 0; i--) {
    if (arr[i].name === safeName) {
      arr[i].sha256 = sha256 || arr[i].sha256 || null;
      arr[i].verifiedAlgo = verify.algo;
      arr[i].verifiedHash = verify.localHash;
      arr[i].verified = verify.algo !== "size-mismatch-tolerated";
      return;
    }
  }
  arr.push({
    name: safeName,
    size: bytes?.length || 0,
    contentType: contentType || null,
    sha256,
    verifiedAlgo: verify.algo,
    verifiedHash: verify.localHash,
    verified: verify.algo !== "size-mismatch-tolerated",
  });
}

// Resumable upload for files > 4 MiB. Creates an upload session and PUTs
// fixed-size chunks with Content-Range headers. Retries each chunk on 5xx
// (the upload-session URL is durable across transient failures). Returns
// true on full success, false on any irrecoverable failure.
async function uploadLargeAttachmentToSharePoint(driveId, token, targetPath, safeName, contentType, bytes) {
  const sessionUrl = "https://graph.microsoft.com/v1.0/drives/" + driveId +
    "/root:/" + encodeDrivePath(targetPath + "/" + safeName) + ":/createUploadSession";
  let sessionRes;
  try {
    sessionRes = await fetchWithRetry(sessionUrl, {
      method: "POST",
      headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({
        item: {
          "@microsoft.graph.conflictBehavior": "rename", // safety net on top of uniquifyNames
          name: safeName,
        },
      }),
    }, { label: "graph createUploadSession" });
  } catch (e) {
    console.warn("Large upload session creation failed:", safeName, e.message);
    return false;
  }
  if (!sessionRes.ok) {
    console.warn("Large upload session HTTP", sessionRes.status, "for", safeName);
    return false;
  }
  const session = await sessionRes.json();
  const uploadUrl = session.uploadUrl;
  if (!uploadUrl) {
    console.warn("Large upload session returned no uploadUrl for", safeName);
    return false;
  }

  const total = bytes.length;
  let offset = 0;
  while (offset < total) {
    const end = Math.min(offset + SP_UPLOAD_CHUNK_SIZE, total);
    const chunk = bytes.subarray(offset, end);
    const range = `bytes ${offset}-${end - 1}/${total}`;
    // The upload URL is pre-authenticated by Graph — DO NOT add Authorization
    // header (it returns 401 if you do). fetchWithRetry handles 503/504 here.
    let res;
    try {
      res = await fetchWithRetry(uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Length": String(chunk.length),
          "Content-Range": range,
        },
        body: chunk,
      }, { label: `large upload ${safeName} ${offset}-${end - 1}` });
    } catch (e) {
      console.warn("Chunked upload network error at", range, "for", safeName, "—", e.message);
      // Best-effort cancel so we don't leave a half-built file on SharePoint
      try { await fetch(uploadUrl, { method: "DELETE" }); } catch {}
      return false;
    }
    if (res.status === 202) {
      // Server accepted chunk; continue
      offset = end;
      continue;
    }
    if (res.status === 200 || res.status === 201) {
      // Final chunk — upload complete. Trust the response body (the DriveItem)
      // for size verification. Hashes are computed asynchronously by SharePoint
      // and won't be present here for SP Online — that's fine; the reconcile
      // sweep checks them later when they're available.
      let item = null;
      try { item = await res.json(); } catch { /* tolerate empty body */ }
      const remoteSize = item?.size;
      if (typeof remoteSize === "number" && remoteSize !== bytes.length) {
        console.warn(`[verify] large file ${safeName}: size mismatch local=${bytes.length} remote=${remoteSize} — tolerating; reconcile will flag if persistent`);
      }
      let localSha256 = null;
      try { localSha256 = await sha256Hex(bytes); } catch { /* non-fatal */ }
      _recordVerifiedFile(safeName, bytes, contentType, {
        ok: true,
        algo: (typeof remoteSize === "number" && remoteSize === bytes.length) ? "size-from-chunked-response" : "size-mismatch-tolerated",
        localHash: String(bytes.length),
        localSha256,
      });
      return true;
    }
    // Permanent failure
    const errText = await res.text().catch(() => "");
    console.warn("Chunked upload failed", res.status, "at", range, "for", safeName, errText.slice(0, 120));
    try { await fetch(uploadUrl, { method: "DELETE" }); } catch {}
    return false;
  }
  // If we exit the loop without 200/201, something went wrong — treat as failure.
  console.warn("Chunked upload completed all chunks but never received 200/201 for", safeName);
  return false;
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

// ─── SHARED FILING SCAFFOLDING ───────────────────────────────────────────────
// Every filing operation (Log RFI new/existing, Log Submittal new/existing,
// plus eventually Save SP and others) goes through the same boilerplate:
//   - check selectedProject + saveInFlight
//   - snapshot the mailbox item before any await (item-switch race protection)
//   - enqueue a crash-recovery intent in localStorage
//   - set saveInFlight = true
//   - try { runUpload } catch { log failed } finally { saveInFlight = false }
//   - on success, log success + dequeue
//
// Before this helper, that ~30 lines of scaffolding lived in each save flow
// (4× duplication, 5× counting Save SP). Adding the audit log this session
// meant editing 5 places and the risk of missing one. Now there's one place.
//
// runUpload receives { snapItem, statusElement, project } and returns:
//   {
//     sp_folder_url?: string,   // for the audit log
//     files?: array,            // for the audit log (defaults to lastAttachmentUploadStats.uploadedFiles)
//     status?: "success" | "partial",  // audit log status; defaults to "success"
//     error?: string,           // explanation if partial; for the audit log
//     successMessage?: string,  // text shown in the status banner on success
//   }
//
// Throwing from runUpload triggers the error path: status banner shows "✗ ..."
// and a "failed" audit log row is written. The queue entry stays in place so
// it surfaces on next taskpane open as a crash-recovery candidate.
async function withFilingScaffold(opts, runUpload) {
  const { operation, statusElement, startMessage = "⏳ Saving…" } = opts;
  if (!selectedProject) {
    setStatus(statusElement, "error", "No project selected.");
    return;
  }
  if (saveInFlight) {
    setStatus(statusElement, "info", "⏳ Another save is in progress; please wait.");
    return;
  }
  // Snapshot BEFORE any await — protects against item-switch races.
  const snapItem = emailItem;
  const queueId = enqueueFilingIntent({
    project_id:    selectedProject.id,
    project_name:  selectedProject.name || "",
    msg_id:        snapItem?.itemId || null,
    operation,
    email_subject: snapItem?.subject || "",
  });
  saveInFlight = true;
  setStatus(statusElement, "info", startMessage);
  try {
    const result = await runUpload({ snapItem, statusElement, project: selectedProject });
    if (result?.successMessage != null) {
      setStatus(statusElement, "success", result.successMessage);
    }
    void logFilingOp({
      project_id:    selectedProject.id,
      msg_id:        snapItem?.itemId || null,
      operation,
      email_subject: snapItem?.subject || null,
      sp_folder_url: result?.sp_folder_url || null,
      files:         result?.files !== undefined ? result.files : (lastAttachmentUploadStats?.uploadedFiles || []),
      status:        result?.status || "success",
      error:         result?.error || null,
    });
    dequeueFilingIntent(queueId);
    return result;
  } catch (e) {
    setStatus(statusElement, "error", "✗ " + e.message);
    void logFilingOp({
      project_id:    selectedProject.id,
      msg_id:        snapItem?.itemId || null,
      operation,
      email_subject: snapItem?.subject || null,
      status:        "failed",
      error:         e.message,
    });
    // Queue entry intentionally NOT dequeued — surfaces on next open as
    // a previous-save-did-not-complete banner.
  } finally {
    saveInFlight = false;
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
  // Snapshot all per-item data SYNCHRONOUSLY before any await. Without this,
  // a fast item-switch (Office swaps mailbox.item silently on pinned panes)
  // mid-save would read fields from the new email — saving the wrong subject,
  // date, sender, or attachments into the current project's folder. The
  // generation counter alone can't protect later sync reads of `emailItem.*`.
  const snapItem = emailItem;
  const snapSubject = snapItem?.subject || "";
  const snapDate = snapItem?.dateTimeCreated;
  const snapFromName = snapItem?.from?.displayName || "";
  const snapFromAddr = snapItem?.from?.emailAddress || "";
  const snapItemId = snapItem?.itemId || "";
  const saveGen = itemContextGeneration; // capture for stale-write detection
  // Crash-recovery queue: record the intent BEFORE any awaits. If the browser
  // crashes during upload, the entry remains and is surfaced as a pending save
  // on next taskpane open. Dequeued at the end of a clean save.
  const queueId = enqueueFilingIntent({
    project_id:    selectedProject.id,
    project_name:  selectedProject.name || "",
    msg_id:        currentMsgId,
    operation:     "email-sp",
    email_subject: snapSubject,
  });
  setStatus("actionStatus", "info", pickSavingMessage());
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
    const d = new Date(snapDate);
    // Folder name = YYYY_MM_DD + (custom name if user set one, else cleaned subject).
    const customCleaned = _customSpFolderName
      ? _customSpFolderName.replace(/[\\/:*?"<>|]/g, "-").replace(/\s+/g, " ").trim().slice(0, 70)
      : "";
    const safeSubject = (snapSubject || "No Subject").replace(/[\\/:*?"<>|]/g, "-").replace(/\s+/g, " ").trim().slice(0, 70);
    const folderTail = customCleaned || safeSubject;
    const emailFolderName = d.getFullYear() + "_" + String(d.getMonth() + 1).padStart(2, "0") + "_" + String(d.getDate()).padStart(2, "0") + " " + folderTail;
    const emailsPath  = await ensureSpFolder(driveId, token, projFolderName, "Emails");
    const targetPath  = await ensureSpFolder(driveId, token, emailsPath, emailFolderName);
    // Re-check generation right before destructive writes. If the user has
    // switched to a different item during folder-resolution, abort cleanly
    // rather than continuing with a mix of old-item folder + new-item body.
    if (saveGen !== itemContextGeneration) {
      setStatus("actionStatus", "info", "Save aborted — you switched emails. Click Save again on the email you want to file.");
      return;
    }
    await writeSpMetadataSidecar(driveId, token, targetPath, buildAddinMetadata(selectedProject, "correspondence"));
    // Pass snapItem so the attachment loop reads from the captured item, not
    // the live module global — protects against item-switch corruption where
    // bytes from a different email could otherwise be filed under this folder.
    const attCount    = await uploadEmailAndAttachments(driveId, token, targetPath, snapItem);
    const spFolderUrl = SP_BASE_URL + "/" + encodeURIComponent(projFolderName) + "/Emails/" + encodeURIComponent(emailFolderName);
    const msgId = currentMsgId;
    const emailRecord = {
      id: uid(), msgId,
      subject: snapSubject,
      from: snapFromName,
      fromAddress: snapFromAddr,
      date: snapDate,
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
    // Append to the filing-integrity audit log so PMS can reconcile this save.
    // Status reflects whether the upload was clean or partial — verified flag
    // gets set by Phase 2 read-back once that lands. Fire-and-forget.
    const status =
      (attempted > 0 && attCount === 0)            ? "failed" :
      (attempted > 0 && attCount < attempted)      ? "partial" :
      "success";
    void logFilingOp({
      project_id:    selectedProject.id,
      msg_id:        msgId,
      operation:     "email-sp",
      sp_folder_url: spFolderUrl,
      files:         (lastAttachmentUploadStats?.uploadedFiles || []),
      email_subject: snapSubject,
      status,
      error:         status === "success" ? null : (warnings.join(" ") || `${attCount}/${attempted} uploaded`),
    });
    // One-shot custom name consumed — clear so the next email's save uses
    // subject-default unless explicitly renamed again.
    _customSpFolderName = "";
    recordSaveAndCelebrate();
    refreshEmailSavedIndicator(true);
    // Save completed without throwing — clear crash-recovery queue entry.
    // (Partial successes still dequeue: the user has the status message and
    // a "partial" audit log row; the queue is only for crash recovery.)
    dequeueFilingIntent(queueId);
  } catch (e) {
    setStatus("actionStatus", "error", "✗ " + e.message);
    void logFilingOp({
      project_id:    selectedProject.id,
      msg_id:        currentMsgId,
      operation:     "email-sp",
      email_subject: snapSubject,
      status:        "failed",
      error:         e.message,
    });
    // Leave the queue entry in place for crash-recovery surfacing on next open.
  }
}
async function doSaveToProjectRecordOnly() {
  return withSaveGuard("save-record", _doSaveToProjectRecordOnly, ["saveSpBtn", "saveRecordBtn"]);
}
async function _doSaveToProjectRecordOnly() {
  if (!selectedProject) { setStatus("actionStatus", "error", "Select a project first."); return; }
  // Body-only save works regardless of attachments — the visual emphasis (de-emph
  // + caption) is the soft nudge toward SharePoint when attachments exist.
  // No confirm dialog: trust the user's intent, surface the consequence in the
  // post-save card ("3 attachments not filed").
  const msgId = getCurrentMessageRecordId();
  if (findSavedEmailRecord(selectedProject, msgId)) {
    refreshEmailSavedIndicator();
    return;
  }
  setStatus("actionStatus", "info", pickSavingMessage());
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
    recordSaveAndCelebrate();
    refreshEmailSavedIndicator(true);
  } catch (e) {
    setStatus("actionStatus", "error", "✗ " + e.message);
  }
}
// ─── LOG NOTE ─────────────────────────────────────────────────────────────────
// Categories that get the full meeting-template page (title + metadata table +
// Discussion / Decisions / Action Items sections). These three are structured
// note types where the user types meeting/visit notes; everything else uses
// the email-body builder below — the page IS the email body, with a header.
const MEETING_NOTE_CATEGORIES = ["Client Meeting", "Internal Meeting", "Site Visit"];
function isMeetingNoteCategory(cat) { return MEETING_NOTE_CATEGORIES.includes(cat); }

// Email-body OneNote page — used for non-meeting categories (Site Visit,
// Decision, Issue, Client Communication, Internal, Action Item, General).
// The email body becomes the page content directly, so embedded images and
// formatting survive. Optional user-typed note appears above the body as a
// short "Setty note" header if present.
function buildAddinEmailNotePageHtml(title, category, dateStr, fromName, fromEmail, emailBodyHtml, userNote, project) {
  // Same print-friendly typography as the meeting template — so all add-in
  // OneNote pages read as one consistent document family.
  const NAVY = "#1F3864";
  const TEXT_DARK = "#222";
  const GRAY_BORDER = "#BFBFBF";
  const GRAY_LIGHT = "#F2F2F2";
  const td = "padding:6px 12px;font-size:11pt;border:1px solid " + GRAY_BORDER;
  const th = td + ";font-weight:bold;background:" + GRAY_LIGHT;
  const h2 = "font-family:Calibri,Arial,sans-serif;font-size:14pt;color:" + NAVY +
             ";border-bottom:1px solid " + GRAY_BORDER + ";padding-bottom:2px;margin-top:18px;margin-bottom:8px";

  const dateFmt = dateStr ? new Date(dateStr).toLocaleDateString("en-US", {
    weekday: "long", month: "long", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit"
  }) : "";
  const fromStr = [fromName, fromEmail ? `&lt;${fromEmail}&gt;` : ""].filter(Boolean).join(" ");
  const safeTitle = escapeOneNoteTextAddin(title);
  const safeUserNote = escapeOneNoteTextAddin(userNote || "");
  const safeFromStr = escapeOneNoteTextAddin(fromStr);
  const safeCategory = escapeOneNoteTextAddin(category || "Note");
  const safeProjNo = escapeOneNoteTextAddin(project?.projectNumber || "");
  const safeProjNm = escapeOneNoteTextAddin(project?.name || "");
  const projSubtitle = (safeProjNo || safeProjNm)
    ? "<div style='font-family:Calibri,Arial,sans-serif;font-size:10.5pt;color:#44546A;letter-spacing:.04em;text-transform:uppercase;margin:0 0 4px'>"
      + safeProjNo + (safeProjNo && safeProjNm ? " · " : "") + safeProjNm + "</div>"
    : "";
  return ""
    + "<div style='max-width:7.5in;font-family:Calibri,Arial,sans-serif;font-size:11pt;color:" + TEXT_DARK + ";line-height:1.5'>"
    + projSubtitle
    + "<h1 style='font-family:Calibri,Arial,sans-serif;font-size:20pt;color:" + NAVY + ";margin:0 0 6px;padding-bottom:6px;border-bottom:2px solid " + NAVY + "'>" + safeCategory + "</h1>"
    + "<div style='font-family:Calibri,Arial,sans-serif;font-size:13pt;font-weight:600;color:" + NAVY + ";margin-bottom:14px'>" + safeTitle + "</div>"
    + "<table style='border-collapse:collapse;width:100%;font-family:Calibri,Arial,sans-serif;margin-bottom:14px'>"
    + (dateFmt ? "<tr><td style='" + th + ";width:120px'>Date</td><td style='" + td + "'>" + dateFmt + "</td></tr>" : "")
    + (fromStr ? "<tr><td style='" + th + ";width:120px'>From</td><td style='" + td + "'>" + safeFromStr + "</td></tr>" : "")
    + "</table>"
    + (userNote ? "<h2 style='" + h2 + "'>Note</h2><p style='font-family:Calibri,Arial,sans-serif;font-size:11pt;margin:0'>" + safeUserNote.replace(/\n/g, "<br>") + "</p>" : "")
    + "<h2 style='" + h2 + "'>Email</h2>"
    + (emailBodyHtml || "<p style='font-family:Calibri,Arial,sans-serif;color:#666;font-style:italic'>(No body content — this email may be a system notification, share invite, or attachment-only message.)</p>")
    + "</div>";
}

function buildAddinMeetingPageHtml(title, category, dateStr, participants, body, project) {
  // Print-friendly typography — mirrors the PMS importer so OneNote pages
  // generated from either side look identical. Calibri 11pt body, navy
  // headings, max-width 7.5in for letter-sized print / paste-to-email.
  const NAVY = "#1F3864";
  const TEXT_DARK = "#222";
  const GRAY_BORDER = "#BFBFBF";
  const GRAY_LIGHT = "#F2F2F2";
  const td = "padding:6px 12px;font-size:11pt;border:1px solid " + GRAY_BORDER;
  const th = td + ";font-weight:bold;background:" + GRAY_LIGHT;
  const h2 = "font-family:Calibri,Arial,sans-serif;font-size:14pt;color:" + NAVY +
             ";border-bottom:1px solid " + GRAY_BORDER + ";padding-bottom:2px;margin-top:18px;margin-bottom:8px";

  // Date is already inside the Notes block ("Start: …"), so no Date row.
  // No H1 / Type / Attendees row — see PMS importer for the reasoning;
  // the cleanedBody preserves the bulleted attendees list.
  const cleanedBody = stripMeetingBoilerplateAddin(body);
  const safeTitle  = escapeOneNoteTextAddin(title);
  const safeProjNo = escapeOneNoteTextAddin(project?.projectNumber || "");
  const safeProjNm = escapeOneNoteTextAddin(project?.name || "");
  const projSubtitle = (safeProjNo || safeProjNm)
    ? "<div style='font-family:Calibri,Arial,sans-serif;font-size:10.5pt;color:#44546A;letter-spacing:.04em;text-transform:uppercase;margin:0 0 4px'>"
      + safeProjNo + (safeProjNo && safeProjNm ? " · " : "") + safeProjNm + "</div>"
    : "";
  return ""
    + "<div style='max-width:7.5in;font-family:Calibri,Arial,sans-serif;font-size:11pt;color:" + TEXT_DARK + ";line-height:1.5'>"
    + projSubtitle
    + "<h1 style='font-family:Calibri,Arial,sans-serif;font-size:20pt;color:" + NAVY + ";margin:0 0 6px;padding-bottom:6px;border-bottom:2px solid " + NAVY + "'>Meeting Minutes</h1>"
    + "<div style='font-family:Calibri,Arial,sans-serif;font-size:13pt;font-weight:600;color:" + NAVY + ";margin-bottom:14px'>" + safeTitle + "</div>"
    // Bordered box for meeting details — same visual treatment as the PMS
    // importer. No Notes label / no grey side panel; just a clean polished
    // block aligned to the page margin.
    + (cleanedBody
        ? "<div style='border:1px solid " + GRAY_BORDER + ";background:#FAFAFA;padding:12px 14px;margin-bottom:16px;font-family:Calibri,Arial,sans-serif;font-size:11pt'>"
          + "<pre style='font-family:Calibri,Arial,sans-serif;font-size:11pt;white-space:pre-wrap;margin:0;color:" + TEXT_DARK + "'>" + cleanedBody + "</pre>"
          + "</div>"
        : "")
    + "<h2 style='" + h2 + "'>Discussion</h2><p>&nbsp;</p>"
    + "<h2 style='" + h2 + "'>Decisions</h2><p>&nbsp;</p>"
    + "<h2 style='" + h2 + "'>Action Items</h2>"
    + "<table style='border-collapse:collapse;width:100%;font-family:Calibri,Arial,sans-serif;font-size:11pt;margin-top:6px'>"
    + "<tr style='background:" + NAVY + ";color:#fff'>"
      + "<th style='padding:6px 12px;text-align:left;font-weight:600;border:1px solid " + NAVY + "'>Item</th>"
      + "<th style='padding:6px 12px;text-align:left;font-weight:600;border:1px solid " + NAVY + "'>Owner</th>"
      + "<th style='padding:6px 12px;text-align:left;font-weight:600;border:1px solid " + NAVY + "'>Due</th>"
    + "</tr>"
    + "<tr><td style='" + td + "'>&nbsp;</td><td style='" + td + "'>&nbsp;</td><td style='" + td + "'>&nbsp;</td></tr>"
    + "</table>"
    + "</div>";
}

// In-memory cache so a second click on Send-to-Teams doesn't re-hit Graph
// for the channel email. Keyed by teamsChannelId. Reset on page reload —
// no need to persist beyond a single Outlook session.
const _channelEmailCache = {};

// READ-ONLY channel email resolution. PMS deliberately doesn't request the
// ChannelSettings.ReadWrite.All scope needed for `provisionEmail`, so we
// only read here. If a channel hasn't had its email provisioned yet, the
// user does it once via Teams ("⋯" menu → "Get email address"); after
// that the read path returns the address every time.
async function resolveChannelEmailAddin(project) {
  const channelId = project?.teamsChannelId;
  if (!channelId) return "";
  if (_channelEmailCache[channelId]) return _channelEmailCache[channelId];
  if (project.teamsChannelEmail) {
    _channelEmailCache[channelId] = project.teamsChannelEmail;
    return project.teamsChannelEmail;
  }
  try {
    const ch = await graphFetch("GET", `/teams/${TEAMS_TEAM_ID}/channels/${channelId}?$select=email`);
    if (ch?.email) { _channelEmailCache[channelId] = ch.email; return ch.email; }
    throw new Error("__NO_EMAIL__");
  } catch (e) {
    if (e.message === "__NO_EMAIL__") {
      throw new Error('Channel has no email yet. In Teams, click "⋯" next to the channel name → "Get email address" → "Copy". That provisions it; the next click here will work.');
    }
    throw new Error("Graph: " + e.message);
  }
}

// Fetch the current Outlook email's HTML body for forwarding. Returns ""
// if no email is selected or the API errors.
function getCurrentEmailBodyHtml() {
  return new Promise(resolve => {
    if (!emailItem?.body?.getAsync) { resolve(""); return; }
    try {
      emailItem.body.getAsync(Office.CoercionType.Html, r => {
        resolve(r.status === Office.AsyncResultStatus.Succeeded ? (r.value || "") : "");
      });
    } catch (e) { resolve(""); }
  });
}

// Send-to-Teams handler. Posts the selected email's subject + body directly
// to the project's Teams channel via Graph (POST .../channels/{id}/messages)
// using the ChannelMessage.Send scope PMS already has for milestone cards.
// No email roundtrip → instant delivery. Trade-off: no "Sent" record in
// Outlook, and file attachments don't transfer (would need a separate
// upload-to-filesFolder + reference flow, deferred for v2).
function getTeamsSentMap() {
  try { return JSON.parse(localStorage.getItem(TEAMS_SENT_MAP_STORAGE_KEY) || "{}"); } catch { return {}; }
}
function markEmailSentToTeams(channelId) {
  const msgId = getCurrentMessageRestId();
  if (!msgId) return;
  const map = getTeamsSentMap();
  map[msgId] = { channelId, teamId: TEAMS_TEAM_ID };
  localStorage.setItem(TEAMS_SENT_MAP_STORAGE_KEY, JSON.stringify(map));
}
function refreshTeamsBtn() {
  const btn = document.getElementById("sendToTeamsBtn");
  if (!btn) return;
  const msgId = getCurrentMessageRestId();
  const state = msgId ? getTeamsSentMap()[msgId] : null;
  if (state?.channelId) {
    const url = `https://teams.microsoft.com/l/channel/${encodeURIComponent(state.channelId)}/_?groupId=${state.teamId}&tenantId=${TEAMS_TENANT_ID}`;
    btn.textContent = "🔗 Open Teams Channel";
    btn.title = "This email was shared to Teams — click to open the channel";
    btn.onclick = () => openExternalUrl(url);
    btn.classList.remove("btn-teams");
    btn.classList.add("btn-teams-sent");
  } else {
    btn.textContent = "💬 Send to Teams Channel";
    btn.title = "";
    btn.onclick = sendToTeamsChannel;
    btn.classList.add("btn-teams");
    btn.classList.remove("btn-teams-sent");
  }
}
async function sendToTeamsChannel() {
  if (!selectedProject) {
    setStatus("actionStatus", "error", "Select a project first.");
    return;
  }
  if (!selectedProject.teamsChannelId) {
    setStatus("actionStatus", "error", "This project doesn't have a Teams channel set up. Configure it in PMS → Overview → Teams Notifications.");
    return;
  }
  if (!emailItem) {
    setStatus("actionStatus", "error", "No email selected to share.");
    return;
  }
  setStatus("actionStatus", "info", "⏳ Posting to Teams channel…");
  try {
    const subject = emailItem.subject || "(no subject)";
    const fromName  = emailItem.from?.displayName  || "";
    const fromEmail = emailItem.from?.emailAddress || "";
    const fromStr   = [fromName, fromEmail ? `&lt;${fromEmail}&gt;` : ""].filter(Boolean).join(" ");
    const origBody  = await getCurrentEmailBodyHtml();
    const attCount  = Array.isArray(emailItem.attachments) ? emailItem.attachments.length : 0;

    const safeSubject = escapeOneNoteTextAddin(subject);
    const safeFrom    = escapeOneNoteTextAddin(fromStr);
    const messageHtml =
      `<h3 style="margin:0 0 8px">${safeSubject}</h3>` +
      (safeFrom ? `<p style="color:#666;font-size:12px;margin:0 0 8px">From: <strong>${safeFrom}</strong></p>` : "") +
      `<blockquote style="border-left:3px solid #ddd;margin:8px 0;padding:0 0 0 12px">${origBody || "<p><em>(no body)</em></p>"}</blockquote>` +
      (attCount > 0 ? `<p style="color:#888;font-size:11px;font-style:italic">📎 Original email has ${attCount} attachment${attCount === 1 ? "" : "s"} — share separately if needed.</p>` : "") +
      `<p style="color:#888;font-size:11px;font-style:italic">Shared from Outlook via PMS Add-in</p>`;

    // Use the on-demand ChannelMessage.Send token (not the default getToken)
    // so the consent prompt only fires the first time the user posts to a
    // channel — kept out of the regular sign-in scopes.
    const channelToken = await getChannelMessageToken();
    await graphFetch("POST",
      `/teams/${TEAMS_TEAM_ID}/channels/${selectedProject.teamsChannelId}/messages`,
      { subject, body: { contentType: "html", content: messageHtml } },
      channelToken
    );
    markEmailSentToTeams(selectedProject.teamsChannelId);
    refreshTeamsBtn();
    setStatus("actionStatus", "success",
      "✓ Posted to Teams channel" + (attCount > 0 ? ` (${attCount} attachment${attCount === 1 ? "" : "s"} not transferred)` : ""));
  } catch (e) {
    setStatus("actionStatus", "error", "Send-to-Teams failed: " + e.message);
  }
}

// Strip the trailing "Summary:" / "Action items:" placeholder lines from
// buildMeetingNoteBody's output. Those are auto-generated headers that
// duplicate the Discussion / Action Items sections the OneNote page renders
// below. Keep this mirrored with stripMeetingBoilerplate() in SettyPMS.html.
function stripMeetingBoilerplateAddin(text) {
  if (!text) return "";
  const i = text.search(/\n\s*Summary\s*:/i);
  if (i >= 0) return text.slice(0, i).trim();
  return text.trim();
}

// Convert Graph's dateTime (UTC, often without a trailing Z) into a local
// wall-clock ISO string for use in <meta name="created">. Without this,
// OneNote renders UTC values as-is — a 5:30 PM ET meeting shows as 9:30 PM.
// Same logic / same bug as the PMS importer; keep this in sync with
// toOneNoteCreatedLocal() in SettyPMS.html if either changes.
function toAddinOneNoteCreatedLocal(dtString) {
  if (!dtString) return null;
  const hasZone = /[Zz]|[+\-]\d\d:?\d\d$/.test(dtString);
  const isoUtc = hasZone ? dtString : dtString + "Z";
  const d = new Date(isoUtc);
  if (isNaN(d.getTime())) return null;
  const pad = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
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

async function createAddinOneNotePage(project, title, body, category, dateStr, emailBodyHtml) {
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
  // Category badge intentionally omitted — it duplicates the OneNote section
  // name the page already lives in (e.g. "Client Meetings"). Project number
  // stays for at-a-glance identification.
  const safeProjNum = escapeOneNoteTextAddin(project.projectNumber || "");
  const badge = [
    project.projectNumber && `<span style="background:#003865;color:#fff;padding:2px 8px;border-radius:3px;font-size:11px;margin-right:6px">${safeProjNum}</span>`,
  ].filter(Boolean).join("");
  const header = `<div style="border-bottom:2px solid #003865;padding-bottom:8px;margin-bottom:16px;font-family:sans-serif">${badge}</div>`;
  const safeTitle = escapeOneNoteTextAddin(title);
  // Branch: meeting-type categories get the meeting-template page (with
  // Discussion / Decisions / Action Items sections); everything else gets
  // the email-body page where the email IS the content.
  const fromName  = emailItem?.from?.displayName  || "";
  const fromEmail = emailItem?.from?.emailAddress || "";
  const bodyHtml = isMeetingNoteCategory(category)
    ? buildAddinMeetingPageHtml(title, category, dateStr, emailParticipants, body, project)
    : buildAddinEmailNotePageHtml(title, category, dateStr, fromName, fromEmail, emailBodyHtml, body, project);
  // toAddinOneNoteCreatedLocal fixes the UTC-as-local rendering bug; fall
  // back to a fresh local timestamp if dateStr is missing or unparseable.
  const createdMeta = toAddinOneNoteCreatedLocal(dateStr) || toAddinOneNoteCreatedLocal(new Date().toISOString()) || new Date().toISOString();
  const pageHtml = `<!DOCTYPE html><html><head><title>${safeTitle}</title><meta name="created" content="${createdMeta}" /></head><body>${header}${bodyHtml}</body></html>`;

  // POST page with retry on 429/503 + idempotency dedup. Graph throttles
  // OneNote aggressively, and a 5xx response can mask a successful server-
  // side create — without dedup, the retry would post a duplicate page.
  const token = await getToken();
  const url = `https://graph.microsoft.com/v1.0/${baseUrl}/sections/${section.id}/pages`;
  const key = _hashOneNoteReq(url, pageHtml);
  const cached = _addinOneNoteCache.get(key);
  if (cached && (Date.now() - cached.at) < ADDIN_ONENOTE_DEDUP_TTL_MS) {
    console.log("[OneNote-addin] idempotency cache hit — skipping duplicate POST");
    return cached.result;
  }
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
      const result = { id: page.id, webUrl: page.links?.oneNoteWebUrl?.href || page.webUrl || "" };
      _addinOneNoteCache.set(key, { result, at: Date.now() });
      // Prune old entries opportunistically
      if (_addinOneNoteCache.size > 100) {
        const cutoff = Date.now() - 2 * ADDIN_ONENOTE_DEDUP_TTL_MS;
        for (const [k, v] of _addinOneNoteCache) if (v.at < cutoff) _addinOneNoteCache.delete(k);
      }
      return result;
    }
    if (res.status === 429 || res.status === 503) {
      const retryAfter = parseInt(res.headers.get("Retry-After") || "0", 10);
      const wait = retryAfter > 0 ? Math.min(15000, retryAfter * 1000) : Math.min(15000, 1000 * Math.pow(2, attempt));
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
  const isMeeting = isMeetingNoteCategory(category);
  // Meeting categories require a typed note body — that's the meeting minutes.
  // Non-meeting categories use the email body as the OneNote content, so a
  // typed note becomes optional context above the email body.
  if (isMeeting && !body) { setStatus("noteStatus", "error", "Note body is empty."); return; }

  // Snapshot per-item identity SYNCHRONOUSLY before any await. If the user
  // switches items during the OneNote round-trip, these are the values that
  // should land on the saved note — not whatever item is selected when the
  // save completes.
  const snapItemId = emailItem?.itemId || "";
  const snapSharedMsgId = getCurrentSharedMessageId() || "";
  const snapICalUId = currentItemICalUId || "";
  const saveGen = itemContextGeneration;

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

        // For non-meeting categories, the email body IS the OneNote page
        // content. Fetched once here so embedded data:-URI images and inline
        // formatting carry into OneNote as-is. cid: references may render
        // broken in OneNote since they reference Outlook attachments — most
        // modern Outlook bodies inline images as data URIs, which work.
        let emailBodyHtml = "";
        if (!isMeeting) {
          try {
            const token = await getToken();
            emailBodyHtml = await getEmailBodyHtml(token);
          } catch (e) {
            console.warn("[note] body fetch failed:", e.message);
          }
        }

        const page = await createAddinOneNotePage(selectedProject, title, body, category, dateStr, emailBodyHtml);
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
      ...(snapItemId ? { sourceItemId: snapItemId } : {}),
      ...(snapSharedMsgId ? { sourceMessageId: snapSharedMsgId } : {}),
      ...(snapICalUId ? { sourceCalendarUId: snapICalUId } : {}),
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
// `itemSnapshot` is the captured Office mailbox item — required to prevent
// attachment-bytes corruption from item-switch races. Callers MUST capture
// emailItem synchronously before any await and pass it in here.
async function uploadEmailUnderFolder(driveId, token, projFolderName, subfolder, recordFolderName, metadata = null, itemSnapshot = null) {
  const subPath    = await ensureSpFolder(driveId, token, projFolderName, subfolder);
  const recordPath = await ensureSpFolder(driveId, token, subPath, recordFolderName);
  await uploadEmailAndAttachments(driveId, token, recordPath, itemSnapshot);
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
// Mirrors the PMS helper of the same name (SettyPMS.html line 156). A project
// is "Setty-prime" when the prime firm name contains "setty" (case-insensitive).
// Used by the add-in to decide whether to offer sub firms as assignment targets.
function isSettyFirm(name) {
  return !!(name && name.toLowerCase().includes("setty"));
}

// Discipline → short folder code. Used by the RFIs/<code>/<RFI-NNN>/IN
// folder structure on SharePoint. Codes follow the standard AEC convention
// (single letter except FP). Unknown disciplines fall back to "GEN" so the
// folder never breaks; the user can later move it manually if needed.
const DISCIPLINE_TO_CODE = {
  "Mechanical":      "M",
  "Electrical":      "E",
  "Plumbing":        "P",
  "Fire Protection": "FP",
  "General":         "GEN",
  "Architectural":   "A",
  "Structural":      "S",
};
function getDisciplineCode(disciplineName) {
  if (!disciplineName) return "GEN";
  return DISCIPLINE_TO_CODE[disciplineName] || "GEN";
}

// Resolve the assignee email + display name from an assignee-dropdown value.
// dropdownValue is "staff:<id>" or "sub:<id>" or "".
// Returns { email, name } or null if not resolvable.
function resolveAssignee(dropdownValue, project) {
  if (!dropdownValue || !project) return null;
  if (dropdownValue.startsWith("staff:")) {
    const staffId = dropdownValue.slice("staff:".length);
    const m = (project.teamMembers || []).find(x => x.id === staffId);
    if (!m) return null;
    return { email: m.email || "", name: m.name || m.role || "Team member" };
  }
  if (dropdownValue.startsWith("sub:")) {
    const subId = dropdownValue.slice("sub:".length);
    const s = (project.subconsultants || []).find(x => x.id === subId);
    if (!s) return null;
    return { email: s.email || "", name: s.contact || s.firm || "Subconsultant" };
  }
  return null;
}

// Build the HTML body of an RFI assignment email. Keeps it short and
// professional. The recipient is being asked to respond to an RFI; the email
// gives them the context, the question, and the SharePoint link to the
// incoming RFI folder (so they can grab the attachments / forwarded email).
function buildRfiAssignmentEmailHtml({ rfi, project, assignee, inFolderUrl }) {
  const esc = (s) => String(s || "").replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" })[c]);
  const projLabel = [project.projectNumber, project.name].filter(Boolean).join(" — ");
  const dueLine = rfi.dueDate ? `<p><strong>Due:</strong> ${esc(rfi.dueDate)}</p>` : "";
  const descLine = rfi.description ? `<p><strong>Question:</strong></p><p style="margin-left:12px">${esc(rfi.description).replace(/\n/g, "<br>")}</p>` : "";
  const folderLine = inFolderUrl ? `<p><a href="${esc(inFolderUrl)}">📁 SharePoint folder (incoming)</a></p>` : "";
  return `
    <p>Hi ${esc(assignee?.name || "")},</p>
    <p>Please review the attached RFI and respond by the due date.</p>
    <p><strong>Project:</strong> ${esc(projLabel)}<br>
       <strong>RFI:</strong> ${esc(rfi.number)} — ${esc(rfi.title || "")}<br>
       <strong>Discipline:</strong> ${esc(rfi.discipline || "—")}<br>
       <strong>From:</strong> ${esc(rfi.from || "—")}</p>
    ${dueLine}
    ${descLine}
    ${folderLine}
    <p>Thanks,<br>${esc(msalAccount?.name || "")}</p>
  `.trim();
}

// Same shape for Submittals. The "review by" framing matches the typical
// submittal workflow (assignee reviews, returns with stamp/comments).
function buildSubAssignmentEmailHtml({ sub, project, assignee, inFolderUrl }) {
  const esc = (s) => String(s || "").replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" })[c]);
  const projLabel = [project.projectNumber, project.name].filter(Boolean).join(" — ");
  const dueLine = sub.dueDate ? `<p><strong>Review by:</strong> ${esc(sub.dueDate)}</p>` : "";
  const descLine = sub.description ? `<p><strong>Description:</strong></p><p style="margin-left:12px">${esc(sub.description).replace(/\n/g, "<br>")}</p>` : "";
  const folderLine = inFolderUrl ? `<p><a href="${esc(inFolderUrl)}">📁 SharePoint folder (incoming)</a></p>` : "";
  return `
    <p>Hi ${esc(assignee?.name || "")},</p>
    <p>Please review the attached submittal and return with stamp/comments by the due date.</p>
    <p><strong>Project:</strong> ${esc(projLabel)}<br>
       <strong>Submittal:</strong> ${esc(sub.number)}${sub.specSection ? " · Spec " + esc(sub.specSection) : ""}<br>
       <strong>Discipline:</strong> ${esc(sub.discipline || "—")}<br>
       <strong>From:</strong> ${esc(sub.from || "—")}</p>
    ${dueLine}
    ${descLine}
    ${folderLine}
    <p>Thanks,<br>${esc(msalAccount?.name || "")}</p>
  `.trim();
}

// ─── DOCX GENERATORS (RFI Response / Submittal Review cover sheets) ─────────
// Both use the same docx library transmittal.html loads. Style mirrors the
// transmittal cover sheet so the documents look like a coherent family.

const DOCX_COLORS = { NAVY: "1e3a8a", RED: "b91c1c", BLUE: "1e40af", GRAY: "475569", LIGHT: "cbd5e1" };

// Returns a Blob of a DOCX cover sheet for an RFI response. Fields used:
// rfi.number, .title, .description, .from, .discipline, .dateReceived, .dueDate.
// project.projectNumber, .name, .prime, .clientName.
async function buildRfiResponseDocx({ rfi, project, response, dateResponded, status }) {
  if (typeof docx === "undefined" || !docx.Packer) {
    throw new Error("DOCX library not loaded — refresh the taskpane and try again.");
  }
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    Header, BorderStyle, WidthType, ShadingType, VerticalAlign,
  } = docx;
  const projLabel = [project.projectNumber, project.name].filter(Boolean).join(" — ");
  const responder = msalAccount?.name || msalAccount?.username || "";

  function infoRow(label, value) {
    return new TableRow({
      children: [
        new TableCell({
          width: { size: 2160, type: WidthType.DXA },
          shading: { type: ShadingType.SOLID, color: "f1f5f9", fill: "f1f5f9" },
          children: [new Paragraph({ children: [new TextRun({ text: label, font: "Calibri", size: 20, bold: true, color: DOCX_COLORS.NAVY })] })],
        }),
        new TableCell({
          width: { size: 7200, type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: String(value || ""), font: "Calibri", size: 20 })] })],
        }),
      ],
    });
  }

  // Paragraph factory for long-form text blocks that need to wrap.
  function block(text, opts = {}) {
    const lines = String(text || "").split(/\r?\n/);
    return lines.map(line =>
      new Paragraph({
        spacing: { after: opts.dense ? 60 : 100 },
        children: [new TextRun({ text: line, font: "Calibri", size: opts.size || 22 })],
      })
    );
  }

  const doc = new Document({
    styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
    sections: [{
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: DOCX_COLORS.RED, space: 1 } },
            spacing: { after: 120 },
            children: [
              new TextRun({ text: "SETTY", font: "Calibri", size: 28, bold: true, color: DOCX_COLORS.RED }),
              new TextRun({ text: "  & Associates    ", font: "Calibri", size: 18, color: DOCX_COLORS.GRAY }),
              new TextRun({ text: "RFI Response", font: "Calibri", size: 18, color: DOCX_COLORS.GRAY }),
            ],
          })],
        }),
      },
      children: [
        new Paragraph({
          spacing: { before: 0, after: 160 },
          children: [new TextRun({ text: "RFI RESPONSE", font: "Calibri", size: 40, bold: true, color: DOCX_COLORS.NAVY })],
        }),
        new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: DOCX_COLORS.NAVY, space: 1 } },
          spacing: { after: 200 },
          children: [new TextRun({ text: (rfi.number || "RFI").toUpperCase(), font: "Calibri", size: 22, bold: true, color: DOCX_COLORS.BLUE })],
        }),

        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2160, 7200],
          rows: [
            infoRow("Project:",       projLabel),
            infoRow("RFI No.:",       rfi.number || ""),
            infoRow("Title:",         rfi.title || ""),
            infoRow("Discipline:",    rfi.discipline || ""),
            infoRow("From:",          rfi.from || ""),
            infoRow("Date Received:", rfi.dateReceived || ""),
            infoRow("Due Date:",      rfi.dueDate || ""),
            infoRow("Responded:",     dateResponded || ""),
            infoRow("Status:",        status || "Responded"),
            infoRow("Responder:",     responder),
          ],
        }),

        new Paragraph({
          spacing: { before: 280, after: 100 },
          children: [new TextRun({ text: "ORIGINAL QUESTION", font: "Calibri", size: 22, bold: true, color: DOCX_COLORS.NAVY })],
        }),
        ...(block(rfi.description || rfi.title || "(no question text on record)", { size: 22 })),

        new Paragraph({
          spacing: { before: 280, after: 100 },
          children: [new TextRun({ text: "RESPONSE", font: "Calibri", size: 22, bold: true, color: DOCX_COLORS.NAVY })],
        }),
        ...(block(response || "(no response text)", { size: 22 })),

        new Paragraph({
          spacing: { before: 400 },
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: DOCX_COLORS.LIGHT, space: 1 } },
          children: [
            new TextRun({ text: `Generated: ${new Date().toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" })}     `, font: "Calibri", size: 18, color: DOCX_COLORS.GRAY }),
            new TextRun({ text: "Setty & Associates", font: "Calibri", size: 18, bold: true, color: DOCX_COLORS.NAVY }),
          ],
        }),
      ],
    }],
  });

  return await Packer.toBlob(doc);
}

// Same shape for Submittal review. Uses stamp + comments instead of free response.
async function buildSubReviewDocx({ sub, project, comments, stamp, dateReturned, status }) {
  if (typeof docx === "undefined" || !docx.Packer) {
    throw new Error("DOCX library not loaded — refresh the taskpane and try again.");
  }
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    Header, BorderStyle, WidthType, ShadingType,
  } = docx;
  const projLabel = [project.projectNumber, project.name].filter(Boolean).join(" — ");
  const reviewer = msalAccount?.name || msalAccount?.username || "";

  function infoRow(label, value) {
    return new TableRow({
      children: [
        new TableCell({
          width: { size: 2160, type: WidthType.DXA },
          shading: { type: ShadingType.SOLID, color: "f1f5f9", fill: "f1f5f9" },
          children: [new Paragraph({ children: [new TextRun({ text: label, font: "Calibri", size: 20, bold: true, color: DOCX_COLORS.NAVY })] })],
        }),
        new TableCell({
          width: { size: 7200, type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: String(value || ""), font: "Calibri", size: 20 })] })],
        }),
      ],
    });
  }
  function block(text, opts = {}) {
    const lines = String(text || "").split(/\r?\n/);
    return lines.map(line =>
      new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: line, font: "Calibri", size: opts.size || 22 })] })
    );
  }

  // Stamp gets visual emphasis — it's the headline finding of the review.
  const stampColor =
    stamp === "Approved"            ? "059669" :
    stamp === "Approved as Noted"   ? "0891b2" :
    stamp === "Revise and Resubmit" ? "ea580c" :
    stamp === "Rejected"            ? "b91c1c" : DOCX_COLORS.NAVY;

  const doc = new Document({
    styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
      headers: {
        default: new Header({
          children: [new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: DOCX_COLORS.RED, space: 1 } },
            spacing: { after: 120 },
            children: [
              new TextRun({ text: "SETTY", font: "Calibri", size: 28, bold: true, color: DOCX_COLORS.RED }),
              new TextRun({ text: "  & Associates    ", font: "Calibri", size: 18, color: DOCX_COLORS.GRAY }),
              new TextRun({ text: "Submittal Review", font: "Calibri", size: 18, color: DOCX_COLORS.GRAY }),
            ],
          })],
        }),
      },
      children: [
        new Paragraph({
          spacing: { before: 0, after: 160 },
          children: [new TextRun({ text: "SUBMITTAL REVIEW", font: "Calibri", size: 40, bold: true, color: DOCX_COLORS.NAVY })],
        }),
        new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: DOCX_COLORS.NAVY, space: 1 } },
          spacing: { after: 100 },
          children: [new TextRun({ text: (sub.number || "SUB").toUpperCase(), font: "Calibri", size: 22, bold: true, color: DOCX_COLORS.BLUE })],
        }),
        new Paragraph({
          spacing: { after: 240 },
          children: [
            new TextRun({ text: "STAMP: ", font: "Calibri", size: 24, bold: true, color: DOCX_COLORS.NAVY }),
            new TextRun({ text: (stamp || "—").toUpperCase(), font: "Calibri", size: 26, bold: true, color: stampColor }),
          ],
        }),

        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2160, 7200],
          rows: [
            infoRow("Project:",        projLabel),
            infoRow("Submittal No.:",  sub.number || ""),
            infoRow("Spec Section:",   sub.specSection || ""),
            infoRow("Description:",    sub.description || ""),
            infoRow("Discipline:",     sub.discipline || ""),
            infoRow("From:",           sub.from || ""),
            infoRow("Date Received:",  sub.dateReceived || ""),
            infoRow("Due Date:",       sub.dueDate || ""),
            infoRow("Returned:",       dateReturned || ""),
            infoRow("Status:",         status || "Returned"),
            infoRow("Reviewer:",       reviewer),
          ],
        }),

        new Paragraph({
          spacing: { before: 280, after: 100 },
          children: [new TextRun({ text: "REVIEW COMMENTS", font: "Calibri", size: 22, bold: true, color: DOCX_COLORS.NAVY })],
        }),
        ...(block(comments || "(no comments)", { size: 22 })),

        new Paragraph({
          spacing: { before: 400 },
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: DOCX_COLORS.LIGHT, space: 1 } },
          children: [
            new TextRun({ text: `Generated: ${new Date().toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" })}     `, font: "Calibri", size: 18, color: DOCX_COLORS.GRAY }),
            new TextRun({ text: "Setty & Associates", font: "Calibri", size: 18, bold: true, color: DOCX_COLORS.NAVY }),
          ],
        }),
      ],
    }],
  });
  return await Packer.toBlob(doc);
}

// Open a new compose window with prefilled recipient, subject, and HTML body.
// Uses the Office.js mailbox API (works in all Outlook clients, unlike mailto:).
// Best-effort — if the API isn't available or fails, we log and return false
// so callers can surface "draft couldn't be opened — open Outlook and forward
// manually" rather than silently failing.
function openComposeDraft({ toEmail, toName, subject, htmlBody }) {
  try {
    if (!Office?.context?.mailbox?.displayNewMessageForm) return false;
    const toRecipients = toEmail ? [{ displayName: toName || toEmail, emailAddress: toEmail }] : [];
    Office.context.mailbox.displayNewMessageForm({
      toRecipients,
      subject: subject || "",
      htmlBody: htmlBody || "",
    });
    return true;
  } catch (e) {
    console.warn("[draft] displayNewMessageForm failed:", e.message);
    return false;
  }
}

// Populate the Assigned-To dropdown for the currently-selected project.
// Internal team always appears. Subconsultants appear only when Setty is the
// prime, since otherwise the user wouldn't be the one assigning RFIs to subs.
// Option value encodes the kind so doSaveRfi can route the write correctly:
//   "staff:<id>"  → assignedTo = [staffId]
//   "sub:<id>"    → subAssigned = subId
function populateRfiAssigneeDropdown() {
  const sel = document.getElementById("rfiAssignedTo");
  const hint = document.getElementById("rfiAssignHint");
  if (!sel) return;
  const project = selectedProject;
  const team = (project?.teamMembers || []).filter(m => m && (m.name || m.role));
  const subs = (project?.subconsultants || []).filter(s => s && s.firm);
  const isPrime = isSettyFirm(project?.prime || "");
  const opts = ['<option value="">— Unassigned —</option>'];
  if (team.length) {
    opts.push('<optgroup label="Team (internal)">');
    for (const m of team) {
      const label = m.name || m.role || "Unnamed";
      const role = m.role && m.name ? ` · ${m.role}` : "";
      opts.push(`<option value="staff:${m.id}">${escHtml(label)}${escHtml(role)}</option>`);
    }
    opts.push('</optgroup>');
  }
  if (isPrime && subs.length) {
    opts.push('<optgroup label="Subconsultants">');
    for (const s of subs) {
      const contact = s.contact ? ` · ${s.contact}` : "";
      opts.push(`<option value="sub:${s.id}">${escHtml(s.firm)}${escHtml(contact)}</option>`);
    }
    opts.push('</optgroup>');
  }
  sel.innerHTML = opts.join("");
  if (hint) {
    hint.textContent = isPrime && subs.length
      ? "(team or sub firm — we're prime on this project)"
      : team.length
      ? ""
      : "(no team members on this project)";
  }
}

// Tiny HTML-escape for the dropdown <option> labels. Subconsultant firm names
// and contact names can technically contain & < > if someone enters them oddly;
// escape defensively rather than trust the values.
function escHtml(s) {
  return String(s || "").replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" })[c]);
}

// Compute the next auto-RFI number for the selected project. Mirrors PMS:
// "RFI-NNN" where NNN = (existingRfis.length + 1) zero-padded. We do this
// against the cached project's rfis array; if a contractor-assigned number
// is provided the user just overwrites the field.
function nextAutoRfiNumber() {
  const count = (selectedProject?.rfis || []).length + 1;
  return "RFI-" + String(count).padStart(3, "0");
}

function prefillRfi() {
  document.getElementById("rfiTitle").value = emailItem?.subject || "";
  document.getElementById("rfiNumber").value = nextAutoRfiNumber();
  document.getElementById("rfiDescription").value = "";
  document.getElementById("rfiNotes").value = "";
  document.getElementById("rfiAssignedTo").value = "";
  // Reset to Email default — most add-in-filed RFIs come via email since the
  // user is in Outlook. They can override per-RFI if the original source was
  // Procore/ACC/etc.
  document.getElementById("rfiReceivedVia").value = "Email";
  populateRfiAssigneeDropdown();
  setRfiMode("new");
  renderRfiPicker();
}
async function doSaveRfi() {
  const title = document.getElementById("rfiTitle").value.trim();
  if (!title) { setStatus("rfiStatus", "error", "Title is required."); return; }
  // Read DOM fields synchronously (no awaits yet) so they're captured before
  // any item-switch race or async drift.
  const userEnteredNumber = (document.getElementById("rfiNumber")?.value || "").trim();
  const description = (document.getElementById("rfiDescription")?.value || "").trim();
  const discipline = document.getElementById("rfiDiscipline").value;
  const fromField = document.getElementById("rfiFrom").value.trim();
  const notesField = document.getElementById("rfiNotes").value.trim();
  const receivedVia = (document.getElementById("rfiReceivedVia")?.value || "Email");
  const assigneeRaw = (document.getElementById("rfiAssignedTo")?.value || "");
  let assignedToStaff = [];
  let subAssigned = "";
  if (assigneeRaw.startsWith("staff:")) {
    assignedToStaff = [assigneeRaw.slice("staff:".length)];
  } else if (assigneeRaw.startsWith("sub:")) {
    subAssigned = assigneeRaw.slice("sub:".length);
  }
  return withFilingScaffold(
    { operation: "rfi-new", statusElement: "rfiStatus" },
    async ({ snapItem }) => {
      // Re-fetch fresh project data so the auto-number (if the user didn't
      // override) reflects what's actually in the cloud — not the add-in's
      // possibly-stale cache.
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
      const autoNum = "RFI-" + String(existingRfis.length + 1).padStart(3, "0");
      const rfiNumber = userEnteredNumber || autoNum;
      const discCode = getDisciplineCode(discipline);
      const received = new Date();

      // Build the new folder structure: <projFolder>/RFIs/<discCode>/<RFI-NNN>/IN
      // The IN folder gets the incoming email. The RFI root folder is what
      // gets stored on the RFI record so users navigate to the RFI level
      // (not the IN subfolder) when clicking SharePoint links in PMS.
      let spFolderUrl = "";
      let inFolderWebUrl = ""; // Specific URL of the /IN folder (for email body link)
      if (freshProject.projectFolderUrl) {
        try {
          const token = await getToken();
          const { driveId } = await resolveSpIds();
          const projFolderName = decodeURIComponent(freshProject.projectFolderUrl.split("/").pop());
          // Sanitize RFI number for folder name (in case user enters "RFI/external#-001" etc.)
          const safeRfiNumber = rfiNumber.replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
          const rfisPath = await ensureSpFolder(driveId, token, projFolderName, "RFIs");
          const discPath = await ensureSpFolder(driveId, token, rfisPath, discCode);
          const rfiPath  = await ensureSpFolder(driveId, token, discPath, safeRfiNumber);
          const inPath   = await ensureSpFolder(driveId, token, rfiPath, "IN");
          // Sidecar at the RFI level so PMS-side reconcile knows what it represents
          await writeSpMetadataSidecar(driveId, token, rfiPath, buildAddinMetadata(freshProject, "rfi"));
          await uploadEmailAndAttachments(driveId, token, inPath, snapItem);
          // Build web URLs. SP_BASE_URL + encoded path segments.
          const rfiPathEncoded = rfiPath.split("/").map(encodeURIComponent).join("/");
          const inPathEncoded  = inPath.split("/").map(encodeURIComponent).join("/");
          spFolderUrl     = SP_BASE_URL + "/" + rfiPathEncoded;
          inFolderWebUrl  = SP_BASE_URL + "/" + inPathEncoded;
        } catch (e) {
          console.warn("RFI SP upload failed:", e.message);
        }
      }

      const rfi = {
        id: uid(), number: rfiNumber, title,
        description,
        discipline,
        from: fromField,
        receivedVia,
        dateReceived: received.toISOString().slice(0, 10),
        dueDate: addBizDays(received, 5),
        status: "Open",
        notes: notesField,
        assignedTo: assignedToStaff,
        subAssigned,
        spFolderUrl, links: [],
        // Source linkage — lets the main view show a "Logged as RFI" chip when
        // the user returns to this email. Mirrors the pattern used for notes.
        sourceItemId:     snapItem?.itemId || "",
        sourceMessageId:  getCurrentSharedMessageId() || "",
        sourceCalendarUId: currentItemICalUId || "",
        createdAt: new Date().toISOString(),
      };
      await applyLocalChangeAndSave(selectedProject.id, fresh => ({
        ...fresh,
        rfis: [...(fresh.rfis || []), rfi],
      }));

      // Open a prefilled draft to the assignee. Mirrors the transmittal flow:
      // user reviews the draft and hits Send. Failure is non-fatal — the RFI
      // is logged regardless and the user can email manually.
      const assignee = resolveAssignee(assigneeRaw, freshProject);
      let draftOpened = false;
      if (assignee) {
        const subjectLine = `${rfiNumber} · ${freshProject.projectNumber || ""} · ${title}`.trim();
        const htmlBody = buildRfiAssignmentEmailHtml({
          rfi, project: freshProject, assignee, inFolderUrl: inFolderWebUrl || spFolderUrl,
        });
        draftOpened = openComposeDraft({
          toEmail: assignee.email,
          toName:  assignee.name,
          subject: subjectLine,
          htmlBody,
        });
      }

      // Clear the form for the next entry
      document.getElementById("rfiTitle").value = "";
      document.getElementById("rfiNumber").value = "";
      document.getElementById("rfiDescription").value = "";
      document.getElementById("rfiNotes").value = "";
      document.getElementById("rfiAssignedTo").value = "";
      document.getElementById("rfiReceivedVia").value = "Email";

      // Refresh the "Logged as RFI" chip so the main view reflects the new state
      // when the user navigates back from the RFI form.
      try { refreshLoggedArtifactChips(); } catch {}

      const baseMsg = "✓ " + rfiNumber + " logged" + (spFolderUrl ? " · filed to SharePoint" : "");
      const draftMsg = draftOpened ? " · ✉️ Draft opened" : (assignee?.email ? "" : (assignee ? " · ⚠ assignee has no email on file" : ""));
      return {
        sp_folder_url:  spFolderUrl || null,
        status:         spFolderUrl ? "success" : "partial",
        error:          spFolderUrl ? null : "RFI logged without SharePoint upload",
        successMessage: baseMsg + draftMsg,
      };
    }
  );
}
// ─── LOG RFI RESPONSE ────────────────────────────────────────────────────────
// Opens the response view for a specific RFI. Pre-fills today's date and the
// current default status ("Responded"). The button that calls this is rendered
// by refreshLoggedArtifactChips when an RFI matching the current email is open.
let _activeResponseRfiId = "";
function openRfiResponseView(rfiId) {
  const rfi = (selectedProject?.rfis || []).find(r => r.id === rfiId);
  if (!rfi) { setStatus("rfiResponseStatusMsg", "error", "RFI not found."); return; }
  _activeResponseRfiId = rfiId;
  // Pre-fill the header chip + form defaults
  document.getElementById("rfiResponseHeader").textContent =
    `${rfi.number}${rfi.title ? " — " + rfi.title : ""}`;
  document.getElementById("rfiResponseText").value = rfi.response || "";
  document.getElementById("rfiResponseDate").value = new Date().toISOString().slice(0, 10);
  document.getElementById("rfiResponseStatus").value = "Responded";
  setStatus("rfiResponseStatusMsg", "", "");
  showView("rfiResponseView");
}

// Submit the response: generate the DOCX, upload it to the RFI's /OUT folder,
// update the project record, and open a prefilled email back to the RFI sender.
// Status flips to whatever the user picked (default "Responded").
async function submitRfiResponse() {
  const rfiId = _activeResponseRfiId;
  if (!rfiId) { setStatus("rfiResponseStatusMsg", "error", "No RFI selected."); return; }
  const rfi = (selectedProject?.rfis || []).find(r => r.id === rfiId);
  if (!rfi) { setStatus("rfiResponseStatusMsg", "error", "RFI no longer in cache. Refresh and try again."); return; }

  // Synchronous DOM reads
  const responseText = (document.getElementById("rfiResponseText")?.value || "").trim();
  const dateResponded = (document.getElementById("rfiResponseDate")?.value || new Date().toISOString().slice(0, 10));
  const newStatus = (document.getElementById("rfiResponseStatus")?.value || "Responded");
  if (!responseText) {
    setStatus("rfiResponseStatusMsg", "error", "Enter the response text first.");
    return;
  }

  const submitBtn = document.getElementById("submitRfiResponseBtn");
  if (submitBtn) submitBtn.disabled = true;
  setStatus("rfiResponseStatusMsg", "info", "⏳ Generating response…");

  try {
    // 1. Build the DOCX cover sheet
    const docxBlob = await buildRfiResponseDocx({
      rfi, project: selectedProject, response: responseText, dateResponded, status: newStatus,
    });

    // 2. Resolve OUT folder path. If the RFI was filed with the new structure
    // its spFolderUrl points at the RFI root; if it predates the new structure
    // the folder will simply be flat — either way, OUT is created inside.
    let outFolderWebUrl = "";
    let inFolderWebUrl = "";
    if (selectedProject?.projectFolderUrl) {
      const token = await getToken();
      const { driveId } = await resolveSpIds();
      let rfiRootPath = spDrivePath(rfi.spFolderUrl);
      if (!rfiRootPath) {
        const projFolderName = decodeURIComponent(selectedProject.projectFolderUrl.split("/").pop());
        const discCode = getDisciplineCode(rfi.discipline);
        const safeRfiNumber = (rfi.number || "RFI-???").replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
        const rfisPath = await ensureSpFolder(driveId, token, projFolderName, "RFIs");
        const discPath = await ensureSpFolder(driveId, token, rfisPath, discCode);
        rfiRootPath    = await ensureSpFolder(driveId, token, discPath, safeRfiNumber);
      }
      const outPath = await ensureSpFolder(driveId, token, rfiRootPath, "OUT");
      // The IN folder may or may not exist; we link to it best-effort. We don't
      // create it here — if the user only filed via PMS and never via the
      // add-in, there might never have been an /IN, and a stub folder would
      // be misleading. The chip link can simply 404 in that case.
      const inPath = rfiRootPath + "/IN";

      // 3. Upload the DOCX to OUT. Uses the same verified upload path as
      // attachments so it gets the same integrity guarantees.
      const safeName = `${(rfi.number || "RFI").replace(/[\\/:*?"<>|]/g, "-")}_Response.docx`;
      const bytes = new Uint8Array(await docxBlob.arrayBuffer());
      await uploadAttachmentToSharePoint(
        driveId, token, outPath, safeName,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        bytes
      );

      outFolderWebUrl = SP_BASE_URL + "/" + outPath.split("/").map(encodeURIComponent).join("/");
      inFolderWebUrl  = SP_BASE_URL + "/" + inPath.split("/").map(encodeURIComponent).join("/");
    }

    // 4. Update the RFI record in Supabase
    await applyLocalChangeAndSave(selectedProject.id, fresh => ({
      ...fresh,
      rfis: (fresh.rfis || []).map(r => r.id === rfi.id ? {
        ...r,
        response: responseText,
        dateResponded,
        status: newStatus,
      } : r),
    }));

    // 5. Audit-log the response generation. Re-using the filing log so PMS
    // reconcile sees the OUT folder activity and can verify the docx is there.
    void logFilingOp({
      project_id:    selectedProject.id,
      msg_id:        "milestone:rfi-response:" + rfi.id, // unique-ish key
      operation:     "rfi-response",
      sp_folder_url: outFolderWebUrl || null,
      files:         [{ name: `${rfi.number}_Response.docx`, verified: true, distributionKind: "save-sp" }],
      email_subject: `${rfi.number} Response`,
      status:        outFolderWebUrl ? "success" : "partial",
      error:         outFolderWebUrl ? null : "RFI response logged without SharePoint upload",
    });

    // 6. Open a prefilled draft to the RFI's original sender (rfi.from) — or
    // fall back to the assignee if there's no usable sender email.
    const recipientEmail = _guessEmailForRfi(rfi, selectedProject);
    const subjectLine = `${rfi.number} Response · ${selectedProject.projectNumber || ""} · ${rfi.title || ""}`.trim();
    const htmlBody = _buildRfiResponseEmailHtml({ rfi, project: selectedProject, response: responseText, dateResponded, outFolderUrl: outFolderWebUrl, inFolderUrl: inFolderWebUrl });
    const draftOpened = openComposeDraft({
      toEmail: recipientEmail?.email || "",
      toName:  recipientEmail?.name  || "",
      subject: subjectLine,
      htmlBody,
    });

    setStatus("rfiResponseStatusMsg", "success",
      `✓ ${rfi.number} response logged (status: ${newStatus})` +
      (outFolderWebUrl ? " · DOCX in /OUT" : "") +
      (draftOpened ? " · ✉️ Draft opened" : ""));

    // Refresh the main view's chip so it reflects the new status (Responded).
    // Slight delay so the success message is visible before nav.
    setTimeout(() => {
      try { refreshLoggedArtifactChips(); } catch {}
      showView("mainView");
    }, 1500);
  } catch (e) {
    setStatus("rfiResponseStatusMsg", "error", "✗ " + e.message);
    console.error("[rfi-response]", e);
  } finally {
    if (submitBtn) submitBtn.disabled = false;
  }
}

// Best-effort recipient resolver for an RFI response. Prefers an explicit
// from-address on the source email; falls back to the assignee. Returns
// { email, name } or null.
function _guessEmailForRfi(rfi, project) {
  // If the original was logged from an email, we may have its from-address on
  // the email record. The project.emails array has msgId → from/fromAddress.
  if (rfi.sourceItemId || rfi.sourceMessageId) {
    const emailRec = (project.emails || []).find(e =>
      (rfi.sourceItemId && e.msgId === rfi.sourceItemId) ||
      (rfi.sourceMessageId && e.msgId === rfi.sourceMessageId)
    );
    if (emailRec?.fromAddress) return { email: emailRec.fromAddress, name: emailRec.from || "" };
  }
  // Fall back to assignee (if a sub) — useful when we're sending the response
  // through the sub who originally received it.
  if (rfi.subAssigned) {
    const sub = (project.subconsultants || []).find(s => s.id === rfi.subAssigned);
    if (sub?.email) return { email: sub.email, name: sub.contact || sub.firm || "" };
  }
  // Internal staff don't usually receive RFI responses externally, so we skip
  // them. The user can type the recipient themselves in the draft.
  return null;
}

function _buildRfiResponseEmailHtml({ rfi, project, response, dateResponded, outFolderUrl, inFolderUrl }) {
  const esc = (s) => String(s || "").replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" })[c]);
  const projLabel = [project.projectNumber, project.name].filter(Boolean).join(" — ");
  const responseEsc = esc(response).replace(/\n/g, "<br>");
  const links = [];
  if (outFolderUrl) links.push(`<a href="${esc(outFolderUrl)}">📁 Response folder (OUT) — formal DOCX</a>`);
  if (inFolderUrl)  links.push(`<a href="${esc(inFolderUrl)}">📁 Original RFI folder (IN)</a>`);
  return `
    <p>Hi,</p>
    <p>Please see our response to <strong>${esc(rfi.number)}</strong> below. The formal response cover sheet is in the SharePoint folder linked at the bottom.</p>
    <p><strong>Project:</strong> ${esc(projLabel)}<br>
       <strong>RFI:</strong> ${esc(rfi.number)} — ${esc(rfi.title || "")}<br>
       <strong>Date Responded:</strong> ${esc(dateResponded)}</p>
    <p><strong>Response:</strong></p>
    <p style="margin-left:12px">${responseEsc}</p>
    ${links.map(l => `<p>${l}</p>`).join("")}
    <p>Thanks,<br>${esc(msalAccount?.name || "")}</p>
  `.trim();
}

async function doFileToExistingRfi() {
  const rfiId = document.getElementById("rfiExistingSelect").value;
  if (!rfiId) { setStatus("rfiExistingStatus", "error", "Select an RFI."); return; }
  const rfi = (selectedProject?.rfis || []).find(r => r.id === rfiId);
  if (!rfi) { setStatus("rfiExistingStatus", "error", "RFI not found."); return; }
  return withFilingScaffold(
    { operation: "rfi-existing", statusElement: "rfiExistingStatus", startMessage: "⏳ Filing email…" },
    async ({ snapItem }) => {
      const token = await getToken();
      const { driveId } = await resolveSpIds();
      // Resolve the RFI's root folder. If the RFI was logged before the new
      // structure was introduced, rfi.spFolderUrl points at the old flat
      // folder (e.g. <proj>/RFIs/RFI-001 Title). Creating "IN" inside it
      // still works — the new layout coexists with the old.
      let rfiRootPath = spDrivePath(rfi.spFolderUrl);
      if (!rfiRootPath) {
        if (!selectedProject.projectFolderUrl) throw new Error("No SharePoint folder on this project. Create one in the PMS first.");
        const projFolderName = decodeURIComponent(selectedProject.projectFolderUrl.split("/").pop());
        const discCode = getDisciplineCode(rfi.discipline);
        const safeRfiNumber = (rfi.number || "RFI-???").replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
        const rfisPath = await ensureSpFolder(driveId, token, projFolderName, "RFIs");
        const discPath = await ensureSpFolder(driveId, token, rfisPath, discCode);
        rfiRootPath    = await ensureSpFolder(driveId, token, discPath, safeRfiNumber);
      }
      // All incoming correspondence for an RFI lands in the IN subfolder.
      const inPath  = await ensureSpFolder(driveId, token, rfiRootPath, "IN");
      const attCount = await uploadEmailAndAttachments(driveId, token, inPath, snapItem);
      let finalUrl = rfi.spFolderUrl;
      if (!rfi.spFolderUrl) {
        finalUrl = SP_BASE_URL + "/" + rfiRootPath.split("/").map(encodeURIComponent).join("/");
        await applyLocalChangeAndSave(selectedProject.id, fresh => ({
          ...fresh,
          rfis: (fresh.rfis || []).map(r => r.id === rfi.id ? { ...r, spFolderUrl: finalUrl } : r),
        }));
      }
      const attMsg = attCount ? " + " + attCount + " attachment" + (attCount > 1 ? "s" : "") : "";
      const attempted = lastAttachmentUploadStats?.attempted || 0;
      const status =
        (attempted > 0 && attCount === 0)       ? "failed" :
        (attempted > 0 && attCount < attempted) ? "partial" :
        "success";
      return {
        sp_folder_url:  finalUrl,
        status,
        error:          status === "success" ? null : `${attCount}/${attempted} uploaded`,
        successMessage: "✓ Filed to " + rfi.number + " · IN" + attMsg,
      };
    }
  );
}
// ─── LOG SUBMITTAL ────────────────────────────────────────────────────────────
// Populate the Submittal assignee dropdown — same logic as RFI: team always,
// subs only when Setty is prime. Option values use the same "staff:<id>" /
// "sub:<id>" encoding so resolveAssignee handles both.
function populateSubAssigneeDropdown() {
  const sel = document.getElementById("subAssignedTo");
  const hint = document.getElementById("subAssignHint");
  if (!sel) return;
  const project = selectedProject;
  const team = (project?.teamMembers || []).filter(m => m && (m.name || m.role));
  const subs = (project?.subconsultants || []).filter(s => s && s.firm);
  const isPrime = isSettyFirm(project?.prime || "");
  const opts = ['<option value="">— Unassigned —</option>'];
  if (team.length) {
    opts.push('<optgroup label="Team (internal)">');
    for (const m of team) {
      const label = m.name || m.role || "Unnamed";
      const role = m.role && m.name ? ` · ${m.role}` : "";
      opts.push(`<option value="staff:${m.id}">${escHtml(label)}${escHtml(role)}</option>`);
    }
    opts.push('</optgroup>');
  }
  if (isPrime && subs.length) {
    opts.push('<optgroup label="Subconsultants">');
    for (const s of subs) {
      const contact = s.contact ? ` · ${s.contact}` : "";
      opts.push(`<option value="sub:${s.id}">${escHtml(s.firm)}${escHtml(contact)}</option>`);
    }
    opts.push('</optgroup>');
  }
  sel.innerHTML = opts.join("");
  if (hint) {
    hint.textContent = isPrime && subs.length
      ? "(team or sub firm — we're prime on this project)"
      : team.length ? "" : "(no team members on this project)";
  }
}

function nextAutoSubNumber() {
  const count = (selectedProject?.submittals || []).length + 1;
  return "SUB-" + String(count).padStart(3, "0");
}

function prefillSub() {
  document.getElementById("subDesc").value = emailItem?.subject || "";
  document.getElementById("subNumber").value = nextAutoSubNumber();
  document.getElementById("subFullDescription").value = "";
  document.getElementById("subNotes").value = "";
  document.getElementById("subAssignedTo").value = "";
  document.getElementById("subReceivedVia").value = "Email";
  populateSubAssigneeDropdown();
  setSubMode("new");
  renderSubPicker();
}
async function doSaveSub() {
  const desc = document.getElementById("subDesc").value.trim();
  if (!desc) { setStatus("subStatus", "error", "Description is required."); return; }
  // Sync DOM reads before any await
  const userEnteredNumber = (document.getElementById("subNumber")?.value || "").trim();
  const fullDescription = (document.getElementById("subFullDescription")?.value || "").trim();
  const specSection = document.getElementById("subSpec").value.trim();
  const discipline = document.getElementById("subDiscipline").value;
  const fromField = document.getElementById("subFrom").value.trim();
  const notesField = document.getElementById("subNotes").value.trim();
  const receivedVia = (document.getElementById("subReceivedVia")?.value || "Email");
  const assigneeRaw = (document.getElementById("subAssignedTo")?.value || "");
  let assignedToStaff = [];
  let subAssigned = "";
  if (assigneeRaw.startsWith("staff:")) {
    assignedToStaff = [assigneeRaw.slice("staff:".length)];
  } else if (assigneeRaw.startsWith("sub:")) {
    subAssigned = assigneeRaw.slice("sub:".length);
  }
  return withFilingScaffold(
    { operation: "sub-new", statusElement: "subStatus" },
    async ({ snapItem }) => {
      // Re-fetch so submittal numbering reflects current cloud state.
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
      const autoNum = "SUB-" + String(existing.length + 1).padStart(3, "0");
      const subNumber = userEnteredNumber || autoNum;
      const discCode = getDisciplineCode(discipline);
      const received = new Date();

      // New folder structure: <projFolder>/Submittals/<discCode>/<SUB-NNN>/IN
      let spFolderUrl = "";
      let inFolderWebUrl = "";
      if (freshProject.projectFolderUrl) {
        try {
          const token = await getToken();
          const { driveId } = await resolveSpIds();
          const projFolderName = decodeURIComponent(freshProject.projectFolderUrl.split("/").pop());
          const safeSubNumber = subNumber.replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
          const subsPath = await ensureSpFolder(driveId, token, projFolderName, "Submittals");
          const discPath = await ensureSpFolder(driveId, token, subsPath, discCode);
          const subPath  = await ensureSpFolder(driveId, token, discPath, safeSubNumber);
          const inPath   = await ensureSpFolder(driveId, token, subPath, "IN");
          await writeSpMetadataSidecar(driveId, token, subPath, buildAddinMetadata(freshProject, "submittal"));
          await uploadEmailAndAttachments(driveId, token, inPath, snapItem);
          const subPathEncoded = subPath.split("/").map(encodeURIComponent).join("/");
          const inPathEncoded  = inPath.split("/").map(encodeURIComponent).join("/");
          spFolderUrl    = SP_BASE_URL + "/" + subPathEncoded;
          inFolderWebUrl = SP_BASE_URL + "/" + inPathEncoded;
        } catch (e) {
          console.warn("Submittal SP upload failed:", e.message);
        }
      }

      const sub = {
        id: uid(), number: subNumber,
        specSection,
        description: desc,
        fullDescription, // longer free-form text (the new field)
        discipline,
        from: fromField,
        receivedVia,
        dateReceived: received.toISOString().slice(0, 10),
        dueDate: addBizDays(received, 10),
        status: "Received",
        notes: notesField,
        assignedTo: assignedToStaff,
        subAssigned,
        spFolderUrl, links: [],
        sourceItemId:      snapItem?.itemId || "",
        sourceMessageId:   getCurrentSharedMessageId() || "",
        sourceCalendarUId: currentItemICalUId || "",
        createdAt: new Date().toISOString(),
      };
      await applyLocalChangeAndSave(selectedProject.id, fresh => ({
        ...fresh,
        submittals: [...(fresh.submittals || []), sub],
      }));

      // Open prefilled draft to the assignee
      const assignee = resolveAssignee(assigneeRaw, freshProject);
      let draftOpened = false;
      if (assignee) {
        const subjectLine = `${subNumber} · ${freshProject.projectNumber || ""} · ${desc}`.trim();
        const htmlBody = buildSubAssignmentEmailHtml({
          sub, project: freshProject, assignee, inFolderUrl: inFolderWebUrl || spFolderUrl,
        });
        draftOpened = openComposeDraft({
          toEmail: assignee.email,
          toName:  assignee.name,
          subject: subjectLine,
          htmlBody,
        });
      }

      document.getElementById("subDesc").value = "";
      document.getElementById("subSpec").value = "";
      document.getElementById("subNumber").value = "";
      document.getElementById("subFullDescription").value = "";
      document.getElementById("subNotes").value = "";
      document.getElementById("subAssignedTo").value = "";
      document.getElementById("subReceivedVia").value = "Email";

      try { refreshLoggedArtifactChips(); } catch {}

      const baseMsg = "✓ " + subNumber + " logged" + (spFolderUrl ? " · filed to SharePoint" : "");
      const draftMsg = draftOpened ? " · ✉️ Draft opened" : (assignee && !assignee.email ? " · ⚠ assignee has no email on file" : "");
      return {
        sp_folder_url:  spFolderUrl || null,
        status:         spFolderUrl ? "success" : "partial",
        error:          spFolderUrl ? null : "Submittal logged without SharePoint upload",
        successMessage: baseMsg + draftMsg,
      };
    }
  );
}
// ─── LOG SUBMITTAL REVIEW ────────────────────────────────────────────────────
let _activeReviewSubId = "";
function openSubReviewView(subId) {
  const sub = (selectedProject?.submittals || []).find(s => s.id === subId);
  if (!sub) { setStatus("subReviewStatusMsg", "error", "Submittal not found."); return; }
  _activeReviewSubId = subId;
  document.getElementById("subReviewHeader").textContent =
    `${sub.number}${sub.description ? " — " + sub.description : ""}`;
  document.getElementById("subReviewStamp").value = sub.stamp || "Approved";
  document.getElementById("subReviewComments").value = sub.comments || "";
  document.getElementById("subReviewDate").value = new Date().toISOString().slice(0, 10);
  document.getElementById("subReviewStatus").value = "Returned";
  setStatus("subReviewStatusMsg", "", "");
  showView("subReviewView");
}

async function submitSubReview() {
  const subId = _activeReviewSubId;
  if (!subId) { setStatus("subReviewStatusMsg", "error", "No submittal selected."); return; }
  const sub = (selectedProject?.submittals || []).find(s => s.id === subId);
  if (!sub) { setStatus("subReviewStatusMsg", "error", "Submittal no longer in cache."); return; }

  const stamp = (document.getElementById("subReviewStamp")?.value || "—");
  const comments = (document.getElementById("subReviewComments")?.value || "").trim();
  const dateReturned = (document.getElementById("subReviewDate")?.value || new Date().toISOString().slice(0, 10));
  const newStatus = (document.getElementById("subReviewStatus")?.value || "Returned");

  const submitBtn = document.getElementById("submitSubReviewBtn");
  if (submitBtn) submitBtn.disabled = true;
  setStatus("subReviewStatusMsg", "info", "⏳ Generating review…");

  try {
    const docxBlob = await buildSubReviewDocx({
      sub, project: selectedProject, comments, stamp, dateReturned, status: newStatus,
    });

    let outFolderWebUrl = "";
    let inFolderWebUrl = "";
    if (selectedProject?.projectFolderUrl) {
      const token = await getToken();
      const { driveId } = await resolveSpIds();
      let subRootPath = spDrivePath(sub.spFolderUrl);
      if (!subRootPath) {
        const projFolderName = decodeURIComponent(selectedProject.projectFolderUrl.split("/").pop());
        const discCode = getDisciplineCode(sub.discipline);
        const safeSubNumber = (sub.number || "SUB-???").replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
        const subsPath = await ensureSpFolder(driveId, token, projFolderName, "Submittals");
        const discPath = await ensureSpFolder(driveId, token, subsPath, discCode);
        subRootPath    = await ensureSpFolder(driveId, token, discPath, safeSubNumber);
      }
      const outPath = await ensureSpFolder(driveId, token, subRootPath, "OUT");
      const inPath  = subRootPath + "/IN";

      const safeName = `${(sub.number || "SUB").replace(/[\\/:*?"<>|]/g, "-")}_Review.docx`;
      const bytes = new Uint8Array(await docxBlob.arrayBuffer());
      await uploadAttachmentToSharePoint(
        driveId, token, outPath, safeName,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        bytes
      );

      outFolderWebUrl = SP_BASE_URL + "/" + outPath.split("/").map(encodeURIComponent).join("/");
      inFolderWebUrl  = SP_BASE_URL + "/" + inPath.split("/").map(encodeURIComponent).join("/");
    }

    await applyLocalChangeAndSave(selectedProject.id, fresh => ({
      ...fresh,
      submittals: (fresh.submittals || []).map(s => s.id === sub.id ? {
        ...s,
        stamp,
        comments,
        dateReturned,
        status: newStatus,
      } : s),
    }));

    void logFilingOp({
      project_id:    selectedProject.id,
      msg_id:        "milestone:sub-review:" + sub.id,
      operation:     "sub-review",
      sp_folder_url: outFolderWebUrl || null,
      files:         [{ name: `${sub.number}_Review.docx`, verified: true, distributionKind: "save-sp" }],
      email_subject: `${sub.number} Review`,
      status:        outFolderWebUrl ? "success" : "partial",
      error:         outFolderWebUrl ? null : "Submittal review logged without SharePoint upload",
    });

    // For submittals, the return recipient is most commonly the original sender
    // (GC/Prime) who submitted it. Same email-record lookup as RFIs.
    const recipientEmail = _guessEmailForSub(sub, selectedProject);
    const subjectLine = `${sub.number} Review · ${selectedProject.projectNumber || ""} · ${sub.description || ""}`.trim();
    const htmlBody = _buildSubReviewEmailHtml({
      sub, project: selectedProject, stamp, comments, dateReturned,
      outFolderUrl: outFolderWebUrl, inFolderUrl: inFolderWebUrl,
    });
    const draftOpened = openComposeDraft({
      toEmail: recipientEmail?.email || "",
      toName:  recipientEmail?.name || "",
      subject: subjectLine,
      htmlBody,
    });

    setStatus("subReviewStatusMsg", "success",
      `✓ ${sub.number} review logged (stamp: ${stamp})` +
      (outFolderWebUrl ? " · DOCX in /OUT" : "") +
      (draftOpened ? " · ✉️ Draft opened" : ""));

    setTimeout(() => {
      try { refreshLoggedArtifactChips(); } catch {}
      showView("mainView");
    }, 1500);
  } catch (e) {
    setStatus("subReviewStatusMsg", "error", "✗ " + e.message);
    console.error("[sub-review]", e);
  } finally {
    if (submitBtn) submitBtn.disabled = false;
  }
}

function _guessEmailForSub(sub, project) {
  if (sub.sourceItemId || sub.sourceMessageId) {
    const emailRec = (project.emails || []).find(e =>
      (sub.sourceItemId && e.msgId === sub.sourceItemId) ||
      (sub.sourceMessageId && e.msgId === sub.sourceMessageId)
    );
    if (emailRec?.fromAddress) return { email: emailRec.fromAddress, name: emailRec.from || "" };
  }
  if (sub.subAssigned) {
    const s = (project.subconsultants || []).find(x => x.id === sub.subAssigned);
    if (s?.email) return { email: s.email, name: s.contact || s.firm || "" };
  }
  return null;
}

function _buildSubReviewEmailHtml({ sub, project, stamp, comments, dateReturned, outFolderUrl, inFolderUrl }) {
  const esc = (s) => String(s || "").replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" })[c]);
  const projLabel = [project.projectNumber, project.name].filter(Boolean).join(" — ");
  const stampColor =
    stamp === "Approved"            ? "#059669" :
    stamp === "Approved as Noted"   ? "#0891b2" :
    stamp === "Revise and Resubmit" ? "#ea580c" :
    stamp === "Rejected"            ? "#b91c1c" : "#1e3a8a";
  const commentsEsc = esc(comments).replace(/\n/g, "<br>");
  const links = [];
  if (outFolderUrl) links.push(`<a href="${esc(outFolderUrl)}">📁 Review folder (OUT) — formal DOCX</a>`);
  if (inFolderUrl)  links.push(`<a href="${esc(inFolderUrl)}">📁 Original submittal folder (IN)</a>`);
  return `
    <p>Hi,</p>
    <p>We've completed our review of <strong>${esc(sub.number)}</strong>. The formal review cover sheet (with stamp + comments) is in the SharePoint folder linked below.</p>
    <p><strong>Project:</strong> ${esc(projLabel)}<br>
       <strong>Submittal:</strong> ${esc(sub.number)}${sub.specSection ? " · Spec " + esc(sub.specSection) : ""}<br>
       <strong>Stamp:</strong> <span style="color:${stampColor};font-weight:bold">${esc(stamp)}</span><br>
       <strong>Date Returned:</strong> ${esc(dateReturned)}</p>
    ${comments ? `<p><strong>Comments:</strong></p><p style="margin-left:12px">${commentsEsc}</p>` : ""}
    ${links.map(l => `<p>${l}</p>`).join("")}
    <p>Thanks,<br>${esc(msalAccount?.name || "")}</p>
  `.trim();
}

async function doFileToExistingSub() {
  const subId = document.getElementById("subExistingSelect").value;
  if (!subId) { setStatus("subExistingStatus", "error", "Select a submittal."); return; }
  const sub = (selectedProject?.submittals || []).find(s => s.id === subId);
  if (!sub) { setStatus("subExistingStatus", "error", "Submittal not found."); return; }
  return withFilingScaffold(
    { operation: "sub-existing", statusElement: "subExistingStatus", startMessage: "⏳ Filing email…" },
    async ({ snapItem }) => {
      const token = await getToken();
      const { driveId } = await resolveSpIds();
      // Resolve the submittal's root folder; new structure or legacy flat layout
      // both work — creating IN inside an existing folder is safe.
      let subRootPath = spDrivePath(sub.spFolderUrl);
      if (!subRootPath) {
        if (!selectedProject.projectFolderUrl) throw new Error("No SharePoint folder on this project. Create one in the PMS first.");
        const projFolderName = decodeURIComponent(selectedProject.projectFolderUrl.split("/").pop());
        const discCode = getDisciplineCode(sub.discipline);
        const safeSubNumber = (sub.number || "SUB-???").replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
        const subsPath = await ensureSpFolder(driveId, token, projFolderName, "Submittals");
        const discPath = await ensureSpFolder(driveId, token, subsPath, discCode);
        subRootPath    = await ensureSpFolder(driveId, token, discPath, safeSubNumber);
      }
      // Incoming correspondence for a submittal lands in the IN subfolder.
      const inPath = await ensureSpFolder(driveId, token, subRootPath, "IN");
      const attCount = await uploadEmailAndAttachments(driveId, token, inPath, snapItem);
      let finalUrl = sub.spFolderUrl;
      if (!sub.spFolderUrl) {
        finalUrl = SP_BASE_URL + "/" + subRootPath.split("/").map(encodeURIComponent).join("/");
        await applyLocalChangeAndSave(selectedProject.id, fresh => ({
          ...fresh,
          submittals: (fresh.submittals || []).map(s => s.id === sub.id ? { ...s, spFolderUrl: finalUrl } : s),
        }));
      }
      const attMsg = attCount ? " + " + attCount + " attachment" + (attCount > 1 ? "s" : "") : "";
      const attempted = lastAttachmentUploadStats?.attempted || 0;
      const status =
        (attempted > 0 && attCount === 0)       ? "failed" :
        (attempted > 0 && attCount < attempted) ? "partial" :
        "success";
      return {
        sp_folder_url:  finalUrl,
        status,
        error:          status === "success" ? null : `${attCount}/${attempted} uploaded`,
        successMessage: "✓ Filed to " + sub.number + " · IN" + attMsg,
      };
    }
  );
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
    const hasKeyword = /\b(due|deadline|by|no later than|nlt|ntp|submit|required|respond|return|need|complete|deliver|before|expected|must have|scheduled|target)\b/.test(before);
    seen.add(iso);
    results.push({ iso, display, ctx, hasKeyword });
  }
  let m;
  // Long month name with optional year: "March 15, 2026" / "May 18th" / "May 18"
  // Year-optional was the missing case — "May 18th" without a year is the dominant
  // form in AEC email; year is implied by the email's received date.
  const p1 = /\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2})(?:st|nd|rd|th)?(?:,?\s+(\d{4}))?\b/gi;
  while ((m = p1.exec(text))) {
    const mo = MONTHS_LONG.findIndex(x => x.toLowerCase() === m[1].toLowerCase()) + 1;
    const dy = +m[2];
    if (m[3]) {
      addResult(toISO(+m[3], mo, dy), m[0], m.index);
    } else {
      const iso = resolveYearlessMonthDay(mo, dy);
      if (iso) addResult(iso, m[0] + "  (" + iso + ")", m.index);
    }
  }
  // Short month name with optional year: "Mar 15, 2026" / "Mar. 15" / "Sep 8"
  const p2 = /\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.?\s+(\d{1,2})(?:st|nd|rd|th)?(?:,?\s+(\d{4}))?\b/gi;
  while ((m = p2.exec(text))) {
    const mo = MONTHS_SHORT.findIndex(x => x.toLowerCase() === m[1].toLowerCase()) + 1;
    const dy = +m[2];
    if (m[3]) {
      addResult(toISO(+m[3], mo, dy), m[0], m.index);
    } else {
      const iso = resolveYearlessMonthDay(mo, dy);
      if (iso) addResult(iso, m[0] + "  (" + iso + ")", m.index);
    }
  }
  // Day-first with optional year: "15 March 2026" / "18 May"
  const p3 = /\b(\d{1,2})(?:st|nd|rd|th)?\s+(January|February|March|April|May|June|July|August|September|October|November|December)(?:\s+(\d{4}))?\b/gi;
  while ((m = p3.exec(text))) {
    const mo = MONTHS_LONG.findIndex(x => x.toLowerCase() === m[2].toLowerCase()) + 1;
    const dy = +m[1];
    if (m[3]) {
      addResult(toISO(+m[3], mo, dy), m[0], m.index);
    } else {
      const iso = resolveYearlessMonthDay(mo, dy);
      if (iso) addResult(iso, m[0] + "  (" + iso + ")", m.index);
    }
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
  // Sort priority: future dates first (past milestones aren't actionable),
  // then keyword-matched, then chronological. With the top-3 slice in the
  // chip render, future dates naturally crowd out past ones.
  const todayISO = new Date().toISOString().slice(0, 10);
  return results.sort((a, b) => {
    const aFuture = a.iso >= todayISO;
    const bFuture = b.iso >= todayISO;
    if (aFuture !== bFuture) return aFuture ? -1 : 1;
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
    const pep = pickQuip(NEW_MILESTONE_QUIPS);
    if (calResult.success) {
      const calLabel = calResult.onShared ? "NYC Shared Calendar" : "your personal calendar";
      setStatus("milestoneStatus", "success", pep + " Saved to " + projLabel + " · synced to " + calLabel);
    } else {
      setStatus("milestoneStatus", "success", pep + " Saved to " + projLabel + " (calendar sync failed: " + calResult.error + ")");
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
  const wrap = document.getElementById("projectQuickLinks");
  const spHint = document.getElementById("projectSpFolderHint");
  if (!pmsBtn || !spBtn) return;
  pmsBtn.disabled = !projectPmsUrl(selectedProject);
  spBtn.disabled = !selectedProject?.projectFolderUrl;
  if (wrap) wrap.style.display = selectedProject ? "grid" : "none";
  // Pre-emptive hint when the selected project has no SharePoint folder yet.
  // Catches the user before they click "Save to SharePoint" and hit the
  // existing block-on-click error.
  if (spHint) {
    spHint.style.display = (selectedProject && !selectedProject.projectFolderUrl) ? "block" : "none";
  }
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
      // Per-contact role removed from the schema (was redundant once company
      // discipline + per-project relationship were split). `lastContacted` is
      // stamped here so every Outlook-driven capture auto-updates recency —
      // the field no longer needs manual entry in PMS.
      const today = new Date().toISOString().slice(0, 10);
      const contact = { id: uid(), name, title, email, phone, lastContacted: today };
      // Find existing client by exact (case-insensitive) name match
      const queryUrl = SUPABASE_URL + "/rest/v1/pms_clients?select=id,client,version";
      const all = await fetch(queryUrl, { headers: SB_HEADERS });
      if (!all.ok) throw new Error("pms_clients GET HTTP " + all.status);
      const rows = await all.json();
      const existing = (rows || []).find(r => r.client?.name && r.client.name.trim().toLowerCase() === targetCompany.toLowerCase());
      if (existing) {
        const ec = existing.client;
        ec.contacts = ec.contacts || [];
        const dupIdx = ec.contacts.findIndex(c => {
          const e = (c.email || "").trim().toLowerCase();
          const n = (c.name || "").trim().toLowerCase();
          const targetE = (email || "").trim().toLowerCase();
          const targetN = (name || "").trim().toLowerCase();
          if (targetE && e && e === targetE) return true;
          return !targetE && targetN && n === targetN;
        });
        if (dupIdx >= 0) {
          // Don't duplicate — just bump lastContacted on the existing record so
          // recency stays accurate even when the same person emails repeatedly.
          ec.contacts = ec.contacts.map((c, i) => i === dupIdx ? { ...c, lastContacted: today } : c);
        } else {
          ec.contacts = [...ec.contacts, contact];
        }
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
        // New client — INSERT. Discipline `types` left empty; the user
        // categorizes in PMS Global Directory (multi-choice picker).
        const newClient = { id: uid(), name: targetCompany, types: [], contacts: [contact], address: "" };
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
      // Also append to the tagged project's per-project directory if a project
      // is currently selected. PMS Directory tab reads project.directory and
      // merges with auto-rolled POCs/subs, so this lands the contact directly
      // in the project's "people on this job" list. Failure here is non-fatal —
      // the global client save already succeeded.
      if (selectedProject) {
        const emailLc = (email || "").toLowerCase();
        // Per-project role left as "Other" — the user assigns the right
        // role (Prime / Client / SUB to Setty / etc.) in PMS Directory tab.
        const dirEntry = {
          id: uid(),
          name,
          title,
          email,
          phone,
          company: targetCompany,
          type: "Other",
          addedAt: new Date().toISOString(),
          addedBy: msalAccount?.username || "",
          addedFromEmail: emailItem?.itemId || "",
          notes: "",
        };
        try {
          await applyLocalChangeAndSave(selectedProject.id, fresh => {
            const dir = fresh.directory || [];
            // Dedup by email when present; if no email, allow add (user can
            // clean up later — the alternative is silently dropping entries
            // that are otherwise valid).
            if (emailLc && dir.some(d => (d.email || "").toLowerCase() === emailLc)) {
              return fresh;
            }
            return { ...fresh, directory: [...dir, dirEntry] };
          });
        } catch (dirErr) {
          // Non-fatal; the client save already succeeded.
          console.warn("[directory] append failed:", dirErr);
        }
      }
    } else {
      if (!selectedProject) { setStatus("contactStatus", "error", "Select a project first."); return; }
      const poc = { id: uid(), name, title, email, phone };
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
    // When saved to a client AND a project is tagged, the contact also lands
    // in the project's directory — surface that in the success message so users
    // know it's findable in PMS without having to check.
    const alsoInDirectory = saveTo === "client" && selectedProject;
    const dirSuffix = alsoInDirectory ? " · added to " + (selectedProject.name || "project") + " directory." : "";
    setStatus("actionStatus", "success", "✓ Saved " + (name || email) + " to " + destLabel + "." + dirSuffix);
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

// ─── QUICK TEXT + EMAIL TEMPLATES ─────────────────────────────────────────────
// Clipboard-based: each button copies its template text so the user can paste
// it into a reply. No Outlook compose API needed — works in read mode without
// any manifest scope changes. Templates carry [Project Name] tokens which get
// replaced from selectedProject when one is selected.
const QUICK_TEMPLATES = {
  // Short one-liners
  "ack":               () => "Received, we will take a look and get back to you.",
  "attached":          () => "Please find the attached. Let me know if you have any questions.",
  "checking-schedule": () => "We will take a look at our schedule and get back to you.",
  "fee-schedule":      () => formatFeeScheduleSnippet(selectedProject),
  // Long form templates — verbatim from PMS EMAIL_TEMPLATES (SettyPMS.html).
  "schedule-update":   () => fillProjectTokens(TPL_SCHEDULE_UPDATE,   selectedProject),
  "intro":             () => fillProjectTokens(TPL_PROJECT_INTRO,     selectedProject),
  "onhold":            () => fillProjectTokens(TPL_PROJECT_ONHOLD,    selectedProject),
  "sitesurvey":        () => fillProjectTokens(TPL_SITE_SURVEY,       selectedProject),
};

const TPL_SCHEDULE_UPDATE = `Dear [Client Contact],

I hope this message finds you well. I wanted to take a moment to share the current project schedule for [Project Name], along with our anticipated milestone dates. Keeping you informed and aligned throughout the project is a priority for our team, and we want to ensure you have full visibility into where things stand.

Below is a summary of the key project milestones:

[Schedule]

Please note that these dates represent our current projections and may be subject to adjustment based on project conditions, design coordination needs, or information received. We will keep you informed of any changes as they arise.

If you have any questions about the schedule, would like to discuss timeline adjustments, or need additional detail on any of the upcoming milestones, please don't hesitate to reach out. We're here to support a smooth and successful project delivery.

Thank you for your continued partnership — we truly value the opportunity to work with you on this project.

Warm regards,
[Your Name]
[Your Title]
Setty
[Phone Number] | [Email Address]`;

const TPL_PROJECT_INTRO = `Dear [Client Contact],

On behalf of our entire team, we would like to thank you for the opportunity to work with you on the [Project Name] project. We are excited to be part of the team and look forward to supporting the project's success from start to finish.

I will be serving as the Project Manager and primary point of contact for the project. Please feel free to reach out to me directly with any questions, coordination items, or project-related needs as we move forward.

Below is the core project team and their respective disciplines:

[Team Members]

Our goal is to provide responsive communication, proactive coordination, and a smooth project experience throughout the duration of the work.

If at any time you have questions, concerns, or need additional assistance, please also feel free to contact Sara Arias or Danny Kang. They are always available to help ensure the project continues moving successfully and that any concerns are addressed promptly.

We appreciate the opportunity to work with you and look forward to a successful collaboration.

Warm regards,
[Your Name]
[Your Title]
Setty
[Phone Number] | [Email Address]`;

const TPL_PROJECT_ONHOLD = `Dear [Client Contact],

I hope you are doing well.

At this time, work on the [Project Name] project is currently on hold pending receipt of the following information/items:

  • [Outstanding Item #1]
  • [Outstanding Item #2]
  • [Outstanding Item #3]

These items are required in order for our team to proceed with the next phase of work and maintain coordination with the project schedule.

Once the requested information is received, we will resume work promptly and provide an updated schedule for remaining deliverables as needed.

Please let us know if you would like to discuss any of the outstanding items or if there are any questions regarding the information required.

Thank you,
[Your Name]
[Title]
Setty
[Phone Number] | [Email Address]`;

const TPL_SITE_SURVEY = `Dear [Client Contact],

Our team would like to schedule our site survey for the [Project Name] project. The pre-design site survey is a critical step for our team; it's crucial that we begin our work with the most accurate and complete view of the existing systems. Below are some of the critical components of our site visit to keep in mind as we schedule the work.

1 – Existing Plans:
Before our team visits the site, we want to verify first that we have every available existing Mechanical, Electrical, Plumbing and Fire Protection Drawing, especially any that might contain riser and one line diagrams and mechanical spaces. If there are any studies, maintenance records or other documentation on the existing systems that could be made available to us to review prior to being onsite, that will be incredibly valuable. In addition, if you have an "AS-BUILT" architectural plan and RCP from any recent site verification work, please share that with us — we can use it to take field notes and denote any critical dimensions. If sufficient existing information is not available, we should discuss options for additional site investigation time for this project and the limits of our investigations.

2 – Limits of Investigation:
Please note, we will not be doing any invasive investigations, such as opening floors and walls or electrical panel covers, testing equipment or tracing. Our team will do limited above-the-ceiling verification only if the ceilings are accessible and a ladder is made available for our use. If the team feels that this project warrants a higher degree of detail, we should discuss options for follow-up invasive field investigation or specialty field documentation options, such as 3D scanning.

3 – On-Site Access:
Our team will need access to all mechanical, electrical and plumbing utility spaces and possibly the roof. Ideally, a maintenance or facilities person who is familiar with the systems can be made available to walk with us, let us in to the needed spaces, and answer our questions while our team is on site.

Warm regards,
[Your Name]
[Your Title]
Setty
[Phone Number] | [Email Address]`;

// Replace [Project Name] / [Project Number] / [Client Contact] tokens when a
// project is selected. Other tokens (e.g. [Your Name]) are left as-is for the
// user to fill in by hand.
function fillProjectTokens(text, project) {
  if (!project) return text;
  return text
    .replace(/\[Project Name\]/g,   project.name || "[Project Name]")
    .replace(/\[Project Number\]/g, project.projectNumber || "[Project Number]");
}

// Build a "Here is the current Fee Schedule" block from project.phases.
// Phases come from PMS as { code, name, fee }; only billable rows are listed
// (matches how the PMS surfaces the fee schedule in its invoice tooling).
function formatFeeScheduleSnippet(project) {
  const intro = "Here is the current Fee Schedule";
  if (!project) {
    return intro + ":\n\n[Select a project in the add-in to auto-fill the fee schedule.]";
  }
  const phases = (project.phases || []).filter(p => p.billable !== false && (p.fee || 0) > 0);
  if (!phases.length) {
    return `${intro} for ${project.name || project.projectNumber || ""}:\n\n[Fee schedule has not been finalized yet in PMS.]`;
  }
  const lines = phases.map(p => {
    const fee = "$" + Math.round(p.fee || 0).toLocaleString();
    return `  • ${p.name || p.code}: ${fee}`;
  });
  const total = phases.reduce((a, p) => a + (p.fee || 0), 0);
  const totalStr = "$" + Math.round(total).toLocaleString();
  const header = project.projectNumber
    ? `${intro} for ${project.projectNumber} – ${project.name || ""}:`
    : `${intro}:`;
  return `${header}\n\n${lines.join("\n")}\n\nTotal: ${totalStr}`;
}

async function copyTemplateToClipboard(key, btnEl) {
  const builder = QUICK_TEMPLATES[key];
  if (!builder) return;
  const text = builder();
  try {
    await navigator.clipboard.writeText(text);
    showQuickStatus(`✓ Copied — paste into your email (Ctrl+V)`, "success");
    // Brief visual confirmation on the button itself — matches the "Copied ✓"
    // affordance the PMS TemplatesPanel uses for the same kind of action.
    if (btnEl) {
      const orig = btnEl.textContent;
      btnEl.textContent = "✓ Copied";
      setTimeout(() => { btnEl.textContent = orig; }, 1200);
    }
  } catch (e) {
    showQuickStatus("Copy failed: " + e.message, "error");
  }
}

function showQuickStatus(msg, type) {
  const el = document.getElementById("quickTextStatus");
  if (!el) return;
  el.textContent = msg;
  el.style.display = "block";
  el.style.padding = "6px 10px";
  el.style.marginTop = "6px";
  el.style.borderRadius = "4px";
  el.style.fontSize = "11px";
  if (type === "success") {
    el.style.background = "#e6f4ea";
    el.style.color = "#137333";
    el.style.border = "1px solid #b6e3b6";
  } else {
    el.style.background = "#fce8e6";
    el.style.color = "#a50e0e";
    el.style.border = "1px solid #f1b0b0";
  }
  // Auto-hide after 2.5s — non-modal, just a confirmation toast.
  clearTimeout(showQuickStatus._t);
  showQuickStatus._t = setTimeout(() => { el.style.display = "none"; }, 2500);
}
