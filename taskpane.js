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
// On-demand too: reading a mailbox that's been shared/delegated to the signed-in
// user (e.g. a departed colleague's) needs Mail.Read.Shared. Kept out of the
// default sign-in so the consent popup only appears the first time someone
// actually scans another person's mailbox.
const SHARED_MAIL_SCOPES = ["Mail.Read.Shared"];
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
const EMAIL_WATCHLIST_TABLE = "pms_email_watchlist";
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
      // Surface the response watchlist once projects are loaded.
      void renderResponseWatchlist();
      const sweepBlock = document.getElementById("sweepBlock");
      if (sweepBlock) sweepBlock.style.display = "block";
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
      // ItemChanged only ever fires while the pane is pinned, so the first one
      // we receive doubles as proof the user found the pin — hide the hint.
      () => { markPanePinned(); showView("mainView"); applyComposeModeUiGuard(); loadItemContext(); }
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
  // NOTE: saveRecordBtn ("Save to Project") is deliberately NOT in this list — it's
  // retired (auto-file on tag covers the record) and stays hidden via its inline
  // display:none. Listing it here would reset display="" in read mode and un-hide it,
  // since nothing else re-applies its visibility afterward.
  const hiddenInCompose = [
    "saveSpBtn", "saveConfirmation",
    "logNoteBtn", "sendToTeamsBtn", "newActionItemBtn",
    "moreActions", "oneNoteLinkBanner",
    "dateSuggestionBlock",
    // Top-level now (no longer inside moreActions) — keep the old compose-
    // hidden behavior. NOTE: in read mode this guard resets display to "";
    // updatePeopleButtonBadge runs after and applies the real visibility.
    "addParticipantBtn",
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
// ── Pin hint banner ──────────────────────────────────────────────────────────
// The pane is far more useful pinned (it follows the user from email to email),
// but the pin icon Outlook renders in the pane chrome is tiny and most users
// never notice it. The banner stays up until either (a) the user dismisses it,
// or (b) an ItemChanged event proves they pinned — whichever comes first.
function markPanePinned() {
  try { localStorage.setItem("setty_addin_pin_seen", "1"); } catch (e) {}
  const b = document.getElementById("pinHintBanner");
  if (b) b.style.display = "none";
}
function initPinHintBanner() {
  const b = document.getElementById("pinHintBanner");
  if (!b) return;
  let hidden = false;
  try {
    hidden = !!(localStorage.getItem("setty_addin_pin_seen") || localStorage.getItem("setty_addin_pin_dismissed"));
  } catch (e) {}
  if (hidden) return;
  b.style.display = "flex";
  const x = document.getElementById("pinHintDismiss");
  if (x) x.onclick = () => {
    try { localStorage.setItem("setty_addin_pin_dismissed", "1"); } catch (e) {}
    b.style.display = "none";
  };
}
function setupEventListeners() {
  document.getElementById("signInBtn").onclick     = doSignIn;
  document.getElementById("signOutBtn").onclick    = doSignOut;
  const wlRefresh = document.getElementById("responseWatchlistRefresh");
  if (wlRefresh) wlRefresh.onclick = () => { void renderResponseWatchlist(); };
  const sweepFileBtn = document.getElementById("sweepFileBtn");
  if (sweepFileBtn) sweepFileBtn.onclick = () => { void sweepRunAndFile(true); };
  const sweepRunAllBtn = document.getElementById("sweepRunAllBtn");
  if (sweepRunAllBtn) sweepRunAllBtn.onclick = () => { void sweepAutoRunToEnd(); };
  const sweepStopBtn = document.getElementById("sweepStopBtn");
  if (sweepStopBtn) sweepStopBtn.onclick = () => {
    _sweepAutoStop = true;
    const s = document.getElementById("sweepStatus");
    if (s) s.textContent += "  ·  ⏹ stopping after this page…";
  };
  const sweepMoreBtn = document.getElementById("sweepMoreBtn");
  if (sweepMoreBtn) sweepMoreBtn.onclick = () => { void sweepRunAndFile(false); };
  document.getElementById("saveSpBtn").onclick     = doSaveToSharePoint;
  document.getElementById("saveRecordBtn").onclick = doSaveToProjectRecordOnly;
  // Pin hint banner — show unless previously dismissed or pinning was detected.
  initPinHintBanner();
  // Version footer — took over the old main-view logo's jobs (hover tooltip +
  // easter egg) when the red header was removed to save vertical space.
  const versionFooter = document.getElementById("versionFooter");
  if (versionFooter) versionFooter.textContent = "v" + (window.__appVersion || "");
  // 5-click easter egg — reveals the cornerstone card. Counter resets after
  // 3 seconds idle so a curious user has time to discover the pattern but
  // doesn't accidentally trigger it across casual clicks.
  // Bound to the signInView logo AND the mainView version footer (the
  // mainView header-logo no longer exists). Counter is shared across both.
  let _logoClickCount = 0;
  let _logoClickTimer = null;
  document.querySelectorAll(".header-logo, #versionFooter").forEach(logoEl => {
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
  // Contact form back returns wherever it was opened from — the participant list
  // normally, or the main screen when reached via the enrich-from-signature shortcut.
  document.getElementById("contactBack").onclick = () => showView(_contactReturnView || "peopleView");
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
  document.getElementById("addParticipantBtn").onclick = onAddParticipantClick;
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
  // Project search — debounced so fast typing doesn't rebuild the dropdown
  // DOM on every keystroke; one delegated click handler instead of re-binding
  // per option on each render.
  const searchInput = document.getElementById("projectSearch");
  const dropdown    = document.getElementById("projectDropdown");
  dropdown.addEventListener("click", (ev) => {
    const opt = ev.target.closest(".proj-option");
    if (!opt) return;
    setSelectedProject(getProjectById(opt.dataset.id), true);
    searchInput.value = "";
    dropdown.style.display = "none";
  });
  let searchDebounce = null;
  searchInput.addEventListener("input", () => {
    clearTimeout(searchDebounce);
    searchDebounce = setTimeout(() => {
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
    }, 150);
  });
  document.addEventListener("click", (e) => {
    if (!e.target.closest(".project-search-wrapper")) dropdown.style.display = "none";
  });
}

// O(1) project lookup. The Map is rebuilt lazily whenever the allProjects
// ARRAY REFERENCE changes (cache hydration and fresh loads always assign a
// new array, they never mutate in place — so reference identity is a valid
// staleness check). Replaces linear .find() scans in render loops.
let _projectMapSource = null;
let _projectMap = null;
function getProjectById(id) {
  if (_projectMapSource !== allProjects) {
    _projectMap = new Map((allProjects || []).map(p => [p.id, p]));
    _projectMapSource = allProjects;
  }
  return _projectMap.get(id);
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
  // Stale people-picker status from the previous email shouldn't carry over.
  try { setStatus("peopleStatus", "", ""); } catch {}
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
          try { updatePeopleButtonBadge(); } catch {}
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
    try { updatePeopleButtonBadge(); } catch {}

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
    try { updatePeopleButtonBadge(); } catch {}
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
    // Kick off async Message-ID header fetch — it'll re-trigger chip detection
    // once it lands, so RFIs whose sourceMessageId was captured (or whose
    // source/linked email records have that ID stored) become matchable.
    try { fetchInternetMessageIdFromHeaders(); } catch {}
    try { refreshLinkToTargetDropdown(); } catch {}
    try { refreshLoggedArtifactChips(); } catch {}
    maybeShowAecQuip();
    // Score this client email; if it looks like it needs a reply, add it to
    // the shared watchlist. Async + silent — never blocks item loading.
    void maybeAddToWatchlist(myGen);
  }
}
function getCurrentMessageRestId() {
  if (!emailItem?.itemId) return "";
  return Office.context.mailbox.convertToRestId(emailItem.itemId, Office.MailboxEnums.RestVersion.v2_0);
}
// Graph webLink for the current message — a URL that opens the exact email in
// Outlook on the web. Stored on a watchlist row so the panel can link back to
// the original. Returns "" if it can't be resolved (link is best-effort).
async function getCurrentMessageWebLink() {
  try {
    const restId = getCurrentMessageRestId();
    if (!restId) return "";
    const data = await graphFetch("GET", "/me/messages/" + restId + "?$select=webLink", null);
    return data?.webLink || "";
  } catch { return ""; }
}
// Fetched-from-headers Message-ID, keyed by itemContextGeneration so it
// invalidates when the user opens a different item. Set asynchronously by
// fetchInternetMessageIdFromHeaders() — `emailItem.internetMessageId` is
// unreliable across Outlook clients (returns "" in new Outlook for Windows),
// so we read the raw RFC-822 Message-ID header as a fallback.
let cachedInternetMessageId = "";
let cachedInternetMessageIdGen = -1;
function fetchInternetMessageIdFromHeaders() {
  const myGen = itemContextGeneration;
  if (!emailItem?.itemId) return;
  // Already have a synchronous value? No need to fetch.
  if (emailItem?.internetMessageId) {
    cachedInternetMessageId = emailItem.internetMessageId;
    cachedInternetMessageIdGen = myGen;
    return;
  }
  if (!emailItem.getAllInternetHeadersAsync) return; // requires Mailbox 1.8+
  emailItem.getAllInternetHeadersAsync(result => {
    if (myGen !== itemContextGeneration) return; // user moved on
    if (result.status !== Office.AsyncResultStatus.Succeeded) return;
    const m = String(result.value || "").match(/^Message-ID:\s*(.+)$/im);
    const id = m ? m[1].trim() : "";
    if (!id) return;
    cachedInternetMessageId = id;
    cachedInternetMessageIdGen = myGen;
    // Now that we have the stable ID, re-run chip detection — earlier calls
    // may have returned nothing because matching depended on this value.
    try { refreshLoggedArtifactChips(); } catch {}
  });
}
function getEffectiveInternetMessageId() {
  if (emailItem?.internetMessageId) return emailItem.internetMessageId;
  if (cachedInternetMessageIdGen === itemContextGeneration) return cachedInternetMessageId;
  return "";
}
// Awaitable form: resolves the STABLE internetMessageId, fetching it from raw
// headers if Outlook didn't expose it synchronously (new Outlook for Windows
// returns "" for internetMessageId). Callers MUST await this before keying a
// record on getCurrentMessageRecordId(), so auto-log/save never fall back to the
// volatile REST/item id — that fallback was drifting across re-opens, producing
// duplicate rows and a disconnected Save-to-SharePoint patch. Resolves "" only if
// headers are unavailable (Mailbox < 1.8) or the fetch times out, in which case
// callers keep the legacy fallback — never worse than before.
function ensureInternetMessageId() {
  const sync = getEffectiveInternetMessageId();
  if (sync) return Promise.resolve(sync);
  const myGen = itemContextGeneration;
  if (!emailItem?.getAllInternetHeadersAsync) return Promise.resolve("");
  return new Promise(resolve => {
    let done = false;
    const finish = v => { if (!done) { done = true; resolve(v); } };
    setTimeout(() => finish(""), 2500); // never hang a save on a stuck header fetch
    try {
      emailItem.getAllInternetHeadersAsync(result => {
        if (myGen !== itemContextGeneration) return finish(""); // user moved on
        if (result.status !== Office.AsyncResultStatus.Succeeded) return finish("");
        const m = String(result.value || "").match(/^Message-ID:\s*(.+)$/im);
        const id = m ? m[1].trim() : "";
        if (id) { cachedInternetMessageId = id; cachedInternetMessageIdGen = myGen; }
        finish(id);
      });
    } catch { finish(""); }
  });
}
function getCurrentMessageRecordId() {
  // Prefer internetMessageId (shared across recipients) for cross-mailbox matching.
  // Keep REST/item IDs as fallbacks so existing records created before this change still resolve.
  return getEffectiveInternetMessageId() || getCurrentMessageRestId() || emailItem?.itemId || "";
}
function getCurrentMessageIdCandidates() {
  return [...new Set([
    getEffectiveInternetMessageId(),
    getCurrentMessageRestId(),
    emailItem?.itemId || "",
  ].filter(Boolean))];
}
function getCurrentSharedMessageId() {
  return getEffectiveInternetMessageId();
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

// Returns RFI/Submittal records related to the currently-open mailbox item.
// Two kinds of relationship are detected:
//   1. SOURCE — the artifact was originally logged FROM this email
//      (matches via sourceItemId / sourceMessageId)
//   2. LINKED — the user later linked this email to the artifact via the
//      "Link to RFI/Sub" dropdown (matches via artifact.links[].targetId
//      pointing at an email record whose msgId matches this item)
//
// Used by refreshLoggedArtifactChips to surface chips on the main view.
// `status` drives whether the action button (Log Response / Log Review)
// shows below the chip.
function getLoggedRfiSubArtifacts(project) {
  if (!project || !emailItem?.itemId) return [];
  const sourceItemId = emailItem.itemId;
  const sourceMessageIds = getCurrentMessageIdCandidates();
  // Build set of email-record IDs for emails on this project that match
  // the currently-open item. Used to detect "linked" relationships below.
  const matchingEmailRecordIds = new Set();
  for (const e of (project.emails || [])) {
    if (!e?.msgId) continue;
    if (e.msgId === sourceItemId || sourceMessageIds.includes(e.msgId)) {
      if (e.id) matchingEmailRecordIds.add(e.id);
    }
  }
  // Supplemental fuzzy match: same email may have multiple records with
  // different EWS IDs (e.g. email moved between folders). Links and source
  // references may point at an older record whose msgId no longer matches
  // the current EWS ID. Match by subject + sender + date (day) instead.
  {
    const curSubject = typeof emailItem.subject === "string" ? emailItem.subject.trim() : "";
    const curFrom = (emailFromAddress || "").toLowerCase().trim();
    const curDateRaw = emailItem.dateTimeCreated;
    const curDate = curDateRaw instanceof Date
      ? curDateRaw.toISOString().slice(0, 10)
      : typeof curDateRaw === "string" ? curDateRaw.slice(0, 10) : "";
    if (curSubject && curFrom) {
      for (const e of (project.emails || [])) {
        if (matchingEmailRecordIds.has(e.id)) continue;
        if ((e.subject || "").trim() !== curSubject) continue;
        if ((e.fromAddress || "").toLowerCase().trim() !== curFrom) continue;
        if (curDate && e.date && String(e.date).slice(0, 10) !== curDate) continue;
        if (e.id) matchingEmailRecordIds.add(e.id);
      }
    }
  }
  // Helper: does this artifact's links[] reference any matching email record?
  const hasLinkToCurrentEmail = (links) =>
    (links || []).some(lk =>
      lk?.targetSystem === "pms" &&
      lk?.targetType === "email" &&
      lk?.targetId &&
      matchingEmailRecordIds.has(lk.targetId)
    );

  // Helper: does the artifact's sourceItemId belong to an email record that
  // the fuzzy matcher recognized as the current email? Handles the case where
  // the email was moved (EWS ID changed) and sourceMessageId was never stored.
  const isSourceViaFuzzy = (srcItemId) => {
    if (!srcItemId) return false;
    return (project.emails || []).some(
      e => e.msgId === srcItemId && matchingEmailRecordIds.has(e.id)
    );
  };

  const matches = [];
  for (const r of (project.rfis || [])) {
    const isSource = r?.sourceItemId === sourceItemId ||
                     (r?.sourceMessageId && sourceMessageIds.includes(r.sourceMessageId)) ||
                     isSourceViaFuzzy(r?.sourceItemId);
    const isLinked = hasLinkToCurrentEmail(r?.links);
    if (isSource || isLinked) {
      matches.push({
        kind: "rfi",
        id: r.id,
        number: r.number || "RFI",
        title: r.title || "",
        date: r.createdAt || r.dateReceived || null,
        spFolderUrl: r.spFolderUrl || "",
        status: r.status || "Open",
        relationship: isSource ? "source" : "linked",
      });
    }
  }
  for (const s of (project.submittals || [])) {
    const isSource = s?.sourceItemId === sourceItemId ||
                     (s?.sourceMessageId && sourceMessageIds.includes(s.sourceMessageId)) ||
                     isSourceViaFuzzy(s?.sourceItemId);
    const isLinked = hasLinkToCurrentEmail(s?.links);
    if (isSource || isLinked) {
      matches.push({
        kind: "sub",
        id: s.id,
        number: s.number || "SUB",
        title: s.description || "",
        date: s.createdAt || s.dateReceived || null,
        spFolderUrl: s.spFolderUrl || "",
        status: s.status || "Received",
        relationship: isSource ? "source" : "linked",
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

// Populate the "Link to RFI/Submittal" dropdown on the main view. Lists only
// OPEN artifacts (the user can't link to a closed RFI in a meaningful way).
// Hides the row entirely when there are no open RFIs/Subs on the selected
// project. Option value encodes the kind: "rfi:<id>" or "sub:<id>".
// Secondary copy: after the primary Save SP completes, optionally copy the
// email + attachments into a linked RFI or Submittal's /IN folder, and update
// the artifact's links[] array to reference the email record.
//
// linkValue is the encoded dropdown value: "rfi:<id>" or "sub:<id>" or "".
// emailRecord is the email object that was just persisted to project.emails
// (we need its id for the bidirectional cross-link).
// snapItem is the captured Office mailbox item (required by the attachment
// upload path to avoid item-switch races).
//
// Returns { ok: bool, label: string } so the caller can append to the status
// message. Failure is non-fatal — the primary save already succeeded; the
// secondary link is a bonus we attempt best-effort.
async function linkEmailToArtifact({ linkValue, emailRecord, snapItem }) {
  if (!linkValue) return { ok: false, label: "" };
  if (!selectedProject?.projectFolderUrl) return { ok: false, label: "" };
  const kind = linkValue.startsWith("rfi:") ? "rfi" :
               linkValue.startsWith("sub:") ? "sub" : null;
  if (!kind) return { ok: false, label: "" };
  const targetId = linkValue.slice(kind === "rfi" ? "rfi:".length : "sub:".length);
  const arr = kind === "rfi" ? (selectedProject.rfis || []) : (selectedProject.submittals || []);
  const target = arr.find(x => x.id === targetId);
  if (!target) return { ok: false, label: "" };

  try {
    const token = await getToken();
    const { driveId } = await resolveSpIds();
    // Resolve the artifact's root path (new structure or legacy flat layout).
    // Mirrors the resolution logic in doFileToExistingRfi / doFileToExistingSub
    // so the link lands in the right place regardless of when the artifact
    // was originally logged.
    let rootPath = spDrivePath(target.spFolderUrl);
    if (!rootPath) {
      const projFolderName = decodeURIComponent(selectedProject.projectFolderUrl.split("/").pop());
      const discCode = getDisciplineCode(target.discipline);
      const safeNumber = (target.number || (kind === "rfi" ? "RFI-???" : "SUB-???"))
        .replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
      const topPath = await ensureSpFolder(driveId, token, projFolderName, kind === "rfi" ? "RFIs" : "Submittals");
      const discPath = await ensureSpFolder(driveId, token, topPath, discCode);
      rootPath = await ensureSpFolder(driveId, token, discPath, safeNumber);
    }
    // Per-email subfolder under /IN, so multiple emails linked to the same
    // RFI/Sub coexist without colliding on "email.html". Same pattern as the
    // project's main Emails folder.
    const uploadResult = await uploadEmailToArtifactInFolder({
      driveId, token, artifactRootPath: rootPath, snapItem,
    });

    // Update the artifact's links[] array to point at the email record.
    // Schema matches the PMS link format (targetSystem/targetType/targetId).
    const linkEntry = {
      id: uid(),
      targetSystem: "pms",
      targetType:   "email",
      targetId:     emailRecord.id,
      label:        "Related correspondence",
      createdAt:    new Date().toISOString(),
    };
    await applyLocalChangeAndSave(selectedProject.id, fresh => {
      const list = kind === "rfi" ? "rfis" : "submittals";
      return {
        ...fresh,
        [list]: (fresh[list] || []).map(x => x.id === targetId ? {
          ...x,
          links: [...((x.links) || []), linkEntry],
        } : x),
      };
    });

    // Audit log: a separate row for the link operation so the reconcile sweep
    // sees the secondary upload as its own filing event. Points at the
    // specific per-email subfolder where the files actually landed.
    void logFilingOp({
      project_id:    selectedProject.id,
      msg_id:        emailRecord.msgId || null,
      operation:     "email-linked-" + kind,
      sp_folder_url: uploadResult.emailFolderUrl,
      files:         (lastAttachmentUploadStats?.uploadedFiles || []),
      email_subject: emailRecord.subject || "",
      status:        "success",
    });

    return { ok: true, label: ` · 📎 linked to ${target.number || (kind === "rfi" ? "RFI" : "Submittal")}` };
  } catch (e) {
    console.warn("[link-to] secondary copy failed:", e.message);
    return { ok: false, label: ` · ⚠ link to RFI/Sub failed: ${e.message.slice(0, 100)}` };
  }
}

function refreshLinkToTargetDropdown() {
  const row = document.getElementById("linkToRow");
  const sel = document.getElementById("linkToTarget");
  if (!row || !sel) return;
  // Remember the user's previous selection so a UI refresh doesn't blow it
  // away. We restore it at the end if the option still exists.
  const prevValue = sel.value;
  if (!selectedProject) {
    row.style.display = "none";
    sel.innerHTML = '<option value="">— Standalone (no linking) —</option>';
    return;
  }
  const openRfis = (selectedProject.rfis || []).filter(r => RFI_OPEN_STATUSES.has(r?.status || "Open"));
  const openSubs = (selectedProject.submittals || []).filter(s => SUB_OPEN_STATUSES.has(s?.status || "Received"));
  if (openRfis.length === 0 && openSubs.length === 0) {
    row.style.display = "none";
    sel.innerHTML = '<option value="">— Standalone (no linking) —</option>';
    return;
  }
  const opts = ['<option value="">— Standalone (no linking) —</option>'];
  if (openRfis.length) {
    opts.push('<optgroup label="Open RFIs">');
    for (const r of openRfis) {
      const label = `${r.number}${r.title ? " — " + r.title.slice(0, 50) : ""}`;
      opts.push(`<option value="rfi:${r.id}">🔵 ${escHtml(label)}</option>`);
    }
    opts.push('</optgroup>');
  }
  if (openSubs.length) {
    opts.push('<optgroup label="Open Submittals">');
    for (const s of openSubs) {
      const label = `${s.number}${s.description ? " — " + s.description.slice(0, 50) : ""}`;
      opts.push(`<option value="sub:${s.id}">📋 ${escHtml(label)}</option>`);
    }
    opts.push('</optgroup>');
  }
  sel.innerHTML = opts.join("");
  // Restore previous selection if still valid
  if (prevValue && sel.querySelector(`option[value="${prevValue}"]`)) {
    sel.value = prevValue;
  }
  row.style.display = "block";
}

// Render the "Logged as RFI-XXX / SUB-XXX" chips on the main view. Chips link
// to the SharePoint folder when one is stored. Open RFIs/Submittals also get
// a "Log Response" / "Log Review" button rendered below the chip so the user
// can finish the workflow in one click. Hidden when nothing applies.
// When an RFI/Submittal chip is showing for the current email, the link-to
// dropdown, save buttons, and Log as RFI/Sub buttons become redundant — the
// email is already filed and linked. Hide them. Restoration is automatic:
// the underlying managers (refreshEmailSavedIndicator, refreshLinkToTargetDropdown)
// always run before this on the next item change, setting their own visibility
// before this function gets a chance to override.
function _applyChipPresenceUiToggles(hasChips) {
  if (!hasChips) return; // no-op — let the upstream managers own the visible state
  const ids = ["linkToRow", "logRfiBtn", "logSubBtn"];
  for (const id of ids) {
    const el = document.getElementById(id);
    if (el) el.style.display = "none";
  }
  for (const sel of [".save-row", ".save-row-caption"]) {
    const el = document.querySelector(sel);
    if (el) el.style.display = "none";
  }
}
function refreshLoggedArtifactChips() {
  const container = document.getElementById("loggedAsArtifactChips");
  if (!container) return;
  if (!selectedProject || !emailItem?.itemId) {
    container.innerHTML = "";
    container.style.display = "none";
    _applyChipPresenceUiToggles(false);
    return;
  }
  const artifacts = getLoggedRfiSubArtifacts(selectedProject);
  if (artifacts.length === 0) {
    container.innerHTML = "";
    container.style.display = "none";
    _applyChipPresenceUiToggles(false);
    return;
  }
  _applyChipPresenceUiToggles(true);
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
    // "Logged as" for the email the RFI/Sub was originally created from;
    // "Linked to" for emails added via the Link to dropdown after the fact.
    const verb = a.relationship === "linked" ? "Linked to" : "Logged as";
    const label = `${icon} ${verb} ${a.number}${a.title ? " — " + a.title.slice(0, 40) : ""} on ${dateStr}${statusLabel}`;
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
    try { refreshLoggedArtifactChips(); } catch {}
    return;
  }
  const existing = findSavedEmailRecord(selectedProject, getCurrentMessageRecordId());
  if (!existing) {
    applyEmailFlowEmphasis();
    try { refreshLoggedArtifactChips(); } catch {}
    return;
  }

  // Record is logged. Keep the Save to SharePoint button ONLY when there are
  // attachments not yet filed there; otherwise collapse to the "✓ logged" card.
  const wasFiledToSharePoint = !!existing.spFolderUrl;
  const hasAttachments = !!existing.hasAttachments || (attCount && attCount > 0);
  if (!wasFiledToSharePoint && hasAttachments) {
    applyEmailFlowEmphasis(); // keeps the SharePoint button visible for attachments
    if (confirmation) confirmation.style.display = "none";
    setStatus("actionStatus", "success", "✓ Logged to project record — Save attachments to SharePoint below.");
    try { refreshLoggedArtifactChips(); } catch {}
    return;
  }
  // Filed to SharePoint, or nothing to file → collapse into the "✓ logged" card.
  if (saveRow) saveRow.style.display = "none";
  if (saveCapRow) saveCapRow.style.display = "none";

  const savedDate = existing.savedAt
    ? new Date(existing.savedAt).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })
    : "earlier";
  const loggedLabels = getLoggedEmailArtifactLabels(selectedProject);

  const primary = wasFiledToSharePoint
    ? "Saved to SharePoint + project record"
    : "Logged to project record";
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
  // Re-apply chip-aware hides last — if an RFI/Sub chip is showing for this
  // email, the save row stays hidden regardless of save state.
  try { refreshLoggedArtifactChips(); } catch {}
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
      // Inline editor: full-width input stacked over Save/Cancel — comfortable
      // to read and tap, instead of a cramped one-line strip. The YYYY_MM_DD
      // prefix is appended automatically at save-time and everyone knows the
      // convention, so it's not repeated in the UI (tooltip mentions it).
      const input = document.createElement("input");
      input.type = "text";
      input.id = "saveSpRenameInput";
      input.value = _customSpFolderName || _getDefaultSpFolderSubject();
      input.maxLength = 70;
      input.title = "Folder name (the date prefix is added automatically)";
      input.style.cssText = "display:block;width:100%;box-sizing:border-box;font-size:12px;padding:5px 8px;margin:2px 0 6px;border:1px solid var(--primary);border-radius:4px;";
      capSp.appendChild(input);

      const btnRow = document.createElement("div");
      btnRow.style.cssText = "display:flex;gap:6px;align-items:center;";
      capSp.appendChild(btnRow);

      const saveBtn = document.createElement("button");
      saveBtn.type = "button";
      saveBtn.textContent = "Save";
      saveBtn.style.cssText = "font-size:11px;padding:4px 14px;border:none;background:var(--primary);color:#fff;border-radius:4px;cursor:pointer;";
      saveBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        commitSpFolderRename(input.value);
      });
      btnRow.appendChild(saveBtn);

      const cancelBtn = document.createElement("button");
      cancelBtn.type = "button";
      cancelBtn.textContent = "Cancel";
      cancelBtn.style.cssText = "font-size:11px;padding:4px 12px;border:1px solid #ccc;background:#fff;color:#555;border-radius:4px;cursor:pointer;";
      cancelBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        cancelSpFolderRename();
      });
      btnRow.appendChild(cancelBtn);

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
        clearBtn.style.cssText = "font-size:11px;padding:4px 12px;border:1px solid #ccc;background:#fff;color:#a00;border-radius:4px;cursor:pointer;margin-left:auto;";
        clearBtn.addEventListener("click", e => {
          e.preventDefault();
          e.stopPropagation();
          commitSpFolderRename("");
        });
        btnRow.appendChild(clearBtn);
      }
    } else {
      // Prominent, always-visible folder control in the pending state. The old
      // faint underlined link was easy to miss, so show the actual folder name
      // that will be created plus a clearly-tappable Rename button. (The
      // YYYY_MM_DD prefix is added automatically at save time, not shown here.)
      const folderName = _customSpFolderName || _getDefaultSpFolderSubject();
      const wrap = document.createElement("span");
      wrap.style.cssText = "display:inline-flex;align-items:center;flex-wrap:wrap;gap:6px 8px;justify-content:center;";

      const label = document.createElement("span");
      label.style.cssText = "font-size:11px;color:var(--muted);";
      label.appendChild(document.createTextNode("📁 Folder: "));
      const strong = document.createElement("strong");
      strong.style.color = "var(--text)";
      strong.textContent = folderName;
      label.appendChild(strong);
      wrap.appendChild(label);

      const renameBtn = document.createElement("button");
      renameBtn.type = "button";
      renameBtn.id = "saveSpRenameLink";
      renameBtn.textContent = _customSpFolderName ? "✏️ Change name" : "✏️ Rename folder";
      renameBtn.title = "Set a custom folder name (the date prefix is added automatically)";
      renameBtn.style.cssText = "font-size:11px;padding:3px 10px;border:1px solid var(--primary);background:#fff;color:var(--primary);border-radius:4px;cursor:pointer;white-space:nowrap;";
      renameBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        openSpFolderRenameEditor();
      });
      wrap.appendChild(renameBtn);
      capSp.appendChild(wrap);
    }
  }
  if (row) row.style.gridTemplateColumns = "1fr";
  if (capRow) capRow.style.gridTemplateColumns = "1fr";
  // "Save to Project" retired (auto-file on tag covers the record). The SharePoint
  // button stays full-width and always available — it files email + attachments
  // into the project folder.
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
    setStatus("signInStatus", "error", "✗ Sign-in failed: " + humanizeError(e));
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
  // loadProjects() hydrates allProjects from the localStorage cache
  // synchronously (before its first await), so when the cache has data the
  // project selection can be restored immediately — the fresh Supabase fetch
  // keeps running in the background instead of gating the pane for the whole
  // network round trip.
  const freshLoad = loadProjects();
  if (allProjects.length) {
    await restoreProjectSelectionForCurrentEmail();
    updateProjectQuickLinks();
    freshLoad.then(async () => {
      // Re-restore only if nothing got selected from cache — never stomp a
      // selection the user made while the fetch was in flight.
      if (!selectedProject) {
        await restoreProjectSelectionForCurrentEmail();
        updateProjectQuickLinks();
      } else {
        // The selection restored from the cached snapshot — re-point it at
        // the fresh object, otherwise directory/RFI/status checks keep
        // reading pre-fetch data for the rest of the session (e.g. contacts
        // added last session showing as "not on project").
        const freshProj = getProjectById(selectedProject.id);
        if (freshProj && freshProj !== selectedProject) {
          setSelectedProject(freshProj, false);
          // If the participants list is on screen, re-render it with the
          // fresh directory.
          if (document.getElementById("peopleView")?.classList.contains("active")) {
            try { showPeopleView(); } catch {}
          }
        }
      }
      // Fresh allClients may change who counts as "new".
      try { updatePeopleButtonBadge(); } catch {}
    }).catch(() => {});
  } else {
    // Cold start (no cache yet) — nothing to restore from, wait for the fetch.
    await freshLoad;
    await restoreProjectSelectionForCurrentEmail();
    updateProjectQuickLinks();
  }
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
// On-demand token for reading a mailbox shared/delegated to the signed-in user.
// Same lazy-consent pattern as getChannelMessageToken: the Mail.Read.Shared
// prompt fires once, the first time someone scans a colleague's mailbox.
async function getSharedMailToken() {
  const account = msalAccount || msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
  if (!account) throw new Error("Not signed in");
  try {
    const r = await msalApp.acquireTokenSilent({ scopes: SHARED_MAIL_SCOPES, account });
    return r.accessToken;
  } catch {
    const r = await msalApp.acquireTokenPopup({ scopes: SHARED_MAIL_SCOPES, account });
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
const PROJECTS_CACHE_KEY = "settyPms:addinProjectsCacheV5"; // V5: adds slim directory/POC emails for people-picker badges
const PROJECTS_CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 24h hard limit; revalidate every open

// Cache-strip strategy v2 (aggressive): keep ONLY the fields the add-in reads
// at pane-open time. Everything else is re-fetched from Supabase via the V2
// fresh fetch (already happens in parallel with cache hydration, ~300-800ms).
//
// Previous "soft strip" (remove just bodyHtmlCompressed) was still hitting
// localStorage quota for users with hundreds of emails — the metadata fields
// alone (msgId, subject, savedAt, etc.) add up at scale. Hard whitelist solves
// this permanently: the cache becomes O(KB) regardless of email volume.
//
// Anything the cache DOESN'T have is fine — the freshly-loaded `allProjects`
// from Supabase overwrites the cached version within 1s of pane open. The
// cache exists only so the project picker renders instantly. Save flows
// re-fetch the FULL project from Supabase via fetchFreshProjectV2 before
// mutating, so they never depend on cache completeness.
function _stripProjectForCache(p) {
  if (!p) return p;
  return {
    // Top-level scalars needed everywhere
    id: p.id,
    projectNumber: p.projectNumber,
    name: p.name,
    client: p.client,
    clientName: p.clientName,
    prime: p.prime,
    settyPm: p.settyPm,
    status: p.status,
    archived: p.archived,
    projectFolderUrl: p.projectFolderUrl,
    teamsChannelId: p.teamsChannelId,
    teamsNotificationsEnabled: p.teamsNotificationsEnabled,
    teamsOneNoteUrl: p.teamsOneNoteUrl,
    teamsOneNoteNotebookId: p.teamsOneNoteNotebookId,
    oneNoteNotebookId: p.oneNoteNotebookId,
    // Small directory arrays — needed for assignee dropdowns
    teamMembers: p.teamMembers || [],
    subconsultants: p.subconsultants || [],
    // Slim people lists — only the emails, used by the people picker's
    // "already on this project?" badge during the cache window before the
    // fresh fetch lands. Save flows always re-fetch the full project.
    directory: (p.directory || []).map(d => ({ email: d.email || "" })),
    projectContacts: { pm: ((p.projectContacts?.pm) || []).map(c => ({ email: c.email || "" })) },
    // Nested record arrays — slim each record to ONLY the fields the add-in
    // reads from the cached project. findSavedEmailRecord needs msgId.
    // refreshOneNoteLinkBanner needs note.{sourceItemId, sourceMessageId,
    // oneNoteUrl}. getLoggedEmailArtifactLabels needs note + milestone +
    // rfi + sub source-id matchers. Pickers need id/number/title/status/etc.
    emails: (p.emails || []).map(e => ({
      id: e.id,
      msgId: e.msgId,
      subject: e.subject,
      from: e.from,
      fromAddress: e.fromAddress,
      date: e.date,
      spFolderUrl: e.spFolderUrl,
      savedAt: e.savedAt,
    })),
    notes: (p.notes || []).map(n => ({
      id: n.id,
      sourceItemId: n.sourceItemId,
      sourceMessageId: n.sourceMessageId,
      sourceCalendarUId: n.sourceCalendarUId,
      oneNoteUrl: n.oneNoteUrl,
      actionItem: n.actionItem,
      category: n.category,
    })),
    milestones: (p.milestones || []).map(m => ({
      id: m.id,
      sourceItemId: m.sourceItemId,
      sourceMessageId: m.sourceMessageId,
      dueDate: m.dueDate,
      name: m.name,
    })),
    rfis: (p.rfis || []).map(r => ({
      id: r.id,
      number: r.number,
      title: r.title,
      status: r.status,
      discipline: r.discipline,
      spFolderUrl: r.spFolderUrl,
      sourceItemId: r.sourceItemId,
      sourceMessageId: r.sourceMessageId,
      // Keep links so the artifact-chip detector can spot emails linked to
      // this RFI after the fact (not just the original source email).
      // Links are small (just metadata) so this doesn't bloat the cache.
      links: r.links || [],
    })),
    submittals: (p.submittals || []).map(s => ({
      id: s.id,
      number: s.number,
      description: s.description,
      status: s.status,
      discipline: s.discipline,
      spFolderUrl: s.spFolderUrl,
      sourceItemId: s.sourceItemId,
      sourceMessageId: s.sourceMessageId,
      links: s.links || [],
    })),
    // Other arrays (changeOrders, attachments, contracts, etc.) — not used at
    // pane-open time, drop entirely. Save flows re-fetch the full project
    // before mutating, so freshness is guaranteed for writes.
  };
}

// Evict prior versions of the projects cache. They occupy localStorage quota
// that may prevent the new (smaller) cache from being writable. Runs on every
// pane open; cheap.
function _pruneStaleProjectsCaches() {
  try {
    const stale = ["settyPms:addinProjectsCache", "settyPms:addinProjectsCacheV2", "settyPms:addinProjectsCacheV3", "settyPms:addinProjectsCacheV4"];
    for (const k of stale) {
      if (k !== PROJECTS_CACHE_KEY && localStorage.getItem(k) != null) {
        try { localStorage.removeItem(k); } catch {}
      }
    }
  } catch {}
}

function loadProjectsCache() {
  _pruneStaleProjectsCaches();
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

// Page past PostgREST's 1000-row response cap for an arbitrary select, returning
// ALL rows. `pathQuery` must already include its query string AND a stable
// &order=<unique-ish key> so offset paging stays disjoint across pages. Throws on
// the first non-ok page so callers can fall back (e.g. to the legacy pms_data row).
async function sbFetchAllRows(pathQuery, label) {
  const out = [];
  const PAGE = 1000;
  for (let offset = 0; offset < 500000; offset += PAGE) {
    const res = await fetchWithRetry(
      SUPABASE_URL + "/rest/v1/" + pathQuery + "&limit=" + PAGE + "&offset=" + offset,
      { headers: SB_HEADERS }, { label });
    if (!res.ok) throw new Error((label || "sbFetchAllRows") + " HTTP " + res.status);
    const rows = await res.json();
    if (rows && rows.length) out.push(...rows);
    if (!rows || rows.length < PAGE) break;   // last page reached
  }
  return out;
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
    // Page both reads past the 1000-row cap. pms_projects (≈130) and pms_clients
    // (≈291) are under it today but both grow; an unpaged read would silently drop
    // everything past row 1000 once they cross it. sbFetchAllRows throws on a failed
    // page, so a transient error still lands in the catch below and falls back.
    const [pRows, cRows] = await Promise.all([
      sbFetchAllRows("pms_projects?select=id,project,version&order=id.asc", "sb loadProjects v2 projects"),
      sbFetchAllRows("pms_clients?select=client&order=id.asc", "sb loadProjects v2 clients"),
    ]);
    if (pRows && cRows) {
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
    // Keep the localStorage snapshot in step with this save — otherwise the
    // next pane open hydrates pre-save data and shows it until the fresh
    // fetch lands (and, before the re-point fix in onSignedIn, for the whole
    // session). Best-effort: cache failure never fails the save.
    try {
      const versionMap = {};
      for (const [id, ver] of _projectVersionCache.entries()) versionMap[id] = ver;
      saveProjectsCache(allProjects, allClients, versionMap);
    } catch {}
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
async function saveProjectEmailRow(projectId, emailRecord, savedToSharePoint, conversationId) {
  if (!projectId || !emailRecord?.msgId) return;
  // Phase 6 dual-write: write both slim index fields AND new body/metadata
  // columns. Some fields (to/cc/preview/hasAttachments/attachmentNames/savedBy)
  // are not yet captured upstream in the add-in's emailRecord — they default
  // to empty/false. A future patch can enrich emailRecord at construction time
  // to fill these from the Outlook item; until then add-in saves will have
  // these fields blank in pms_project_emails. PMS-side saves already populate
  // all of them.
  const row = {
    record_id: emailRecord.id,
    project_id: projectId,
    msg_id: emailRecord.msgId,
    conversation_id: conversationId ?? (currentConversationId || null),
    subject: emailRecord.subject || "",
    from_name: emailRecord.from || "",
    from_address: emailRecord.fromAddress || "",
    to_addresses: emailRecord.to || "",
    cc_addresses: emailRecord.cc || "",
    // Explicit direction so Claude can reason about "what we sent" vs "what we received"
    // without inferring from the address. Firm domain (setty.com) sender => outgoing.
    direction: String(emailRecord.fromAddress || "").toLowerCase().endsWith("@setty.com") ? "outgoing" : "incoming",
    email_date: emailRecord.date || null,
    saved_at: emailRecord.savedAt || new Date().toISOString(),
    saved_by: emailRecord.savedBy || null,
    sp_folder_url: emailRecord.spFolderUrl || "",
    saved_to_sharepoint: !!savedToSharePoint,
    body_html_compressed: emailRecord.bodyHtmlCompressed || null,
    body_html_size: emailRecord.bodyHtmlSize || 0,
    body_text: emailRecord.bodyText || "",
    preview: emailRecord.preview || "",
    has_attachments: !!emailRecord.hasAttachments,
    attachment_names: emailRecord.attachmentNames || [],
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

// Active (Mon–Fri) milliseconds between two instants. All of Saturday and
// Sunday (local time) are skipped, so a Friday-evening client email doesn't
// trip the response clock until the equivalent point on Monday. Walks the span
// one local day at a time, counting only weekday segments — handles partial
// first/last days correctly.
function businessMsBetween(from, to) {
  const end = new Date(to).getTime();
  let cursor = new Date(from).getTime();
  if (!(end > cursor)) return 0;
  let total = 0;
  while (cursor < end) {
    const d = new Date(cursor);
    const nextMidnight = new Date(d);
    nextMidnight.setHours(24, 0, 0, 0);
    const segEnd = Math.min(nextMidnight.getTime(), end);
    const day = d.getDay();                 // 0 Sun … 6 Sat
    if (day !== 0 && day !== 6) total += segEnd - cursor;
    cursor = segEnd;
  }
  return total;
}
// A flagged email stays hidden until this much business time has elapsed since
// the latest unanswered client message — "24 hours, minus the weekend".
const RESPONSE_GRACE_MS = 24 * 60 * 60 * 1000;

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

// ── JOBCARD ──────────────────────────────────────────────────────────────────
// At-a-glance "what we know about this job" lines, shown inside the green project
// badge. Pure client-side render from the already-loaded project object (status,
// milestones, RFIs, submittals, action-item notes) — no AI, no network call, $0.
function renderJobcard(project) {
  const el = document.getElementById("jobcardBody");
  if (!el) return;
  if (!project) { el.style.display = "none"; el.innerHTML = ""; return; }

  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, c =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const today = new Date().toISOString().slice(0, 10);
  const ok = (d) => typeof d === "string" && d >= "2015-01-01" && d <= "2100-12-31";
  const fmt = (d) => { try { return new Date(d + "T00:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric" }); } catch { return esc(d); } };

  const liveMs = (project.milestones || []).filter(m =>
    m && ok(m.dueDate) && !m.cancelled && (m.status || "") !== "Completed");
  const allMs = (project.milestones || []).filter(m =>
    m && ok(m.dueDate) && !m.cancelled);
  // Next = soonest still-open milestone strictly ahead of us. Previous = the most
  // recent checkpoint reached, today included (completed or not). Today belongs in
  // Previous — it's been reached — so the boundary is > today / <= today, not
  // >= today / < today, which dropped a milestone dated today from both rows.
  // We dropped the "overdue" count: past milestones rarely get marked Completed,
  // so nearly every job looked perpetually overdue. Next/Previous reads as a clean
  // "where we are" timeline.
  const nextM = liveMs.filter(m => m.dueDate > today)
    .sort((a, b) => String(a.dueDate).localeCompare(String(b.dueDate)))[0] || null;
  const prevM = allMs.filter(m => m.dueDate <= today)
    .sort((a, b) => String(b.dueDate).localeCompare(String(a.dueDate)))[0] || null;

  const actions = (project.notes || []).filter(n =>
    n && (n.actionItem || n.category === "Action Item") && (n.actionStatus || "open") !== "done");
  const nextAction = actions.filter(a => ok(a.actionDueDate))
    .sort((a, b) => String(a.actionDueDate).localeCompare(String(b.actionDueDate)))[0] || null;

  const openRfis = (project.rfis || []).filter(r => RFI_OPEN_STATUSES.has(r?.status || "Open")).length;
  const openSubs = (project.submittals || []).filter(s => SUB_OPEN_STATUSES.has(s?.status || "Received")).length;

  const rows = [];
  if (project.status) {
    rows.push('<div style="margin-bottom:4px;"><span style="display:inline-block;font-size:11px;font-weight:600;padding:1px 8px;border-radius:10px;background:var(--bg);border:1px solid #b6e3b6;color:var(--success);">' + esc(project.status) + '</span></div>');
  }
  const row = (label, valueHtml) =>
    '<div style="display:flex;gap:8px;font-size:12px;line-height:1.45;color:var(--text);margin-top:4px;">' +
    '<span style="color:var(--text-soft);min-width:64px;flex:none;font-weight:600;">' + label + '</span>' +
    '<span style="min-width:0;">' + valueHtml + '</span></div>';

  if (nextM) {
    rows.push(row("Next", esc(nextM.name || "Milestone") + " · due " + fmt(nextM.dueDate)));
  }
  if (prevM) {
    rows.push(row("Previous", esc(prevM.name || "Milestone") + " · " + fmt(prevM.dueDate)));
  }

  if (actions.length > 0) {
    let v = actions.length + " action item" + (actions.length > 1 ? "s" : "");
    if (nextAction) v += " · 1 due " + fmt(nextAction.actionDueDate);
    rows.push(row("Actions", esc(v)));
  }

  if (openRfis > 0 || openSubs > 0) {
    const parts = [];
    if (openRfis > 0) parts.push(openRfis + " RFI" + (openRfis > 1 ? "s" : "") + " open");
    if (openSubs > 0) parts.push(openSubs + " submittal" + (openSubs > 1 ? "s" : "") + " in review");
    rows.push(row("Technical", esc(parts.join(" · "))));
  }

  // Ask Claude about this job — opens the user's OWN Claude (their existing seat) with a
  // prefilled prompt. The job snapshot is embedded INLINE so the answer works EVERYWHERE
  // (web, desktop, mobile, even a free account) with no connector dependency — the connector
  // is desktop-only and not yet approved org-wide, so a connector-only prompt would come back
  // empty on web. A soft connector line lets desktop users who do have it pull live email/
  // document detail too. $0: no API key, no metering, no new service.
  const askLines = [
    "Brief me on " + (project.projectNumber ? project.projectNumber + " " : "") +
    (project.name || "this project") +
    ". Here is the current snapshot from our PMS — tell me what needs my attention, what's at risk, and what to do next:",
  ];
  if (project.status) askLines.push("- Status: " + project.status);
  if (nextM) askLines.push("- Next milestone: " + (nextM.name || "Milestone") + " due " + nextM.dueDate);
  if (prevM) askLines.push("- Previous milestone: " + (prevM.name || "Milestone") + " on " + prevM.dueDate);
  if (actions.length > 0) askLines.push("- Open action items: " + actions.length +
    (nextAction ? " (next due " + nextAction.actionDueDate + ")" : ""));
  if (openRfis > 0) askLines.push("- Open RFIs: " + openRfis);
  if (openSubs > 0) askLines.push("- Submittals in review: " + openSubs);
  askLines.push(
    "If you have the Setty PMS connector enabled, use it to pull the latest emails and documents " +
    "(including anything I still owe a reply on) and cite sources; otherwise brief me from the snapshot above.");
  const askPrompt = askLines.join("\n");
  const askUrl = "https://claude.ai/new?q=" + encodeURIComponent(askPrompt);
  rows.push('<a href="#" id="jcAskClaude" style="display:inline-flex;align-items:center;gap:5px;margin-top:8px;font-size:12px;font-weight:600;color:var(--primary);text-decoration:none;">💬 Ask Claude about this job →</a>');

  el.innerHTML = rows.join("");
  el.style.display = "block"; // the Ask-Claude link always renders, so the body always shows

  const askEl = document.getElementById("jcAskClaude");
  if (askEl) askEl.onclick = (e) => {
    e.preventDefault();
    try { openExternalUrl(askUrl); } catch (err) { console.warn("[jobcard] open Claude failed:", err); }
  };
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
  // Re-populate the "Link to" dropdown — open RFIs/Subs differ per project.
  try { refreshLinkToTargetDropdown(); } catch (e) { console.warn("[link-to] refresh failed:", e.message); }
  // "N new" count depends on the selected project's directory.
  try { updatePeopleButtonBadge(); } catch {}
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
  // Once a project is tagged, hide the search box (the badge shows the project +
  // a × to clear). Show it again when there's no selection.
  const searchWrap = document.getElementById("projectSearchWrapper");
  if (searchWrap) searchWrap.style.display = selectedProject ? "none" : "";
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
  try { renderJobcard(selectedProject); } catch (e) { console.warn("[jobcard] render failed:", e.message); }
  refreshActionItemOwnerOptions();
  refreshEmailSavedIndicator();
  refreshOneNoteLinkBanner();
  refreshCalendarStatus();
  applyPipelineUiRules();
  renderProjectSuggestions();
  void renderDateSuggestions();
  // Auto-capture the email to the project log whenever it resolves to a project —
  // explicit tag OR inherited from the conversation tag (restored on open) — so the
  // whole tagged chain gets logged as you read it. Quiet + deduped inside.
  void autoSaveEmailToRecord();
}

// Auto-save the open email to the project log the moment a project is tagged —
// no "save to record" click needed (capture everything; the SharePoint button is
// still used to file attachments down). Read-mode + dedup guarded.
let _autoSavingMsgId = null; // in-flight guard so rapid restores don't double-save
async function autoSaveEmailToRecord() {
  try {
    if (!selectedProject) return;
    if (typeof emailItem?.subject !== "string") return; // compose / appointment — skip
    await ensureInternetMessageId(); // resolve the STABLE id before keying the record (prevents re-log drift)
    const msgId = getCurrentMessageRecordId();
    if (!msgId) return;
    if (_autoSavingMsgId === msgId) return;                    // a save for this email is already running
    if (findSavedEmailRecord(selectedProject, msgId)) return;  // already filed
    _autoSavingMsgId = msgId;
    try { await _doSaveToProjectRecordOnly(true); }            // quiet — no celebrate/status
    finally { _autoSavingMsgId = null; }
  } catch (e) {
    console.warn("auto-save failed:", e);
  }
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
      "Calendar event detected: use 'Log as Note' (under More actions) for meetings/site visits and 'Add Participant to Contacts' for attendees.");
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
  const project = getProjectById(projectId);
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
// signals from the subject, the sender, the project number, and — now that each
// job carries its own directory — how many of that job's contacts are on the
// To/CC line and whether the email came from the job's designated PM contact.
// We *suggest* (never
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
  // Directory signals — the project directory lists who's on each job, so the
  // people on an email can point straight at a project. Scored in
  // directorySignalScore(). TUNE these against real mail; see the SCORING
  // POLICY note there for the "how many is 'lots'?" decision.
  // Set to 1 (not higher) because Setty's contacts often sit on several active
  // jobs at once — so one shared name on CC stays below SUGGESTION_MIN_SCORE
  // and can't surface a job alone; it takes 2+ contacts lining up.
  perDirectoryContactOnThread: 1, // each project contact found on the To/CC line
  senderIsDesignatedPm:        6, // FROM the job's designated PM (moderate: one PM often runs several jobs)
};
const SUGGESTION_MIN_SCORE = 2;
const SUGGESTION_MAX_RESULTS = 3;
// Learned sender→project signal (the same filed-log RPC the sweep uses) cached for
// the SYNCHRONOUS chips (suggestProjects). Loaded once per session by
// ensureSenderSignalCache(); empty until it lands, then read synchronously so each
// email you file makes the chips — not just the sweep — smarter about who files where.
let _senderSignalCache = new Map();
let _senderSignalPromise = null;
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
// Directory-based signal for one project. Two facts about an email point hard
// at a job: how many of the job's OWN contacts are on the To/CC line, and
// whether the email is FROM the job's designated PM contact. Both are computed
// from data already in the cached project (directory[].email + the PM contact).
//
// We deliberately ignore our own firm's addresses — a Setty teammate is on
// every internal thread, so their presence tells us nothing about WHICH job.
//
// KNOWN LIMITATION: a contact who sits on many active jobs (e.g. a county
// facilities director) is less discriminating than one unique to a single job.
// A later refinement could down-weight a contact by how many projects they
// appear on (an inverse-frequency weight); for now every project contact on
// the thread counts equally. Returns the { points, reasons } shape the caller
// already accumulates.
const FIRM_EMAIL_DOMAIN = "@setty.com";
function directorySignalScore(project, participants, senderEmail) {
  const norm = e => (e || "").trim().toLowerCase();
  const isExternal = e => e && !e.endsWith(FIRM_EMAIL_DOMAIN);

  // This job's discriminating contacts: external directory people + the
  // designated PM, with our own firm's addresses filtered out.
  const contactEmails = new Set(
    [
      ...((project.directory || []).map(d => norm(d.email))),
      ...((project.projectContacts?.pm || []).map(c => norm(c.email))),
    ].filter(isExternal)
  );
  if (!contactEmails.size) return { points: 0, reasons: [] };

  // How many of them are on the To/CC line? (The sender is scored separately.)
  const onThread = new Set(
    (participants || [])
      .filter(p => p.label === "To" || p.label === "CC")
      .map(p => norm(p.emailAddress))
      .filter(Boolean)
  );
  let overlap = 0;
  for (const e of contactEmails) if (onThread.has(e)) overlap++;

  // Is the email FROM this job's designated PM contact?
  const pmEmails = new Set((project.projectContacts?.pm || []).map(c => norm(c.email)).filter(Boolean));
  const fromPm = pmEmails.has(norm(senderEmail));

  // ── SCORING POLICY ─────────────────────────────────────────────────────────
  // LINEAR: every project contact on the thread is worth
  // perDirectoryContactOnThread, so "lots of them" stack up. Tuned to Setty's
  // reality (contacts often span several active jobs): one lone contact scores
  // 1 — below SUGGESTION_MIN_SCORE, so it can't surface a job alone — while 2+
  // lining up clears the bar. "From the designated PM" is a moderate nudge (6)
  // that surfaces all of that PM's jobs for you to pick from, since one PM
  // often runs several. Re-tune the weights above if real mail says otherwise.
  let points = 0;
  const reasons = [];
  if (overlap > 0) {
    points += overlap * SUGGESTION_WEIGHTS.perDirectoryContactOnThread;
    reasons.push(overlap === 1 ? "1 project contact on thread"
                               : overlap + " project contacts on thread");
  }
  if (fromPm) {
    points += SUGGESTION_WEIGHTS.senderIsDesignatedPm;
    reasons.push("from project's PM contact");
  }
  return { points, reasons };
}
function suggestProjects(subject, senderEmail, participants = emailParticipants) {
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

    // Directory signals: project contacts on the To/CC line + sender = PM contact.
    const dirSignal = directorySignalScore(p, participants, senderEmail);
    if (dirSignal.points) {
      score += dirSignal.points;
      reasons.push(...dirSignal.reasons);
    }

    if (score >= SUGGESTION_MIN_SCORE) {
      scored.push({ project: p, score, reasons });
    }
  }

  // Fold in the LEARNED sender→project signal (same filed-log data the sweep uses,
  // cached on email open). Beyond boosting an existing subject match, mergeSenderSignal
  // ADDS a candidate for a project this sender reliably files to — so a known sender
  // surfaces the right job even when the subject names no project at all.
  const learned = _senderSignalCache.get((senderEmail || "").trim().toLowerCase());
  if (learned && learned.length) mergeSenderSignal(scored, learned);

  scored.sort((a, b) => b.score - a.score || (a.project.name || "").localeCompare(b.project.name || ""));
  return scored.slice(0, SUGGESTION_MAX_RESULTS);
}

// ─── EMAIL SWEEP (dry-run preview) ────────────────────────────────────────────
// Scans the user's recent mail and classifies each message with suggestProjects:
//   file   → confident single match (project # in subject, OR top score
//            >= SWEEP_AUTOFILE_MIN and beats the runner-up by SWEEP_AUTOFILE_MARGIN)
//   review → plausible but ambiguous (>= SUGGESTION_MIN_SCORE) → human queue
//   skip   → nothing matched
// PREVIEW WRITES NOTHING. It exists to calibrate the thresholds against real mail
// before auto-filing is enabled. Filing + the review queue come next once the
// numbers look right.
const SWEEP_PAGE_SIZE = 250;          // messages per Graph page — a lean $select keeps this well under the 504 ceiling
const SWEEP_TARGET_ACTIONABLE = 300;  // keep paging within one batch until this many file+review surface. Raised from 50 for backfill: each Run & file / Load more now walks far deeper before pausing (older/sparser mail hits the SWEEP_MAX_PAGES_PER_BATCH ceiling instead) — fewer clicks to get through a mailbox
const SWEEP_MAX_PAGES_PER_BATCH = 30; // safety cap: at most 30×250 = 7,500 messages scanned per Preview / Load-more press
const SWEEP_REVIEW_MIN  = 4;   // min name/acronym score to ENTER review (≥2 distinctive signals)

// Common AEC / institutional words that appear across many project names — too
// generic to identify a project alone (data-driven from the filed-email subjects).
const SWEEP_EXTRA_STOPWORDS = new Set([
  "nyc", "ny", "new", "york", "city", "state", "county", "public", "authority", "dormitory",
  "suny", "sucf", "ogs", "dasny", "sca", "cuny", "dcas", "ddc", "dsny", "nycha", "mta",
  "university", "college", "school", "schools", "campus", "hall", "building", "bldg", "center", "centre",
  "design", "services", "service", "engineering", "consulting", "construction", "architectural",
  "replacement", "replace", "upgrade", "upgrades", "improvements", "improvement",
  "rehabilitation", "rehab", "roof", "hvac", "mep", "electrical", "mechanical", "plumbing", "fire", "alarm",
  "study", "feasibility", "assessment", "survey", "report", "garage", "facility", "facilities",
  "department", "dept", "office", "task", "order",
]);
function sweepTokenize(s) {
  // Drop generic words AND short bare numbers ("23", "06") that coincidentally
  // match dates/counts; keep 3+ digit identifiers (280, 292, 578) and alphanumerics (9r).
  return suggestionTokenize(s)
    .filter((t) => !SWEEP_EXTRA_STOPWORDS.has(t))
    .filter((t) => !(/^\d+$/.test(t) && t.length < 3));
}

// Bulk/marketing/automated sender check — no-reply, alerts, ESP & bid/event domains.
function looksPromotional(addr) {
  const a = (addr || "").toLowerCase();
  const local = a.split("@")[0] || "";
  const domain = a.split("@")[1] || "";
  if (/(^|[._-])(no-?reply|do-?not-?reply|donotreply|noreply|newsletter|marketing|mailer|notifications?|alerts?|bounce|invitations?|projecttracker|tracker)([._-]|$)/.test(local)) return true;
  if (/(mailchimp|constantcontact|sendgrid|eventbrite|substack|hubspot|marketo|pardot|sparkpostmail|rsgsv|mcsv|cmail\d|bidnet|nyscr|proest|crewnetwork)/.test(domain)) return true;
  return false;
}

// Is a project ID present in the subject? The SAPX number (full or sans .00
// suffix) OR the client's prime number — both appear in ~7-9% of filed subjects
// and are the strongest possible signal. Guards against junk/short prime values.
function subjectHasProjectId(subjLower, p) {
  // Normalize separators so "2022.37-01" (prime) matches "2022.37.01" in a subject.
  const norm = (s) => s.replace(/[_\-\s]+/g, ".");
  const subjN = norm(subjLower);
  const ids = [];
  const sapx = String(p.projectNumber || "").trim().toLowerCase();
  if (sapx.length >= 5) { ids.push(sapx); ids.push(sapx.split(".")[0]); }
  const prime = String(p.primeProjectNumber || "").trim().toLowerCase();
  if (prime.length >= 5 && /\d/.test(prime)) ids.push(prime);
  return ids.some((id) => id && (subjLower.includes(id) || subjN.includes(norm(id))));
}

// Subject-only project matcher for the sweep. Ignores sender-domain / shared-
// contact signals (noise in a batch). Signals (data-driven from filed emails):
// project ID in subject (strong), DISTINCTIVE name words, acronym.
function sweepSuggest(subject) {
  const subj = (subject || "").toLowerCase();
  if (!subj) return [];
  const subjTokens = new Set(sweepTokenize(subj));
  const scored = [];
  for (const p of (allProjects || [])) {
    if (!p || p.archived || (!p.name && !p.projectNumber)) continue;
    let score = 0;
    const reasons = [];
    if (subjectHasProjectId(subj, p)) {
      score += SUGGESTION_WEIGHTS.projectNumberInSubject; // 10 — near-certain
      reasons.push("project # in subject");
    }
    const hits = sweepTokenize(p.name || "").filter((t) => subjTokens.has(t));
    if (hits.length) {
      score += hits.length * SUGGESTION_WEIGHTS.perNameTokenInSubject;
      reasons.push(hits.length + " name word" + (hits.length > 1 ? "s" : "") + " match");
    }
    if ([...suggestionAcronyms(p.name || "")].some((a) => subjTokens.has(a))) {
      score += SUGGESTION_WEIGHTS.acronymInSubject;
      reasons.push("acronym match");
    }
    if (score >= SUGGESTION_WEIGHTS.perNameTokenInSubject) scored.push({ project: p, score, reasons });
  }
  scored.sort((a, b) => b.score - a.score || (a.project.name || "").localeCompare(b.project.name || ""));
  return scored.slice(0, 3);
}

// To/CC list in the shape suggestProjects expects, from a Graph message.
function sweepParticipants(msg) {
  const out = [];
  for (const r of (msg.toRecipients || [])) {
    const a = r.emailAddress?.address; if (a) out.push({ label: "To", emailAddress: a });
  }
  for (const r of (msg.ccRecipients || [])) {
    const a = r.emailAddress?.address; if (a) out.push({ label: "CC", emailAddress: a });
  }
  return out;
}

// Confidence policy (tightened after live testing — auto-file must be near-certain):
//   auto-file ONLY when a project ID (SAPX or client prime #) is in the subject;
//   review only when the name/acronym score clears SWEEP_REVIEW_MIN (≥2 distinctive
//   signals); otherwise skip. A single coincidental word no longer files anything.
function classifySweep(candidates) {
  if (!candidates.length) return { action: "skip" };
  const top = candidates[0];
  const hasId = (top.reasons || []).some((r) => r.includes("project #"));
  if (hasId) return { action: "file", project: top.project, score: top.score, reasons: top.reasons };
  if (top.score >= SWEEP_REVIEW_MIN) return { action: "review", candidates: candidates.slice(0, 3) };
  return { action: "skip" };
}

// msg_ids already in the email log, so the preview ignores already-filed mail.
async function sweepLoadFiledIds() {
  // PostgREST caps every response at 1000 rows, so a plain ?select=msg_id only ever
  // returned the first 1000 of however many are filed — the rest looked "never filed"
  // and got re-processed (and re-saved, only to be rejected by the unique index) on
  // every run. Page through ALL rows so the dedup set is complete.
  const ids = new Set();
  try {
    const PAGE = 1000;
    for (let offset = 0; offset < 500000; offset += PAGE) {
      const res = await fetchWithRetry(
        SUPABASE_URL + "/rest/v1/" + PROJECT_EMAILS_TABLE + "?select=msg_id&limit=" + PAGE + "&offset=" + offset,
        { headers: SB_HEADERS }, { label: "sb sweep filed ids" });
      if (!res.ok) break;
      const rows = await res.json();
      for (const r of (rows || [])) if (r.msg_id) ids.add(r.msg_id);
      if (!rows || rows.length < PAGE) break;   // last page reached
    }
  } catch { /* return whatever we gathered */ }
  return ids;
}

// conversation_id -> projectId for threads that already have a filed email, so a
// thread tied to a job auto-files the rest instead of re-surfacing each message.
// Only UNAMBIGUOUS threads (filed to exactly one project) auto-file; a thread
// split across projects is left for normal per-message classification.
async function sweepLoadConvoProjectMap() {
  try {
    const byConv = new Map();
    const PAGE = 1000;                          // page past the 1000-row PostgREST cap
    for (let offset = 0; offset < 500000; offset += PAGE) {
      const res = await fetchWithRetry(
        SUPABASE_URL + "/rest/v1/" + PROJECT_EMAILS_TABLE + "?select=conversation_id,project_id&conversation_id=not.is.null&limit=" + PAGE + "&offset=" + offset,
        { headers: SB_HEADERS }, { label: "sb sweep convo map" });
      if (!res.ok) break;
      const rows = await res.json();
      for (const r of (rows || [])) {
        const c = r.conversation_id;
        if (!c || !r.project_id) continue;
        if (!byConv.has(c)) byConv.set(c, new Set());
        byConv.get(c).add(r.project_id);
      }
      if (!rows || rows.length < PAGE) break;    // last page reached
    }
    const map = new Map();
    for (const [c, set] of byConv) if (set.size === 1) map.set(c, [...set][0]);
    return map;
  } catch { return new Map(); }
}

// Learned sender→project signal from the filed log (RPC sender_project_signals).
// Self-improving: as more emails get filed, more senders become project-specific.
// Fetched once per sweep. Returns Map(lower-address -> [{projectId, n, projects}]).
async function sweepLoadSenderMap() {
  // sender_project_signals RETURNS TABLE, so its RPC response is capped at 1000
  // rows just like a plain table read. The result is one row per sender→project
  // signal; that count climbs as the email backfill adds external senders (ceiling
  // ≈ 2× distinct external senders), so page past the cap instead of trusting one
  // call. PostgREST applies limit/offset/order to a table-valued function's result;
  // the explicit order on a stable key keeps successive pages disjoint.
  const map = new Map();
  try {
    const PAGE = 1000;
    for (let offset = 0; offset < 500000; offset += PAGE) {
      const res = await fetchWithRetry(
        SUPABASE_URL + "/rest/v1/rpc/sender_project_signals?limit=" + PAGE + "&offset=" + offset +
          "&order=from_address.asc,project_id.asc",
        { method: "POST", headers: { ...SB_HEADERS, "Content-Type": "application/json" }, body: "{}" },
        { label: "sb sender signals" });
      if (!res.ok) break;
      const rows = await res.json();
      for (const r of (rows || [])) {
        const addr = (r.from_address || "").toLowerCase();
        if (!addr) continue;
        if (!map.has(addr)) map.set(addr, []);
        map.get(addr).push({ projectId: r.project_id, n: r.n, projects: r.projects });
      }
      if (!rows || rows.length < PAGE) break;   // last page reached
    }
  } catch { /* return whatever we gathered */ }
  return map;
}

// Merge the learned sender signal into a message's subject candidates. A sender
// who historically files to ONE project is strong evidence (even with no subject
// match); two projects is moderate. This is what lets sender-only emails (no
// project name in the subject) still surface for review.
function mergeSenderSignal(candidates, senderProjects) {
  if (!senderProjects || !senderProjects.length) return candidates;
  const byId = new Map(candidates.map((c) => [c.project.id, c]));
  for (const sp of senderProjects) {
    const proj = (allProjects || []).find((p) => p && p.id === sp.projectId);
    if (!proj) continue;
    const boost = sp.projects === 1 ? 8 : 4;
    if (byId.has(sp.projectId)) {
      const c = byId.get(sp.projectId);
      c.score += boost;
      if (!c.reasons.includes("known sender")) c.reasons.push("known sender");
    } else {
      const c = { project: proj, score: boost, reasons: ["known sender"] };
      byId.set(sp.projectId, c);
      candidates.push(c);
    }
  }
  candidates.sort((a, b) => b.score - a.score);
  return candidates;
}

// Load the learned sender→project map ONCE per session and cache it for the
// synchronous chips (suggestProjects). Reuses the sweep's loader. onReady fires
// only on the initial load so the chips can re-rank the moment the signal arrives;
// after that every email reads _senderSignalCache synchronously. Refreshes on
// taskpane reload (so a session's new filings show up next time the pane opens).
function ensureSenderSignalCache(onReady) {
  if (_senderSignalPromise) return _senderSignalPromise;
  _senderSignalPromise = (async () => {
    try { _senderSignalCache = await sweepLoadSenderMap(); }
    catch { _senderSignalCache = new Map(); }
  })();
  if (typeof onReady === "function") _senderSignalPromise.then(onReady);
  return _senderSignalPromise;
}

async function sweepRecentMail() {
  const btn = document.getElementById("sweepRunBtn");
  const statusEl = document.getElementById("sweepStatus");
  const resultsEl = document.getElementById("sweepResults");
  if (!allProjects || !allProjects.length) {
    if (statusEl) statusEl.textContent = "Projects still loading — try again in a moment.";
    return;
  }
  if (btn) btn.disabled = true;
  if (statusEl) statusEl.textContent = "⏳ Scanning " + sweepResolveMailbox().label + " (all folders)…";
  if (resultsEl) resultsEl.innerHTML = "";
  try {
    const r = await sweepResolveScan(true); // fresh batch; paginates to ~target actionable across all folders
    if (statusEl) {
      statusEl.textContent =
        "Scanned " + _sweepTotals.scanned + " in " + _sweepSource.label +
        " · would auto-file " + r.file.length + " · review " + r.review.length +
        " · skip " + _sweepTotals.skip + " · already filed " + _sweepTotals.alreadyFiled +
        (_sweepCursor ? " · more await Run & file" : " · end of mailbox") +
        "  — preview only, nothing saved.";
    }
    const grp = (t) => '<div style="font-weight:600;margin:8px 0 4px;">' + t + "</div>";
    const meta = (t) => '<span style="color:#888;">' + t + "</span>";
    const num = (t) => t ? ' <span style="color:#888;font-family:Consolas,monospace;font-size:11px;">' + sweepEsc(t) + "</span>" : "";
    const rowCss = 'style="padding:4px 0;border-top:1px solid #eee;font-size:12px;"';
    const lines = [];
    if (r.file.length) {
      lines.push(grp("✅ Would auto-file (" + r.file.length + ")"));
      for (const e of r.file) {
        lines.push("<div " + rowCss + "><b>" + sweepEsc(e.project.name || e.project.projectNumber) +
          "</b>" + num(e.project.name ? e.project.projectNumber : "") + " — " + sweepEsc(e.subject) + "<br>" +
          meta(sweepEsc(e.from) + " · " + sweepEsc((e.date || "").slice(0, 10)) + " · score " + e.score) + "</div>");
      }
    }
    if (r.review.length) {
      lines.push(grp("🟡 Needs review (" + r.review.length + ")"));
      for (const e of r.review) {
        const opts = (e.candidates || [])
          .map((c) => sweepEsc(c.project.name || c.project.projectNumber) + " (" + c.score + ")").join(" · ");
        lines.push("<div " + rowCss + ">" + sweepEsc(e.subject) + "<br>" +
          meta(sweepEsc(e.from) + " · " + sweepEsc((e.date || "").slice(0, 10)) + " → " + opts) + "</div>");
      }
    }
    if (resultsEl) resultsEl.innerHTML = lines.join("") || meta("No new emails matched a project — try Load more or a different mailbox.");
  } catch (e) {
    if (statusEl) statusEl.textContent = "✗ " + humanizeError(e);
  } finally {
    if (btn) btn.disabled = false;
  }
}

function renderSweepResults(scanned, b) {
  const statusEl = document.getElementById("sweepStatus");
  const resultsEl = document.getElementById("sweepResults");
  if (statusEl) {
    statusEl.textContent =
      "Scanned " + scanned + " · would auto-file " + b.file.length +
      " · review " + b.review.length + " · skip " + b.skip +
      " · already filed " + b.alreadyFiled + "  — preview only, nothing saved.";
  }
  if (!resultsEl) return;
  const esc = (s) => String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const grp = (t) => '<div style="font-weight:600;margin:8px 0 4px;">' + t + "</div>";
  const meta = (t) => '<span style="color:#888;">' + t + "</span>";
  const rowCss = 'style="padding:4px 0;border-top:1px solid #eee;font-size:12px;"';
  const out = [];
  if (b.file.length) {
    out.push(grp("✅ Would auto-file (" + b.file.length + ")"));
    for (const e of b.file) {
      out.push("<div " + rowCss + "><b>" + esc(e.project.projectNumber || e.project.name) + "</b> — " +
        esc(e.subject) + "<br>" + meta(esc(e.from) + " · " + esc(e.date) + " · score " + e.score +
        " (" + esc((e.reasons || []).join(", ")) + ")") + "</div>");
    }
  }
  if (b.review.length) {
    out.push(grp("🟡 Needs review (" + b.review.length + ")"));
    for (const e of b.review) {
      const opts = (e.candidates || [])
        .map((c) => esc(c.project.projectNumber || c.project.name) + " (" + c.score + ")").join(" · ");
      out.push("<div " + rowCss + ">" + esc(e.subject) + "<br>" +
        meta(esc(e.from) + " · " + esc(e.date) + " → " + opts) + "</div>");
    }
  }
  resultsEl.innerHTML = out.join("") || meta("No new emails matched a project.");
}

// ─── EMAIL SWEEP v2: actual filing + review queue ─────────────────────────────
// "Run & file" auto-files the confident bucket and queues ambiguous ones for a
// one-click human confirm. Filing reuses the per-email recipe from
// _doSaveToProjectRecordOnly (getToken → fetch body → compressHtmlAddin) and ALSO
// stores stripped body_text so swept mail is full-text searchable immediately
// (the manual flow leaves body_text empty — that was the backfill root cause).
let _sweepReview = []; // ambiguous items awaiting confirm after a Run & file
const sweepEsc = (s) => String(s == null ? "" : s)
  .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

// ── Mailbox source resolution ────────────────────────────────────────────────
// Empty field (or your own address) → "/me"; a colleague's address → "/users/{addr}"
// so the same scan code can target a departed PM's shared mailbox.
function sweepResolveMailbox() {
  // Shared-mailbox entry retired (the box was removed) — the sweep always scans
  // your own mailbox. To bulk-file a colleague's mail, open their mailbox in
  // Outlook and file those emails directly.
  return { base: "/me", shared: false, label: "your mailbox" };
}
// Module-level scan source + cursor, shared across Preview / Run & file / Load more.
// _sweepSource carries the messages-collection base ("/me" or "/users/{addr}") and
// the token (a Mail.Read.Shared token for a colleague's mailbox); _sweepCursor is
// the @odata.nextLink for the next batch (null = start fresh or mailbox exhausted).
let _sweepSource = { base: "/me", token: null, label: "your mailbox", shared: false };
let _sweepCursor = null;
let _sweepTotals = { scanned: 0, skip: 0, alreadyFiled: 0, filed: 0 };
// Conversation-level dedup for the current run: convId -> projectId (the thread is
// tied to a job) or "review" (already queued once this run). Reset each fresh run
// so a thread is surfaced/auto-filed ONCE, not per-message. Persists across
// Load-more batches so a long thread doesn't re-appear page after page.
let _sweepConvoSeen = new Map();
// Auto-run ("Run to end") state: _sweepAutoRunning guards against a double start;
// _sweepAutoStop is set by the Stop button and checked between pages and batches.
let _sweepAutoRunning = false;
let _sweepAutoStop = false;

// Resolve mailbox + token + cursor, then scan one batch. reset=true starts a new
// run from the chosen mailbox; reset=false continues the saved cursor (Load more).
async function sweepResolveScan(reset) {
  if (reset) {
    const mb = sweepResolveMailbox();
    _sweepSource = { base: mb.base, label: mb.label, shared: mb.shared, token: null };
    _sweepCursor = null;
    _sweepTotals = { scanned: 0, skip: 0, alreadyFiled: 0, filed: 0 };
    _sweepConvoSeen = new Map();
  }
  // Refresh the token at the start of EVERY batch (incl. Load more) so a long
  // backfill never carries a stale token into graphFetch's 401 path — which would
  // otherwise retry a shared-mailbox call with the wrong (own-mailbox) token.
  _sweepSource.token = _sweepSource.shared ? await getSharedMailToken() : await getToken();
  const batch = await sweepScanBatch({ base: _sweepSource.base, token: _sweepSource.token, cursor: _sweepCursor });
  _sweepCursor = batch.nextLink; // null once the mailbox is fully walked
  _sweepTotals.scanned += batch.scanned;
  _sweepTotals.skip += batch.skip;
  _sweepTotals.alreadyFiled += batch.alreadyFiled;
  return batch;
}
// Fetch + classify mail into rich items carrying the Graph id + conversationId
// needed to file later. Writes nothing. Paginates within the batch — following
// @odata.nextLink verbatim, per Graph's paging rules — until enough actionable
// items surface (so skipped / already-filed messages no longer cap the run), the
// mailbox is exhausted, or the per-batch page cap is hit. Scans ALL folders
// (incl. Sent): no Focused filter, since a backfill wants the whole mailbox.
async function sweepScanBatch({ base, token, cursor }) {
  const filed = await sweepLoadFiledIds();
  const senderMap = await sweepLoadSenderMap();
  const convoMap = await sweepLoadConvoProjectMap();
  const out = { scanned: 0, file: [], review: [], skip: 0, alreadyFiled: 0, nextLink: null };
  // First page builds the query; later pages follow the (base-stripped) nextLink.
  let path = cursor || (base + "/messages?$top=" + SWEEP_PAGE_SIZE +
    "&$select=id,internetMessageId,inferenceClassification,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments" +
    "&$orderby=receivedDateTime%20desc");
  for (let page = 0; page < SWEEP_MAX_PAGES_PER_BATCH; page++) {
    const data = await graphFetch("GET", path, null, token);
    for (const m of (data?.value || [])) {
      out.scanned++;
      const sender = m.from?.emailAddress?.address || "";
      if (looksPromotional(sender)) { out.skip++; continue; } // newsletters / marketing
      const mid = m.internetMessageId || "";
      if (mid && filed.has(mid)) { out.alreadyFiled++; continue; }
      const item = {
        gid: m.id, convId: m.conversationId || "", msgId: mid,
        subject: m.subject || "(no subject)",
        from: m.from?.emailAddress?.name || sender, fromAddress: sender,
        to: (m.toRecipients || []).map((r) => r.emailAddress?.address).filter(Boolean).join(", "),
        cc: (m.ccRecipients || []).map((r) => r.emailAddress?.address).filter(Boolean).join(", "),
        date: m.receivedDateTime || "", hasAttachments: !!m.hasAttachments,
      };
      // Conversation-level dedup — one decision per thread, not per message. If the
      // thread is already tied to a job (filed in a past run, or auto-filed earlier
      // this run), file the rest of it there instead of surfacing each message. If
      // it's already queued for review this run, skip the extras so it doesn't pop
      // up over and over.
      if (item.convId) {
        const seen = _sweepConvoSeen.get(item.convId);
        const tiedProj = convoMap.get(item.convId) || (seen && seen !== "review" ? seen : null);
        if (tiedProj) {
          const proj = (allProjects || []).find((p) => p && p.id === tiedProj);
          if (proj) {
            _sweepConvoSeen.set(item.convId, tiedProj);
            out.file.push({ ...item, project: proj, score: 99, reasons: ["same thread already filed to this job"] });
            continue;
          }
        }
        if (seen === "review") { out.alreadyFiled++; continue; } // thread already queued once this run
      }
      const verdict = classifySweep(mergeSenderSignal(sweepSuggest(m.subject || ""), senderMap.get(sender.toLowerCase())));
      if (verdict.action === "skip") { out.skip++; continue; }
      if (verdict.action === "file") {
        if (item.convId) _sweepConvoSeen.set(item.convId, verdict.project.id);
        out.file.push({ ...item, project: verdict.project, score: verdict.score, reasons: verdict.reasons });
      } else {
        if (item.convId) _sweepConvoSeen.set(item.convId, "review");
        out.review.push({ ...item, candidates: verdict.candidates });
      }
    }
    // Live progress so a deep scan doesn't look frozen — refreshes once per page (~250 msgs).
    const _se = document.getElementById("sweepStatus");
    if (_se) _se.textContent = "⏳ Reading mailbox… " + out.scanned + " scanned · " +
      out.file.length + " to file · " + out.review.length + " to review…";
    const next = data?.["@odata.nextLink"] || null;
    out.nextLink = next ? next.replace("https://graph.microsoft.com/v1.0", "") : null;
    if (!out.nextLink) break;                                                  // mailbox fully walked
    if (_sweepAutoStop) break;                                                 // user hit Stop — cursor preserved so they can resume
    if (out.file.length + out.review.length >= SWEEP_TARGET_ACTIONABLE) break; // enough to act on this batch
    path = out.nextLink;
  }
  return out;
}

// Strip HTML to plain text for searchable body_text (mirrors the server side).
function stripHtmlToText(html) {
  return (html || "")
    .replace(new RegExp("<script[^>]*>[^]*?</script>", "gi"), " ")
    .replace(new RegExp("<style[^>]*>[^]*?</style>", "gi"), " ")
    .replace(new RegExp("<br[^>]*>", "gi"), "\n")
    .replace(new RegExp("</(p|div|tr|li|h[1-6])>", "gi"), "\n")
    .replace(new RegExp("<[^>]+>", "g"), " ")
    .replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<").replace(/&gt;/gi, ">")
    .replace(/&#39;/gi, "'").replace(/&apos;/gi, "'").replace(/&quot;/gi, '"')
    .replace(/ +/g, " ").replace(/\n +/g, "\n").trim();
}

// File sweep items to item.projectId: writes the search-index row (with stripped
// body_text) and dual-writes project.emails[] grouped per project. {filed, failed}.
async function sweepFileItems(items) {
  if (!items.length) return { filed: 0, failed: 0 };
  // Read bodies from whatever mailbox this batch was scanned from — a colleague's
  // shared mailbox needs its Mail.Read.Shared token + /users/{addr} message path.
  // Acquire fresh (silent, cached) so a later review-confirm can't fail on a stale
  // token routing through graphFetch's own-mailbox 401 fallback.
  const token = _sweepSource?.shared ? await getSharedMailToken() : await getToken();
  const base = _sweepSource?.base || "/me";
  const byProject = new Map();
  let filed = 0, failed = 0;
  let done = 0;
  for (const it of items) {
    done++;
    const _fe = document.getElementById("sweepStatus");
    if (_fe) _fe.textContent = "⏳ Filing email " + done + " of " + items.length + "… (" + filed + " saved)";
    try {
      let bodyHtml = "";
      try {
        const d = await graphFetch("GET", base + "/messages/" + it.gid + "?$select=body", null, token);
        bodyHtml = d?.body?.content || "";
      } catch (_) { /* body fetch is best-effort */ }
      const rec = {
        id: uid(), msgId: it.msgId,
        subject: it.subject, from: it.from, fromAddress: it.fromAddress,
        to: it.to, cc: it.cc, date: it.date,
        bodyText: bodyHtml ? stripHtmlToText(bodyHtml) : "",
        bodyHtmlCompressed: bodyHtml ? compressHtmlAddin(bodyHtml) : "",
        bodyHtmlSize: bodyHtml.length,
        hasAttachments: it.hasAttachments, attachmentNames: [],
        spFolderUrl: "", links: [],
        savedAt: new Date().toISOString(), savedBy: _getCurrentUserEmail() || "",
        savedToSharePoint: false,
      };
      await saveProjectEmailRow(it.projectId, rec, false, it.convId || null);
      if (!byProject.has(it.projectId)) byProject.set(it.projectId, []);
      byProject.get(it.projectId).push(rec);
      filed++;
    } catch (e) {
      console.warn("sweep file failed:", e);
      failed++;
    }
  }
  // Keep project.emails[] in sync — one save per project (avoids version churn).
  for (const [pid, recs] of byProject) {
    try {
      await applyLocalChangeAndSave(pid, (fresh) => {
        // Dedup against FRESH data, mirroring the DB's (project_id, msg_id)
        // uniqueness: drop any copy already filed here since our last sync so the
        // JSON array can't drift even though the table now can't.
        const have = new Set((fresh.emails || []).map(e => e.msgId).filter(Boolean));
        const add = recs.filter(r => !r.msgId || !have.has(r.msgId));
        return add.length ? { ...fresh, emails: [...(fresh.emails || []), ...add] } : fresh;
      });
    } catch (e) {
      console.warn("sweep dual-write to project.emails failed for", pid, e);
    }
  }
  return { filed, failed };
}

// RUN & FILE — auto-file the confident bucket, queue the ambiguous ones.
async function sweepRunAndFile(reset = true) {
  const btn = document.getElementById("sweepFileBtn");
  const moreBtn = document.getElementById("sweepMoreBtn");
  const statusEl = document.getElementById("sweepStatus");
  if (!allProjects || !allProjects.length) { if (statusEl) statusEl.textContent = "Projects still loading…"; return; }
  if (btn) btn.disabled = true;
  if (moreBtn) moreBtn.disabled = true;
  if (statusEl) statusEl.textContent = reset ? "⏳ Scanning and filing confident matches…" : "⏳ Loading the next batch…";
  try {
    if (reset) _sweepReview = [];
    const r = await sweepResolveScan(reset);
    for (const it of r.file) it.projectId = it.project.id;
    const { filed, failed } = await sweepFileItems(r.file);
    _sweepTotals.filed += filed;
    _sweepReview.push(...r.review);
    const pending = _sweepReview.filter((e) => !e._done && !(e._filedTo && e._filedTo.length)).length;
    if (statusEl) statusEl.textContent =
      "✓ Filed " + _sweepTotals.filed + (failed ? (" · " + failed + " failed this batch") : "") +
      " · " + pending + " to review · scanned " + _sweepTotals.scanned +
      " · skipped " + _sweepTotals.skip + " · already filed " + _sweepTotals.alreadyFiled +
      " · [" + _sweepSource.label + "]" + (_sweepCursor ? " · more available" : " · end of mailbox");
    renderSweepReviewQueue();
  } catch (e) {
    if (statusEl) statusEl.textContent = "✗ " + humanizeError(e);
  } finally {
    if (btn) btn.disabled = false;
    updateSweepMoreBtn();
  }
}
// Show the Load-more button only while the chosen mailbox still has pages left.
function updateSweepMoreBtn() {
  const moreBtn = document.getElementById("sweepMoreBtn");
  if (!moreBtn) return;
  moreBtn.style.display = _sweepCursor ? "" : "none";
  moreBtn.disabled = false;
}

// Swap the run/more buttons for a Stop button while an auto-run is going.
function _setSweepAutoUI(running) {
  for (const id of ["sweepFileBtn", "sweepRunAllBtn", "sweepMoreBtn"]) {
    const el = document.getElementById(id);
    if (el) el.style.display = running ? "none" : "";
  }
  const stop = document.getElementById("sweepStopBtn");
  if (stop) stop.style.display = running ? "" : "none";
}

// "Run to end" — keep filing batches until the mailbox is exhausted or the user
// hits Stop. sweepRunAndFile already shows per-page/per-email progress and the
// cumulative totals after each batch, so the numbers climb as it goes. We reassert
// the auto UI after each batch because sweepRunAndFile re-enables the normal
// buttons in its own finally block.
async function sweepAutoRunToEnd() {
  if (_sweepAutoRunning) return;
  _sweepAutoRunning = true;
  _sweepAutoStop = false;
  _setSweepAutoUI(true);
  try {
    await sweepRunAndFile(true);          // first batch — fresh scan from the top
    _setSweepAutoUI(true);
    while (_sweepCursor && !_sweepAutoStop) {
      await sweepRunAndFile(false);       // next batch, continuing the cursor
      _setSweepAutoUI(true);
    }
    const s = document.getElementById("sweepStatus");
    if (s) s.textContent += _sweepAutoStop ? "  ·  ⏹ stopped" : "  ·  ✅ reached the end of the mailbox";
  } catch (e) {
    const s = document.getElementById("sweepStatus");
    if (s) s.textContent = "✗ " + humanizeError(e);
  } finally {
    _sweepAutoRunning = false;
    _setSweepAutoUI(false);
    updateSweepMoreBtn();
  }
}

// Review queue — each ambiguous email gets a button per candidate project. Filing
// to one does NOT close the row: an email can belong to two jobs, so once filed it
// collapses to a ✓ confirmation with "＋ also file to another job", which re-opens
// the candidates to copy the SAME email into a second project. One file resolves
// the item (drops it from the "left" count); Skip dismisses an unmatched one.
function renderSweepReviewQueue() {
  const resultsEl = document.getElementById("sweepResults");
  if (!resultsEl) return;
  const isFiled = (e) => !!(e._filedTo && e._filedTo.length);
  const needsAttention = (e) => !e._done && !isFiled(e);
  const visible = _sweepReview.filter((e) => e._done !== "skip"); // hide skipped; keep ✓ confirmations
  if (!visible.length) { resultsEl.innerHTML = '<span style="color:#888;">Review queue clear.</span>'; return; }
  const bcss = 'style="margin:2px 4px 2px 0;padding:2px 8px;font-size:12px;cursor:pointer;"';
  const rows = ['<div style="font-weight:600;margin:8px 0 4px;">🟡 Review (' + _sweepReview.filter(needsAttention).length + ' left)</div>'];
  _sweepReview.forEach((e, i) => {
    if (e._done === "skip") return;
    const subjLine = sweepEsc(e.subject) + '<br><span style="color:#888;">' +
      sweepEsc(e.from) + " · " + sweepEsc((e.date || "").slice(0, 10)) + "</span>";
    if (e._done === "filing") {
      rows.push('<div style="padding:6px 0;border-top:1px solid #eee;font-size:12px;">' + subjLine +
        '<br><span style="color:#888;">⏳ filing…</span></div>');
      return;
    }
    const filed = e._filedTo || [];
    const filedNames = filed.map((f) => f.name).join(", ");
    const filedIds = new Set(filed.map((f) => f.id));
    const moreToFile = filed.length < (e.candidates || []).length;
    // Filed and not re-expanded → compact ✓ confirmation + opt-in to copy elsewhere.
    if (filed.length && !e._expand) {
      rows.push('<div style="padding:6px 0;border-top:1px solid #eee;font-size:12px;opacity:0.9;">' + subjLine +
        '<br><span style="color:#107c10;">✓ Filed to ' + sweepEsc(filedNames) + "</span>" +
        (moreToFile ? ' <button type="button" data-sw="expand" data-i="' + i + '" ' + bcss + ">＋ also file to another job</button>" : "") +
        "</div>");
      return;
    }
    // Buttons lead with the project NAME — people recognize "CHCR Grove and New
    // Dorp", not "SAPX226014.00". The number rides along as a muted subtitle so
    // same-named candidates (e.g. two phases of one job) stay distinguishable.
    // A project already filed shows as a disabled green ✓ chip.
    const cands = (e.candidates || []).map((c, ci) => {
      const name = (c.project.name || "").trim();
      const num  = (c.project.projectNumber || "").trim();
      const primary  = name || num;                 // always show something
      const subtitle = (name && num) ? num : "";    // number as a subtitle, never the sole label
      const already  = filedIds.has(c.project.id);
      const style = "margin:2px 4px 2px 0;padding:3px 8px;font-size:12px;text-align:left;line-height:1.25;vertical-align:top;" +
        (already ? "cursor:default;background:#dff6dd;border:1px solid #9fd89f;color:#0b5a0b;" : "cursor:pointer;");
      const opener = already
        ? '<button type="button" disabled style="' + style + '">'
        : '<button type="button" data-sw="file" data-i="' + i + '" data-ci="' + ci + '" style="' + style + '">';
      return opener +
        '<span style="font-weight:600;">' + (already ? "✓ " : "") + sweepEsc(primary) + "</span>" +
        (subtitle ? '<br><span style="color:#888;font-size:11px;font-family:Consolas,monospace;">' + sweepEsc(subtitle) + "</span>" : "") +
        "</button>";
    }).join("");
    const lead = filed.length
      ? '<span style="color:#107c10;">✓ Filed to ' + sweepEsc(filedNames) + "</span> · also file to:"
      : "File to:";
    const tail = filed.length
      ? '<button type="button" data-sw="done" data-i="' + i + '" ' + bcss + ">✓ Done</button>"
      : '<button type="button" data-sw="skip" data-i="' + i + '" ' + bcss + ">Skip</button>";
    rows.push('<div style="padding:6px 0;border-top:1px solid #eee;font-size:12px;">' + subjLine +
      "<br>" + lead + " " + cands + tail + "</div>");
  });
  resultsEl.innerHTML = rows.join("");
  resultsEl.querySelectorAll("button[data-sw]").forEach((b) => {
    b.onclick = () => {
      const i = +b.dataset.i; const e = _sweepReview[i]; if (!e) return;
      switch (b.dataset.sw) {
        case "skip":   e._done = "skip";  renderSweepReviewQueue(); break;
        case "done":   e._expand = false; renderSweepReviewQueue(); break;  // collapse to the ✓ confirmation
        case "expand": e._expand = true;  renderSweepReviewQueue(); break;  // re-open candidates to copy elsewhere
        case "file":   void sweepConfirmReview(i, +b.dataset.ci); break;
      }
    };
  });
}

async function sweepConfirmReview(i, ci) {
  const e = _sweepReview[i];
  if (!e || e._done === "skip" || e._done === "filing") return;
  const proj = e.candidates?.[ci]?.project;
  if (!proj) return;
  e._filedTo = e._filedTo || [];
  if (e._filedTo.some((f) => f.id === proj.id)) return; // don't double-file the same job
  e._done = "filing";
  renderSweepReviewQueue();
  const { filed } = await sweepFileItems([{ ...e, projectId: proj.id }]);
  e._done = null;
  if (filed) {
    e._filedTo.push({ id: proj.id, name: (proj.name || proj.projectNumber || "project") });
    e._expand = false; // collapse to the ✓ confirmation; "＋ also file" re-expands to copy to another job
  } else {
    const s = document.getElementById("sweepStatus"); if (s) s.textContent = "✗ Filing failed — try again.";
  }
  renderSweepReviewQueue();
}

function renderProjectSuggestions() {
  const block = document.getElementById("suggestionBlock");
  const chips = document.getElementById("suggestionChips");
  const labelText = document.getElementById("suggestionLabelText");
  if (!block || !chips) return;
  if (selectedProject) { block.style.display = "none"; chips.innerHTML = ""; return; }

  // First email of the session: load the learned sender signal, then re-rank the
  // chips once it lands (a known sender can surface a job with no subject match).
  // No-op on later calls — the cache is warm and read synchronously by suggestProjects.
  ensureSenderSignalCache(renderProjectSuggestions);

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
      const proj = getProjectById(el.dataset.id);
      if (proj) setSelectedProject(proj, true);
    };
  });
  block.style.display = "block";
}

// ─── RESPONSE-NEEDED CLASSIFIER ──────────────────────────────────────────────
// Scores a CLIENT email by how likely it is to need a reply from us. Same
// weighted-signal shape as suggestProjects(): accumulate points, collect the
// reasons that fired, flag when the total clears RESPONSE_MIN_SCORE. We never
// auto-act — this only decides "is this worth putting on the watchlist".
//
// TUNE: the phrase lists below are the part to adjust after real-world use.
// Setty clients use a mix of blunt and polite phrasing, so MIN_SCORE is low
// (3) — a single softer signal can still flag. Raise it if the panel gets
// noisy; expand RESPONSE_POLITE_PHRASES if soft asks slip through.
const RESPONSE_WEIGHTS = {
  questionMarkInSubject:  4,
  questionMarkInBody:     2,
  interrogativeOpener:    3,
  politeRequest:          2,
  hasDueDate:             3,
  urgencyWord:            2,
  addressedToMe:          2,   // signed-in user is in the To field, not just CC
  namedInSalutation:      3,   // the greeting names the signed-in user
  exclusionPhrase:       -5,   // explicit opt-out beats inferred signals
};
const RESPONSE_MIN_SCORE = 3;

// Sentence-opening interrogatives — matched only at the start of a sentence so
// a mid-sentence "how" ("…explained how we did it") doesn't count.
const RESPONSE_INTERROGATIVES = [
  "can you", "could you", "would you", "will you", "do you", "are you",
  "is it", "have you", "when ", "what ", "where ", "how ", "why ",
  "which ", "who ",
];
const RESPONSE_POLITE_PHRASES = [
  "please confirm", "please advise", "please let me know", "please send",
  "please provide", "let me know", "let us know", "get back to me",
  "your thoughts", "any update", "awaiting your", "looking forward to your",
  "circle back", "need your",
];
const RESPONSE_URGENCY_WORDS = [
  "asap", "urgent", "time-sensitive", "time sensitive", "by eod",
  "end of day", "no later than", "right away", "first thing",
];
const RESPONSE_EXCLUSIONS = [
  "no response needed", "no reply needed", "no action required",
  "no action needed", "for your records", "just an fyi", "no need to reply",
];

// Strip URLs so a "?utm=…" query string doesn't read as a question.
function stripUrlsForScan(s) {
  return (s || "").replace(/https?:\/\/\S+/gi, " ");
}

// subject, bodyText: raw strings from the email. emailReceivedDate: Date|string
// used by extractDueDates to resolve year-less dates. Returns the same
// { score, reasons } shape as suggestProjects(), plus a needsReply boolean and
// any ISO dates found (so the watchlist row can store a due date).
function needsResponse(subject, bodyText, emailReceivedDate, opts = {}) {
  const subj = (subject || "").toLowerCase();
  const body = stripUrlsForScan(trimToCurrentMessage(bodyText || "")).toLowerCase();
  let score = 0;
  const reasons = [];
  // A "content signal" is real evidence of an ASK — a question, a request, a
  // deadline. The recipient signals (addressed to you / named in greeting) are
  // NOT content: a thank-you note is still addressed to you. needsReply
  // requires at least one content signal, so recipient signals can only
  // amplify a genuine ask — never originate a flag on their own.
  let contentSignal = false;

  if (subj.includes("?")) {
    score += RESPONSE_WEIGHTS.questionMarkInSubject;
    reasons.push("question in subject");
    contentSignal = true;
  }
  if (body.includes("?")) {
    score += RESPONSE_WEIGHTS.questionMarkInBody;
    reasons.push("question in body");
    contentSignal = true;
  }
  // Interrogative openers — split into sentences, check start-of-sentence only.
  const sentences = body.split(/[.!?\n]+/).map(s => s.trim()).filter(Boolean);
  if (sentences.some(s => RESPONSE_INTERROGATIVES.some(q => s.startsWith(q)))) {
    score += RESPONSE_WEIGHTS.interrogativeOpener;
    reasons.push("direct question");
    contentSignal = true;
  }
  if (RESPONSE_POLITE_PHRASES.some(p => body.includes(p) || subj.includes(p))) {
    score += RESPONSE_WEIGHTS.politeRequest;
    reasons.push("request phrase");
    contentSignal = true;
  }
  // Reuse the existing date detector — any due date is a deadline signal.
  const dueDates = extractDueDates(bodyText || "", emailReceivedDate);
  if (dueDates.length) {
    score += RESPONSE_WEIGHTS.hasDueDate;
    reasons.push("mentions a date");
    contentSignal = true;
  }
  if (RESPONSE_URGENCY_WORDS.some(w => body.includes(w) || subj.includes(w))) {
    score += RESPONSE_WEIGHTS.urgencyWord;
    reasons.push("urgency");
    contentSignal = true;
  }
  // Recipient signals — computed by the caller (needs the recipient list and
  // the signed-in user's name, which a pure subject+body function can't see).
  // These are boosters only: they add weight but never set contentSignal.
  if (opts.addressedToMe) {
    score += RESPONSE_WEIGHTS.addressedToMe;
    reasons.push("addressed to you");
  }
  if (opts.namedInSalutation) {
    score += RESPONSE_WEIGHTS.namedInSalutation;
    reasons.push("greets you by name");
  }
  if (RESPONSE_EXCLUSIONS.some(x => body.includes(x) || subj.includes(x))) {
    score += RESPONSE_WEIGHTS.exclusionPhrase;
    reasons.push("explicit no-reply");
  }

  return {
    score,
    reasons,
    // Both gates: a genuine ask must be present AND the weighted score must
    // clear the bar. A thank-you note addressed to you fails the first gate.
    needsReply: contentSignal && score >= RESPONSE_MIN_SCORE,
    dueDates: dueDates.map(d => d.iso),
  };
}

// ─── RESPONSE WATCHLIST — DATA LAYER ─────────────────────────────────────────
// A client email that needsResponse() flags gets one row in the shared
// pms_email_watchlist table, keyed on conversationId. Whether the row is still
// "open" is NOT trusted from this table — it's re-verified against Graph at
// render time (isThreadAwaitingReply). The table is just the watchlist.

// Identifies the current user for watchlist ownership. Used for BOTH the
// added_by written on insert and the filter on read, so they always agree —
// the panel only shows emails this user flagged.
function watchlistUserKey() {
  return msalAccount?.username || msalAccount?.name || "unknown";
}

// Looks up THIS user's row for a thread. Rows are keyed (conversation_id,
// added_by), so a colleague's row for the same thread is invisible here — that
// is what lets two people track the same thread independently.
async function getMyWatchlistRow(conversationId) {
  if (!conversationId) return null;
  const url = SUPABASE_URL + "/rest/v1/" + EMAIL_WATCHLIST_TABLE +
    "?conversation_id=eq." + encodeURIComponent(conversationId) +
    "&added_by=eq." + encodeURIComponent(watchlistUserKey()) +
    "&select=conversation_id,status&limit=1";
  try {
    const res = await fetch(url, { headers: SB_HEADERS });
    if (!res.ok) return null;
    const rows = await res.json();
    return rows?.[0] || null;
  } catch { return null; }
}

// True if ANY Setty user's row for this thread is 'dismissed'. Used to honor a
// shared dismiss: once someone clears a thread, it stays cleared for everyone —
// including a colleague who only opens the email afterward.
async function isConversationDismissed(conversationId) {
  if (!conversationId) return false;
  const url = SUPABASE_URL + "/rest/v1/" + EMAIL_WATCHLIST_TABLE +
    "?conversation_id=eq." + encodeURIComponent(conversationId) +
    "&status=eq.dismissed&select=conversation_id&limit=1";
  try {
    const res = await fetch(url, { headers: SB_HEADERS });
    if (!res.ok) return false;
    return ((await res.json())?.length || 0) > 0;
  } catch { return false; }
}

// Inserts a watchlist row only if THIS user isn't already tracking the thread.
// This is what makes the classify-on-open hook safe to fire on every email
// view: re-opening a thread you already dismissed (or that's been answered)
// must NOT resurrect your row, and a thread any colleague has dismissed stays
// dismissed for everyone. Returns true only when a new row was created.
async function addWatchlistEntryIfNew(entry) {
  if (!entry?.conversation_id) return false;
  if (await getMyWatchlistRow(entry.conversation_id)) return false;
  if (await isConversationDismissed(entry.conversation_id)) return false;
  try {
    const res = await fetch(SUPABASE_URL + "/rest/v1/" + EMAIL_WATCHLIST_TABLE, {
      method: "POST",
      headers: { ...SB_HEADERS, "Prefer": "return=minimal" },
      body: JSON.stringify(entry),
    });
    if (!res.ok) {
      console.warn("Watchlist insert failed:", res.status, await res.text());
      return false;
    }
    return true;
  } catch (e) {
    console.warn("Watchlist insert failed:", e);
    return false;
  }
}

async function getWatchlistOpenItems() {
  const url = SUPABASE_URL + "/rest/v1/" + EMAIL_WATCHLIST_TABLE +
    "?status=eq.watching&added_by=eq." + encodeURIComponent(watchlistUserKey()) +
    "&select=*&order=added_at.asc";
  try {
    const res = await fetch(url, { headers: SB_HEADERS });
    if (!res.ok) return [];
    return (await res.json()) || [];
  } catch { return []; }
}

// Updates watchlist rows for a thread. By default only THIS user's row (the
// 'answered' flip is per-row — but behaves as shared because every participant's
// row independently flips on their next Graph check). Pass { allUsers: true } to
// update EVERY Setty row for the thread at once — that's how a dismiss clears the
// flag for every colleague who had the same email flagged, not just the clicker.
async function setWatchlistStatus(conversationId, status, opts = {}) {
  if (!conversationId) return;
  let url = SUPABASE_URL + "/rest/v1/" + EMAIL_WATCHLIST_TABLE +
    "?conversation_id=eq." + encodeURIComponent(conversationId);
  if (!opts.allUsers) {
    url += "&added_by=eq." + encodeURIComponent(watchlistUserKey());
  }
  try {
    await fetch(url, {
      method: "PATCH",
      headers: { ...SB_HEADERS, "Prefer": "return=minimal" },
      body: JSON.stringify({ status, last_checked_at: new Date().toISOString() }),
    });
  } catch (e) { console.warn("Watchlist status update failed:", e); }
}

// The "answered?" check. Asks Graph for the single most recent message in the
// thread: if its sender is outside Setty, we still owe a reply. A reply from
// ANY @setty.com address — including a colleague's — counts as answered, so a
// teammate handling it clears the item for everyone. Returns { awaiting, since }
// where `since` is the latest message's timestamp (ms) — when awaiting, that's
// the unanswered client message, i.e. when our response clock starts. Falls back
// to { awaiting: true, since: null } when Graph can't tell, so a flagged email
// is never silently lost.
async function isThreadAwaitingReply(conversationId) {
  if (!conversationId) return { awaiting: true, since: null, webLink: null };
  // Fetch the conversation's messages and pick the newest CLIENT-SIDE. We must
  // NOT add $orderby here: Graph rejects $filter + $orderby on /me/messages when
  // they reference different properties (conversationId vs receivedDateTime),
  // returning a 400 — which would land in the catch below and make every thread
  // look permanently unanswered. Sorting in JS sidesteps the limitation.
  const filter = "conversationId eq '" + conversationId.replace(/'/g, "''") + "'";
  const path = "/me/messages?$filter=" + encodeURIComponent(filter) +
    "&$top=50&$select=from,receivedDateTime,sentDateTime,webLink";
  try {
    const data = await graphFetch("GET", path);
    const msgs = data?.value || [];
    if (!msgs.length) return { awaiting: true, since: null, webLink: null };
    const when = m => new Date(m.receivedDateTime || m.sentDateTime || 0).getTime();
    msgs.sort((a, b) => when(b) - when(a));
    const latestSender = (msgs[0].from?.emailAddress?.address || "").toLowerCase();
    return {
      awaiting: !latestSender.endsWith("@setty.com"),   // latest from Setty = answered
      since: when(msgs[0]) || null,
      // The latest message's link + id — used as fallback "open original"
      // targets for rows flagged before web_link/item_id were stored. restId is
      // the Graph id (always returned); the render converts it to an EWS id.
      webLink: msgs[0].webLink || null,
      restId: msgs[0].id || null,
    };
  } catch (e) {
    console.warn("[watchlist] thread check failed:", e?.message || e);
    return { awaiting: true, since: null, webLink: null };   // couldn't tell — keep watching
  }
}

// Classify-on-open hook. Runs when a client email is viewed: scores it, and if
// it looks like it needs a reply, adds it to the watchlist. Best-effort and
// silent — any failure here must never block the rest of the taskpane.
async function maybeAddToWatchlist(myGen) {
  try {
    if (currentItemKind !== "message" || !emailItem) return;
    const from = (emailFromAddress || "").toLowerCase();
    if (!from) return;
    // Hard gate: never watch internal mail. A reply we owe is ALWAYS to an
    // outside party, so anything from a @setty.com address is out — including a
    // colleague's message inside a thread. This guard comes first because the
    // getClientByEmail domain fallback below could otherwise match a @setty.com
    // sender if a colleague was ever saved under a client's contact record.
    if (isSettyInternalEmail(from)) return;
    // Gate: the sender must be a company on the global client list. This is
    // what keeps automated mail (Microsoft 365 notifications, newsletters,
    // vendors) off the watchlist — getClientByEmail matches a known client by
    // contact address or shared email domain.
    const client = getClientByEmail(emailFromAddress);
    if (!client) return;

    const token = await getToken();
    const html  = await getEmailBodyHtml(token);
    if (myGen !== itemContextGeneration) return;          // user moved on
    const tmp = document.createElement("div");
    tmp.innerHTML = html || "";
    const text = (tmp.innerText || tmp.textContent || "").replace(/\s+/g, " ");

    // Recipient signals — an email sent directly TO you, or one whose greeting
    // names you, is far more likely to need your reply than a mass CC. Computed
    // here because needsResponse() can't see the recipient list or your name.
    const myAddr = (msalAccount?.username || "").toLowerCase();
    const addressedToMe = !!myAddr &&
      (emailItem.to || []).some(r => (r.emailAddress || "").toLowerCase() === myAddr);
    const myFirstName = (msalAccount?.name || "").trim().split(/\s+/)[0].replace(/[^\w]/g, "");
    const namedInSalutation = myFirstName.length >= 2 &&
      new RegExp("\\b" + myFirstName + "\\b", "i").test(text.slice(0, 60));

    const subject = (typeof emailItem.subject === "string") ? emailItem.subject : "";
    const verdict = needsResponse(subject, text, emailItem.dateTimeCreated,
      { addressedToMe, namedInSalutation });
    if (!verdict.needsReply) return;

    const conversationId = await getCurrentConversationId();
    if (!conversationId || myGen !== itemContextGeneration) return;

    // Don't flag a thread Setty has already handled: if the most recent message
    // in the conversation is from a @setty.com address, someone has replied and
    // we owe nothing — so it never goes on the watchlist in the first place. The
    // render-time check still enforces the 24-business-hour grace before a
    // genuinely unanswered thread actually surfaces in the panel.
    const threadState = await isThreadAwaitingReply(conversationId);
    if (myGen !== itemContextGeneration) return;
    if (!threadState.awaiting) return;

    // Best-effort deep link back to this email — lets the panel row reopen the
    // original. Captured here (not at render) because it needs the live item.
    const webLink = await getCurrentMessageWebLink();
    if (myGen !== itemContextGeneration) return;

    const added = await addWatchlistEntryIfNew({
      conversation_id: conversationId,
      subject: subject || "(No subject)",
      client_email: emailFromAddress,
      client_name: client.name || emailFrom || "",
      project_id: selectedProject?.id || null,
      due_date: verdict.dueDates[0] || null,
      score: verdict.score,
      reasons: verdict.reasons,
      web_link: webLink || null,
      item_id: emailItem?.itemId || null,   // EWS id → native open in desktop Outlook
      added_by: watchlistUserKey(),
    });
    if (added) void renderResponseWatchlist();
  } catch (e) {
    console.warn("[watchlist] add check failed:", e?.message || e);
  }
}

// Renders the "needs a reply" panel: fetches open rows, verifies each thread
// against Graph (flipping answered ones to 'answered'), shows the rest oldest
// first. Called on add-in load and from the panel's ↻ button — deliberately
// NOT on every email open, since each render costs one Graph call per item.
let _watchlistRendering = false;
async function renderResponseWatchlist() {
  const block = document.getElementById("responseWatchlistBlock");
  const list  = document.getElementById("responseWatchlistList");
  const count = document.getElementById("responseWatchlistCount");
  if (!block || !list || _watchlistRendering) return;
  _watchlistRendering = true;
  try {
    const open = await getWatchlistOpenItems();
    // Verify each thread in parallel — the open list is small in practice.
    // TODO if it grows past ~15 items: switch to a single Graph $batch request.
    const states = await Promise.all(open.map(r => isThreadAwaitingReply(r.conversation_id)));
    const now = Date.now();
    const stillOpen = [];
    open.forEach((row, i) => {
      const { awaiting, since, webLink, restId } = states[i];
      if (!awaiting) { void setWatchlistStatus(row.conversation_id, "answered"); return; }
      // Awaiting a reply — but hold it back until a full business day (24h minus
      // weekends) has passed without a response. Measure from the latest client
      // message; fall back to when we first flagged it if Graph couldn't date
      // the thread. Below the threshold, the row stays 'watching' — just hidden.
      const startedAt = since ?? new Date(row.added_at).getTime();
      if (businessMsBetween(startedAt, now) < RESPONSE_GRACE_MS) return;
      // Backfill "open original" targets for rows flagged before web_link/item_id
      // were stored, using the latest message from the thread check above. The
      // EWS id (for native desktop open) is converted from the Graph id.
      if (!row.web_link && webLink) row.web_link = webLink;
      if (!row.item_id && restId) {
        try { row.item_id = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0); } catch { /* keep web_link fallback */ }
      }
      stillOpen.push(row);
    });
    if (!stillOpen.length) { block.style.display = "none"; list.innerHTML = ""; return; }
    if (count) count.textContent = String(stillOpen.length);
    list.innerHTML = stillOpen.map(r => {
      const proj    = getProjectById(r.project_id);
      const reasons = Array.isArray(r.reasons) ? r.reasons.join(" · ") : "";
      const due     = r.due_date
        ? `<span class="wl-due">due ${escHtml(new Date(r.due_date + "T12:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric" }))}</span>`
        : "";
      const subjHtml = (r.web_link || r.item_id)
        ? `<button type="button" class="wl-subject wl-open" data-itemid="${escHtml(r.item_id || "")}" data-url="${escHtml(r.web_link || "")}" title="Open the original email">${escHtml(r.subject || "(No subject)")}</button>`
        : `<div class="wl-subject">${escHtml(r.subject || "(No subject)")}</div>`;
      return `
        <div class="wl-item">
          <div class="wl-main">
            ${subjHtml}
            <div class="wl-meta">${escHtml(r.client_name || r.client_email || "")}${proj ? " · " + escHtml(proj.name) : ""} · flagged ${escHtml(_relativeTime(r.added_at))}</div>
            <div class="wl-reasons">${escHtml(reasons)} ${due}</div>
          </div>
          <button type="button" class="wl-dismiss" data-cid="${escHtml(r.conversation_id)}" title="Not awaiting a reply — dismiss">×</button>
        </div>`;
    }).join("");
    list.querySelectorAll(".wl-open").forEach(el => {
      el.onclick = () => openWatchlistEmail(el.dataset.itemid, el.dataset.url);
    });
    list.querySelectorAll(".wl-dismiss").forEach(el => {
      el.onclick = () => {
        // Optimistic UI: drop the chip immediately so the click feels instant.
        // The Supabase write runs in the background — a dismiss is low-stakes,
        // so we don't block on it or roll back on failure (a failed write just
        // means the chip reappears on the next reload, and the user re-clicks).
        const cid  = el.dataset.cid;
        const item = el.closest(".wl-item");
        if (item) item.remove();
        const remaining = list.querySelectorAll(".wl-item").length;
        if (count) count.textContent = String(remaining);
        if (!remaining) block.style.display = "none";
        // Shared dismiss: clear this thread for every Setty colleague who flagged
        // it, not just me — keeps the panel quiet across the team.
        void setWatchlistStatus(cid, "dismissed", { allUsers: true });
      };
    });
    block.style.display = "block";
  } finally {
    _watchlistRendering = false;
  }
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

// ── DIRECTORY-STATUS LOOKUPS (people picker) ─────────────────────────────────
// Exact email matches only — getClientByEmail's domain fallback above is for
// guessing the company on the contact form, not for claiming a person is
// already saved. All checks run against in-memory data (allClients +
// selectedProject), so they cost nothing per render.
function findGlobalContact(email) {
  const emailLc = (email || "").trim().toLowerCase();
  if (!emailLc) return null;
  for (const c of (allClients || [])) {
    const hit = (c.contacts || []).find(ct => (ct.email || "").trim().toLowerCase() === emailLc);
    if (hit) return { client: c, contact: hit };
  }
  return null;
}
// "Already on this job" = in project.directory (contact-save path) OR in the
// project POC list (save-to-project path) — the two places the add-in writes.
function isInProjectDirectory(email) {
  const emailLc = (email || "").trim().toLowerCase();
  if (!emailLc || !selectedProject) return false;
  if ((selectedProject.directory || []).some(d => (d.email || "").trim().toLowerCase() === emailLc)) return true;
  return ((selectedProject.projectContacts?.pm) || []).some(c => (c.email || "").trim().toLowerCase() === emailLc);
}
// "Already on SOME job" — is this email in ANY project's directory or PM list,
// not just the selected one. Memoized against the allProjects reference
// (loadProjects reassigns the array on refresh), so the set rebuilds only when
// the project data changes.
let _allDirEmailsCache = { ref: null, set: new Set() };
function isInAnyProjectDirectory(email) {
  const emailLc = (email || "").trim().toLowerCase();
  if (!emailLc) return false;
  if (_allDirEmailsCache.ref !== allProjects) {
    const set = new Set();
    for (const p of (allProjects || [])) {
      for (const d of (p.directory || [])) {
        const e = (d.email || "").trim().toLowerCase();
        if (e) set.add(e);
      }
      for (const c of ((p.projectContacts?.pm) || [])) {
        const e = (c.email || "").trim().toLowerCase();
        if (e) set.add(e);
      }
    }
    _allDirEmailsCache = { ref: allProjects, set };
  }
  return _allDirEmailsCache.set.has(emailLc);
}
// Setty staff are managed via the project's Teams tab in PMS, not the
// contact directories — so they're excluded from the "new" nudge entirely.
function isSettyInternalEmail(email) {
  return (email || "").trim().toLowerCase().endsWith("@setty.com");
}
// Enrichment opportunity: the SENDER is already filed (global directory or this
// project's directory) but their saved record has NEITHER a title NOR a phone —
// both of which this email's signature can fill in one tap. We require both to
// be blank on purpose: titles scrape unreliably (the parser often grabs the firm
// name), so a title-only gap isn't worth nagging about. Returns the saved record
// when it's enrichable, otherwise null.
function senderNeedsEnrichment() {
  const addr = (emailFromAddress || "").trim().toLowerCase();
  if (!addr) return null;
  // Setty staff are managed via the project's Teams tab, not the contact
  // directories (same exclusion getParticipantDirectoryStatus applies). Never
  // nudge to enrich an internal sender from their signature — even if a stale
  // @setty.com record lingers in a directory from before that rule existed.
  if (isSettyInternalEmail(addr)) return null;
  // Saved/enriched this session — stop nudging immediately, even before the
  // in-memory directory cache reflects the backfilled title/phone.
  if (_sessionSavedContactEmails.has(addr)) return null;
  const rec = findGlobalContact(addr)?.contact
    || (selectedProject?.directory || []).find(d => (d.email || "").trim().toLowerCase() === addr)
    || null;
  if (!rec) return null;
  const noTitle = !(rec.title || "").trim();
  const noPhone = !(rec.phone || "").trim();
  return (noTitle && noPhone) ? rec : null;
}
function getParticipantDirectoryStatus(p) {
  const emailLc = (p?.emailAddress || "").trim().toLowerCase();
  const internal = isSettyInternalEmail(emailLc);
  const sessionSaved = !!emailLc && _sessionSavedContactEmails.has(emailLc);
  const globalHit = internal ? null : findGlobalContact(emailLc);
  const inProject = internal ? false : isInProjectDirectory(emailLc);
  // Filed on SOME job's directory even if not in the global directory or this
  // project — still "known", so it must not read as a brand-new contact. Stops
  // the nudge firing for people you've already got on another project. (inProject
  // ⊂ inAnyProject, but the selected project may be a fresher object than the
  // cached allProjects copy, so check both.)
  const inAnyProject = internal ? false : (inProject || isInAnyProjectDirectory(emailLc));
  // "New" = external, nowhere in the firm's directories, and not just saved
  // this session.
  return { internal, globalHit, inProject, inAnyProject, sessionSaved, isNew: !internal && !globalHit && !inProject && !inAnyProject && !sessionSaved };
}
// Nudge on the main-view button: surface how many of this email's
// participants aren't in any directory yet, BEFORE the user opens the list.
// When there's nothing actionable at all — every participant is Setty staff
// or already filed everywhere relevant — the button disappears entirely.
// "Actionable" includes the global-but-not-on-this-project case, since the
// people view is where the one-click "+ project" add lives.
// Single source of truth for why the main-view button is showing, shared by the
// badge renderer and the click handler so they never disagree about whether
// this is an "enrich the sender from signature" click vs. the participant list.
function peopleButtonMode() {
  const statuses = (emailParticipants || []).map(getParticipantDirectoryStatus);
  const newCount = statuses.filter(s => s.isNew).length;
  // "+ project" only makes sense once the email is tagged to a project — without
  // one there's nowhere to add them, so a global-directory hit alone must NOT
  // keep the button on screen. (Previously this counted regardless of project,
  // which is why the button always showed on untagged emails.)
  const addableToProject = selectedProject
    ? statuses.filter(s => s.globalHit && !s.inProject && !s.sessionSaved).length
    : 0;
  // Even when everyone's already filed, surface the button if the sender's
  // record can be enriched from this email's signature.
  const enrichRec = senderNeedsEnrichment();
  // "Enrich-only" = the signature opportunity is the SOLE reason to show — no new
  // contacts to capture, nothing to add to a project. That's when the button
  // jumps straight to the form and relabels itself.
  const enrichOnly = newCount === 0 && addableToProject === 0 && !!enrichRec;
  return { newCount, addableToProject, enrichRec, enrichOnly };
}
function updatePeopleButtonBadge() {
  const btn = document.getElementById("addParticipantBtn");
  if (!btn) return;
  const { newCount, addableToProject, enrichRec, enrichOnly } = peopleButtonMode();
  // Hidden in compose (nothing filed yet — mirrors applyComposeModeUiGuard)
  // and whenever there's no one actionable.
  const compose = (typeof isComposeMode === "function") && isComposeMode();
  btn.style.display = (!compose && (newCount > 0 || addableToProject > 0 || !!enrichRec)) ? "" : "none";
  // Label: new contacts take priority. When the ONLY reason to show is the
  // signature-enrichment opportunity, say so on the button itself.
  if (newCount > 0) {
    btn.textContent = "👥 Add Participant to Contacts (" + newCount + " new)";
  } else if (enrichOnly) {
    const who = (enrichRec.name || "").trim().split(/\s+/)[0];
    btn.textContent = who
      ? "✍️ Complete " + who + "'s contact from signature"
      : "✍️ Complete a contact from its signature";
  } else {
    btn.textContent = "👥 Add Participant to Contacts";
  }
  // Visually promote the button when there's actually someone to capture or enrich.
  btn.classList.toggle("btn-has-new", newCount > 0 || enrichOnly);
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
// The email's true send date. For ARCHIVED items, emailItem.dateTimeCreated is the
// archive-copy time (a 2021 email can read as 2023), so prefer the internet "Date:"
// header — the real send date Outlook shows. Falls back to dateTimeCreated.
async function getEmailDateReliable() {
  const headerDate = await new Promise(resolve => {
    try {
      if (!emailItem?.getAllInternetHeadersAsync) return resolve("");
      emailItem.getAllInternetHeadersAsync(r => {
        if (r.status !== Office.AsyncResultStatus.Succeeded) return resolve("");
        const m = String(r.value || "").match(/^Date:\s*(.+)$/im);
        resolve(m ? m[1].trim() : "");
      });
    } catch { resolve(""); }
  });
  if (headerDate) {
    const d = new Date(headerDate);
    if (!isNaN(d.getTime())) return d.toISOString();
  }
  return emailItem?.dateTimeCreated || null;
}

async function getEmailBodyHtml(token) {
  const item = Office.context.mailbox.item;
  const msgId = item?.itemId;
  if (msgId && _emailBodyCache.has(msgId)) return _emailBodyCache.get(msgId);

  let body = "";
  // 1) Graph — full-fidelity HTML body.
  try {
    const restId = Office.context.mailbox.convertToRestId(msgId, Office.MailboxEnums.RestVersion.v2_0);
    const data = await graphFetch("GET", "/me/messages/" + restId + "?$select=body", null, token);
    body = data?.body?.content || "";
  } catch { /* fall through to the Office.js route */ }

  // 2) Office.js fallback — reads the OPEN item's HTML directly, no network, so it
  //    works when the Graph /me/messages fetch fails (the cause of "Email body
  //    could not be retrieved" → empty records). getAsync can fail transiently
  //    right after an item opens, so retry with short backoff.
  if (!body && item?.body?.getAsync) {
    for (let i = 0; i < 3 && !body; i++) {
      body = await new Promise(resolve => {
        try {
          item.body.getAsync(Office.CoercionType.Html, r =>
            resolve(r.status === Office.AsyncResultStatus.Succeeded ? (r.value || "") : ""));
        } catch { resolve(""); }
      });
      if (!body && i < 2) await new Promise(res => setTimeout(res, 300 * (i + 1)));
    }
  }

  if (body && msgId) _emailBodyCache.set(msgId, body);
  return body;
}

// ── Inline image resolution (for OneNote pages + the PMS email log) ──────────
// Email HTML references embedded images as cid: pointers; the bytes live as
// inline file attachments. OneNote can't read cid: (it needs multipart name:
// parts) and the PMS log needs them as data: URIs — both start from this fetch.
// Everything here is best-effort: on any failure callers keep the original body.
const _inlineImgCache = new Map(); // itemId -> Map<cidLower, {contentType, bytes}>
async function getInlineImagesForEmail(token) {
  const item = Office.context.mailbox.item;
  const itemId = item?.itemId;
  if (itemId && _inlineImgCache.has(itemId)) return _inlineImgCache.get(itemId);
  const map = new Map();
  try {
    const restId = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
    const attData = await graphFetch("GET", "/me/messages/" + restId + "/attachments", null, token);
    for (const att of (attData?.value || [])) {
      if (att["@odata.type"] !== "#microsoft.graph.fileAttachment") continue;
      if (!att.isInline || !(att.contentType || "").startsWith("image/")) continue;
      const cid = String(att.contentId || "").replace(/[<>]/g, "").trim().toLowerCase();
      if (!cid) continue;
      let bytes = att.contentBytes ? toBytesFromBase64(att.contentBytes) : null;
      if (!bytes && att.id) {
        const r = await fetchWithRetry(
          "https://graph.microsoft.com/v1.0/me/messages/" + restId + "/attachments/" + att.id + "/$value",
          { headers: { "Authorization": "Bearer " + token } }, { label: "graph inline img" });
        if (r.ok) bytes = new Uint8Array(await r.arrayBuffer());
      }
      if (bytes) map.set(cid, { contentType: att.contentType, bytes });
    }
  } catch (e) { console.warn("[inline-img] fetch failed:", e.message); }
  if (itemId) _inlineImgCache.set(itemId, map);
  return map;
}

// Best-effort raster downscale to keep the email log light + fit OneNote's 4MB
// cap. Returns the ORIGINAL bytes if the canvas path is unavailable or doesn't help.
const INLINE_IMG_MAX_DIM = 1500;
async function downscaleInlineImage(bytes, contentType) {
  try {
    if (typeof createImageBitmap !== "function") return { bytes, contentType };
    const bmp = await createImageBitmap(new Blob([bytes], { type: contentType }));
    const longEdge = Math.max(bmp.width, bmp.height);
    const scale = Math.min(1, INLINE_IMG_MAX_DIM / longEdge);
    if (scale >= 1 && bytes.length <= 300000) { if (bmp.close) bmp.close(); return { bytes, contentType }; }
    const w = Math.max(1, Math.round(bmp.width * scale));
    const h = Math.max(1, Math.round(bmp.height * scale));
    let blob;
    if (typeof OffscreenCanvas === "function") {
      const cv = new OffscreenCanvas(w, h);
      cv.getContext("2d").drawImage(bmp, 0, 0, w, h);
      blob = await cv.convertToBlob({ type: "image/jpeg", quality: 0.72 });
    } else {
      const cv = document.createElement("canvas");
      cv.width = w; cv.height = h;
      cv.getContext("2d").drawImage(bmp, 0, 0, w, h);
      blob = await new Promise(res => cv.toBlob(res, "image/jpeg", 0.72));
    }
    if (bmp.close) bmp.close();
    if (!blob) return { bytes, contentType };
    const out = new Uint8Array(await blob.arrayBuffer());
    return out.length < bytes.length ? { bytes: out, contentType: "image/jpeg" } : { bytes, contentType };
  } catch (e) {
    console.warn("[inline-img] downscale failed:", e.message);
    return { bytes, contentType };
  }
}

function _bytesToBase64(bytes) {
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  }
  return btoa(bin);
}

// Run an async replacer over every regex match (String.replace can't await).
async function replaceAsyncAddin(str, regex, asyncFn) {
  const collected = [];
  str.replace(regex, (...a) => { collected.push(a); return ""; });
  const reps = await Promise.all(collected.map(a => asyncFn(...a)));
  let i = 0;
  return str.replace(regex, () => reps[i++]);
}

// Matches src="cid:..." / src='cid:...' (quote-aware via the \1 backreference).
const cidSrcRe = () => /src\s*=\s*(["'])cid:([^"']+)\1/gi;

// Email log: replace cid: refs with downscaled data: URIs so images render in the
// PMS email viewer. Caps total embedded size; anything over stays a cid (the
// full-res original is in SharePoint).
const INLINE_LOG_TOTAL_CAP = 6 * 1024 * 1024;
async function embedInlineImagesAsDataUris(html, token) {
  if (!html || !/src\s*=\s*["']cid:/i.test(html)) return html;
  const imgs = await getInlineImagesForEmail(token);
  if (!imgs || imgs.size === 0) return html;
  let total = 0;
  const done = new Map();
  return await replaceAsyncAddin(html, cidSrcRe(), async (m, q, rawCid) => {
    const cid = String(rawCid).replace(/[<>]/g, "").trim().toLowerCase();
    const img = imgs.get(cid);
    if (!img) return m;
    if (done.has(cid)) return done.get(cid);
    const ds = await downscaleInlineImage(img.bytes, img.contentType);
    if (total + ds.bytes.length > INLINE_LOG_TOTAL_CAP) return m;
    total += ds.bytes.length;
    const rep = `src=${q}data:${ds.contentType};base64,${_bytesToBase64(ds.bytes)}${q}`;
    done.set(cid, rep);
    return rep;
  });
}

// OneNote: rewrite cid: -> name:imgN and return the binary parts for a multipart
// create-page POST. Honors OneNote's ~4MB request / 30-image limits.
const ONENOTE_TOTAL_CAP = 3.5 * 1024 * 1024;
async function rewriteCidToOneNoteParts(html, imgs) {
  const parts = [];
  if (!html || !imgs || imgs.size === 0) return { html, parts };
  let total = 0, idx = 0;
  const named = new Map();
  const out = await replaceAsyncAddin(html, cidSrcRe(), async (m, q, rawCid) => {
    const cid = String(rawCid).replace(/[<>]/g, "").trim().toLowerCase();
    const img = imgs.get(cid);
    if (!img) return m;
    if (named.has(cid)) return `src=${q}name:${named.get(cid)}${q}`;
    if (parts.length >= 28) return m;
    const ds = await downscaleInlineImage(img.bytes, img.contentType);
    if (total + ds.bytes.length > ONENOTE_TOTAL_CAP) return m;
    total += ds.bytes.length;
    const name = "img" + (idx++);
    named.set(cid, name);
    parts.push({ name, bytes: ds.bytes, contentType: ds.contentType });
    return `src=${q}name:${name}${q}`;
  });
  return { html: out, parts };
}

// ── RELIABLE PLAIN-TEXT BODY ─────────────────────────────────────────────────
// Office's item.body.getAsync fails transiently right after an item opens
// (status Failed, especially in new Outlook / OWA) — the cause of "signature
// extraction only works the second or third time I open the email". Retry
// with short backoff, then fall back to the Graph body (already cached per
// item for the filing flow) stripped to plain text. Returns null when every
// route failed; "" is a genuinely empty body.
async function getBodyTextReliable(item, attempts = 3) {
  for (let i = 0; i < attempts; i++) {
    const text = await new Promise(resolve => {
      try {
        if (!item?.body?.getAsync) return resolve(null);
        item.body.getAsync(Office.CoercionType.Text, r =>
          resolve(r.status === Office.AsyncResultStatus.Succeeded ? (r.value || "") : null));
      } catch { resolve(null); }
    });
    if (text !== null) return text;
    if (i < attempts - 1) await new Promise(res => setTimeout(res, 300 * (i + 1)));
  }
  // Graph fallback — getEmailBodyHtml reads the CURRENT mailbox item, so only
  // use it if the pane hasn't switched to a different email meanwhile.
  try {
    if (Office.context.mailbox.item?.itemId !== item?.itemId) return null;
    const token = await getToken();
    const html = await getEmailBodyHtml(token);
    if (html) return htmlToPlainText(html);
  } catch (e) {
    console.info("[body] Graph fallback failed:", e.message);
  }
  return null;
}

function htmlToPlainText(html) {
  const t = String(html || "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|tr|li|h[1-6]|table)>/gi, "\n")
    .replace(/<[^>]+>/g, "");
  const decoder = document.createElement("textarea");
  decoder.innerHTML = t; // entity decode — tags already stripped above
  return decoder.value.replace(/ /g, " ");
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
  lastAttachmentUploadStats = { attempted: 0, uploaded: 0, failed: [], uploadedFiles: [], attemptedNames: [] };
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
      // Multi-attachment saves can run 30s+ on big files; surface progress in
      // whatever busy-status slot the calling flow already opened so users can
      // tell "working" from "stuck" (and don't re-click Save).
      if (items.length > 1 && _busyStatusElId) {
        setStatus(_busyStatusElId, "info", "⏳ Uploading attachments… " + Math.min(i + batch.length, items.length) + "/" + items.length);
      }
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
      lastAttachmentUploadStats.attemptedNames = uniqueNames;
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
    const fileAtts = (attData?.value || []).filter(att =>
      att["@odata.type"] === "#microsoft.graph.fileAttachment" &&
      // Skip small inline images (signature logos, social icons, banners).
      // att.size here is MIME-encoded (~+33%), so 40KB encoded ≈ 30KB raw —
      // comfortably above logo size, below real pasted screenshots/photos.
      !(att.isInline && att.contentType && att.contentType.startsWith("image/") && att.size < 40000));
    lastAttachmentUploadStats.attempted = fileAtts.length;
    const uniqueGraphNames = uniquifyNames(fileAtts.map(a => a.name));
    lastAttachmentUploadStats.attemptedNames = uniqueGraphNames;

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
// Returns only the names of file-type attachments — no byte fetch. Cheap.
// Used by "save to project record only" path where we don't upload anything
// but still want PMS to display the attachment list.
//
// TWO PATHS, in order:
//   1. Office.js getAttachmentsAsync — local, no network. Works most of the time.
//   2. Graph /me/messages/{id}/attachments — fallback when Office.js silently
//      returns 0 attachments despite the email actually having some. This
//      happens in some account configurations and Office versions; mirrors
//      the same fallback uploadEmailAndAttachments uses.
//
// Silent on failure (returns []) since names are metadata, not data.
async function getAttachmentNamesOnly(item, token) {
  item = item || emailItem;
  // Try Office.js first
  if (item?.getAttachmentsAsync) {
    const officeNames = await new Promise((resolve) => {
      item.getAttachmentsAsync((res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          const names = (res.value || [])
            .filter(a => a.attachmentType === Office.MailboxEnums.AttachmentType.File)
            .map(a => a.name)
            .filter(Boolean);
          resolve(names);
        } else {
          resolve([]);
        }
      });
    });
    if (officeNames.length > 0) return officeNames;
  }
  // Graph fallback — only if Office.js gave us nothing AND caller provided a token
  if (token && item?.itemId) {
    try {
      const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
      const data = await graphFetch("GET", "/me/messages/" + restId + "/attachments?$select=name,contentType,size", null, token);
      return (data?.value || [])
        .filter(a => a["@odata.type"] === "#microsoft.graph.fileAttachment")
        .map(a => a.name)
        .filter(Boolean);
    } catch (e) {
      console.warn("Graph attachment list failed:", e.message);
      return [];
    }
  }
  return [];
}

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
    // Skip small inline images (signature logos, social icons, banners) — same
    // rule as the Graph fallback path. bytes.length is the true raw size here,
    // so the 40KB cutoff is a hair stricter than the Graph path's encoded size,
    // but both reliably catch logos and keep real pasted screenshots/photos.
    if (att.isInline && (att.contentType || "").startsWith("image/") && bytes.length < 40000) continue;
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
    setStatus(statusElement, "error", "✗ " + humanizeError(e));
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
// Update an already-saved email's SharePoint + body fields in pms_project_emails
// (used when Save to SharePoint runs on a record that auto-save created earlier —
// avoids a duplicate row). PATCH by record_id; no unique-constraint dependency.
async function patchEmailSpFields(recordId, fields) {
  const url = SUPABASE_URL + "/rest/v1/" + PROJECT_EMAILS_TABLE + "?record_id=eq." + encodeURIComponent(recordId);
  const res = await fetchWithRetry(url, {
    method: "PATCH",
    headers: { ...SB_HEADERS, "Content-Type": "application/json", Prefer: "return=minimal" },
    body: JSON.stringify(fields),
  }, { label: "sb project_emails sp-update" });
  if (!res.ok) throw new Error("pms_project_emails PATCH HTTP " + res.status + ": " + (await res.text()).slice(0, 150));
}
// Patch by (project_id, msg_id) instead of the surrogate record_id. msg_id is the
// stable business key, so the SharePoint link lands on the correct row even if the
// inline record's id ever drifted — this is the direct cure for "saved to
// SharePoint but the link didn't stick."
async function patchEmailSpFieldsByMsg(projectId, msgId, fields) {
  if (!projectId || !msgId) return;
  const url = SUPABASE_URL + "/rest/v1/" + PROJECT_EMAILS_TABLE +
    "?project_id=eq." + encodeURIComponent(projectId) +
    "&msg_id=eq." + encodeURIComponent(msgId);
  const res = await fetchWithRetry(url, {
    method: "PATCH",
    headers: { ...SB_HEADERS, "Content-Type": "application/json", Prefer: "return=minimal" },
    body: JSON.stringify(fields),
  }, { label: "sb project_emails sp-update-by-msg" });
  if (!res.ok) throw new Error("pms_project_emails PATCH(by msg) HTTP " + res.status + ": " + (await res.text()).slice(0, 150));
}
async function doSaveToSharePoint() {
  return withSaveGuard("save-sp", _doSaveToSharePoint, ["saveSpBtn", "saveRecordBtn"]);
}
async function _doSaveToSharePoint() {
  if (!selectedProject) { setStatus("actionStatus", "error", "Select a project first."); return; }
  if (!selectedProject.projectFolderUrl) { setStatus("actionStatus", "error", "No SharePoint folder on this project. Create one in the PMS first."); return; }
await ensureInternetMessageId(); // STABLE id so existingRecord + the SP patch resolve to the right row
const currentMsgId = getCurrentMessageRecordId();
const existingRecord = findSavedEmailRecord(selectedProject, currentMsgId);
// Read the Link To target up front so we can branch on it BEFORE the
// already-saved short-circuit below. Without this, the link operation
// is silently skipped when the user is adding a link to an email that
// was previously saved standalone.
const earlyLinkValue = (document.getElementById("linkToTarget")?.value || "");
if (existingRecord) {
  // If a Link To target is picked, run JUST the link operation against the
  // existing email record. This is the "I forgot to link the first time"
  // recovery path — the most common reason a user re-clicks Save SP on an
  // already-filed email is to add a link they couldn't pick before.
  if (earlyLinkValue) {
    setStatus("actionStatus", "info", "📎 Linking to RFI/Submittal…");
    const linkSnapItem = emailItem;
    try {
      const linkResult = await linkEmailToArtifact({
        linkValue: earlyLinkValue,
        emailRecord: existingRecord,
        snapItem: linkSnapItem,
      });
      if (linkResult.ok) {
        setStatus("actionStatus", "success", "✓ Linked existing email" + linkResult.label);
      } else {
        setStatus("actionStatus", "error", "✗ Link failed" + (linkResult.label || ""));
      }
    } catch (e) {
      setStatus("actionStatus", "error", "✗ Link failed: " + humanizeError(e));
    }
    // Reset link dropdown so subsequent saves don't re-link
    const linkToEl = document.getElementById("linkToTarget");
    if (linkToEl) linkToEl.value = "";
    try { refreshLoggedArtifactChips(); } catch {}
    try { refreshLinkToTargetDropdown(); } catch {}
    return;
  }
  // Already filed to SharePoint → nothing to upload; just refresh the indicator.
  if (existingRecord.spFolderUrl) { refreshEmailSavedIndicator(); return; }
  // Otherwise this email was logged to the project record (e.g. by auto-save on
  // tag) but never pushed to SharePoint. Fall through to upload the attachments
  // and UPDATE this same record below — don't bail, and don't duplicate the row.
}
  // Snapshot all per-item data SYNCHRONOUSLY before any await. Without this,
  // a fast item-switch (Office swaps mailbox.item silently on pinned panes)
  // mid-save would read fields from the new email — saving the wrong subject,
  // date, sender, or attachments into the current project's folder. The
  // generation counter alone can't protect later sync reads of `emailItem.*`.
  const snapItem = emailItem;
  const snapSubject = snapItem?.subject || "";
  const snapDate = await getEmailDateReliable();
  const snapFromName = snapItem?.from?.displayName || "";
  const snapFromAddr = snapItem?.from?.emailAddress || "";
  // Capture recipients in the same snapshot so an item-switch mid-save can't
  // record one email's TO list under a different email's record. Format
  // matches PMS-side: comma-separated email addresses.
  const snapTo = (snapItem?.to || []).map(r => r.emailAddress).filter(Boolean).join(", ");
  const snapCc = (snapItem?.cc || []).map(r => r.emailAddress).filter(Boolean).join(", ");
  const snapItemId = snapItem?.itemId || "";
  const saveGen = itemContextGeneration; // capture for stale-write detection
  // Read the "Link to RFI/Sub" dropdown synchronously so it can't drift if
  // the user switches projects mid-save.
  const linkToValue = (document.getElementById("linkToTarget")?.value || "");
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
  setStatus("actionStatus", "info", "⏳ " + pickSavingMessage());
  try {
    const token = await getToken();
    const { driveId } = await resolveSpIds();
    // Phase 3: fetch body HTML once up front so we can both upload to SharePoint
    // AND store the compressed version on the project record.
    // Track body-fetch failure separately so we can surface it in the success
    // message — previously a silent "" fallback meant the user thought everything
    // worked but the email record had no readable body.
    let bodyHtml = await getEmailBodyHtml(token);
    const bodyFetchFailed = !bodyHtml || bodyHtml.length === 0;
    // Inline images arrive as cid: refs (bytes live as inline attachments) and
    // don't render in the PMS log — re-embed them as downscaled data: URIs so they
    // show. Best-effort; on failure the original body is kept (SharePoint has the
    // full-res originals regardless).
    if (bodyHtml) {
      try { bodyHtml = await embedInlineImagesAsDataUris(bodyHtml, token); }
      catch (e) { console.warn("[inline-img] log embed failed:", e.message); }
    }
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
    // Capture attachment names from the upload stats so PMS displays them on
    // the saved email card. Names are populated by uploadEmailAndAttachments
    // even if some uploads later failed — the email still HAS the attachments.
    const attachmentNames = (lastAttachmentUploadStats?.attemptedNames || []).slice();
    const emailRecord = {
      id: existingRecord ? existingRecord.id : uid(), msgId,
      subject: snapSubject,
      from: snapFromName,
      fromAddress: snapFromAddr,
      to: snapTo,
      cc: snapCc,
      date: snapDate,
      bodyText: bodyHtml ? stripHtmlToText(bodyHtml) : "",
      bodyHtmlCompressed: compressedBody,
      bodyHtmlSize: bodyHtml.length,
      hasAttachments: attachmentNames.length > 0,
      attachmentNames,
      spFolderUrl, links: [],
      savedAt: new Date().toISOString(),
      savedBy: _getCurrentUserEmail() || "",
    };
    // If this email was already logged (auto-save) and THIS pass couldn't fetch a
    // body, keep the body the earlier save captured — don't blank it out.
    if (existingRecord && !(emailRecord.bodyHtmlSize > 0)) {
      emailRecord.bodyText = existingRecord.bodyText || emailRecord.bodyText;
      emailRecord.bodyHtmlCompressed = existingRecord.bodyHtmlCompressed || emailRecord.bodyHtmlCompressed;
      emailRecord.bodyHtmlSize = existingRecord.bodyHtmlSize || emailRecord.bodyHtmlSize;
    }
    // Re-fetch latest projects, then append email to the FRESH copy of this project.
    // Prevents the add-in from overwriting concurrent PMS edits made during this session.
    await applyLocalChangeAndSave(selectedProject.id, fresh => {
      // Replace any existing entry for this email (auto-save may have logged it
      // already) so the SharePoint URL lands on ONE record instead of duplicating.
      const others = (fresh.emails || []).filter(e => e.id !== emailRecord.id && e.msgId !== emailRecord.msgId);
      return { ...fresh, emails: [...others, emailRecord] };
    });
    let indexSaveFailed = false;
    try {
      if (existingRecord) {
        // Auto-save already inserted the search-index row; just patch in the
        // SharePoint URL + attachment names (and refresh the body if we got one)
        // so we update that one row instead of inserting a duplicate.
        // Patch by the record's STORED msg_id (not its surrogate id), so the link
        // lands on the right row whether it was keyed by internetMessageId or a
        // legacy REST id.
        await patchEmailSpFieldsByMsg(selectedProject.id, existingRecord.msgId || currentMsgId, {
          sp_folder_url: spFolderUrl || "",
          saved_to_sharepoint: true,
          attachment_names: attachmentNames || [],
          body_text: emailRecord.bodyText || "",
          body_html_compressed: emailRecord.bodyHtmlCompressed || null,
          body_html_size: emailRecord.bodyHtmlSize || 0,
        });
      } else {
        await saveProjectEmailRow(selectedProject.id, emailRecord, true);
      }
    } catch (idxErr) {
      console.warn("project_emails index write failed:", idxErr);
      indexSaveFailed = true;
    }
    // Optional: if the user picked a "Link to RFI/Sub" target, copy the email
    // into that artifact's /IN folder and add a links[] entry. Best-effort —
    // primary save already succeeded so we never throw out of this branch.
    const linkResult = await linkEmailToArtifact({ linkValue: linkToValue, emailRecord, snapItem });
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

    const linkSuffix = linkResult.label || "";
    if (attempted > 0 && attCount === 0) {
      const sample = (lastAttachmentUploadStats?.failed || []).slice(0, 2).join("; ");
      setStatus("actionStatus", "error", "Email saved, but 0/" + attempted + " attachments uploaded. " + (sample ? "Failed: " + sample + ". " : "") + "Try saving again — if it keeps failing, check your connection (VPN?)." + (warnings.length ? " " + warnings.join(" ") : "") + linkSuffix);
    } else if (warnings.length > 0) {
      setStatus("actionStatus", "info", "✓ Saved to SharePoint" + attMsg + " and project record. " + warnings.join(" ") + linkSuffix);
    } else if (attempted === 0) {
      setStatus("actionStatus", "info", "Email saved to SharePoint, but no attachments were detected by Outlook/Graph for this message." + linkSuffix);
    } else {
      setStatus("actionStatus", "success", "✓ Saved to SharePoint" + attMsg + " and project record." + linkSuffix);
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
    // Same one-shot semantics for the Link To target: don't carry over from
    // one save to the next. Refresh chips so the new artifact-source-id
    // chip + the just-linked RFI both show up.
    const linkToEl = document.getElementById("linkToTarget");
    if (linkToEl) linkToEl.value = "";
    try { refreshLoggedArtifactChips(); } catch {}
    try { refreshLinkToTargetDropdown(); } catch {}
    recordSaveAndCelebrate();
    refreshEmailSavedIndicator(true);
    // Save completed without throwing — clear crash-recovery queue entry.
    // (Partial successes still dequeue: the user has the status message and
    // a "partial" audit log row; the queue is only for crash recovery.)
    dequeueFilingIntent(queueId);
  } catch (e) {
    setStatus("actionStatus", "error", "✗ " + humanizeError(e));
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
async function _doSaveToProjectRecordOnly(quiet = false) {
  if (!selectedProject) { if (!quiet) setStatus("actionStatus", "error", "Select a project first."); return; }
  // Body-only save works regardless of attachments — the visual emphasis (de-emph
  // + caption) is the soft nudge toward SharePoint when attachments exist.
  // No confirm dialog: trust the user's intent, surface the consequence in the
  // post-save card ("3 attachments not filed").
  await ensureInternetMessageId(); // STABLE id before we key/look-up the record
  const msgId = getCurrentMessageRecordId();
  if (findSavedEmailRecord(selectedProject, msgId)) {
    refreshEmailSavedIndicator();
    return;
  }
  // Always show an in-progress indicator (even on quiet auto-save) so users don't
  // click off the email mid-capture. Quiet only suppresses the celebration.
  setStatus("actionStatus", "info", quiet ? "⏳ Logging email to project…" : "⏳ " + pickSavingMessage());
  let ok = false;
  try {
    // Phase 3: capture and compress body so PMS can render it without a Graph round-trip.
    const token = await getToken();
    let bodyHtml = await getEmailBodyHtml(token);
    const bodyFetchFailed = !bodyHtml || bodyHtml.length === 0;
    // Inline images arrive as cid: refs (bytes live as inline attachments) and
    // don't render in the PMS log — re-embed them as downscaled data: URIs so they
    // show. Best-effort; on failure the original body is kept (SharePoint has the
    // full-res originals regardless).
    if (bodyHtml) {
      try { bodyHtml = await embedInlineImagesAsDataUris(bodyHtml, token); }
      catch (e) { console.warn("[inline-img] log embed failed:", e.message); }
    }
    const compressedBody = bodyHtml ? compressHtmlAddin(bodyHtml) : "";
    const from = emailItem.from;
    const to = (emailItem.to || []).map(r => r.emailAddress).filter(Boolean).join(", ");
    const cc = (emailItem.cc || []).map(r => r.emailAddress).filter(Boolean).join(", ");
    // Lightweight metadata fetch — names only, no byte download. Tries
    // Office.js first, falls back to Graph if Office.js reports 0 attachments
    // (which it sometimes silently does even when attachments exist).
    const attachmentNames = await getAttachmentNamesOnly(emailItem, token);
    const emailRecord = {
      id: uid(), msgId,
      subject: emailItem.subject || "",
      from: from?.displayName || "",
      fromAddress: from?.emailAddress || "",
      to,
      cc,
      date: await getEmailDateReliable(),
      bodyText: bodyHtml ? stripHtmlToText(bodyHtml) : "",
      bodyHtmlCompressed: compressedBody,
      bodyHtmlSize: bodyHtml.length,
      hasAttachments: attachmentNames.length > 0,
      attachmentNames,
      spFolderUrl: "", links: [],
      savedAt: new Date().toISOString(),
      savedBy: _getCurrentUserEmail() || "",
      savedToSharePoint: false,
    };
    await applyLocalChangeAndSave(selectedProject.id, fresh => {
      // Dedup against FRESH data, mirroring the DB's (project_id, msg_id)
      // uniqueness: if this email was filed here since our pane last synced
      // (another user, or a prior auto-save), don't append a second copy.
      if ((fresh.emails || []).some(e => e.msgId && e.msgId === emailRecord.msgId)) return fresh;
      return { ...fresh, emails: [...(fresh.emails || []), emailRecord] };
    });
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
    } else if (!quiet) {
      setStatus("actionStatus", "success", "✓ Saved to project record (no SharePoint upload).");
    }
    if (!quiet) recordSaveAndCelebrate();
    refreshEmailSavedIndicator(!quiet);
    ok = true;
  } catch (e) {
    if (!quiet) setStatus("actionStatus", "error", "✗ " + humanizeError(e));
    else console.warn("auto-save record failed:", e);
  } finally {
    // Never strand the "⏳ Logging email to project…" transient as a perpetual
    // spinner. If it's STILL the active status after we finish — a quiet auto-save
    // that failed on a transient conflict (common while a sweep is churning project
    // versions), or a post-save refresh that didn't find the record — resolve it:
    // clear on success, or point at the manual Save button on failure. Guarded on
    // the in-progress text, so real "Saved"/"Logged"/"Auto-tagged" messages survive.
    const el = document.getElementById("actionStatus");
    if (el && (el.textContent || "").includes("Logging email to project")) {
      if (ok) setStatus("actionStatus", "", "");
      else setStatus("actionStatus", "info", "Couldn't auto-log this email — reopen it or use 🗂️ Save to Project.");
    }
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
    setStatus("actionStatus", "error", "Send-to-Teams failed: " + humanizeError(e));
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
  // Resolve inline (cid:) images so they render on the OneNote page. OneNote can't
  // read cid: or data: URIs — only public URLs or name: parts backed by binary
  // parts in a multipart POST. Best-effort: on failure we fall back to the plain
  // HTML page (images just won't show, exactly like before this change).
  let oneNoteImgParts = [];
  let emailBodyForNote = emailBodyHtml;
  if (!isMeetingNoteCategory(category) && emailBodyHtml && /src\s*=\s*["']cid:/i.test(emailBodyHtml)) {
    try {
      const imgs = await getInlineImagesForEmail(await getToken());
      const r = await rewriteCidToOneNoteParts(emailBodyHtml, imgs);
      emailBodyForNote = r.html;
      oneNoteImgParts = r.parts;
    } catch (e) { console.warn("[OneNote inline-img] resolve failed:", e.message); }
  }
  const bodyHtml = isMeetingNoteCategory(category)
    ? buildAddinMeetingPageHtml(title, category, dateStr, emailParticipants, body, project)
    : buildAddinEmailNotePageHtml(title, category, dateStr, fromName, fromEmail, emailBodyForNote, body, project);
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
    // With inline images, send a multipart create-page (HTML "Presentation" part +
    // one binary part per image, referenced as name:imgN). Otherwise plain text/html
    // exactly as before. (Don't set Content-Type for FormData — the browser adds the
    // multipart boundary.)
    let res;
    if (oneNoteImgParts.length) {
      const form = new FormData();
      form.append("Presentation", new Blob([pageHtml], { type: "text/html" }));
      for (const part of oneNoteImgParts) {
        form.append(part.name, new Blob([part.bytes], { type: part.contentType }), part.name);
      }
      res = await fetch(url, { method: "POST", headers: { "Authorization": "Bearer " + token }, body: form });
    } else {
      res = await fetch(url, {
        method: "POST",
        headers: { "Authorization": "Bearer " + token, "Content-Type": "text/html" },
        body: pageHtml,
      });
    }
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
    setStatus("noteStatus", "error", "✗ " + humanizeError(e));
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
    setStatus("actionItemStatus", "error", "✗ " + humanizeError(e));
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
// Layout mirrors the Newforma RFI Transmittal format (two-page structured doc:
// metadata grid + boxed answer with boilerplate notations on page 1, FROM/TO/
// contents tables on page 2). This is a deliberate format choice — the legal
// "General Notations" boilerplate is risk management language that's standard
// in the AEC industry, and the structured TO/FROM blocks make the email an
// auditable artifact.

const DOCX_COLORS = { NAVY: "1e3a8a", RED: "b91c1c", BLUE: "1e40af", GRAY: "475569", LIGHT: "cbd5e1", BOX_BORDER: "94a3b8" };

// Standard "General Notations" boilerplate for RFI responses. Tells the
// contractor that the RFI response is not authorization to incur additional
// cost — they must issue a PCO (Proposed Change Order) for that. This text
// is unchanged across all Setty RFI responses; if it ever needs to evolve,
// only this constant needs updating.
const RFI_GENERAL_NOTATIONS = [
  "This RFI, RFI response, and related correspondence, written, verbal, or other; is not an authorization to proceed with work which requires additional costs. If any RFI reply requires a change to the contract documents, the contractor shall issue for review, a Proposed Change Order (PCO). Additional costs are authorized only after a PCO has been reviewed, and formally issued as an Accepted Change Order (ACO).",
  "If the contractor believes an RFI response has materially changed the contract documents, or would subsequently create conflicts in work, or conflicts in coordination with their own or an associated trade; the contractor must state such concerns formally and promptly, prior to acting upon the RFI response.",
];

// Submittal review boilerplate. Standard "review for general conformance"
// language — Setty's review doesn't relieve the contractor of responsibility
// for accuracy of fabrication, coordination with the work, etc.
const SUB_GENERAL_NOTATIONS = [
  "This review is for general conformance with the design intent of the contract documents only. Markings and comments do not relieve the contractor of responsibility for accuracy of fabrication, dimensions, quantities, coordination with other trades, performance of any equipment, or means and methods of construction.",
  "The contractor is responsible for confirming all dimensions, field conditions, and coordination requirements prior to fabrication or installation. Any deviations from the contract documents must be brought to the Architect/Engineer's attention in writing prior to proceeding.",
];

// Format helper used by both DOCX generators for the metadata grid rows.
// Returns a TableRow with a label cell (gray background, bold blue text)
// and a value cell.
function _docxInfoRow(docxLib, label, value, opts = {}) {
  const { Paragraph, TextRun, TableRow, TableCell, WidthType, ShadingType } = docxLib;
  const labelW = opts.labelW || 1600;
  const valueW = opts.valueW || 3080;
  return new TableRow({
    children: [
      new TableCell({
        width: { size: labelW, type: WidthType.DXA },
        shading: { type: ShadingType.SOLID, color: "f8fafc", fill: "f8fafc" },
        children: [new Paragraph({ spacing: { before: 30, after: 30 }, children: [new TextRun({ text: label, font: "Calibri", size: 16, bold: true, color: DOCX_COLORS.GRAY })] })],
      }),
      new TableCell({
        width: { size: valueW, type: WidthType.DXA },
        children: [new Paragraph({ spacing: { before: 30, after: 30 }, children: [new TextRun({ text: String(value || ""), font: "Calibri", size: 20 })] })],
      }),
    ],
  });
}

// Build the recipients list for the response. Defaults: original RFI sender
// (from rfi.from + email lookup) + any assigned subconsultant POC + the
// responder themselves on the COPIES line. Optional override via opts.
function _resolveRfiResponseRecipients(rfi, project) {
  const recipients = [];
  // Original sender: look up in project.emails by sourceItemId/sourceMessageId
  if (rfi.sourceItemId || rfi.sourceMessageId) {
    const emailRec = (project.emails || []).find(e =>
      (rfi.sourceItemId && e.msgId === rfi.sourceItemId) ||
      (rfi.sourceMessageId && e.msgId === rfi.sourceMessageId)
    );
    if (emailRec?.fromAddress) {
      recipients.push({ name: emailRec.from || "", company: "", email: emailRec.fromAddress, phone: "" });
    }
  }
  // If no email lookup, fall back to rfi.from as a name-only entry
  if (recipients.length === 0 && rfi.from) {
    recipients.push({ name: rfi.from, company: "", email: "", phone: "" });
  }
  // Assigned subconsultant POC
  if (rfi.subAssigned) {
    const sub = (project.subconsultants || []).find(s => s.id === rfi.subAssigned);
    if (sub) {
      recipients.push({
        name: sub.contact || sub.firm || "",
        company: sub.firm || "",
        email: sub.email || "",
        phone: sub.phone || "",
      });
    }
  }
  return recipients;
}

// Returns a Blob of a DOCX cover sheet for an RFI response. Two-page layout
// matching the Newforma RFI Transmittal format:
//   Page 1: SETTY header, metadata grid (PROJECT / DATE SENT / SUBJECT / RFI ID /
//     TYPE / TRANSMITTAL ID / PURPOSE / VIA), QUESTION section, ANSWER section
//     in a bordered box with "Response (Answered) from", "General Notations"
//     boilerplate, and the response body.
//   Page 2: FROM table (name, company, email, phone), TO table, DESCRIPTION OF
//     CONTENTS table, COPIES section.
async function buildRfiResponseDocx({ rfi, project, response, dateResponded, status, recipients }) {
  if (typeof docx === "undefined" || !docx.Packer) {
    throw new Error("DOCX library not loaded — refresh the taskpane and try again.");
  }
  const lib = docx;
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    Header, BorderStyle, WidthType, ShadingType, AlignmentType, PageBreak,
  } = lib;

  const projLabel = project.name || "";
  const projNumber = project.projectNumber || "";
  const responder = msalAccount?.name || msalAccount?.username || "";
  const responderEmail = msalAccount?.username || "";
  const dateSent = dateResponded || new Date().toISOString().slice(0, 10);
  const subject = rfi.title || "";
  const rfiId = rfi.number || "";
  const purpose = status === "Responded" ? "Answered" : (status || "Answered");
  // Transmittal ID: a per-project counter would require schema work. For now,
  // use the RFI number + date — auditable enough.
  const transmittalId = `${rfiId}-${dateSent.replace(/-/g, "")}`;
  const via = rfi.receivedVia || "Email";
  const toList = recipients && recipients.length ? recipients : _resolveRfiResponseRecipients(rfi, project);

  // Bullet-point detection: if the response uses "- " or "• " prefixes, treat
  // each line as a bullet so the DOCX renders them as a list. Otherwise plain
  // paragraphs.
  function responseBodyParagraphs() {
    const raw = String(response || "(no response text)").trim();
    const lines = raw.split(/\r?\n/).filter(l => l.trim().length > 0);
    const hasBullets = lines.some(l => /^[-•*]\s+/.test(l.trim()));
    if (!hasBullets) {
      return lines.map(line =>
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun({ text: line, font: "Calibri", size: 22 })],
        })
      );
    }
    return lines.map(line => {
      const stripped = line.trim().replace(/^[-•*]\s+/, "");
      return new Paragraph({
        spacing: { after: 100 },
        bullet: { level: 0 },
        children: [new TextRun({ text: stripped, font: "Calibri", size: 22 })],
      });
    });
  }

  function sectionLabel(text) {
    return new Paragraph({
      spacing: { before: 200, after: 80 },
      children: [new TextRun({ text, font: "Calibri", size: 18, bold: true, color: DOCX_COLORS.GRAY })],
    });
  }

  // The "ANSWER" boxed section. The Newforma version puts a bordered box
  // around everything from "Response (Answered) from" through the response
  // body. We approximate with a single-cell table that has visible borders.
  function buildAnswerBox() {
    const innerChildren = [
      // "Response (Answered) from: <name>" with bottom border like the PDF
      new Paragraph({
        spacing: { before: 100, after: 100 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: DOCX_COLORS.GRAY, space: 1 } },
        children: [new TextRun({ text: `Response (${purpose}) from: ${responder}`, font: "Calibri", size: 22, bold: true })],
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [new TextRun({ text: "Remarks:", font: "Calibri", size: 18, italics: true, color: DOCX_COLORS.GRAY })],
      }),
      new Paragraph({
        spacing: { before: 100, after: 80 },
        children: [new TextRun({ text: "General Notations:", font: "Calibri", size: 22, bold: true })],
      }),
      ...RFI_GENERAL_NOTATIONS.map(p =>
        new Paragraph({
          spacing: { after: 140 },
          children: [new TextRun({ text: p, font: "Calibri", size: 20 })],
        })
      ),
      new Paragraph({
        spacing: { before: 200, after: 80 },
        children: [new TextRun({ text: "Response:", font: "Calibri", size: 22, bold: true })],
      }),
      ...responseBodyParagraphs(),
      // Signature block — name / company / date
      new Paragraph({
        spacing: { before: 280, after: 0 },
        children: [new TextRun({ text: responder, font: "Calibri", size: 20 })],
      }),
      new Paragraph({
        spacing: { after: 0 },
        children: [new TextRun({ text: "Setty & Associates", font: "Calibri", size: 20 })],
      }),
      new Paragraph({
        spacing: { after: 100 },
        children: [new TextRun({ text: dateSent, font: "Calibri", size: 20 })],
      }),
    ];
    return new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [9360],
      rows: [new TableRow({
        children: [new TableCell({
          width: { size: 9360, type: WidthType.DXA },
          // Visible box border on all sides
          borders: {
            top:    { style: BorderStyle.SINGLE, size: 4, color: DOCX_COLORS.BOX_BORDER },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: DOCX_COLORS.BOX_BORDER },
            left:   { style: BorderStyle.SINGLE, size: 4, color: DOCX_COLORS.BOX_BORDER },
            right:  { style: BorderStyle.SINGLE, size: 4, color: DOCX_COLORS.BOX_BORDER },
          },
          children: innerChildren,
        })],
      })],
    });
  }

  // FROM table (Setty responder) — page 2
  function buildFromTable() {
    const colHeader = (label, w) => new TableCell({
      width: { size: w, type: WidthType.DXA },
      shading: { type: ShadingType.SOLID, color: "f1f5f9", fill: "f1f5f9" },
      children: [new Paragraph({ spacing: { before: 30, after: 30 }, children: [new TextRun({ text: label, font: "Calibri", size: 16, bold: true, color: DOCX_COLORS.GRAY })] })],
    });
    const colData = (text, w) => new TableCell({
      width: { size: w, type: WidthType.DXA },
      children: [new Paragraph({ spacing: { before: 30, after: 30 }, children: [new TextRun({ text: String(text || ""), font: "Calibri", size: 20 })] })],
    });
    return new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2200, 2200, 3300, 1660],
      rows: [
        new TableRow({ tableHeader: true, children: [
          colHeader("NAME", 2200), colHeader("COMPANY", 2200), colHeader("EMAIL", 3300), colHeader("PHONE", 1660),
        ]}),
        new TableRow({ children: [
          colData(responder, 2200), colData("Setty & Associates", 2200), colData(responderEmail, 3300), colData("", 1660),
        ]}),
      ],
    });
  }

  function buildToTable() {
    const colHeader = (label, w) => new TableCell({
      width: { size: w, type: WidthType.DXA },
      shading: { type: ShadingType.SOLID, color: "f1f5f9", fill: "f1f5f9" },
      children: [new Paragraph({ spacing: { before: 30, after: 30 }, children: [new TextRun({ text: label, font: "Calibri", size: 16, bold: true, color: DOCX_COLORS.GRAY })] })],
    });
    const colData = (text, w) => new TableCell({
      width: { size: w, type: WidthType.DXA },
      children: [new Paragraph({ spacing: { before: 30, after: 30 }, children: [new TextRun({ text: String(text || ""), font: "Calibri", size: 20 })] })],
    });
    const rows = [
      new TableRow({ tableHeader: true, children: [
        colHeader("NAME", 2200), colHeader("COMPANY", 2200), colHeader("EMAIL", 3300), colHeader("PHONE", 1660),
      ]}),
    ];
    if (toList.length === 0) {
      rows.push(new TableRow({ children: [
        colData("(no recipients on record)", 2200), colData("", 2200), colData("", 3300), colData("", 1660),
      ]}));
    } else {
      for (const r of toList) {
        rows.push(new TableRow({ children: [
          colData(r.name, 2200), colData(r.company, 2200), colData(r.email, 3300), colData(r.phone, 1660),
        ]}));
      }
    }
    return new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2200, 2200, 3300, 1660],
      rows,
    });
  }

  function buildContentsTable() {
    const colHeader = (label, w) => new TableCell({
      width: { size: w, type: WidthType.DXA },
      shading: { type: ShadingType.SOLID, color: "f1f5f9", fill: "f1f5f9" },
      children: [new Paragraph({ spacing: { before: 30, after: 30 }, children: [new TextRun({ text: label, font: "Calibri", size: 16, bold: true, color: DOCX_COLORS.GRAY })] })],
    });
    const colData = (text, w) => new TableCell({
      width: { size: w, type: WidthType.DXA },
      children: [new Paragraph({ spacing: { before: 30, after: 30 }, children: [new TextRun({ text: String(text || ""), font: "Calibri", size: 20 })] })],
    });
    const docxFileName = `${rfiId} Response ${dateSent.replace(/-/g, "")}.pdf`;
    return new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [800, 1400, 3760, 1100, 1100, 1200],
      rows: [
        new TableRow({ tableHeader: true, children: [
          colHeader("QTY", 800), colHeader("DATED", 1400), colHeader("TITLE", 3760), colHeader("NUMBER", 1100), colHeader("SCALE", 1100), colHeader("SIZE", 1200),
        ]}),
        new TableRow({ children: [
          colData("1", 800), colData(dateSent, 1400), colData(docxFileName, 3760), colData("", 1100), colData("", 1100), colData("", 1200),
        ]}),
      ],
    });
  }

  const headerBlock = new Header({
    children: [new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: DOCX_COLORS.RED, space: 1 } },
      spacing: { after: 120 },
      children: [
        new TextRun({ text: "SETTY", font: "Calibri", size: 36, bold: true, color: DOCX_COLORS.RED }),
        new TextRun({ text: "                                                                ", font: "Calibri", size: 16 }),
        new TextRun({ text: "RFI Transmittal", font: "Calibri", size: 24, bold: true, color: DOCX_COLORS.NAVY }),
      ],
    })],
  });

  const doc = new Document({
    styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
      headers: { default: headerBlock },
      children: [
        // Discipline label (matches the "SME" label at the top of Newforma's page 1)
        new Paragraph({
          spacing: { before: 0, after: 200 },
          children: [new TextRun({ text: (rfi.discipline || "").toUpperCase(), font: "Calibri", size: 18, color: DOCX_COLORS.GRAY })],
        }),
        // Metadata grid — 4 columns: LABEL | VALUE | LABEL | VALUE.
        // We build cells directly here rather than via the _docxInfoRow helper
        // because that helper returns a 2-cell row; the grid needs 4 cells per row.
        (function() {
          const labelCell = (text, w) => new TableCell({
            width: { size: w, type: WidthType.DXA },
            shading: { type: ShadingType.SOLID, color: "f8fafc", fill: "f8fafc" },
            children: [new Paragraph({ spacing: { before: 40, after: 40 }, children: [new TextRun({ text, font: "Calibri", size: 16, bold: true, color: DOCX_COLORS.GRAY })] })],
          });
          const valueCell = (text, w) => new TableCell({
            width: { size: w, type: WidthType.DXA },
            children: [new Paragraph({ spacing: { before: 40, after: 40 }, children: [new TextRun({ text: String(text || ""), font: "Calibri", size: 20 })] })],
          });
          const W = [1600, 3080, 1600, 3080];
          return new Table({
            width: { size: 9360, type: WidthType.DXA },
            columnWidths: W,
            rows: [
              new TableRow({ children: [
                labelCell("PROJECT:", W[0]),
                valueCell(`${projLabel}${projNumber ? "\n" + projNumber : ""}`, W[1]),
                labelCell("DATE SENT:", W[2]),
                valueCell(dateSent, W[3]),
              ]}),
              new TableRow({ children: [
                labelCell("SUBJECT:", W[0]),
                valueCell(subject, W[1]),
                labelCell("RFI ID:", W[2]),
                valueCell(rfiId, W[3]),
              ]}),
              new TableRow({ children: [
                labelCell("TYPE:", W[0]),
                valueCell("RFI", W[1]),
                labelCell("TRANSMITTAL ID:", W[2]),
                valueCell(transmittalId, W[3]),
              ]}),
              new TableRow({ children: [
                labelCell("PURPOSE:", W[0]),
                valueCell(purpose, W[1]),
                labelCell("VIA:", W[2]),
                valueCell(via, W[3]),
              ]}),
            ],
          });
        })(),

        sectionLabel("QUESTION:"),
        ...(String(rfi.description || rfi.title || "").trim() ? [
          new Paragraph({
            spacing: { after: 200 },
            indent: { left: 1600 },
            children: [new TextRun({ text: rfi.description || rfi.title || "", font: "Calibri", size: 22 })],
          })
        ] : [
          new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: " ", font: "Calibri", size: 22 })] }),
        ]),

        sectionLabel("SUGGESTION:"),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: " ", font: "Calibri", size: 22 })] }),

        sectionLabel("ANSWER:"),
        buildAnswerBox(),

        // Page break before the FROM/TO/CONTENTS tables (page 2)
        new Paragraph({
          children: [new TextRun({ text: "FROM", font: "Calibri", size: 16, color: DOCX_COLORS.GRAY })],
          spacing: { before: 240, after: 60 },
          pageBreakBefore: true,
        }),
        buildFromTable(),

        new Paragraph({
          children: [new TextRun({ text: "TO", font: "Calibri", size: 16, color: DOCX_COLORS.GRAY })],
          spacing: { before: 240, after: 60 },
        }),
        buildToTable(),

        new Paragraph({
          children: [new TextRun({ text: "DESCRIPTION OF CONTENTS", font: "Calibri", size: 16, color: DOCX_COLORS.GRAY })],
          spacing: { before: 240, after: 60 },
        }),
        buildContentsTable(),

        new Paragraph({
          children: [new TextRun({ text: "COPIES:", font: "Calibri", size: 16, color: DOCX_COLORS.GRAY })],
          spacing: { before: 240, after: 60 },
        }),
        new Paragraph({
          spacing: { after: 0 },
          children: [new TextRun({ text: responder, font: "Calibri", size: 20 })],
        }),
      ],
    }],
  });

  return await Packer.toBlob(doc);
}

function _resolveSubReviewRecipients(sub, project) {
  const recipients = [];
  if (sub.sourceItemId || sub.sourceMessageId) {
    const emailRec = (project.emails || []).find(e =>
      (sub.sourceItemId && e.msgId === sub.sourceItemId) ||
      (sub.sourceMessageId && e.msgId === sub.sourceMessageId)
    );
    if (emailRec?.fromAddress) {
      recipients.push({ name: emailRec.from || "", company: "", email: emailRec.fromAddress, phone: "" });
    }
  }
  if (recipients.length === 0 && sub.from) {
    recipients.push({ name: sub.from, company: "", email: "", phone: "" });
  }
  if (sub.subAssigned) {
    const s = (project.subconsultants || []).find(x => x.id === sub.subAssigned);
    if (s) {
      recipients.push({
        name: s.contact || s.firm || "",
        company: s.firm || "",
        email: s.email || "",
        phone: s.phone || "",
      });
    }
  }
  return recipients;
}

// Submittal Review DOCX — same two-page Newforma-style layout as the RFI
// version, with stamp + comments substituted for the answer/response, and
// submittal-appropriate boilerplate.
async function buildSubReviewDocx({ sub, project, comments, stamp, dateReturned, status, recipients }) {
  if (typeof docx === "undefined" || !docx.Packer) {
    throw new Error("DOCX library not loaded — refresh the taskpane and try again.");
  }
  const lib = docx;
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    Header, BorderStyle, WidthType, ShadingType,
  } = lib;
  const projLabel = project.name || "";
  const projNumber = project.projectNumber || "";
  const reviewer = msalAccount?.name || msalAccount?.username || "";
  const reviewerEmail = msalAccount?.username || "";
  const dateSent = dateReturned || new Date().toISOString().slice(0, 10);
  const subject = sub.description || "";
  const subId = sub.number || "";
  const purpose = stamp || "Reviewed";
  const transmittalId = `${subId}-${dateSent.replace(/-/g, "")}`;
  const via = sub.receivedVia || "Email";
  const toList = recipients && recipients.length ? recipients : _resolveSubReviewRecipients(sub, project);

  const stampColor =
    stamp === "Approved"            ? "059669" :
    stamp === "Approved as Noted"   ? "0891b2" :
    stamp === "Revise and Resubmit" ? "ea580c" :
    stamp === "Rejected"            ? "b91c1c" : DOCX_COLORS.NAVY;

  function commentsBodyParagraphs() {
    const raw = String(comments || "(no comments)").trim();
    const lines = raw.split(/\r?\n/).filter(l => l.trim().length > 0);
    const hasBullets = lines.some(l => /^[-•*]\s+/.test(l.trim()));
    if (!hasBullets) {
      return lines.map(line =>
        new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: line, font: "Calibri", size: 22 })] })
      );
    }
    return lines.map(line => {
      const stripped = line.trim().replace(/^[-•*]\s+/, "");
      return new Paragraph({
        spacing: { after: 100 }, bullet: { level: 0 },
        children: [new TextRun({ text: stripped, font: "Calibri", size: 22 })],
      });
    });
  }

  function sectionLabel(text) {
    return new Paragraph({
      spacing: { before: 200, after: 80 },
      children: [new TextRun({ text, font: "Calibri", size: 18, bold: true, color: DOCX_COLORS.GRAY })],
    });
  }

  function buildReviewBox() {
    const innerChildren = [
      new Paragraph({
        spacing: { before: 100, after: 100 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: DOCX_COLORS.GRAY, space: 1 } },
        children: [new TextRun({ text: `Review (${purpose}) by: ${reviewer}`, font: "Calibri", size: 22, bold: true })],
      }),
      new Paragraph({
        spacing: { after: 100 },
        children: [
          new TextRun({ text: "Stamp: ", font: "Calibri", size: 22, bold: true }),
          new TextRun({ text: (stamp || "—").toUpperCase(), font: "Calibri", size: 22, bold: true, color: stampColor }),
        ],
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [new TextRun({ text: "Remarks:", font: "Calibri", size: 18, italics: true, color: DOCX_COLORS.GRAY })],
      }),
      new Paragraph({
        spacing: { before: 100, after: 80 },
        children: [new TextRun({ text: "General Notations:", font: "Calibri", size: 22, bold: true })],
      }),
      ...SUB_GENERAL_NOTATIONS.map(p =>
        new Paragraph({ spacing: { after: 140 }, children: [new TextRun({ text: p, font: "Calibri", size: 20 })] })
      ),
      new Paragraph({
        spacing: { before: 200, after: 80 },
        children: [new TextRun({ text: "Comments:", font: "Calibri", size: 22, bold: true })],
      }),
      ...commentsBodyParagraphs(),
      new Paragraph({
        spacing: { before: 280, after: 0 },
        children: [new TextRun({ text: reviewer, font: "Calibri", size: 20 })],
      }),
      new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: "Setty & Associates", font: "Calibri", size: 20 })] }),
      new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: dateSent, font: "Calibri", size: 20 })] }),
    ];
    return new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [9360],
      rows: [new TableRow({
        children: [new TableCell({
          width: { size: 9360, type: WidthType.DXA },
          borders: {
            top:    { style: BorderStyle.SINGLE, size: 4, color: DOCX_COLORS.BOX_BORDER },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: DOCX_COLORS.BOX_BORDER },
            left:   { style: BorderStyle.SINGLE, size: 4, color: DOCX_COLORS.BOX_BORDER },
            right:  { style: BorderStyle.SINGLE, size: 4, color: DOCX_COLORS.BOX_BORDER },
          },
          children: innerChildren,
        })],
      })],
    });
  }

  // FROM/TO/CONTENTS table builders — same shape as RFI version
  const colHeader = (label, w) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: { type: ShadingType.SOLID, color: "f1f5f9", fill: "f1f5f9" },
    children: [new Paragraph({ spacing: { before: 30, after: 30 }, children: [new TextRun({ text: label, font: "Calibri", size: 16, bold: true, color: DOCX_COLORS.GRAY })] })],
  });
  const colData = (text, w) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    children: [new Paragraph({ spacing: { before: 30, after: 30 }, children: [new TextRun({ text: String(text || ""), font: "Calibri", size: 20 })] })],
  });

  const fromTable = new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2200, 2200, 3300, 1660],
    rows: [
      new TableRow({ tableHeader: true, children: [colHeader("NAME", 2200), colHeader("COMPANY", 2200), colHeader("EMAIL", 3300), colHeader("PHONE", 1660)] }),
      new TableRow({ children: [colData(reviewer, 2200), colData("Setty & Associates", 2200), colData(reviewerEmail, 3300), colData("", 1660)] }),
    ],
  });

  const toRows = [
    new TableRow({ tableHeader: true, children: [colHeader("NAME", 2200), colHeader("COMPANY", 2200), colHeader("EMAIL", 3300), colHeader("PHONE", 1660)] }),
  ];
  if (toList.length === 0) {
    toRows.push(new TableRow({ children: [colData("(no recipients on record)", 2200), colData("", 2200), colData("", 3300), colData("", 1660)] }));
  } else {
    for (const r of toList) {
      toRows.push(new TableRow({ children: [colData(r.name, 2200), colData(r.company, 2200), colData(r.email, 3300), colData(r.phone, 1660)] }));
    }
  }
  const toTable = new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2200, 2200, 3300, 1660], rows: toRows });

  const docxFileName = `${subId} Review ${dateSent.replace(/-/g, "")}.pdf`;
  const contentsTable = new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [800, 1400, 3760, 1100, 1100, 1200],
    rows: [
      new TableRow({ tableHeader: true, children: [colHeader("QTY", 800), colHeader("DATED", 1400), colHeader("TITLE", 3760), colHeader("NUMBER", 1100), colHeader("SCALE", 1100), colHeader("SIZE", 1200)] }),
      new TableRow({ children: [colData("1", 800), colData(dateSent, 1400), colData(docxFileName, 3760), colData("", 1100), colData("", 1100), colData("", 1200)] }),
    ],
  });

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
              new TextRun({ text: "SETTY", font: "Calibri", size: 36, bold: true, color: DOCX_COLORS.RED }),
              new TextRun({ text: "                                                                ", font: "Calibri", size: 16 }),
              new TextRun({ text: "Submittal Transmittal", font: "Calibri", size: 24, bold: true, color: DOCX_COLORS.NAVY }),
            ],
          })],
        }),
      },
      children: [
        new Paragraph({
          spacing: { before: 0, after: 200 },
          children: [new TextRun({ text: (sub.discipline || "").toUpperCase(), font: "Calibri", size: 18, color: DOCX_COLORS.GRAY })],
        }),
        // Metadata grid
        (function() {
          const labelCell = (text, w) => new TableCell({
            width: { size: w, type: WidthType.DXA },
            shading: { type: ShadingType.SOLID, color: "f8fafc", fill: "f8fafc" },
            children: [new Paragraph({ spacing: { before: 40, after: 40 }, children: [new TextRun({ text, font: "Calibri", size: 16, bold: true, color: DOCX_COLORS.GRAY })] })],
          });
          const valueCell = (text, w) => new TableCell({
            width: { size: w, type: WidthType.DXA },
            children: [new Paragraph({ spacing: { before: 40, after: 40 }, children: [new TextRun({ text: String(text || ""), font: "Calibri", size: 20 })] })],
          });
          const W = [1600, 3080, 1600, 3080];
          return new Table({
            width: { size: 9360, type: WidthType.DXA },
            columnWidths: W,
            rows: [
              new TableRow({ children: [
                labelCell("PROJECT:", W[0]), valueCell(`${projLabel}${projNumber ? "\n" + projNumber : ""}`, W[1]),
                labelCell("DATE SENT:", W[2]), valueCell(dateSent, W[3]),
              ]}),
              new TableRow({ children: [
                labelCell("SUBJECT:", W[0]), valueCell(subject, W[1]),
                labelCell("SUBMITTAL ID:", W[2]), valueCell(subId, W[3]),
              ]}),
              new TableRow({ children: [
                labelCell("SPEC SECTION:", W[0]), valueCell(sub.specSection || "", W[1]),
                labelCell("TRANSMITTAL ID:", W[2]), valueCell(transmittalId, W[3]),
              ]}),
              new TableRow({ children: [
                labelCell("STAMP:", W[0]),
                new TableCell({
                  width: { size: W[1], type: WidthType.DXA },
                  children: [new Paragraph({
                    spacing: { before: 40, after: 40 },
                    children: [new TextRun({ text: (stamp || "—").toUpperCase(), font: "Calibri", size: 22, bold: true, color: stampColor })],
                  })],
                }),
                labelCell("VIA:", W[2]), valueCell(via, W[3]),
              ]}),
            ],
          });
        })(),

        sectionLabel("DESCRIPTION:"),
        new Paragraph({
          spacing: { after: 200 },
          indent: { left: 1600 },
          children: [new TextRun({ text: sub.fullDescription || sub.description || "", font: "Calibri", size: 22 })],
        }),

        sectionLabel("REVIEW:"),
        buildReviewBox(),

        new Paragraph({
          spacing: { before: 240, after: 60 },
          pageBreakBefore: true,
          children: [new TextRun({ text: "FROM", font: "Calibri", size: 16, color: DOCX_COLORS.GRAY })],
        }),
        fromTable,
        new Paragraph({ spacing: { before: 240, after: 60 }, children: [new TextRun({ text: "TO", font: "Calibri", size: 16, color: DOCX_COLORS.GRAY })] }),
        toTable,
        new Paragraph({ spacing: { before: 240, after: 60 }, children: [new TextRun({ text: "DESCRIPTION OF CONTENTS", font: "Calibri", size: 16, color: DOCX_COLORS.GRAY })] }),
        contentsTable,
        new Paragraph({ spacing: { before: 240, after: 60 }, children: [new TextRun({ text: "COPIES:", font: "Calibri", size: 16, color: DOCX_COLORS.GRAY })] }),
        new Paragraph({ children: [new TextRun({ text: reviewer, font: "Calibri", size: 20 })] }),
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

// Seed a description textarea with the email's plain-text body. Async —
// Office hands the body back on a callback — so the fill is guarded: only
// applied if the field is still empty when the text arrives (never overwrite
// something the user started typing) and only if the pinned pane hasn't
// switched to a different email mid-fetch. Quoted reply history is cut at the
// first "From:" / "On … wrote:" marker since an RFI/submittal description
// wants the new ask, not the whole thread.
function seedDescriptionFromEmailBody(fieldId, maxChars = 600) {
  const item = emailItem;
  void (async () => {
    const raw = await getBodyTextReliable(item);
    if (!raw) return;
    if (emailItem !== item) return;
    const el = document.getElementById(fieldId);
    if (!el || el.value.trim()) return;
    let text = raw.replace(/\r\n/g, "\n").trim();
    const cut = text.search(/\n\s*(From:|-{3,}\s*Original Message|On .{10,80} wrote:)/i);
    if (cut > 0) text = text.slice(0, cut).trim();
    if (text.length > maxChars) text = text.slice(0, maxChars).trimEnd() + "…";
    if (text) el.value = text;
  })();
}

// Default an assignee <select> to the project's Setty PM when they're on the
// team roster — the most common routing for incoming RFIs/submittals. The
// dropdown options are "staff:<member.id>", so match the PM by name.
function defaultAssigneeToPm(selectId) {
  const pmName = (selectedProject?.settyPm || "").trim().toLowerCase();
  if (!pmName) return;
  const pm = (selectedProject?.teamMembers || []).find(m =>
    m?.id && (m.name || "").trim().toLowerCase() === pmName
  );
  const sel = document.getElementById(selectId);
  if (pm && sel && sel.querySelector(`option[value="staff:${pm.id}"]`)) {
    sel.value = "staff:" + pm.id;
  }
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
  defaultAssigneeToPm("rfiAssignedTo");
  seedDescriptionFromEmailBody("rfiDescription");
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

      // Build the new folder structure: <projFolder>/RFIs/<discCode>/<RFI-NNN>/IN/<date subject>
      // The per-email subfolder under /IN means multiple emails to the same RFI
      // coexist without colliding on "email.html". The RFI root URL is what
      // gets stored on the record so PMS navigation lands at the RFI level.
      let spFolderUrl = "";    // RFI root URL (stored on the record)
      let inFolderWebUrl = ""; // /IN folder URL (for the assignment email body link)
      let emailFolderUrl = ""; // specific per-email subfolder (for audit log)
      let spUploadError = "";
      if (freshProject.projectFolderUrl) {
        try {
          const token = await getToken();
          const { driveId } = await resolveSpIds();
          const projFolderName = decodeURIComponent(freshProject.projectFolderUrl.split("/").pop());
          const safeRfiNumber = rfiNumber.replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
          const rfisPath = await ensureSpFolder(driveId, token, projFolderName, "RFIs");
          const discPath = await ensureSpFolder(driveId, token, rfisPath, discCode);
          const rfiPath  = await ensureSpFolder(driveId, token, discPath, safeRfiNumber);
          await writeSpMetadataSidecar(driveId, token, rfiPath, buildAddinMetadata(freshProject, "rfi"));
          // Upload via the shared helper — per-email subfolder, real error
          // propagation. Throws on actual upload failure so we know.
          const uploadResult = await uploadEmailToArtifactInFolder({
            driveId, token, artifactRootPath: rfiPath, snapItem,
          });
          spFolderUrl    = SP_BASE_URL + "/" + rfiPath.split("/").map(encodeURIComponent).join("/");
          inFolderWebUrl = uploadResult.inFolderUrl;
          emailFolderUrl = uploadResult.emailFolderUrl;
        } catch (e) {
          console.warn("RFI SP upload failed:", e.message);
          spUploadError = e.message;
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
    try { refreshLinkToTargetDropdown(); } catch {}

      const baseMsg = "✓ " + rfiNumber + " logged" + (spFolderUrl ? " · filed to SharePoint" : "");
      const draftMsg = draftOpened ? " · ✉️ Draft opened" : (assignee?.email ? "" : (assignee ? " · ⚠ assignee has no email on file" : ""));
      const errMsg = spUploadError ? ` · ⚠ SP upload failed: ${spUploadError.slice(0, 120)}` : "";
      return {
        // sp_folder_url points at the per-email subfolder so the reconcile
        // sweep checks the right place. Falls back to the RFI root, then null,
        // depending on which step succeeded.
        sp_folder_url:  emailFolderUrl || spFolderUrl || null,
        status:         emailFolderUrl ? "success" : (spFolderUrl ? "partial" : "partial"),
        error:          emailFolderUrl ? null : (spUploadError || "RFI logged without SharePoint upload"),
        successMessage: baseMsg + draftMsg + errMsg,
      };
    }
  );
}
// Upload an email + its attachments INTO an RFI/Submittal's /IN folder, using
// a per-email subfolder so multiple emails coexist without colliding on
// "email.html". Matches the structure the project's main Emails folder uses.
//
// Returns { emailPath, emailFolderUrl, inPath, inFolderUrl, attCount } so
// callers can audit-log the precise upload location AND store the artifact
// root URL on the record (for navigation).
//
// Throws on creation failure so callers can surface the real error instead
// of silently filing a "logged without SharePoint upload" status.
async function uploadEmailToArtifactInFolder({ driveId, token, artifactRootPath, snapItem }) {
  if (!artifactRootPath) throw new Error("artifactRootPath missing");
  const inPath = await ensureSpFolder(driveId, token, artifactRootPath, "IN");
  // Per-email subfolder. Format mirrors the project Emails folder:
  // "YYYY-MM-DD <sanitized subject>". Capped at 60 chars to keep the path
  // length comfortably under SharePoint's 400-char URL ceiling.
  const dateObj = new Date(snapItem?.dateTimeCreated || Date.now());
  const datePart = dateObj.getFullYear() + "-" +
                   String(dateObj.getMonth() + 1).padStart(2, "0") + "-" +
                   String(dateObj.getDate()).padStart(2, "0");
  const subject = (snapItem?.subject || "Email")
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 60);
  const emailFolderName = (datePart + " " + subject).trim() || (datePart + " Email");
  const emailPath = await ensureSpFolder(driveId, token, inPath, emailFolderName);
  const attCount = await uploadEmailAndAttachments(driveId, token, emailPath, snapItem);
  return {
    emailPath,
    emailFolderUrl: SP_BASE_URL + "/" + emailPath.split("/").map(encodeURIComponent).join("/"),
    inPath,
    inFolderUrl:    SP_BASE_URL + "/" + inPath.split("/").map(encodeURIComponent).join("/"),
    attCount,
  };
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
      // ALWAYS construct the new-structure path for response OUT regardless of
      // where the RFI was originally filed. Legacy RFIs (logged before the
      // RFIs/<DiscCode>/RFI-NNN convention) keep their spFolderUrl pointing
      // at the old flat folder, but the response transmittal should land in
      // the canonical new path so the team can find it predictably. The new
      // discipline-coded path is created if it doesn't exist.
      const projFolderName = decodeURIComponent(selectedProject.projectFolderUrl.split("/").pop());
      const discCode = getDisciplineCode(rfi.discipline);
      const safeRfiNumber = (rfi.number || "RFI-???").replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
      const rfisPath = await ensureSpFolder(driveId, token, projFolderName, "RFIs");
      const discPath = await ensureSpFolder(driveId, token, rfisPath, discCode);
      const rfiRootPath = await ensureSpFolder(driveId, token, discPath, safeRfiNumber);
      const outPath = await ensureSpFolder(driveId, token, rfiRootPath, "OUT");
      // /IN is referenced for the email link, not created here — if no IN
      // folder exists yet (e.g. PMS-only filings), the chip link 404s, which
      // is the honest signal.
      const inPath = rfiRootPath + "/IN";

      // Upload the DOCX to OUT via the verified attachment path so it gets the
      // same integrity guarantees as everything else.
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
    const htmlBody = _buildRfiResponseEmailHtml({ rfi, project: selectedProject, response: responseText, dateResponded, outFolderUrl: outFolderWebUrl, inFolderUrl: inFolderWebUrl, status: newStatus });
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
    try { refreshLinkToTargetDropdown(); } catch {}
      showView("mainView");
    }, 1500);
  } catch (e) {
    setStatus("rfiResponseStatusMsg", "error", "✗ " + humanizeError(e));
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

function _buildRfiResponseEmailHtml({ rfi, project, response, dateResponded, outFolderUrl, inFolderUrl, status }) {
  const esc = (s) => String(s || "").replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" })[c]);
  const projLabel = [project.projectNumber, project.name].filter(Boolean).join(" ");
  const responderName = msalAccount?.name || msalAccount?.username || "";
  const dateDisplay = dateResponded || new Date().toISOString().slice(0, 10);
  const purpose = status === "Responded" ? "Answered" : (status || "Answered");
  const senderId = rfi.number || "";

  // Format response as bullet list if the user used "- " / "• " / "* " line
  // prefixes; otherwise render line breaks as <br>.
  function formatResponse(text) {
    const raw = String(text || "").trim();
    const lines = raw.split(/\r?\n/).filter(l => l.trim().length > 0);
    const hasBullets = lines.some(l => /^[-•*]\s+/.test(l.trim()));
    if (!hasBullets) {
      return lines.map(l => esc(l)).join("<br>");
    }
    return "<ul style='margin:6px 0 6px 18px;padding:0'>" +
      lines.map(l => {
        const stripped = l.trim().replace(/^[-•*]\s+/, "");
        return `<li style='margin-bottom:4px'>${esc(stripped)}</li>`;
      }).join("") +
      "</ul>";
  }

  const responseHtml = formatResponse(response);
  const generalNotationsHtml = RFI_GENERAL_NOTATIONS.map(p => `<p style="margin:8px 0">${esc(p)}</p>`).join("");

  // Recipients block — best-effort from rfi.from + assignee/sub
  const recipients = _resolveRfiResponseRecipients(rfi, project);
  const toBlock = recipients.length
    ? recipients.map(r => {
        const company = r.company ? ` (${esc(r.company)})` : "";
        return esc(r.name) + company;
      }).join("; ")
    : esc(rfi.from || "(no recipient on record)");

  const senderName = recipients[0]?.name || rfi.from || "";
  const senderCompany = recipients[0]?.company || "";

  return `
<div style="font-family: Calibri, Arial, sans-serif; font-size: 14px; color: #1f2937;">

  <p style="margin:6px 0"><strong>Project:</strong> ${esc(projLabel)}</p>

  <p style="margin:14px 0 8px 0; font-weight:bold;">Notification about RFI ${esc(rfi.title || "")}</p>

  <p style="margin:6px 0">This email contains the response for an RFI.</p>

  <hr style="border:none; border-top:1px solid #e5e7eb; margin:18px 0">

  <p style="margin:6px 0; font-weight:bold; font-size:16px;">Answer</p>

  <p style="margin:10px 0"><strong>Response (${esc(purpose)}) from:</strong> ${esc(responderName)}</p>

  <p style="margin:10px 0 4px 0"><strong>Remarks:</strong></p>

  <p style="margin:10px 0; font-weight:bold;">General Notations:</p>
  ${generalNotationsHtml}

  <p style="margin:14px 0 4px 0; font-weight:bold;">Response:</p>
  <div style="margin-left:6px">${responseHtml}</div>

  <p style="margin:20px 0 4px 0">${esc(responderName)}<br>
  Setty &amp; Associates<br>
  ${esc(dateDisplay)}</p>

  <hr style="border:none; border-top:1px solid #e5e7eb; margin:18px 0">

  <p style="margin:6px 0; font-weight:bold;">RFI Info</p>
  <p style="margin:4px 0"><strong>To:</strong> ${toBlock}</p>
  <p style="margin:4px 0"><strong>From:</strong> ${esc(senderName)}${senderCompany ? " (" + esc(senderCompany) + ")" : ""}</p>
  <p style="margin:4px 0"><strong>Purpose:</strong> ${esc(purpose)}</p>
  <p style="margin:4px 0"><strong>Sender ID:</strong> ${esc(senderId)}</p>
  <p style="margin:4px 0"><strong>Expiration Date:</strong> None</p>

  ${outFolderUrl || inFolderUrl ? `
    <hr style="border:none; border-top:1px solid #e5e7eb; margin:18px 0">
    <p style="margin:6px 0; font-weight:bold;">Transferred Files / Links</p>
    ${outFolderUrl ? `<p style="margin:4px 0"><a href="${esc(outFolderUrl)}">📁 RFI Response folder (OUT) — cover sheet DOCX</a></p>` : ""}
    ${inFolderUrl  ? `<p style="margin:4px 0"><a href="${esc(inFolderUrl)}">📁 Original RFI folder (IN)</a></p>` : ""}
  ` : ""}

</div>
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
      // Resolve the RFI's root folder. If the RFI predates the new structure,
      // rfi.spFolderUrl points at the old flat folder; creating IN + per-email
      // subfolders inside still works.
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
      // Per-email subfolder under /IN so multiple filed emails coexist.
      const uploadResult = await uploadEmailToArtifactInFolder({
        driveId, token, artifactRootPath: rfiRootPath, snapItem,
      });
      const attCount = uploadResult.attCount;
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
        // sp_folder_url points at the per-email subfolder so reconcile finds
        // the files in the right place.
        sp_folder_url:  uploadResult.emailFolderUrl,
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
  defaultAssigneeToPm("subAssignedTo");
  seedDescriptionFromEmailBody("subFullDescription");
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

      // New folder structure: <projFolder>/Submittals/<discCode>/<SUB-NNN>/IN/<date subject>
      let spFolderUrl = "";
      let inFolderWebUrl = "";
      let emailFolderUrl = "";
      let spUploadError = "";
      if (freshProject.projectFolderUrl) {
        try {
          const token = await getToken();
          const { driveId } = await resolveSpIds();
          const projFolderName = decodeURIComponent(freshProject.projectFolderUrl.split("/").pop());
          const safeSubNumber = subNumber.replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
          const subsPath = await ensureSpFolder(driveId, token, projFolderName, "Submittals");
          const discPath = await ensureSpFolder(driveId, token, subsPath, discCode);
          const subPath  = await ensureSpFolder(driveId, token, discPath, safeSubNumber);
          await writeSpMetadataSidecar(driveId, token, subPath, buildAddinMetadata(freshProject, "submittal"));
          const uploadResult = await uploadEmailToArtifactInFolder({
            driveId, token, artifactRootPath: subPath, snapItem,
          });
          spFolderUrl    = SP_BASE_URL + "/" + subPath.split("/").map(encodeURIComponent).join("/");
          inFolderWebUrl = uploadResult.inFolderUrl;
          emailFolderUrl = uploadResult.emailFolderUrl;
        } catch (e) {
          console.warn("Submittal SP upload failed:", e.message);
          spUploadError = e.message;
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
    try { refreshLinkToTargetDropdown(); } catch {}

      const baseMsg = "✓ " + subNumber + " logged" + (spFolderUrl ? " · filed to SharePoint" : "");
      const draftMsg = draftOpened ? " · ✉️ Draft opened" : (assignee && !assignee.email ? " · ⚠ assignee has no email on file" : "");
      const errMsg = spUploadError ? ` · ⚠ SP upload failed: ${spUploadError.slice(0, 120)}` : "";
      return {
        sp_folder_url:  emailFolderUrl || spFolderUrl || null,
        status:         emailFolderUrl ? "success" : "partial",
        error:          emailFolderUrl ? null : (spUploadError || "Submittal logged without SharePoint upload"),
        successMessage: baseMsg + draftMsg + errMsg,
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
      // ALWAYS construct the new-structure path for review OUT regardless of
      // where the submittal was originally filed. Matches the RFI Log Response
      // behavior — review transmittal lands at Submittals/<DiscCode>/SUB-NNN/OUT/.
      const projFolderName = decodeURIComponent(selectedProject.projectFolderUrl.split("/").pop());
      const discCode = getDisciplineCode(sub.discipline);
      const safeSubNumber = (sub.number || "SUB-???").replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
      const subsPath = await ensureSpFolder(driveId, token, projFolderName, "Submittals");
      const discPath = await ensureSpFolder(driveId, token, subsPath, discCode);
      const subRootPath = await ensureSpFolder(driveId, token, discPath, safeSubNumber);
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
      status: newStatus,
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
    try { refreshLinkToTargetDropdown(); } catch {}
      showView("mainView");
    }, 1500);
  } catch (e) {
    setStatus("subReviewStatusMsg", "error", "✗ " + humanizeError(e));
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

function _buildSubReviewEmailHtml({ sub, project, stamp, comments, dateReturned, outFolderUrl, inFolderUrl, status }) {
  const esc = (s) => String(s || "").replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" })[c]);
  const projLabel = [project.projectNumber, project.name].filter(Boolean).join(" ");
  const reviewerName = msalAccount?.name || msalAccount?.username || "";
  const dateDisplay = dateReturned || new Date().toISOString().slice(0, 10);
  const purpose = stamp || "Reviewed";
  const subId = sub.number || "";
  const stampColor =
    stamp === "Approved"            ? "#059669" :
    stamp === "Approved as Noted"   ? "#0891b2" :
    stamp === "Revise and Resubmit" ? "#ea580c" :
    stamp === "Rejected"            ? "#b91c1c" : "#1e3a8a";

  function formatBody(text) {
    const raw = String(text || "").trim();
    const lines = raw.split(/\r?\n/).filter(l => l.trim().length > 0);
    const hasBullets = lines.some(l => /^[-•*]\s+/.test(l.trim()));
    if (!hasBullets) return lines.map(l => esc(l)).join("<br>");
    return "<ul style='margin:6px 0 6px 18px;padding:0'>" +
      lines.map(l => {
        const stripped = l.trim().replace(/^[-•*]\s+/, "");
        return `<li style='margin-bottom:4px'>${esc(stripped)}</li>`;
      }).join("") + "</ul>";
  }
  const commentsHtml = formatBody(comments);
  const generalNotationsHtml = SUB_GENERAL_NOTATIONS.map(p => `<p style="margin:8px 0">${esc(p)}</p>`).join("");

  const recipients = _resolveSubReviewRecipients(sub, project);
  const toBlock = recipients.length
    ? recipients.map(r => esc(r.name) + (r.company ? ` (${esc(r.company)})` : "")).join("; ")
    : esc(sub.from || "(no recipient on record)");
  const senderName = recipients[0]?.name || sub.from || "";
  const senderCompany = recipients[0]?.company || "";

  return `
<div style="font-family: Calibri, Arial, sans-serif; font-size: 14px; color: #1f2937;">

  <p style="margin:6px 0"><strong>Project:</strong> ${esc(projLabel)}</p>

  <p style="margin:14px 0 8px 0; font-weight:bold;">Notification about Submittal ${esc(sub.description || "")}</p>

  <p style="margin:6px 0">This email contains the review for a submittal.</p>

  <hr style="border:none; border-top:1px solid #e5e7eb; margin:18px 0">

  <p style="margin:6px 0; font-weight:bold; font-size:16px;">Review</p>

  <p style="margin:10px 0"><strong>Review (${esc(purpose)}) by:</strong> ${esc(reviewerName)}</p>

  <p style="margin:10px 0"><strong>Stamp:</strong> <span style="color:${stampColor};font-weight:bold;text-transform:uppercase">${esc(stamp || "—")}</span></p>

  <p style="margin:10px 0 4px 0"><strong>Remarks:</strong></p>

  <p style="margin:10px 0; font-weight:bold;">General Notations:</p>
  ${generalNotationsHtml}

  ${comments ? `
    <p style="margin:14px 0 4px 0; font-weight:bold;">Comments:</p>
    <div style="margin-left:6px">${commentsHtml}</div>
  ` : ""}

  <p style="margin:20px 0 4px 0">${esc(reviewerName)}<br>
  Setty &amp; Associates<br>
  ${esc(dateDisplay)}</p>

  <hr style="border:none; border-top:1px solid #e5e7eb; margin:18px 0">

  <p style="margin:6px 0; font-weight:bold;">Submittal Info</p>
  <p style="margin:4px 0"><strong>To:</strong> ${toBlock}</p>
  <p style="margin:4px 0"><strong>From:</strong> ${esc(senderName)}${senderCompany ? " (" + esc(senderCompany) + ")" : ""}</p>
  <p style="margin:4px 0"><strong>Purpose:</strong> ${esc(purpose)}</p>
  <p style="margin:4px 0"><strong>Sender ID:</strong> ${esc(subId)}</p>
  ${sub.specSection ? `<p style="margin:4px 0"><strong>Spec Section:</strong> ${esc(sub.specSection)}</p>` : ""}

  ${outFolderUrl || inFolderUrl ? `
    <hr style="border:none; border-top:1px solid #e5e7eb; margin:18px 0">
    <p style="margin:6px 0; font-weight:bold;">Transferred Files / Links</p>
    ${outFolderUrl ? `<p style="margin:4px 0"><a href="${esc(outFolderUrl)}">📁 Submittal Review folder (OUT) — cover sheet DOCX</a></p>` : ""}
    ${inFolderUrl  ? `<p style="margin:4px 0"><a href="${esc(inFolderUrl)}">📁 Original submittal folder (IN)</a></p>` : ""}
  ` : ""}

</div>
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
      // Per-email subfolder under /IN so multiple filed emails coexist.
      const uploadResult = await uploadEmailToArtifactInFolder({
        driveId, token, artifactRootPath: subRootPath, snapItem,
      });
      const attCount = uploadResult.attCount;
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
        // Audit log points at the per-email subfolder so reconcile finds files there.
        sp_folder_url:  uploadResult.emailFolderUrl,
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
  const prefix  = project.projectNumber ? project.projectNumber + " " : "";
  const subject = prefix + project.name + " — " + milestone.name;
  const pmName  = (project.settyPm || "").trim();
  const event = {
    subject,
    isAllDay: true,
    // Surface the PM on the calendar's Location line so a glance at the day
    // view shows who owns the milestone. Graph's `location` is an object.
    ...(pmName ? { location: { displayName: "PM: " + pmName } } : {}),
    // No reminder — these are reference markers on the calendar, not things to
    // be alerted about. Must be explicit: omitting it lets Graph apply the
    // mailbox's default reminder, which was spamming everyone.
    isReminderOn: false,
    start: { dateTime: milestone.dueDate + "T00:00:00", timeZone: "Eastern Standard Time" },
    end:   { dateTime: endStr          + "T00:00:00", timeZone: "Eastern Standard Time" },
    categories: ["PMS Milestone"],
  };
  try {
    const token = await getToken();
    // Carry the source email's body into the event so the calendar entry has
    // the context it came from. Best-effort: skipped for manual entries,
    // appointments, or if the body can't be fetched (getEmailBodyHtml → "").
    try {
      const bodyHtml = (currentItemKind === "message") ? await getEmailBodyHtml(token) : "";
      if (bodyHtml) {
        const srcSubject = (emailItem?.subject || "").trim();
        const srcFrom    = (emailFrom || emailFromAddress || "").trim();
        const header = (srcSubject || srcFrom)
          ? `<p style="margin:0 0 8px"><b>From email:</b> ${escHtml(srcSubject)}${srcFrom ? " — " + escHtml(srcFrom) : ""}</p><hr>`
          : "";
        event.body = { contentType: "HTML", content: header + bodyHtml };
      }
    } catch { /* body is a nice-to-have; never block the event on it */ }
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
    const calLabel = calResult.onShared ? "NYC Shared Calendar" : "your personal calendar";
    const successMsg = calResult.success
      ? pep + " Saved to " + projLabel + " · synced to " + calLabel
      : pep + " Saved to " + projLabel + " (calendar sync failed: " + calResult.error + ")";
    setStatus("milestoneStatus", "success", successMsg);
    // Auto-return to the main pane after a beat — staying on a bare status +
    // back button cost an extra click on every milestone save. The success
    // message carries over to the main status slot so it isn't lost.
    setTimeout(() => {
      // Skip if the user already navigated somewhere else in the meantime.
      if (!document.getElementById("datesView")?.classList.contains("active")) return;
      const form = document.getElementById("milestoneForm");
      if (form) form.style.display = "none";
      setStatus("milestoneStatus", "", "");
      showView("mainView");
      setStatus("actionStatus", "success", successMsg);
    }, 1200);
  } catch(e) {
    setStatus("milestoneStatus", "error", "✗ " + humanizeError(e));
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
// Where the contact form should return to on Back / after save. Defaults to the
// participant list; the enrich-from-main shortcut sets it to "mainView" so the
// user lands back where they started with the (now-cleared) nudge.
let _contactReturnView = "peopleView";
// Main-view button click: when the button exists only to enrich the sender from
// their signature, skip the participant list and open their form directly.
// Otherwise behave as before and show the list.
function onAddParticipantClick() {
  const { enrichOnly } = peopleButtonMode();
  if (enrichOnly) {
    const senderAddr = (emailFromAddress || "").trim().toLowerCase();
    const sender = (emailParticipants || []).find(p => (p.emailAddress || "").trim().toLowerCase() === senderAddr);
    if (sender) {
      _contactReturnView = "mainView";
      prefillContactFromParticipant(sender);
      return;
    }
  }
  _contactReturnView = "peopleView";
  showPeopleView();
}
// After a contact save, return where the user came from. From the enrich-from-
// main shortcut that's the main screen; the badge refresh clears the nudge
// (the sender is now in _sessionSavedContactEmails). From the list it's the list,
// preserving the "work through several participants" flow.
function returnAfterContactSave() {
  if (_contactReturnView === "mainView") {
    showView("mainView");
    updatePeopleButtonBadge();
  } else {
    showPeopleView();
  }
}
function showPeopleView() {
  // Rows opened from here return to the list (not main) — the enrich-from-main
  // shortcut overrides this before opening the form.
  _contactReturnView = "peopleView";
  const list = document.getElementById("participantList");
  // Setty staff are managed via the project's Teams tab in PMS — leave them
  // out of the list entirely instead of rendering inert rows.
  const externalRows = (emailParticipants || [])
    .map((p, i) => ({ p, i, st: getParticipantDirectoryStatus(p) }))
    .filter(r => !r.st.internal);
  if (!externalRows.length) {
    list.innerHTML = '<p style="font-size:12px;color:var(--text-soft);">No external participants on this email.</p>';
  } else {
    const labelColor = { From: "#c50f1f", To: "#0f6cbd", CC: "#0e6d5c", Required: "#0f6cbd", Optional: "#616161", Organizer: "#c50f1f" };
    const labelBg    = { From: "#fde7e9", To: "#eaf3fb", CC: "#e0f5f0", Required: "#eaf3fb", Optional: "#f3f2f1", Organizer: "#fde7e9" };
    // Sort the people worth capturing to the top: unknown first, then known-
    // globally-but-new-to-this-project, then fully filed. Stable within
    // groups (From/To/CC order preserved via original index).
    const rows = externalRows;
    const rank = r => r.st.isNew ? 0 : (r.st.globalHit && !r.st.inProject && !r.st.sessionSaved ? 1 : 2);
    rows.sort((a, b) => rank(a) - rank(b) || a.i - b.i);
    list.innerHTML = rows.map(({ p, i, st }) => {
      const fullyFiled = st.sessionSaved || st.inProject;
      let statusHtml = "";
      if (st.sessionSaved) {
        statusHtml = '<span class="pill added">✓ Added</span>';
      } else if (st.inProject) {
        statusHtml = '<span class="pill added" title="Already in this project\'s directory">✓ On project</span>';
      } else if (st.globalHit) {
        const company = escHtml(st.globalHit.client?.name || "directory");
        statusHtml = `<span class="pill added company" title="Already in the global directory under ${company}">✓ ${company}</span>`
          + (selectedProject
            ? `<button type="button" class="quick-add-proj" data-idx="${i}" title="Add to ${escHtml(selectedProject.name || "this project")}'s directory — no retyping">+ project</button>`
            : "");
      } else {
        statusHtml = '<span class="pill" style="background:var(--primary);color:#fff;" title="Not in any directory yet — click the row to add">+ Add</span>';
      }
      return `
      <div class="participant-row${fullyFiled ? ' added' : ''}" data-idx="${i}"${fullyFiled ? ' style="opacity:0.8;"' : ''}>
        <div class="participant-id">
          <div style="font-size:13px;font-weight:600;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
            ${escHtml(p.displayName || p.emailAddress)}
          </div>
          ${p.displayName && p.displayName !== p.emailAddress ? `<div style="font-size:11px;color:var(--text-soft);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
            ${escHtml(p.emailAddress || "")}
          </div>` : ""}
        </div>
        ${statusHtml}
        <span class="pill" style="background:${labelBg[p.label]||'var(--surface-2)'};color:${labelColor[p.label]||'var(--text-soft)'};">
          ${escHtml(p.label || "")}
        </span>
      </div>`;
    }).join("");
    list.querySelectorAll(".participant-row").forEach(el => {
      // data-idx is the participant's index in emailParticipants (assigned
      // before the internal-staff filter), so lookups stay correct.
      el.onclick = () => prefillContactFromParticipant(emailParticipants[+el.dataset.idx]);
    });
    list.querySelectorAll(".quick-add-proj").forEach(btn => {
      btn.onclick = (e) => {
        e.stopPropagation();
        quickAddToProjectDirectory(+btn.dataset.idx, btn);
      };
    });
    // Nudge: the sender is already filed but their record has neither title nor
    // phone — this email's signature can fill both in one tap. Skipped when
    // another action just posted a status (e.g. quick-add success).
    const senderRec = senderNeedsEnrichment();
    const statusEl = document.getElementById("peopleStatus");
    if (senderRec && statusEl && !statusEl.textContent) {
      setStatus("peopleStatus", "info", "💡 " + (senderRec.name || emailFrom || "The sender") + "'s saved contact has no title or phone — tap their row to fill it from this email's signature.");
    }
  }
  updatePeopleButtonBadge();
  showView("peopleView");
}

// One-click "+ project": the person is already in the global directory, so
// copy their existing record into the selected project's directory without
// making the user re-type anything on the contact form.
async function quickAddToProjectDirectory(idx, btnEl) {
  const p = emailParticipants[idx];
  const hit = findGlobalContact(p?.emailAddress || "");
  if (!hit || !selectedProject) return;
  if (btnEl) { btnEl.disabled = true; btnEl.textContent = "adding…"; }
  const ct = hit.contact;
  const dirEntry = {
    id: uid(),
    name: ct.name || p.displayName || "",
    title: ct.title || "",
    email: ct.email || p.emailAddress || "",
    phone: ct.phone || "",
    company: hit.client?.name || "",
    type: "Other",
    addedAt: new Date().toISOString(),
    addedBy: msalAccount?.username || "",
    addedFromEmail: emailItem?.itemId || "",
    notes: "",
  };
  const emailLc = (dirEntry.email || "").toLowerCase();
  try {
    await applyLocalChangeAndSave(selectedProject.id, fresh => {
      const dir = fresh.directory || [];
      if (emailLc && dir.some(d => (d.email || "").toLowerCase() === emailLc)) return fresh;
      return { ...fresh, directory: [...dir, dirEntry] };
    });
    setStatus("peopleStatus", "success", "✓ Added " + (dirEntry.name || dirEntry.email) + " to " + (selectedProject.name || "project") + " directory.");
    // Re-render: applyLocalChangeAndSave updated selectedProject locally, so
    // the row flips to "✓ On project" and re-sorts on its own.
    showPeopleView();
  } catch (e) {
    if (btnEl) { btnEl.disabled = false; btnEl.textContent = "+ project"; }
    setStatus("peopleStatus", "error", "✗ " + humanizeError(e));
  }
}
function prefillContactFromParticipant(p) {
  const matchedClient = getClientByEmail(p.emailAddress || "");
  // Known contact → show what's already on file, so opening the form doubles
  // as a record check. Signature parsing then fills only what's still blank.
  const existing = findGlobalContact(p.emailAddress || "")?.contact || null;
  document.getElementById("contactName").value    = existing?.name  || p.displayName || "";
  document.getElementById("contactTitle").value   = existing?.title || "";
  document.getElementById("contactCompany").value = matchedClient?.name || "";
  document.getElementById("contactEmail").value   = p.emailAddress || "";
  document.getElementById("contactPhone").value   = existing?.phone || "";
  setStatus("contactStatus", "", "");
  maybePrefillFromSignature(p);
  showView("contactView");
}

// ── SIGNATURE PARSING (title + phone prefill) ────────────────────────────────
// When the person being added is the SENDER, their signature block is at the
// bottom of this email — scrape title and phone from it so the contact form
// is a confirm, not a typing exercise. To/CC participants are skipped: their
// signatures only appear in quoted history (if at all), and pulling a title
// from the wrong person's signature is worse than leaving the field blank.
function maybePrefillFromSignature(p) {
  const pAddr = (p?.emailAddress || "").trim().toLowerCase();
  const senderAddr = (emailFromAddress || "").trim().toLowerCase();
  if (!pAddr || pAddr !== senderAddr) return;
  const item = emailItem;
  void (async () => {
    const bodyText = await getBodyTextReliable(item);
    if (bodyText === null) { console.info("[signature] body unavailable after retries — prefill skipped"); return; }
    if (emailItem !== item) return; // pinned pane switched emails mid-fetch
    // Bail if the form has moved on to a different person meanwhile.
    const emailField = document.getElementById("contactEmail");
    if (!emailField || (emailField.value || "").trim().toLowerCase() !== pAddr) return;
    const company = document.getElementById("contactCompany")?.value || "";
    const sig = parseSenderSignature(bodyText, p.displayName || emailFrom || "", company);
    const titleEl = document.getElementById("contactTitle");
    const phoneEl = document.getElementById("contactPhone");
    if (sig.title && titleEl && !titleEl.value.trim()) titleEl.value = sig.title;
    if (sig.phone && phoneEl && !phoneEl.value.trim()) phoneEl.value = sig.phone;
  })();
}

// Job-title vocabulary for the fallback scan. AEC-heavy on purpose — extend
// this list if a common title at your consultants/GCs slips through.
const SIGNATURE_TITLE_WORDS = /\b(engineer|architect|manager|director|principal|president|vice president|associate|designer|coordinator|administrator|estimator|superintendent|partner|owner|founder|specialist|consultant|surveyor|planner|drafter|technician|officer|executive|chief|lead|vp|ceo|coo|cfo|pm|leed)\b/i;
const SIGNATURE_PHONE_RE = /(\+?\d{1,2}[\s.\-])?(\(\d{3}\)|\d{3})[\s.\-]?\d{3}[\s.\-]?\d{4}(\s?(x|ext\.?)\s?\d{1,5})?/;

function parseSenderSignature(bodyText, senderName, companyName) {
  const out = { title: "", phone: "" };
  let text = (bodyText || "").replace(/\r\n/g, "\n");
  // Only the current message — signatures in quoted history belong to other
  // people (or older versions of this one).
  const cut = text.search(/\n\s*(From:|-{3,}\s*Original Message|On .{10,80} wrote:)/i);
  if (cut > 0) text = text.slice(0, cut);
  // Signature = the tail of the message.
  const lines = text.split("\n").map(l => l.trim()).filter(Boolean).slice(-15);
  if (!lines.length) return out;

  const companyLc = (companyName || "").trim().toLowerCase();
  const nameTokens = (senderName || "").toLowerCase().split(/[\s,]+/).filter(t => t.length > 1 && !/^(jr|sr|ii|iii|iv|pe|aia|leed)\.?$/.test(t));
  const isNameLine  = l => { const lc = l.toLowerCase(); return nameTokens.length >= 1 && nameTokens.filter(t => lc.includes(t)).length >= Math.min(2, nameTokens.length); };
  const isNoiseLine = l => /@/.test(l) || /\b(www\.|https?:)/i.test(l) || SIGNATURE_PHONE_RE.test(l);
  const isCompanyLine = l => !!companyLc && l.toLowerCase().includes(companyLc);
  const titleOk = l => l.length >= 3 && l.length <= 70 && /[a-z]/i.test(l) && !isNoiseLine(l) && !isCompanyLine(l) && !isNameLine(l);
  // Multi-part lines ("Senior PM | RPH Architects") — judge each segment.
  const titleFromLine = l => {
    for (const seg of l.split(/\s*[|•·]\s*/)) {
      if (titleOk(seg)) return seg.trim();
    }
    return "";
  };

  // Primary: the line right after the sender's name in the signature block
  // (Name / Title / Company is the overwhelmingly common layout). Search from
  // the bottom up so a "Hi Bob," greeting line can't be mistaken for the sig.
  for (let i = lines.length - 1; i >= 0 && !out.title; i--) {
    if (!isNameLine(lines[i])) continue;
    for (const next of [lines[i + 1], lines[i + 2]]) {
      if (!next) break;
      const t = titleFromLine(next);
      if (t) { out.title = t; break; }
    }
    break; // bottom-most name line is the signature; don't keep walking up
  }
  // Fallback: any tail line that uses job-title vocabulary.
  if (!out.title) {
    for (const l of lines) {
      const t = titleFromLine(l);
      if (t && SIGNATURE_TITLE_WORDS.test(t)) { out.title = t; break; }
    }
  }

  // Phone: prefer cell/mobile, then direct, then anything. Label is whatever
  // short tag precedes the number ("C:", "Cell", "M.", "Direct") — a line can
  // carry several numbers ("T 212… | C 917…"), so every match is scored.
  let best = null;
  for (const l of lines) {
    const re = new RegExp(SIGNATURE_PHONE_RE.source, "g");
    let m;
    while ((m = re.exec(l)) !== null) {
      const prefix = l.slice(0, m.index).toLowerCase();
      const score = /\b(c|cell|m|mob|mobile)[.: )]*$/.test(prefix) ? 3
                  : /\b(d|direct|dd)[.: )]*$/.test(prefix) ? 2 : 1;
      if (!best || score > best.score) best = { value: m[0].trim(), score };
    }
  }
  if (best) out.phone = best.value;
  return out;
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
// Opens the original email behind a watchlist row. Prefers Outlook's native
// message form so desktop users stay in their client; falls back to the web
// link (OWA in a browser) when there's no item id or the native call fails.
function openWatchlistEmail(itemId, webLink) {
  if (itemId) {
    try {
      if (Office?.context?.mailbox?.displayMessageForm) {
        Office.context.mailbox.displayMessageForm(itemId);
        return;
      }
    } catch (e) {
      console.warn("displayMessageForm failed, falling back to web link:", e);
    }
  }
  if (webLink) openExternalUrl(webLink);
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
  let resolvedCompany = ""; // which company the contact actually landed under (for the success message)
  try {
    if (saveTo === "client") {
      // V2: write to pms_clients (per-client rows). Previously this PATCHed
      // pms_data.clients (legacy singleton blob), which post-migration is
      // never re-read by PMS — so the contact silently disappeared. Now we
      // upsert the client row directly.
      //
      // Company resolution order:
      //   1. what the user typed in the Company field
      //   2. an existing directory company whose contacts share this email's
      //      domain — prevents auto-creating a company named after the PERSON
      //      when the field is blank (June 2026: "Kunwar Rana" became a
      //      Global Directory company even though Code Green already held
      //      @codegreen.com contacts)
      //   3. the person's name (legacy last resort)
      const domainClient = !company ? getClientByEmail(email) : null;
      const targetCompany = (company || domainClient?.name || name).trim();
      // Per-contact role removed from the schema (was redundant once company
      // discipline + per-project relationship were split). `lastContacted` is
      // stamped here so every Outlook-driven capture auto-updates recency —
      // the field no longer needs manual entry in PMS.
      const today = new Date().toISOString().slice(0, 10);
      const contact = { id: uid(), name, title, email, phone, lastContacted: today };
      // Find existing client by exact (case-insensitive) name match.
      // Narrow server-side with ilike so we don't download the whole clients
      // table on every contact save — escape LIKE metacharacters so names
      // containing % _ * match literally, and wrap in wildcards so DB names
      // with stray whitespace padding still come back. The .find() below
      // remains the authority (same trim/lowercase equality as before).
      const likeSafe = targetCompany.replace(/\\/g, "\\\\").replace(/[%_*]/g, m => "\\" + m);
      const queryUrl = SUPABASE_URL + "/rest/v1/pms_clients?select=id,client,version&client->>name=ilike." + encodeURIComponent("*" + likeSafe + "*");
      const all = await fetch(queryUrl, { headers: SB_HEADERS });
      if (!all.ok) throw new Error("pms_clients GET HTTP " + all.status);
      const rows = await all.json();
      let existing = (rows || []).find(r => r.client?.name && r.client.name.trim().toLowerCase() === targetCompany.toLowerCase());
      // Exact match failed → normalized fallbacks before concluding "new
      // company". Squash punctuation/spacing/legal suffixes (LLP, PC, …),
      // then allow prefix containment: "OGS" ≈ "OGS - New York State",
      // "Code Green" ≈ "Code Green Solutions". Exact-only matching is what
      // created the second OGS row in the Global Directory (June 2026).
      // Mirrors companySquash()/findDuplicateCompanies() in SettyPMS.html.
      if (!existing) {
        const squashName = s => (s || "").toLowerCase()
          .replace(/\b(llp|llc|inc|pc|pllc|dpc|corp|co|ltd|pa)\b\.?/g, "")
          .replace(/[^a-z0-9]/g, "");
        const targetSquash = squashName(targetCompany);
        existing = (rows || []).find(r => targetSquash && squashName(r.client?.name) === targetSquash)
          || (rows || []).find(r => {
               const n = squashName(r.client?.name);
               return targetSquash.length >= 3 && n.length >= 3 && (n.startsWith(targetSquash) || targetSquash.startsWith(n));
             });
      }
      resolvedCompany = existing?.client?.name || targetCompany;
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
          // Don't duplicate — bump lastContacted, and backfill any fields the
          // existing record is missing (older contacts predate signature
          // extraction, so title/phone are often blank). Existing values
          // always win: this enriches, never overwrites.
          ec.contacts = ec.contacts.map((c, i) => i === dupIdx ? {
            ...c,
            name:  c.name  || name,
            title: c.title || title,
            phone: c.phone || phone,
            lastContacted: today,
          } : c);
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
        // Sync the in-memory directory with what we just persisted. allClients is
        // loaded once per session (and cached), NOT refetched per email — without
        // this, findGlobalContact keeps returning the stale, pre-enrichment record,
        // so the signature nudge re-fires every time the email is reopened even
        // though the title/phone are now saved.
        const _cid = ec.id || existing.id;
        const _savedLc = (email || "").trim().toLowerCase();
        let _synced = false;
        allClients = (allClients || []).map(c => {
          if (_synced) return c;
          const idHit = _cid && c?.id === _cid;
          const emailHit = _savedLc && (c?.contacts || []).some(x => (x.email || "").trim().toLowerCase() === _savedLc);
          if (idHit || emailHit) { _synced = true; return ec; }
          return c;
        });
        if (!_synced) allClients = [...(allClients || []), ec];
        renderCompanySuggestions();
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
            // Already on this project? Backfill only the blank fields (enrich,
            // never overwrite) so a captured title/phone reaches the project's
            // directory too — not just the global one. Without this, the
            // project "people on this job" entry stays blank and the enrichment
            // nudge keeps re-firing for project-filed senders.
            const existingIdx = emailLc ? dir.findIndex(d => (d.email || "").toLowerCase() === emailLc) : -1;
            if (existingIdx >= 0) {
              const cur = dir[existingIdx];
              if ((cur.title || "").trim() && (cur.phone || "").trim()) return fresh; // nothing to fill
              const nextDir = dir.slice();
              nextDir[existingIdx] = {
                ...cur,
                name:  cur.name  || name,
                title: cur.title || title,
                phone: cur.phone || phone,
              };
              return { ...fresh, directory: nextDir };
            }
            // No email to dedup on → allow the add (user can clean up later; the
            // alternative is silently dropping otherwise-valid entries).
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
    // resolvedCompany may differ from what was typed (domain match or
    // normalized-name match) — show where the contact ACTUALLY landed so a
    // silent redirect is never invisible to the user.
    const destLabel = saveTo === "client"
      ? (resolvedCompany || company || name)
      : ((selectedProject?.projectNumber ? selectedProject.projectNumber + " — " : "") + (selectedProject?.name || "project POC"));
    // When saved to a client AND a project is tagged, the contact also lands
    // in the project's directory — surface that in the success message so users
    // know it's findable in PMS without having to check.
    const alsoInDirectory = saveTo === "client" && selectedProject;
    const dirSuffix = alsoInDirectory ? " · added to " + (selectedProject.name || "project") + " directory." : "";
    setStatus("actionStatus", "success", "✓ Saved " + (name || email) + " to " + destLabel + "." + dirSuffix);
    setStatus("contactStatus", "", "");
    returnAfterContactSave();
    return;
  } catch (e) {
    if (e.message === "__DUP__") {
      // Treat dup as a benign success for the "add another" flow — mark them
      // as ✓ and return to the list rather than stranding the user on the form.
      const savedEmailKey = (email || "").toLowerCase();
      if (savedEmailKey) _sessionSavedContactEmails.add(savedEmailKey);
      setStatus("actionStatus", "info", "Already in this project's POC list — no duplicate added.");
      setStatus("contactStatus", "", "");
      returnAfterContactSave();
      return;
    }
    setStatus("contactStatus", "error", "✗ " + humanizeError(e));
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
// While a "⏳ …" busy message is showing, this holds its element id so the
// attachment upload loop can push progress updates ("2/5") into the same slot.
let _busyStatusElId = null;

function setStatus(elId, type, msg) {
  const el = document.getElementById(elId);
  if (!el) return;
  el.className = "status-msg" + (msg ? " show " + type : "");
  // "⏳" prefix = busy message. Swap the static emoji for the animated
  // spinner so users can tell a working save from a stuck one — every
  // existing call site opts in just by keeping the emoji convention.
  if (msg && msg.startsWith("⏳")) {
    el.textContent = msg.slice(1).trim();
    const spin = document.createElement("span");
    spin.className = "spinner";
    spin.style.cssText = "vertical-align:-2px;margin-right:6px;";
    el.prepend(spin);
    _busyStatusElId = elId;
  } else {
    el.textContent = msg;
    if (_busyStatusElId === elId) _busyStatusElId = null;
  }
}

// Map raw API/SDK failures to plain-language guidance. The raw message still
// goes to the console for debugging — the line the user sees should say what
// happened and what to DO, never an HTTP status or a JSON dump.
function humanizeError(err) {
  const raw = (err?.message || String(err || "")).trim();
  console.warn("[error shown to user]", raw);
  const m = raw.toLowerCase();
  if (m.includes("graph 401") || m.includes("interaction_required") || m.includes("login_required") || m.includes("not signed in"))
    return "Your session expired — sign out (⋯ menu) and back in, then retry.";
  if (m.includes("graph 403"))
    return "You don't have permission for that SharePoint location — check your access with IT.";
  if (m.includes("graph 404"))
    return "Outlook hasn't synced this message to the cloud yet — wait a few seconds and retry.";
  if (m.includes("graph 429") || m.includes("graph 503") || m.includes("graph 504"))
    return "Microsoft is throttling requests right now — wait a moment and retry.";
  if (m.includes("save conflict") || m.includes("modified by someone else"))
    return "Someone else updated this project at the same moment. Retry — the add-in re-reads the latest version before saving.";
  if (m.includes("failed to fetch") || m.includes("networkerror") || m.includes("network error") || m.includes("load failed"))
    return "Network hiccup — check your connection (VPN?) and retry.";
  if (m.includes("quota"))
    return "Browser storage is full — the save went through, but offline caching is paused.";
  // Unknown failure: show the message but capped so a stray API body can't
  // flood the status line.
  return raw.length > 160 ? raw.slice(0, 160) + "…" : raw;
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