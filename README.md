# Setty PMS — Outlook Add-in

## Files
- `manifest.xml` — Outlook add-in manifest (upload this to deploy)
- `taskpane.html` — task pane UI
- `taskpane.js` — all logic (auth, Supabase, Graph API)
- `commands.html` — required stub

## Setup Steps

### 1. Create GitHub Repo
1. Create a new repo at github.com: **smartias/setty-pms-addin** (must be exactly this name)
2. Upload all four files to the repo root
3. Go to Settings → Pages → Source: Deploy from branch → Branch: main → / (root)
4. Wait ~2 minutes for GitHub Pages to activate
5. Verify: visit `https://smartias.github.io/setty-pms-addin/taskpane.html` — you should see the sign-in screen

### 2. Create icon files
You need four icon PNG files in an `/assets/` folder in the repo:
- `icon-16.png` (16×16)
- `icon-32.png` (32×32)
- `icon-64.png` (64×64)
- `icon-80.png` (80×80)
- `icon-128.png` (128×128)

These can be simple SETTY logo PNGs. The add-in will work without them but Outlook may show a default icon.

### 3. Azure App Registration — Add Redirect URI
1. Go to: portal.azure.com → Azure Active Directory → App registrations
2. Find app with Client ID: **c4739c11-e89b-4a04-9580-f2d886356301**
3. Click Authentication → Add a platform → Single-page application
4. Redirect URI: `https://smartias.github.io/setty-pms-addin/taskpane.html`
5. Click Configure, then Save

### 4. Deploy the Manifest to Outlook
**Option A — Admin deployment (recommended, deploys to whole team):**
1. Go to Microsoft 365 Admin Center (admin.microsoft.com)
2. Settings → Integrated apps → Upload custom apps
3. Upload `manifest.xml`
4. Assign to: All users (or specific users/groups)
5. Takes ~24 hours to propagate to everyone

**Option B — Individual sideload (for testing):**
1. In Outlook desktop: File → Manage Add-ins → "My add-ins" → Add a custom add-in → Add from file
2. Select `manifest.xml`
3. Appears immediately in your Outlook only

### 5. Supabase setup for shared thread tagging (required for cross-user chain tags)
To allow project tags to follow an email conversation across different Outlook accounts, create a small mapping table in Supabase.

Run this SQL in the Supabase SQL editor:

```sql
create table if not exists public.pms_email_thread_tags (
  conversation_id text primary key,
  project_id text not null,
  tagged_by text,
  updated_at timestamptz not null default now(),
  created_at timestamptz not null default now()
);

create index if not exists idx_pms_email_thread_tags_project_id
  on public.pms_email_thread_tags (project_id);
```

If Row Level Security (RLS) is enabled for your project, make sure your existing add-in role can `select` and `insert`/`upsert` on this table (the add-in uses Supabase REST with the same anon key already configured in `taskpane.js`).

Example RLS policy setup (for current add-in architecture using the anon key from the client):

```sql
alter table public.pms_email_thread_tags enable row level security;

drop policy if exists pms_email_thread_tags_select_anon on public.pms_email_thread_tags;
create policy pms_email_thread_tags_select_anon
on public.pms_email_thread_tags
for select
to anon
using (true);

drop policy if exists pms_email_thread_tags_insert_anon on public.pms_email_thread_tags;
create policy pms_email_thread_tags_insert_anon
on public.pms_email_thread_tags
for insert
to anon
with check (true);

drop policy if exists pms_email_thread_tags_update_anon on public.pms_email_thread_tags;
create policy pms_email_thread_tags_update_anon
on public.pms_email_thread_tags
for update
to anon
using (true)
with check (true);
```

> Security note: this mirrors the current front-end-only model already used by this add-in. If you later move Supabase writes behind a server/edge function, tighten these policies accordingly.

---

## How It Works

Open any email in Outlook. Click **Setty PMS** in the ribbon.

The task pane opens on the right with:
1. The email subject and from/date shown at top
2. A project search box — type to search by name or project number
3. Once a project is selected:
   - **📁 Save to SharePoint + Project Record** — uploads email as HTML to `/Emails/YYYY_MM_DD Subject/` and saves to the project's Emails tab in PMS
   - **📝 Log as Note** — pre-fills subject as note body, pick category (Meeting, Site Visit, Client Communication, etc.), saves to Notes tab
   - **🔵 Log as RFI** — pre-fills title from subject and From as contractor, auto-assigns RFI number and 5-day due date
   - **📋 Log as Submittal** — same pattern, 10-day due date
   - **👤 Extract Contact** — parses the email signature for name, title, company, email, phone; lets you save to the Client Directory or project POC list

---

## Known Limitations
- The add-in reads emails from your mailbox only (not shared mailboxes) unless you have delegate access
- Contact extraction uses heuristic parsing — works well for standard email signatures but may need manual cleanup for unusual formats
- Admin deployment requires M365 Global Admin or Exchange Admin role
