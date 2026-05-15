# Setty PMS — Word Add-in (Print to OneNote)

Sidebar inside Word that exports the current document as a PDF printout into a project's OneNote notebook, and optionally saves the original `.docx` to a chosen subfolder of the project's SharePoint library.

## What it does

1. Loads the project list from the Setty PMS (Supabase).
2. Reads the current document's filename + first ~2000 chars of body text.
3. Suggests the most likely project (by filename / project number / client name match).
4. User confirms project, edits the OneNote page title, and clicks **Send to OneNote**.
5. PDF lands in the project's OneNote notebook → `Documents` section (auto-created).
6. (Optional) `.docx` lands in the chosen SharePoint subfolder, with `conflictBehavior=rename` so re-saves become `(1)`, `(2)`, etc.

## Files

| File | Purpose |
| --- | --- |
| `manifest.xml` | Word add-in manifest. Targets `Document` host, `ReadWriteDocument` permission. Source URL is `https://smartias.github.io/setty-pms-addin/word/taskpane.html`. |
| `taskpane.html` | Sidebar UI (project picker, OneNote title, SharePoint checkbox, folder picker). |
| `taskpane.js` | Auth (MSAL), project load, auto-suggest, PDF/.docx export, OneNote multipart POST, SharePoint PUT. |

The MSAL bundle lives one level up at `../msal-browser.min.js` (shared with the Outlook add-in).

## Hosting

Files must be hosted at the same GitHub Pages site as the Outlook add-in:

```
https://smartias.github.io/setty-pms-addin/word/taskpane.html
https://smartias.github.io/setty-pms-addin/word/taskpane.js
https://smartias.github.io/setty-pms-addin/word/manifest.xml
```

Push the `word/` directory to the `setty-pms-addin` repo and Pages will serve it.

## Sideloading for testing (single user)

1. **Word Desktop**: `Insert` → `Get Add-ins` → `My Add-ins` → `Upload My Add-in` → pick `manifest.xml`.
2. **Word for Web**: `Insert` → `Add-ins` → `Upload My Add-in`.

The button appears on the Home tab in a new "Setty PMS" group, labelled **Print to OneNote**.

## Tenant-wide deployment

Microsoft 365 Admin Center → `Settings` → `Integrated apps` → `Upload custom apps` → upload `manifest.xml`. Assign to the marketing/contracts/PM groups (or whoever needs it).

## Permissions requested

- `User.Read` — sign-in identity
- `Notes.ReadWrite` — create OneNote pages and sections (no admin consent)
- `Files.ReadWrite.All` — write the .docx to SharePoint when checkbox is ticked

No admin-consent scopes — same security posture as the Outlook add-in.

## Things to know

- **Project must have a OneNote notebook linked** in the PMS, otherwise the OneNote post fails with "no notebook linked". The sidebar surfaces that as an error.
- **SharePoint save requires `projectFolderUrl` on the project**. Same field the Outlook add-in uses — projects created via PMS already have it.
- **PDF rendering quality** is whatever Word's built-in PDF export produces. OneNote re-renders each PDF page as an image inside the page — same as the desktop "Send to OneNote" printer.
- **Auto-suggest threshold**: top match needs score ≥ 5 to auto-select. Tune in `applyAutoSuggest` if you want it more/less aggressive.
- **Score weights** live in `scoreProjectMatch` (taskpane.js). Adjust if the auto-suggest is wrong too often — most likely tuning lever is the `name.length > 4` guard (raise to 5 or 6 if short project names cause false positives).
