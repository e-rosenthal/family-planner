# Family Planner

The family planner, events and grocery list as an offline-capable PWA, backed by
a Google Sheet. Built for two phones that often have no signal.

- **App:** https://e-rosenthal.github.io/family-planner/
- **Data:** the planner Google Sheet
- **Server:** the Sheet's bound Apps Script project, deployed as a web app

## How it works

The page used to be served by Apps Script and talked to the server with
`google.script.run`. That only works inside an Apps Script page and only when
online. Now the page is static on GitHub Pages and talks to a JSON API:

```
POST https://script.google.com/macros/s/.../exec
{ "token": "...", "action": "sync", "ops": [...], "want": ["planner","groceries"] }
```

One `sync` call sends everything queued and returns fresh state, so a phone
that has been offline catches up in a single round trip.

Three things make that safe:

**Edits are small actions, not row writes.** Changing the Lead on an event
sends `{kind:"saveEvent", id, fields:{lead:"Sam"}}` and the server writes that
one cell. If the other phone changed the Notes on the same event meanwhile,
both survive.

**Events have stable IDs.** The Events tab previously had no ID, so edits were
addressed by sheet row. That is fine for a live page but wrong for an edit
queued on a phone: by the time it syncs, rows may have moved. Column I now
holds an ID that stays with the event. Existing rows were backfilled and the
column is filled in automatically for anything added later.

**Creates carry the ID the phone generated.** A retried "add" lands on the row
it already created instead of adding a duplicate.

Edits are also folded together while queued: setting the same cell twice
offline reaches the sheet once, with the later value.

### CORS

Apps Script web apps cannot answer a CORS preflight, so requests must stay
"simple": `text/plain`, no custom headers. That is why the token travels in the
request body rather than an `Authorization` header.

## Layout

```
docs/                     what GitHub Pages serves
  index.html              shell and styles
  app.js                  API client, offline queue, all views
  sw.js                   service worker, caches the app shell
  manifest.webmanifest
  icons/
apps-script/              pushed with clasp, not served
  Code.js                 original server logic, onEdit, the old page
  Api.js                  doPost, token check, event IDs
  appsscript.json
tools/make-icons.py       regenerates the icon set
```

Two files are deliberately gitignored and must not be committed:

- `apps-script/.clasp.json` — the script ID
- `apps-script/Config.js` — the household's names, field list and labels

Both are pushed to Apps Script by clasp, which is the system of record for
them; `clasp pull` brings them back on a new machine. Copy the matching
`.example` files to get started. The repo holds the engine, the script project
holds the specifics, so no names appear here.

The client takes its field order, labels and people from the server with each
sync, so adding a column or renaming a person means editing `Config.js` and the
sheet, never this repo.

## The access token

The script URL and token are **never** in this repo. They are typed into each
phone once and kept in that phone's local storage.

The token lives in Script Properties. To create or read it, open the Apps
Script editor and run `setupApiToken()`, then check the execution log.
`rotateApiToken()` issues a new one; every phone then has to be re-paired.

The deployment is open to "Anyone" because the PWA cannot sign in to Google, so
the token is the only thing protecting the data. The old `doGet` page is gated
on the same token and needs `?t=YOUR_TOKEN` to open.

## Deploying a change

To the app:

```bash
git push                 # GitHub Pages republishes within a minute or so
```

Bump `CACHE` in `docs/sw.js` when the shell changes, otherwise phones keep
serving the old copy from the service worker cache.

To the server:

```bash
cd apps-script && clasp push
```

Then in the Apps Script editor: **Deploy > Manage deployments**, edit the
existing deployment, set Version to **New version**, and Deploy. Editing the
existing deployment keeps the URL, so the phones do not need re-pairing.
Creating a *new* deployment would issue a new URL and break both phones.

## Setting up a phone

1. Open the app URL in Safari (iPhone) or Chrome (Android).
2. Paste the web app URL and the access code, then tap Connect.
3. Share > **Add to Home Screen**.

Open it from the home screen icon rather than the browser: that is what gives
it the standalone window and its own storage.
