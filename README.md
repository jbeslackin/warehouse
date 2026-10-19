# Warehouse Commander: Ranked Play: setup for IT

A small web app for the Thorofare warehouse crew. Each employee signs in with their own Wyanoke Microsoft 365 account, claims and completes logistics tasks ("quests"), logs proof of completion, and earns XP on a monthly crew leaderboard. Jordon Slack is the admin.

- **Sign-in:** Microsoft Entra ID (MSAL.js, single tenant). No new passwords.
- **Data:** two JSON files, both read and written through Microsoft Graph as the signed-in user. Saves use each file's eTag, so two people saving at the same moment never overwrite each other.
  - **Warehouse board** (`warehouse-commander.json`, in SharePoint): what the whole crew sees. Its quests come from the Warehouse mailbox, plus anything an admin posts there.
  - **Admin's personal list** (`my-personal-quests.json`, in the admin's own OneDrive): quests from the admin's own inbox. Only that admin can see it; crew members never load it. The app creates this file automatically on the admin's first sign-in.
- **Hosting:** four static files, no server code, no database.
- **Inbox to game:** Claude (Jordon's approved Claude workspace, through its Microsoft 365 connector) acts as Jordon and needs no extra permissions beyond Jordon's own access. It turns Warehouse mailbox emails into quests on the shared board, and Jordon's own emails into quests on his private list. It also flags quests that look finished, such as when a UPS "delivered" email arrives, for a person to confirm.

Files in this package:

| File | What it is |
|---|---|
| `index.html` | The page |
| `app.js` | App logic |
| `config.js` | Settings to fill in |
| `warehouse-commander.json` | Warehouse board starting data (current warehouse quests, recent UPS packages, two paused example recurring missions; crew list starts empty) |
| `my-personal-quests.json` | Jordon's current personal quests. Optional: upload it to Jordon's OneDrive at `Warehouse Commander/my-personal-quests.json` before his first sign-in to keep them; otherwise the app starts him with an empty personal list |

Estimated time: about 30–45 minutes.

## 1. Put the data file in SharePoint

1. Pick (or create) a SharePoint site the warehouse crew can access, e.g. `https://wyanokegroup.sharepoint.com/sites/Warehouse`.
2. In its **Documents** library, create a folder named `Warehouse Commander` and upload `warehouse-commander.json` into it. Leave version history on (the default) so any bad change can be rolled back.
3. Permissions on that folder:
   - Warehouse crew (a group such as "Warehouse Crew"): **Edit**. People who can't edit it can't play.
   - Jordon Slack: **Owner** (or Full Control), so he can add and remove crew without a ticket.
4. Put the site's hostname, site path and file path into `config.js` (`SITE_HOSTNAME`, `SITE_PATH`, `DATA_FILE`).

## 2. Host the files so Jordon can update them

Recommended: **GitHub repository + Azure Static Web Apps (free tier)**.
1. Create a private GitHub repository (in the company's GitHub organization if there is one) containing the four app files, and give Jordon write access. The data file stays in SharePoint, not in the repository.
2. Azure portal → Create a resource → Static Web App → Plan: Free → Source: GitHub → pick that repository (app location `/`, no build step).
3. Every change committed to the repository redeploys the site automatically within a few minutes, so app updates don't need a ticket.
4. Note the site URL, e.g. `https://<name>.azurestaticapps.net/`.

Any other internal HTTPS static host also works, as long as Jordon can publish updated files.

## 3. Register the app in Entra ID

1. Entra admin center → **App registrations** → **New registration**.
   - Name: `Warehouse Commander: Ranked Play`
   - Supported account types: **Accounts in this organizational directory only**
   - Redirect URI: platform **Single-page application (SPA)**, value = the exact site URL from step 2 (include the trailing `/` if the URL has one).
2. **API permissions** → Add → Microsoft Graph → **Delegated**: `User.Read`, `Sites.ReadWrite.All`, `Files.ReadWrite` (the signed-in user's own OneDrive, used only for an admin's private list). Then **Grant admin consent for Wyanoke Group**, so employees don't see a consent prompt.
   - The app only ever acts as the signed-in user, so people can only reach files they can already open. To narrow it further, use `Sites.Selected` scoped to the one site.
3. Copy **Directory (tenant) ID** and **Application (client) ID** from the Overview page into `config.js`, and commit it.

## 4. Test

1. Open the site URL and sign in as Jordon. The admin form, plus the **Debrief** and **Recurring missions** views, should appear.
2. Click **Join the crew**, then claim and complete a quest.
3. Sign in as a second employee in another browser. They should see Jordon on the leaderboard, and their own changes should show up for Jordon within about 20 seconds.

**Warehouse TV:** open `<site URL>#tv` on the warehouse monitor, signed in as any crew member. It shows priority missions and the monthly leaderboard full-screen and refreshes itself.

Optional: add the site URL as a **Website** tab in the warehouse Teams channel.

## What's in the app

| Area | Features |
|---|---|
| Work tracking | Proof of completion (a tracking number or note; required on quests flagged "needs proof"), notes on each quest, activity history, a link back to the source email |
| Game layer | XP by tier, +25% for clearing on or before the due date, workday streaks, 8 badges, monthly seasons with a hall of fame, an all-time leaderboard |
| Warehouse ops | Recurring missions (every workday, weekly or monthly; they post themselves), checklists, squad quests (several people, XP split), TV mode |
| UPS Tracker | Exceptions, claims, deliveries and labels from the Warehouse mailbox, with timelines, acknowledge and one-click quests |
| Claude intel | "New from inbox" markers for quests Claude added, a "Confirm & clear" banner when Claude spots a delivery confirmation, a weekly or monthly debrief with a copy-ready summary |

## Admins

Admins are the emails in `ADMIN_EMAILS` in `config.js`. They see the warehouse board plus their own private list (Personal tab), can post to either and move quests between them, and can post, assign and delete quests, manage recurring missions, mark inbox quests reviewed, dismiss intel, remove players, and see the Debrief.

The admin tools are hidden from everyone else on the page. The real permission boundary is SharePoint: anyone with Edit access to the data file could change it directly. That's fine for an internal crew game; to lock it down further, give the crew a custom permission level without delete, or move the data to a backend service.

## Data format (version 2)

```json
{
  "version": 2,
  "crew": [{ "id": "<Entra object id>", "name": "…", "email": "…", "role": "…", "color": "#ff8a1f", "joinedAt": "…" }],
  "quests": [{
    "id": "q123", "title": "Ship H737 boxes to San Diego (2-day)", "giver": "Maura Conroy (Vindico)",
    "tier": "urgent | main | side | lore", "due": "2026-10-19 or null", "details": "…",
    "source": "…", "sourceLink": "<Outlook link or null>",
    "status": "open | done", "assignees": ["<id>"], "credit": ["<id>"], "doneAt": null,
    "squad": false, "requiresProof": true, "proof": { "tracking": "…", "note": "…", "by": "<id>", "at": "…" },
    "checklist": [{ "id": "c1", "text": "…", "done": false, "by": null }],
    "notes": [{ "id": "n1", "by": "<id>", "at": "…", "text": "…" }],
    "history": [{ "at": "…", "by": "<id or claude>", "action": "…", "detail": "…" }],
    "fromInbox": true, "reviewed": false,
    "intel": { "text": "UPS shows delivered Fri 10/16 9:45 AM", "link": "<Outlook link>", "at": "…" },
    "createdAt": "…", "createdBy": "<id | claude | recurring>", "recurringId": null
  }],
  "recurring": [{ "id": "r1", "title": "…", "tier": "side", "cadence": "daily | weekly | monthly", "weekday": 5, "day": 1,
                  "details": "…", "checklist": ["…"], "active": true, "lastSpawned": "2026-10-19" }]
}
```

XP, levels, streaks, badges and seasons are all computed from completed quests, so they can't drift out of sync. Finished quests older than 45 days are trimmed to the fields XP needs, which keeps the file small (well under 1 MB).

### How Claude updates the files

- Emails from the **Warehouse mailbox** become quests in `warehouse-commander.json` (SharePoint). Emails from **Jordon's own inbox** become quests in Jordon's `my-personal-quests.json` (OneDrive), which only he sees. Jordon can move any quest between the two with one click.
- Personal quests are private and don't count on the crew leaderboard; warehouse quests earn XP as usual.
- New quests get `fromInbox: true`, `reviewed: false`, `createdBy: "claude"`, a `sourceLink`, and a `history` entry. Claude skips anything already on the board (matched by job number or email thread).
- When it finds evidence that a quest is finished, Claude sets `intel` on that quest. It never marks a quest done by itself; a person confirms with **Confirm & clear**.
- Claude never edits `crew`, `credit`, `assignees`, `notes` or `proof`.

### UPS Tracker

The **UPS Tracker** view lists every UPS package the Warehouse mailbox has heard about, with tabs for Needs attention (open exceptions and claims), On the way, Delivered and All, plus search by tracking number, reference or destination. Each package shows its timeline and links to UPS tracking and the source email. Anyone can **Acknowledge** an exception; admins can turn one into a quest with one click. Open alerts show as a count in the header and on the tab.

Claude fills it from UPS emails in the Warehouse mailbox (`pkginfo@ups.com`, `auto-notify@ups.com`, `preferred.us@ups.com`): one entry per tracking number in `shipments`, each new email added as an event. Status follows the newest event: label created → in transit → exception / exception resolved → delivered; claim emails mark it "claim open". Packages are linked to a quest by reference or job number (`questId`), and a delivery on a linked quest also sets that quest's `intel` for Confirm & clear. Delivered packages drop off after 30 days.

```json
"shipments": [{ "tracking": "1Z…", "ref": "ITC 1", "shipTo": "…", "direction": "outbound | inbound", "service": "…",
  "status": "exception | claim | resolved | transit | label | delivered", "lastAt": "…", "lastText": "…",
  "events": [{ "at": "…", "type": "…", "text": "…", "link": "<Outlook link>" }],
  "questId": "<quest id or null>", "ack": false, "ackBy": null, "ackAt": null }]
```
