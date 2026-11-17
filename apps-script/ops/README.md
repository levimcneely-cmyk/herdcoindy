# Herd Co. Operations — backend for `/admin`

`herdcoindy.com/admin` is a static page. It holds **no data**. Everything lives in a private
Google Sheet owned by the Herd Co. Google account. This Apps Script checks the password on every call.

## One-time setup (~10 min)

1. Signed in as **levi@herdcoindy.com**, create a Google Sheet named **Herd Co. Operations**.
2. **Extensions → Apps Script**. Delete the starter code, paste in `Code.gs` from this folder, and save.
3. Select `setup` in the function dropdown and click **Run**. Approve the permissions (Sheets + Drive).
   This creates the `Projects`, `Log` and `Files` tabs, plus a private Drive folder:
   *Herd Co. Operations — Contracts & Files (private)*.
4. Reload the Sheet and pick **Herd Ops → Set password…**. Use 10+ characters. It's stored only as a salted hash.
5. In Apps Script, go to **Deploy → New deployment → Web app**. Set *Execute as*: **Me**, and *Who has access*: **Anyone**.
   Deploy, then copy the `/exec` URL.
6. In this repo, paste that URL into `assets/ops-config.js`:
   `window.HERD_OPS = { api: 'https://script.google.com/macros/s/…/exec' };`
   Commit and push. `/admin` now unlocks, and website consult requests show up as **Leads**.

## Later

- **Changing Code.gs:** go to Deploy → Manage deployments → ✏️ → Version: **New version**. This keeps the same URL.
- **Change password:** use ⋯ → Change password in `/admin`, or Herd Ops → Set password… in the Sheet. Both sign out every other device.
- **Lost phone / sign out everywhere:** Herd Ops → Sign out all devices.
- **10 wrong passwords** trigger a 15-minute lockout.
- Sessions last 12 hours, or 30 days if you tick "Keep me signed in on this device".

## Upgrading to v2 (automations, payments, contracts)

1. In the Sheet: **Extensions → Apps Script**, replace all of `Code.gs` with the new version, and save.
2. Run `setup` once. Google asks for new permissions (Gmail read, send email, Docs, external requests). Approve them.
   This adds the `Payments` and `Health` tabs and new Project columns. It also sorts existing projects into Client / Internal / Personal.
3. **Deploy → Manage deployments → ✏️ → Version: New version → Deploy.** The URL stays the same.
4. In /admin, open **⋯ → Settings & automations** and click **Turn on automations**.
   Then click **Create starter agreement**. This makes a complete, branded Client Project Agreement as a Google Doc in the contracts folder. Review it, edit it in Google Docs, and it becomes the template. You can paste a link to your own agreement Doc instead.

### What runs on its own
| When | What |
|---|---|
| Nightly, 3 AM | Checks every Live / Under-contract project's site, admin and backend URL. Emails you when something goes down or comes back up. Then logs new emails with each Client's address. |
| Mondays, 7 AM | Week-ahead digest: sites down, follow-ups, deliveries, renewals, clients gone quiet, balances due |

To skip monitoring for a project (for example, an app only your client's domain can open), set **Watch nightly = no** on its Architecture tab.

### Contract template placeholders
`{{provider}} {{client}} {{org}} {{contactName}} {{email}} {{phone}} {{service}} {{scope}} {{value}} {{deposit}} {{balance}} {{monthly}} {{deliveryDue}} {{date}} {{liveUrl}} {{domain}}`.
**Generate contract draft** (Contracts & files tab) copies the template into the private contracts folder and fills it in.
**Email to client** (on the draft) sends it from your account as a PDF with an editable message. It saves a "SENT" copy of exactly what went out, logs the email, and sets the next step to "Get the signed agreement back".
When the signed copy comes back, upload it as **Executed contract**. That fills in the contract date and offers to move the project to Under contract.

### Client handoff guide
**Spec / handoff → Client handoff** leaves out money, contacts and anything under an architecture heading that contains "Internal".

## What's where

| Tab / place | Holds |
|---|---|
| `Projects` | One row per client/project: stage, contact, money, dates, URLs, architecture notes |
| `Log` | Notes, calls, emails and automatic stage-change history |
| `Files` | Contracts and docs: kind, date, Drive link |
| `Payments` | Each payment: date, amount, method, invoice # |
| `Health` | Latest result for each monitored URL (rewritten each run) |
| Drive folder | The uploaded files themselves (private to your account) |

You can also edit the Sheet directly. Add columns if you like, and the app keeps them.
The stage keys are `lead`, `discussion`, `contract`, `live`, `future` and `closed`.
