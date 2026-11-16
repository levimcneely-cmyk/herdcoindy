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

## What's where

| Tab / place | Holds |
|---|---|
| `Projects` | One row per client/project: stage, contact, money, dates, URLs, architecture notes |
| `Log` | Notes, calls, emails and automatic stage-change history |
| `Files` | Contracts and docs: kind, date, Drive link |
| Drive folder | The uploaded files themselves (private to your account) |

You can also edit the Sheet directly. Add columns if you like, and the app keeps them.
The stage keys are `lead`, `discussion`, `contract`, `live`, `future` and `closed`.
