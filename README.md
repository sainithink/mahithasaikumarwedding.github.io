# Mahitha & Sai Kumar — #SaHiLoveStory

Wedding website for Sunday, 22 November 2026, Aria University, Milpitas, CA.
Static files, hosted on GitHub Pages. RSVPs are saved to a Google Sheet through a
small Google Apps Script. That's free, and there's no server to run.

| File | What it is |
|---|---|
| `index.html` | The wedding site. Date, times and venue are in `CONFIG` at the top of its script. |
| `admin.html` | Password-protected guest list: totals, filters, search, CSV, print, delete. |
| `config.js` | The one setting you must fill in: your Apps Script URL. |
| `apps-script/Code.gs` | The RSVP backend. Paste it into Google Apps Script (step 1). |

## 1. Create the RSVP sheet (about 5 minutes)

1. Go to [sheets.new](https://sheets.new) and name the sheet e.g. **SaHi RSVPs**.
2. **Extensions → Apps Script**. Delete what's there, paste all of `apps-script/Code.gs`, and save.
3. **Project Settings** (gear icon) → **Script Properties** → **Add script property**:
   - Property: `ADMIN_PASSWORD`
   - Value: the password you'll use on the guest-list page
4. **Deploy → New deployment**. Click the gear next to "Select type" and choose **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy**, then allow the permissions. Google warns the app is unverified because it's your own script. Choose *Advanced → Go to project*.
5. Copy the **Web app URL**. It ends in `/exec`.

The `RSVPs` tab appears in the sheet with the first reply. You can open the sheet at any time; it's the same list as `admin.html`.

## 2. Connect the site

Paste the URL into `config.js`:

```js
window.RSVP_ENDPOINT = "https://script.google.com/macros/s/AKfy…/exec";
```

## 3. Publish on GitHub Pages

1. Create a new **public** repository on GitHub, e.g. `sahi-wedding`. GitHub Pages on a free account needs a public repo.
2. Push this folder:
   ```bash
   git remote add origin https://github.com/<you>/sahi-wedding.git
   git push -u origin main
   ```
3. Go to the repo's **Settings → Pages → Build and deployment**: Source **Deploy from a branch**, branch **main**, folder **/ (root)**, then **Save**.
4. After about a minute the site is live at `https://<you>.github.io/sahi-wedding/`.
   The guest list is at `…/sahi-wedding/admin.html`.

A custom domain can be added in the same Pages settings.

## Good to know

- **The repo is public, and that's fine.** Replies live only in your Google Sheet. The password lives only in Script Properties, never in these files.
- **Guests replying again** with the same phone or email (or the same name if they gave neither) update their earlier reply instead of being counted twice.
- **Changing `Code.gs` later:** edit the script, then **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**. That keeps the same URL.
- **Wrong passwords:** after 10 wrong tries the guest list locks for 10 minutes.
