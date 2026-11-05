# Sitter Hours

A tiny clock-in app for babysitting. It goes on an iPhone home screen like a real app and saves everything to a Google Sheet.

It has two views, and each person gets their own link:

| | Sitter's link | Parent's link |
|---|---|---|
| Clock in / out | ✅ | (sees a live timer) |
| Add or edit shifts | ✅ | view only |
| See hours, owed, paid | ✅ | ✅ |
| Log a payment | ✅ | ✅ |
| Delete a payment | any | only ones the parent added |

```
iPhone home screen icon ──► index.html (GitHub Pages) ──► Apps Script web app ──► Google Sheet
```

Each link has a secret key in it. The script checks the key on every request and decides what that person is allowed to do. Changing the page can't get around this.

---

## Setup (about 20 minutes, done once)

### 1. Spreadsheet

1. **Sitter:** create a blank sheet at [sheets.new](https://sheets.new) and share it with you as an **Editor**.
2. **You:** open it and go to **Extensions → Apps Script**. Replace `Code.gs` with [`apps-script/Code.gs`](apps-script/Code.gs) and click **Save**.
3. Choose **`setup`** in the function dropdown and click **▶ Run**. When Google asks for permission, click **Review permissions**, then **Advanced → Go to … (unsafe)**, then **Allow**. The warning is expected because this is your own script.
4. On the **Summary** tab, fill in the yellow cells: **hourly rate** (starts at $25) and **sitter's first name**.

### 2. Publish the script

**Deploy → New deployment →** ⚙️ **Web app**
- **Execute as:** Me
- **Who has access:** Anyone

Click **Deploy** and copy the **Web app URL** (it ends in `/exec`).

### 3. Put the app page online

Push this folder to a **public** GitHub repo, then go to **Settings → Pages**, choose **Deploy from a branch**, and set it to `main` and `/ (root)`. The app will be at `https://YOUR-USERNAME.github.io/REPO/`.

The page contains no data or keys, so it's safe for the repo to be public.

### 4. Get the two links

1. Reload the spreadsheet. A **Sitter Hours** menu appears.
2. Click **Sitter Hours → Get app links**.
3. Paste in the app address from step 3 and the Web app URL from step 2, then click **Save & make links**.
4. Copy the **sitter's link** and the **parent's link**.

### 5. Install on each phone

Text each person their own link. They open it in **Safari** and tap **Share → Add to Home Screen → Add**.

> Don't send the sitter's link to the parent, because it has full access.

---

## Everyday use

**Sitter:**
- Tap **Clock In** when she arrives and **Clock Out** when she leaves.
- Tap a shift to fix its times. Use **+ Add a shift I forgot** if she never clocked in.
- **Got paid?** records a payment.
- When the parent logs a payment, she sees a notice the next time the app refreshes.

**Parent:**
- Sees whether the sitter is on the clock, with a live timer.
- Sees what they owe, hours this week, and the full shift history.
- **I paid** logs a payment.

**Changing the rate:** edit the yellow cell on the Summary tab. Only new shifts use the new rate.

## Turning a link off

Use **Sitter Hours → Turn off the parent link** (or the sitter's). The old link stops working right away, and the menu shows a new one to send.

## Updating the script later

Paste in the new code, then **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**. The URL stays the same.

Don't create a new deployment, because that changes the URL and breaks the links.

## Troubleshooting

- **"This link doesn't work anymore"**: the key was changed. Send a fresh link from the menu.
- **"Couldn't reach the spreadsheet"**: check that *Who has access* is set to **Anyone**.
- **Times are off by a few hours**: in the spreadsheet, go to **File → Settings → Time zone**.
- **Sitter Hours menu is missing**: reload the sheet. If it's still missing, run `setup` again (it won't erase anything).

## Try it locally

```bash
python3 -m http.server 8765
```

Open http://localhost:8765 and try the demo as the sitter or the parent. Demo data is fake and stays in that browser.

## Files

```
index.html             The whole phone app (both views, no build step)
manifest.webmanifest   Home screen name and icon
icons/                 Home screen icons
apps-script/Code.gs    Google Apps Script backend (paste into the Sheet)
```
