# Job Tracker
A clean, lightweight web app for tracking job applications. Built for people who want something more pleasant than a spreadsheet, but who still want their data in a Google Sheet they fully own and control.

https://bishwas1113.github.io/job-tracker/

---

## ✨ Features

- **Clean, calm interface** — Japandi-inspired design with 5 themes (Day, Night, Sand, Sage, Linen)
- **Click-to-edit** any field — title, company, status, dates, fit rating
- **One-click status changes** via colored pills
- **Smart sorting** — by date, fit, company, title, or pipeline status
- **Works on phone and laptop** — fully responsive, add to home screen for app-like experience
- **Your data, your Google Sheet** — sync to a sheet you own, edit directly anytime
- **CSV import / export** — bring in old data or take it elsewhere
- **Free forever** — no accounts, no subscriptions, no third-party servers

---

## 📋 Setup (5–10 minutes, one time)

You'll need a Google account. That's the only prerequisite.

### Step 1 — Create your Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com)
2. Click **Blank spreadsheet**
3. Rename it something like **"Job Tracker Data"** (top-left, click the title)

That's it for now — leave the sheet empty. The app will set up its structure automatically.

### Step 2 — Add the backend script

The "script" is a small piece of code that connects the app to your sheet. Don't worry — you just paste it in, you don't have to write anything.

1. In your sheet, click **Extensions** in the top menu → **Apps Script**
2. A new tab opens with a code editor. Delete the placeholder code (`function myFunction() {}`) so the editor is empty.
3. Copy the entire script below and paste it in:

```javascript
// === Job Tracker — Backend Sync Script ===

const SHEET_NAME = "Jobs";
const HEADERS = ["id","title","company","link","code","applied","deadline","username","notes","fit","status","archived","archivedDate","updatedAt"];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#f5f5f7");
    sheet.setFrozenRows(1);
  }
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (firstRow.length < HEADERS.length || firstRow.join(",") !== HEADERS.join(",")) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sheet;
}

function sanitize(v) {
  if (v === undefined || v === null) return "";
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v;
  return String(v).replace(/[\r\n\t]+/g, " ").trim();
}

function readAll() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return values
    .filter(row => row[0])
    .map(row => {
      const obj = {};
      HEADERS.forEach((h, i) => { obj[h] = row[i]; });
      obj.fit = obj.fit === "" ? 0 : Number(obj.fit) || 0;
      obj.archived = obj.archived === true || obj.archived === "TRUE" || obj.archived === "true";
      ["applied","deadline","archivedDate"].forEach(k => {
        if (obj[k] instanceof Date) {
          obj[k] = Utilities.formatDate(obj[k], Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          obj[k] = obj[k] ? String(obj[k]) : "";
        }
      });
      obj.updatedAt = obj.updatedAt ? String(obj.updatedAt) : "";
      return obj;
    });
}

function writeAll(records) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
  }
  if (!records.length) return;
  const rows = records.map(r => HEADERS.map(h => sanitize(r[h])));
  sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
}

function isRecentlyCreated(rec) {
  if (!rec || !rec.id) return false;
  const m = String(rec.id).match(/^rec_(\d+)_/);
  if (!m) return false;
  const created = parseInt(m[1], 10);
  return (Date.now() - created) < 5 * 60 * 1000;
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || "sync";

    if (action === "sync") {
      const serverRecs = readAll();
      const clientRecs = body.records || [];

      const serverById = {};
      serverRecs.forEach(r => serverById[r.id] = r);
      const clientById = {};
      clientRecs.forEach(r => clientById[r.id] = r);

      const knownClientIds = new Set(body.knownIds || clientRecs.map(r => r.id));

      const merged = {};

      clientRecs.forEach(c => {
        const s = serverById[c.id];
        if (!s) {
          merged[c.id] = c;
        } else {
          const cTime = c.updatedAt || "";
          const sTime = s.updatedAt || "";
          if (sTime && sTime > cTime) {
            merged[c.id] = s;
          } else {
            merged[c.id] = c;
          }
        }
      });

      serverRecs.forEach(s => {
        if (!clientById[s.id] && !knownClientIds.has(s.id)) {
          merged[s.id] = s;
        }
      });

      Object.keys(merged).forEach(id => {
        const inClient = clientById[id];
        const inServer = serverById[id];
        if (inClient && !inServer && knownClientIds.has(id)) {
          const rec = clientById[id];
          if (!isRecentlyCreated(rec)) {
            delete merged[id];
          }
        }
      });

      const finalRecs = Object.values(merged);
      writeAll(finalRecs);

      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, records: finalRecs, total: finalRecs.length }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "read") {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, records: readAll() }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: "Unknown action" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, records: readAll() }))
    .setMimeType(ContentService.MimeType.JSON);
}
```

4. Click the **💾 Save** icon (or `Cmd+S` / `Ctrl+S`). When prompted, name the project **"Job Tracker Backend"**.

### Step 3 — Deploy the script as a web app

This is the only slightly fiddly step. Follow carefully.

1. In the Apps Script editor, click the **Deploy** button in the top-right (blue) → **New deployment**
2. Click the **⚙️ gear icon** next to "Select type" → choose **Web app**
3. Fill in:
   - **Description:** `Job Tracker API` (optional)
   - **Execute as:** **Me** (your Google account)
   - **Who has access:** **Anyone** ⚠️ — this is required for the app to talk to the script. Don't worry: the URL contains a long random string, so nobody can find it without you sharing it. Your data stays private as long as you don't share the URL.
4. Click **Deploy**
5. Google will ask for permissions. Click **Authorize access** → choose your Google account → if you see a "Google hasn't verified this app" warning, click **Advanced** → **Go to Job Tracker Backend (unsafe)**. (It says "unsafe" because *you* are the developer, not because anything is actually unsafe. You wrote this. Sort of.) → Click **Allow**.
6. You'll see a **Web app URL** ending in `/exec`. **Copy this URL** — you'll need it in the next step.

> 💡 **Save this URL somewhere safe** (notes app, password manager). You'll need it again if you set up the app on another device.

### Step 4 — Open the app and connect it

1. Open the **live app URL** at the top of this README in your browser
2. Click **⚙️ Settings** in the top-right
3. Paste your **Web app URL** into the "Apps Script Web App URL" field
4. Click **Save & Connect**
5. The status dot at the top-right should turn green and say "Connected"

🎉 **You're done.** Add a new job application using the **+ New Entry** tab, and within a couple of seconds it'll show up in your Google Sheet too.

---

## 📱 Add to your phone home screen (recommended)

To use the app like a native phone app:

**iPhone (Safari):**
1. Open the app URL in Safari
2. Tap the **Share** button (square with arrow at the bottom)
3. Scroll down → tap **Add to Home Screen**
4. Name it "Job Tracker" → Add
5. Open from your home screen → ⚙️ Settings → paste your Apps Script URL again (each device needs this set once)

**Android (Chrome):**
1. Open the app URL in Chrome
2. Tap the three-dot menu → **Add to Home screen**
3. Confirm
4. Open from your home screen → ⚙️ Settings → paste your Apps Script URL

---

## 🖥 Add to your Mac as an app (optional)

**Using Chrome:**
1. Open the app URL in Chrome
2. Three-dot menu → **Cast, save, and share** → **Install page as app**
3. Name it "Job Tracker" → Install
4. You'll find Job Tracker.app in Applications, can be added to Dock and launched via Spotlight

---

## 🎨 Themes

In Settings, pick from 5 Japandi-inspired themes:
- **Shoji** — warm off-white, soft indigo (default day mode)
- **Sumi** — deep charcoal, moss green (night mode)
- **Suna** — sandy beige with terracotta accent
- **Wabi** — pale sage with forest accent
- **Asa** — cool linen with slate-blue accent

---

## 📊 How the sync works

- The app saves locally first (so it's instant)
- Any change is pushed to your Google Sheet within ~1.5 seconds
- You can edit the sheet directly anytime — changes flow back to the app
- If you delete a row in the sheet, it disappears from the app on next sync
- Each device keeps its own local copy synced through the sheet

---

## 🛠 Troubleshooting

**"Sync failed" error**
- Make sure your Apps Script Web App URL is correct (it should end in `/exec`)
- Check that you set "Who has access" to **Anyone** during deployment
- If you edited the script after deploying, you need to **Deploy → Manage deployments → ✏️ edit → New version → Deploy**

**Phantom rows appearing in the sheet**
- Edit the row directly in the sheet to fix it
- Use the **Sync Now** button in the app to refresh

**App opens but nothing happens when I click things**
- Make sure you're opening the app in a real browser (Safari, Chrome, etc.), not as a file preview
- On iPhone, don't try to open the HTML directly from Files — use the hosted URL

**I made changes to the sheet but they're not in the app**
- Click **Sync Now** in Settings to pull the latest from the sheet

---

## 🔒 Privacy

- The app's HTML code is public (anyone can view it on GitHub), but your data is not
- Your data lives only in:
  - Your Google Sheet (private to you and anyone you share it with)
  - Your browser's localStorage on each device (local to that device)
- The Apps Script Web App URL is the "key" to your data. Treat it like a password — don't share it, don't post it publicly.

---

## 🤝 Sharing with friends

If you want a friend or family member to use this:

1. Send them the live app URL (the GitHub Pages link at the top)
2. Tell them to follow this README's setup steps — they'll create **their own** Google Sheet and **their own** Apps Script URL
3. Their data stays in their sheet, yours in yours. No data is shared between users.

---

## 📝 License

MIT — feel free to fork, modify, and use however you'd like.
