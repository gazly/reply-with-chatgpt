# Reply with ChatGPT (Outlook add-in for Mac)

A small Outlook add-in that helps you write email replies with ChatGPT, using
**your own ChatGPT login** (the ChatGPT Mac app or chatgpt.com).

How it works:

1. Select an email and click **Reply with ChatGPT** in the toolbar. A side panel opens.
2. Pick a tone (and an optional instruction), then click **Copy prompt & open ChatGPT**.
3. Paste into ChatGPT (⌘V), adjust the answer until you like it, and copy it.
4. Back in the panel, paste it into the big box (⌘V) and click **Create reply draft**.
5. Outlook opens a normal reply with your signature and the quoted thread.
   Review it and send it yourself.

What it does **not** do:

- No OpenAI API, no API keys, no server. The add-in is a few static files
  (HTML, CSS, JavaScript) and everything runs in the panel on your Mac.
- It never reads or automates the ChatGPT website. You use ChatGPT yourself.
- It never sends email. It only opens a draft.
- No analytics or tracking. The only network requests are the add-in's own
  files and Microsoft's Office.js library.

## What's in this folder

| Path | What it is |
|---|---|
| `manifest.xml` | Tells Outlook about the add-in. Points at `https://localhost:3000` for testing. |
| `docs/` | The add-in itself: the files you host. (GitHub Pages needs the folder to be called `docs`.) |
| `docs/taskpane.html`, `.css`, `.js` | The side panel. |
| `docs/text-utils.js` | Prompt building, quote trimming and text→HTML conversion. |
| `scripts/serve.js` | Local HTTPS server for testing (`npm start`). |
| `scripts/make-manifest.js` | Makes `manifest.prod.xml` with your hosting URL (`npm run manifest`). |
| `scripts/make-icons.js` | Redraws the icons in `docs/assets/` (`npm run icons`). |
| `test/` | Unit tests for the text helpers (`npm test`). |
| `TESTING.md` | Manual test checklist. |

---

## 1. Install Node.js (Apple Silicon Mac)

You only need Node.js for local testing and the helper scripts. The add-in
itself doesn't use it.

1. Go to <https://nodejs.org> and download the **LTS** version for macOS.
   The installer (`.pkg`) works on Apple Silicon (M1/M2/M3/M4).
2. Open the downloaded file and click through the installer.
3. Open **Terminal** (⌘Space, type "Terminal", press Return) and check it worked:

   ```bash
   node -v
   ```

   You should see something like `v22.x.x`.

(If you use Homebrew you can instead run `brew install node`.)

## 2. Test it locally

In Terminal, go to this folder and install the two helper tools:

```bash
cd ~/reply-with-chatgpt
```

```bash
npm install
```

Install the local HTTPS certificate (one time only; macOS asks for your password
so it can trust the certificate):

```bash
npm run certs
```

Start the local server:

```bash
npm start
```

Leave that Terminal window open. Check that it works by opening
<https://localhost:3000/taskpane.html> in Safari or Chrome. You should see the
panel with the message "Open this panel from Outlook to use it", and **no
certificate warning**. (If you see a warning, see Troubleshooting below.)

Now install the add-in in Outlook using `manifest.xml` (see section 4). It loads
from your Mac while `npm start` is running.

Run the unit tests any time with:

```bash
npm test
```

## 3. Put it online (free static hosting)

Outlook needs the files on a public HTTPS address so the add-in works without
your Mac running a server. Pick **one** of these.

### Option A: GitHub Pages

1. Create a free account at <https://github.com> if you don't have one.
2. Click **+ → New repository**. Name it `reply-with-chatgpt`. Public is fine:
   the add-in contains no secrets and no email content. Click **Create repository**.
3. Upload the files. The easiest way without Git: on the new repository page click
   **uploading an existing file**, then drag in the whole contents of this
   folder **except** `node_modules`. Click **Commit changes**.
4. Go to **Settings → Pages**. Under **Build and deployment** choose
   **Deploy from a branch**, branch **main**, folder **/docs**, then **Save**.
5. Wait a minute. Your site is at
   `https://YOUR-USERNAME.github.io/reply-with-chatgpt/`.
   Check that `https://YOUR-USERNAME.github.io/reply-with-chatgpt/taskpane.html` opens.

### Option B: Cloudflare Pages

1. Create a free account at <https://dash.cloudflare.com>.
2. Go to **Workers & Pages → Create → Pages → Upload assets**.
3. Name the project (for example `reply-with-chatgpt`) and upload the **`docs`**
   folder only.
4. Your site is at `https://reply-with-chatgpt.pages.dev` (or similar).
   Check that `/taskpane.html` opens.

### Update the manifest with your URL

In Terminal, in this folder, run (with your own address, no trailing slash needed):

```bash
npm run manifest -- https://YOUR-USERNAME.github.io/reply-with-chatgpt
```

This writes **`manifest.prod.xml`**, a copy of `manifest.xml` with every
`https://localhost:3000` replaced by your address. Use this file in the next step.

When you change the add-in later, upload the changed `docs/` files again.
You only need to reinstall the manifest if you change the manifest itself.

## 4. Install the add-in for yourself

1. Open Outlook on the web in a browser and sign in with the same account you use
   in Outlook for Mac: <https://aka.ms/olksideload>
   (This opens Outlook with the **Add-ins for Outlook** window.)
2. Choose **My add-ins**.
3. Scroll to **Custom Addins** and click **Add a custom add-in → Add from File…**
4. Choose `manifest.prod.xml` (hosted version) or `manifest.xml` (local testing
   with `npm start`). Click **Install** if asked.
5. Open Outlook for Mac. Select an email: the **Reply with ChatGPT** button appears
   in the toolbar (in new Outlook it may be under **Apps** or the **…** menu;
   you can pin it from there). If it doesn't show up, quit Outlook (⌘Q) and open
   it again; it can take a few minutes to sync.

Only install one of the two manifests at a time: they share the same add-in ID.
To switch from local to hosted, remove the old one first (same **My add-ins**
page → **…** → **Remove**), then add the other.

## Settings

Click the gear icon in the panel. Settings are saved with your mailbox
(Outlook's roaming settings), so they follow you to Outlook on the web.

- **Prompt template**: the text sent to ChatGPT. Placeholders:
  `{tone} {instruction} {language} {senderName} {senderEmail} {subject} {date} {to} {cc} {body}`.
  It must contain `{body}`.
- **Default tone** and **Default reply** (Reply or Reply All).
- **Max email length**: long emails are cut to this many characters (default 6000).
  Old quoted history ("On … wrote:", "From: … Sent: …", "-----Original Message-----")
  is removed first.
- **I use the ChatGPT Mac app**: skips opening the browser; just copies the prompt so
  you can switch to the app and press ⌘V.

Tip: the panel can be **pinned** (pin icon at the top of the panel) so it stays
open while you move between emails. When you select another email, the reply box
is cleared so a reply can't go to the wrong person.

## Troubleshooting

**The button doesn't appear**
- Make sure you selected an email you *received* (it's for reading mail, not for
  drafts, meetings or appointments).
- In new Outlook for Mac, look under **Apps** in the toolbar or the **…** (more) menu.
- Quit Outlook completely (⌘Q) and reopen it. New add-ins can take a few minutes
  to appear.
- Check it's installed: open <https://aka.ms/olksideload> → **My add-ins** →
  **Custom Addins**. If it isn't there, add the manifest again.
- Your work or school account's admin may block custom add-ins. Try a personal
  Outlook.com account to confirm.

**The panel is blank or says it can't load**
- Local testing: is `npm start` still running in Terminal?
- Open the panel's address in Safari: `https://localhost:3000/taskpane.html`
  (or your hosted address). If Safari shows a certificate warning, run
  `npm run certs` again, then quit and reopen Outlook.
  To start fresh: `npx office-addin-dev-certs uninstall` then `npm run certs`.
- Hosted: check the URLs in `manifest.prod.xml` exactly match your site
  (open one of them in a browser). GitHub Pages addresses are case-sensitive.
- After changing files, quit and reopen Outlook so it doesn't use a cached copy.

**Copy or paste doesn't work**
- If the prompt can't be copied automatically, the panel shows it in a box: click
  **Select all & copy** (or click in it and press ⌘A then ⌘C).
- **Paste from clipboard** can be blocked by Outlook. Click in the reply box and
  press ⌘V instead (that always works).
- If ChatGPT's browser tab doesn't open, click **Show prompt → Open ChatGPT**, or
  open chatgpt.com yourself. If you use the Mac app, turn on
  **I use the ChatGPT Mac app** in Settings.

**The reply draft looks wrong**
- Paragraphs come from blank lines in ChatGPT's answer; single line breaks are kept.
- `**bold**` from ChatGPT's copy button becomes bold. Other Markdown (lists,
  headings) stays as plain text, which reads fine in email.

## Privacy

The email text goes only to your clipboard, and from there wherever *you* paste it.
The add-in has no server, sends nothing over the network, and stores only your
settings (in your own mailbox).
