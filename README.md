# Reply with AI (Outlook add-in for Mac)

An Outlook add-in that reads the email thread you have open and writes your
reply, as you, with **Google Gemini (free)** or **Claude (paid)**. It shows up in
the toolbar as **Reply with ChatGPT**.

How it works:

1. Select an email and click **Reply with ChatGPT** in the toolbar. A side panel opens.
2. Optionally pick a tone, or type what the reply should say ("confirm Thursday 12pm").
3. Click **Write reply**. The reply appears in the panel in a few seconds.
4. Edit it if you like, then click **Create reply draft**. Outlook opens a normal
   reply with your signature and the quoted thread. Review it and send it yourself.

Good to know:

- It reads the **whole thread** (turn this off in Settings to send only the newest message).
- It writes as you (your Outlook name), answers what was asked of you, and puts
  `[placeholders]` where it needs a fact it doesn't have instead of making one up.
- If the newest message is one **you** sent, it writes a follow-up to everyone
  and switches to Reply All.
- It never sends email. It only opens a draft.
- ChatGPT still works as a manual option: **Use ChatGPT instead** copies a prompt
  for you to paste into ChatGPT, and **Paste answer** brings the answer back.

## Choose the AI (one-time setup)

Click the **gear icon** in the panel and pick **AI that writes the reply**.
Keys are saved only on that device (the panel's local storage), not in your
mailbox, so add them again if you also use Outlook on the web.

### Google Gemini (free) - the default

1. Go to <https://aistudio.google.com> and sign in with a Google account.
2. Click **Get API key → Create API key**, and copy it (it starts with `AIza`).
3. Paste it into **Gemini API key** in Settings and click **Save**.

- Free, with no card needed. The free tier has per-minute and daily request
  limits (shown in AI Studio); if you hit one, the panel says so. Wait, or switch
  to Gemini 3.5 Flash-Lite.
- **On the free tier, Google may use what you send to improve its products, and
  people may review it.** Don't use the free tier for confidential client email.
  Turning on billing in AI Studio switches to the paid tier, where it isn't used.

### Claude (paid)

1. Go to <https://console.anthropic.com>, sign up, and add credit under **Billing**.
2. Go to **API keys → Create key**, and copy it (it starts with `sk-ant-`).
3. In Settings pick **Claude**, paste the key into **Claude API key**, and click **Save**.

- About 2-4 US cents per reply with Claude Opus 5.5 (the default), less with
  Sonnet or Haiku. Long threads cost more.
- Anthropic doesn't train on API data by default.

**Privacy (both):** only when you click **Write reply**, the thread text, subject,
names and email addresses are sent over HTTPS to the AI provider you chose. Check
your company's policy on sending client emails to an AI service. The add-in has no
server of its own and no analytics.

## What's in this folder

| Path | What it is |
|---|---|
| `manifest.xml` | Tells Outlook about the add-in. Points at `https://localhost:3000` for testing. |
| `docs/` | The add-in itself: the files you host. (GitHub Pages needs the folder to be called `docs`.) |
| `docs/taskpane.html`, `.css`, `.js` | The side panel. |
| `docs/text-utils.js` | Prompt building (Claude and ChatGPT), quote trimming and text→HTML conversion. |
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

Click the gear icon in the panel. The API key stays on this device; everything
else is saved with your mailbox (Outlook's roaming settings).

- **AI that writes the reply**: Gemini (free) or Claude (paid), each with its own
  API key and model: Gemini 3.8 Flash or 3.5 Flash-Lite; Claude Opus 5.5,
  Sonnet 5.5 or Haiku 5.5.
- **About you and how you write**: optional notes Claude follows every time, e.g.
  your role, company, and style ("short and direct, UK spelling").
- **Read the full thread**: on by default. Off sends only the newest message
  (older quoted history like "On … wrote:" is cut).
- **Default tone** and **Default reply** (Reply or Reply All).
- **Max email length**: longer threads are shortened to this many characters
  (default 30000) and the panel tells you when that happens.
- **ChatGPT settings**: the prompt template for the ChatGPT option (must contain
  `{body}`), and **I use the ChatGPT Mac app** to skip opening the browser.

Tip: the panel can be **pinned** (pin icon at the top of the panel) so it stays
open while you move between emails. When you select another email, the reply box
is cleared so a reply can't go to the wrong person.

## Troubleshooting

**The button doesn't appear**
- Make sure you selected an email (it's for reading mail, not for drafts,
  meetings or appointments).
- In new Outlook for Mac, look under **Apps** in the toolbar or the **…** (more) menu.
- Quit Outlook completely (⌘Q) and reopen it. New add-ins can take a few minutes
  to appear.
- Check it's installed: open <https://aka.ms/olksideload> → **My add-ins** →
  **Custom Addins**. If it isn't there, add the manifest again.

**"Sorry, we can't load the add-in"**
- If you installed `manifest.xml` (local testing), `npm start` must be running.
  Switch to `manifest.prod.xml` so it loads from GitHub Pages instead.
- Hosted: check `https://YOUR-SITE/taskpane.html` opens in a browser.
- After changing files, quit and reopen Outlook so it doesn't use a cached copy.

**AI errors**
- "API key isn't valid": paste the key again in Settings (no spaces).
- Gemini "used up the free requests": wait a minute (or until the next day for the
  daily limit, which resets at midnight Pacific time), or switch model.
- Claude "out of credit": add credit at console.anthropic.com → Billing.
- "Couldn't reach …": check your internet connection; some company networks block
  generativelanguage.googleapis.com or api.anthropic.com.
- "… declined": rephrase the instruction, switch AI in Settings, or use ChatGPT.

**Copy or paste doesn't work (ChatGPT option)**
- If the prompt can't be copied automatically, click **Select all & copy** (or click
  in the prompt box and press ⌘A then ⌘C).
- **Paste answer** can be blocked by Outlook. Click in the reply box and press ⌘V.

**The reply draft looks wrong**
- Paragraphs come from blank lines in the reply; single line breaks are kept.
- `**bold**` becomes bold. Other Markdown (lists, headings) stays as plain text.

## Privacy

- **Gemini / Claude:** only when you click **Write reply**, the thread text, subject,
  names and addresses are sent over HTTPS to Google (`generativelanguage.googleapis.com`)
  or Anthropic (`api.anthropic.com`). Gemini's free tier may use it to improve Google's products.
- **ChatGPT option:** the prompt goes only to your clipboard, then wherever you paste it.
- No server of our own, no analytics. Settings live in your mailbox; the API key
  lives only on your device.
