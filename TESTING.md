# Test checklist

Run `npm test` first (unit tests for quote trimming, length limit, template and HTML conversion).
Then check each item in **new Outlook for Mac** and in **Outlook on the web**.

| # | Test | Expected |
|---|---|---|
| 1 | **Plain-text email**: open it, click *Reply with ChatGPT* | Sender and subject show; status says "Ready". |
| 2 | Click *Copy prompt & open ChatGPT* | "Prompt copied ✓"; chatgpt.com opens; ⌘V in ChatGPT pastes the prompt with the email body. |
| 3 | **HTML email** (newsletter, formatted mail) | Prompt contains readable plain text, no HTML tags. |
| 4 | **Long thread** (many "On … wrote:" / "From: … Sent:" blocks) | Prompt contains only the newest message. Very long single messages end with "[… email shortened]". |
| 5 | Paste a multi-paragraph answer, *Create reply draft* with **Reply** | Reply form opens to the sender only; paragraphs and line breaks kept; signature and quoted thread present; nothing sent. |
| 6 | Same with **Reply All** | Reply form opens with all original recipients. |
| 7 | **Sinhala / Tamil** email (e.g. "ආයුබෝවන්", "வணக்கம்") | Text shows correctly in the prompt and in the draft; Language "Same as the email" keeps the reply in that language. |
| 8 | Choose Language **Tamil** on an English email | Prompt says "Language: Tamil". |
| 9 | **Dark mode** (macOS dark appearance / Outlook dark theme) | Panel switches to dark colours; all text readable. |
| 10 | Settings: change tone, Reply All default, max length, Mac app option, Save; reopen the panel | Choices kept. With the Mac app option on, no browser opens and status says "paste it into ChatGPT (⌘V)". |
| 11 | Settings: remove `{body}` from the template and Save | Warning; not saved. *Reset template to default* restores it. |
| 12 | *Create reply draft* with an empty box | "Paste ChatGPT's reply into the box first." |
| 13 | *Paste from clipboard* | Pastes, or explains to press ⌘V if blocked. |
| 14 | Pin the panel, switch to another email | Sender/subject update; reply box clears. |
| 15 | Open the panel on a meeting invite or with nothing selected | Friendly message, no error. |
| 16 | Count clicks: select email → button → Copy prompt → (paste/copy in ChatGPT) → paste → Create reply draft | 3 clicks in the panel plus one paste each way. |
| 17 | Network: Safari Web Inspector (or browser dev tools on the web) → Network tab | Only the add-in's own files and `appsforoffice.microsoft.com/…/office.js`. |
