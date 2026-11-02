/*
 * Reply with AI: task pane logic.
 * Gemini or Claude writes the reply straight into the panel (the email thread is
 * sent to that provider's API with your own key). ChatGPT stays available as a
 * manual copy/paste option. It never sends email: it only opens a reply draft.
 */
(function () {
  "use strict";

  var T = window.RWCText;
  var CHATGPT_URL = "https://chatgpt.com/";
  var CLAUDE_URL = "https://api.anthropic.com/v1/messages";
  var GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/";
  var SETTINGS_KEY = "rwcSettings";
  // API keys are kept in this device's localStorage, not in the mailbox.
  var KEY_STORES = { claude: "rwcClaudeKey", gemini: "rwcGeminiKey" };
  var PROVIDER_NAMES = { claude: "Claude", gemini: "Gemini" };
  var MODELS = {
    "claude-opus-5-5": "Claude Opus 5.5 (best writing)",
    "claude-sonnet-5-5": "Claude Sonnet 5.5 (faster, cheaper)",
    "claude-haiku-5-5": "Claude Haiku 5.5 (fastest, cheapest)",
  };
  var GEMINI_MODELS = {
    "gemini-3.8-flash": "Gemini 3.8 Flash (best free)",
    "gemini-3.5-flash-lite": "Gemini 3.5 Flash-Lite (fastest)",
  };
  var DEFAULTS = {
    template: T.DEFAULT_TEMPLATE,
    tone: "Professional",
    replyMode: "reply",
    maxBody: 30000,
    useMacApp: false,
    provider: "gemini",
    model: "claude-opus-5-5",
    geminiModel: "gemini-3.8-flash",
    styleNotes: "",
    fullThread: true,
  };

  var settings = Object.assign({}, DEFAULTS);
  var email = null; // the email as read when the panel opened or the item changed
  var emailState = "loading"; // loading | ready | none | error
  var loadToken = 0;
  var writing = null; // AbortController while the AI is writing
  var els = {};

  // ---------- small helpers ----------

  function $(id) {
    return document.getElementById(id);
  }

  function setStatus(message, kind) {
    var el = els.status || $("status");
    if (!el) return;
    el.textContent = message || "";
    el.className = "status" + (kind ? " " + kind : "");
  }

  // Wraps event handlers so a bug shows a friendly message, never an unhandled error.
  function safe(fn) {
    return function (event) {
      try {
        return fn.call(this, event);
      } catch (err) {
        console.error(err);
        setStatus("Something went wrong: " + (err && err.message ? err.message : err), "err");
      }
    };
  }

  function mailbox() {
    return window.Office && Office.context && Office.context.mailbox;
  }

  function currentItem() {
    var mb = mailbox();
    return mb ? mb.item : null;
  }

  function formatPerson(p) {
    if (!p) return "";
    if (p.displayName && p.emailAddress && p.displayName !== p.emailAddress) {
      return p.displayName + " <" + p.emailAddress + ">";
    }
    return p.emailAddress || p.displayName || "";
  }

  function replyMode() {
    var checked = document.querySelector('input[name="replyMode"]:checked');
    return checked ? checked.value : "reply";
  }

  function setReplyMode(mode) {
    var radio = document.querySelector('input[name="replyMode"][value="' + (mode === "replyAll" ? "replyAll" : "reply") + '"]');
    if (radio) radio.checked = true;
  }

  function fillSelect(select, options, value, fallback) {
    select.innerHTML = "";
    Object.keys(options).forEach(function (key) {
      var opt = document.createElement("option");
      opt.value = key;
      opt.textContent = options[key];
      select.appendChild(opt);
    });
    select.value = options[value] ? value : fallback;
  }

  function fillToneSelect(select, value) {
    var names = {};
    Object.keys(T.TONES).forEach(function (name) {
      names[name] = name;
    });
    fillSelect(select, names, value, DEFAULTS.tone);
  }

  function provider() {
    return PROVIDER_NAMES[settings.provider] ? settings.provider : DEFAULTS.provider;
  }

  function readApiKey(which) {
    try {
      return window.localStorage.getItem(KEY_STORES[which || provider()]) || "";
    } catch (err) {
      return "";
    }
  }

  function writeApiKey(which, key) {
    try {
      if (key) window.localStorage.setItem(KEY_STORES[which], key);
      else window.localStorage.removeItem(KEY_STORES[which]);
      return true;
    } catch (err) {
      return false;
    }
  }

  function showProviderSettings() {
    var p = els.setProvider.value;
    els.geminiSettings.hidden = p !== "gemini";
    els.claudeSettings.hidden = p !== "claude";
  }

  function updateWriteLabel() {
    els.writeBtn.textContent = "Write reply with " + PROVIDER_NAMES[provider()];
  }

  // ---------- settings (roamingSettings + this device's API key) ----------

  function loadSettings() {
    try {
      var saved = Office.context.roamingSettings.get(SETTINGS_KEY);
      if (saved && typeof saved === "object") settings = Object.assign({}, DEFAULTS, saved);
    } catch (err) {
      console.warn("Could not read settings", err);
    }
  }

  function saveSettings(next, done) {
    try {
      var rs = Office.context.roamingSettings;
      rs.set(SETTINGS_KEY, next);
      rs.saveAsync(function (result) {
        var ok = result.status === Office.AsyncResultStatus.Succeeded;
        if (ok) settings = Object.assign({}, DEFAULTS, next);
        done(ok, ok ? null : result.error && result.error.message);
      });
    } catch (err) {
      done(false, err.message);
    }
  }

  function applyDefaultsToForm() {
    fillToneSelect(els.tone, settings.tone);
    setReplyMode(settings.replyMode);
    updateWriteLabel();
  }

  function openSettings() {
    els.setProvider.value = provider();
    els.setGeminiKey.value = readApiKey("gemini");
    fillSelect(els.setGeminiModel, GEMINI_MODELS, settings.geminiModel, DEFAULTS.geminiModel);
    els.setApiKey.value = readApiKey("claude");
    fillSelect(els.setModel, MODELS, settings.model, DEFAULTS.model);
    showProviderSettings();
    els.setStyleNotes.value = settings.styleNotes || "";
    els.setFullThread.checked = !!settings.fullThread;
    els.setTemplate.value = settings.template;
    fillToneSelect(els.setTone, settings.tone);
    els.setReplyMode.value = settings.replyMode === "replyAll" ? "replyAll" : "reply";
    els.setMaxBody.value = settings.maxBody;
    els.setMacApp.checked = !!settings.useMacApp;
    els.mainView.hidden = true;
    els.settingsView.hidden = false;
    els.settingsBtn.setAttribute("aria-expanded", "true");
    setStatus("");
  }

  function closeSettings() {
    els.settingsView.hidden = true;
    els.mainView.hidden = false;
    els.settingsBtn.setAttribute("aria-expanded", "false");
  }

  function onSaveSettings() {
    var max = parseInt(els.setMaxBody.value, 10);
    if (!(max >= 500 && max <= 200000)) {
      setStatus("Max email length must be between 500 and 200000.", "warn");
      return;
    }
    var template = els.setTemplate.value.trim() ? els.setTemplate.value : T.DEFAULT_TEMPLATE;
    if (template.indexOf("{body}") === -1) {
      setStatus("The ChatGPT template needs {body} so ChatGPT sees the email.", "warn");
      return;
    }
    var claudeKey = els.setApiKey.value.trim();
    var geminiKey = els.setGeminiKey.value.trim();
    if (claudeKey && claudeKey.indexOf("sk-ant-") !== 0) {
      setStatus("That doesn't look like a Claude API key (they start with sk-ant-).", "warn");
      return;
    }
    if (geminiKey && /\s/.test(geminiKey)) {
      setStatus("The Gemini API key has a space in it. Paste it again.", "warn");
      return;
    }
    if (!writeApiKey("claude", claudeKey) || !writeApiKey("gemini", geminiKey)) {
      setStatus("Couldn't save the API key on this device.", "err");
      return;
    }
    var next = {
      provider: els.setProvider.value,
      geminiModel: els.setGeminiModel.value,
      template: template,
      tone: els.setTone.value,
      replyMode: els.setReplyMode.value,
      maxBody: max,
      useMacApp: els.setMacApp.checked,
      model: els.setModel.value,
      styleNotes: els.setStyleNotes.value.trim(),
      fullThread: els.setFullThread.checked,
    };
    els.saveSettingsBtn.disabled = true;
    saveSettings(next, function (ok, message) {
      els.saveSettingsBtn.disabled = false;
      if (ok) {
        applyDefaultsToForm();
        if (email && email.fromMe) setReplyMode("replyAll");
        closeSettings();
        setStatus("Settings saved ✓", "ok");
      } else {
        setStatus("Could not save settings" + (message ? ": " + message : "."), "err");
      }
    });
  }

  // ---------- reading the email ----------

  function loadEmail() {
    var token = ++loadToken;
    email = null;
    var item = currentItem();

    if (!item) {
      emailState = "none";
      els.emailFrom.textContent = "-";
      els.emailSubject.textContent = "-";
      setStatus("No email selected. Select a single email to reply to.", "warn");
      return;
    }
    if (item.itemType && item.itemType !== Office.MailboxEnums.ItemType.Message) {
      emailState = "none";
      setStatus("This works with emails only, not meetings or appointments.", "warn");
      return;
    }

    var from = item.from || item.sender || {};
    var me = mailbox().userProfile || {};
    var fromMe = !!(from.emailAddress && me.emailAddress &&
      from.emailAddress.toLowerCase() === me.emailAddress.toLowerCase());
    els.emailFrom.textContent = (formatPerson(from) || "(unknown sender)") + (fromMe ? " (you)" : "");
    els.emailFrom.title = formatPerson(from);
    els.emailSubject.textContent = item.subject || "(no subject)";
    els.emailSubject.title = item.subject || "";
    // Replying to your own message only makes sense as a follow-up to everyone on it.
    setReplyMode(fromMe ? "replyAll" : settings.replyMode);
    emailState = "loading";
    setStatus("Reading the email…", "info");

    try {
      item.body.getAsync(Office.CoercionType.Text, function (result) {
        if (token !== loadToken) return; // a newer email was selected meanwhile
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          emailState = "error";
          setStatus("Couldn't read this email's text" + (result.error ? " (" + result.error.message + ")" : "") + ".", "err");
          return;
        }
        email = {
          senderName: from.displayName || from.emailAddress || "",
          senderEmail: from.emailAddress || "",
          senderLine: formatPerson(from),
          fromMe: fromMe,
          myName: me.displayName || "",
          myEmail: me.emailAddress || "",
          to: (item.to || []).map(formatPerson).join(", "),
          cc: (item.cc || []).map(formatPerson).join(", "),
          subject: item.subject || "",
          date: item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString() : "",
          rawBody: result.value || "",
        };
        emailState = "ready";
        if (!email.rawBody.trim()) setStatus("This email has no text body, but you can still write a reply.", "warn");
        else if (readApiKey()) setStatus("Ready. Click Write reply.", "info");
        else setStatus("Add your " + PROVIDER_NAMES[provider()] + " API key in Settings (gear icon) to start.", "info");
      });
    } catch (err) {
      emailState = "error";
      setStatus("Couldn't read this email: " + err.message, "err");
    }
  }

  function onItemChanged() {
    // Pinned panel: a different email was selected. Clear the old reply so it can't go to the wrong person.
    if (writing) writing.abort();
    els.reply.value = "";
    els.promptText.value = "";
    loadEmail();
  }

  function requireEmail() {
    if (emailState === "loading") {
      setStatus("Still reading the email, try again in a moment.", "info");
      return false;
    }
    if (emailState !== "ready" || !email) {
      setStatus("No readable email selected. Select an email first.", "warn");
      return false;
    }
    return true;
  }

  function promptContext() {
    var cleaned = T.cleanEmailBody(email.rawBody, settings.maxBody, settings.fullThread);
    return {
      cleaned: cleaned,
      tone: T.TONES[els.tone.value] || els.tone.value,
      language: els.language.value || "the same language as the email",
      instruction: els.instruction.value.trim(),
    };
  }

  // ---------- AI providers (Claude, Gemini) ----------

  function errorDetail(bodyText) {
    try {
      var parsed = JSON.parse(bodyText);
      return (parsed.error && parsed.error.message) || "";
    } catch (err) {
      return "";
    }
  }

  function claudeErrorMessage(status, bodyText) {
    var detail = errorDetail(bodyText);
    if (status === 401) return "Your Claude API key isn't valid. Check it in Settings.";
    if (status === 403) return "This API key isn't allowed to use that model. " + detail;
    if (status === 429) return "Too many requests right now. Wait a minute and try again.";
    if (status === 529 || status === 503) return "Claude is busy right now. Try again in a moment.";
    if (/credit balance/i.test(detail)) {
      return "Your Anthropic account is out of credit. Add credit at console.anthropic.com → Billing.";
    }
    return "Claude returned an error" + (status ? " (" + status + ")" : "") + (detail ? ": " + detail : ".");
  }

  function geminiErrorMessage(status, bodyText) {
    var detail = errorDetail(bodyText);
    if (/api key/i.test(detail) && (status === 400 || status === 401 || status === 403)) {
      return "Your Gemini API key isn't valid. Check it in Settings.";
    }
    if (status === 429) {
      return "You've used up Gemini's free requests for now. Wait a minute (or until tomorrow for the daily limit), or switch model in Settings.";
    }
    if (status === 503 || status === 500) return "Gemini is busy right now. Try again in a moment.";
    if (status === 404) return "That Gemini model isn't available to your key. Pick another model in Settings.";
    return "Gemini returned an error" + (status ? " (" + status + ")" : "") + (detail ? ": " + detail : ".");
  }

  // Reads a server-sent-events body, passing each event's parsed JSON data to onData.
  function readSse(res, onData) {
    var reader = res.body.getReader();
    var decoder = new TextDecoder();
    var buffer = "";

    function handleEvent(chunk) {
      var data = chunk
        .split("\n")
        .filter(function (line) { return line.indexOf("data:") === 0; })
        .map(function (line) { return line.slice(5).trim(); })
        .join("");
      if (!data) return;
      var event;
      try {
        event = JSON.parse(data);
      } catch (err) {
        return;
      }
      onData(event);
    }

    function pump() {
      return reader.read().then(function (part) {
        if (part.done) {
          if (buffer.trim()) handleEvent(buffer);
          return;
        }
        buffer += decoder.decode(part.value, { stream: true }).replace(/\r/g, "");
        var events = buffer.split("\n\n");
        buffer = events.pop();
        events.forEach(handleEvent);
        return pump();
      });
    }
    return pump();
  }

  function postStream(url, headers, body, signal, errorMessage, onData) {
    return fetch(url, { method: "POST", headers: headers, body: JSON.stringify(body), signal: signal })
      .then(function (res) {
        if (!res.ok) {
          return res.text().then(function (text) {
            throw new Error(errorMessage(res.status, text));
          });
        }
        return readSse(res, onData);
      });
  }

  // Each stream function calls onText with pieces of reply text and resolves with
  // "refusal", "max_tokens" or another stop reason.
  function streamClaude(apiKey, model, prompt, signal, onText) {
    var body = {
      model: model,
      max_tokens: 16000,
      stream: true,
      output_config: { effort: "medium" },
      system: prompt.system,
      messages: [{ role: "user", content: prompt.user }],
    };
    var headers = {
      "content-type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2026-10-21",
      "anthropic-dangerous-direct-browser-access": "true",
    };
    // If a safety check declines the request, let the API retry it on its recommended model.
    if (model !== "claude-haiku-5-5") {
      headers["anthropic-beta"] = "server-side-fallback-2026-07-01";
      body.fallbacks = "default";
    }

    var stopReason = null;
    return postStream(CLAUDE_URL, headers, body, signal, claudeErrorMessage, function (event) {
      if (event.type === "content_block_delta" && event.delta && event.delta.type === "text_delta") {
        onText(event.delta.text);
      } else if (event.type === "message_delta" && event.delta && event.delta.stop_reason) {
        stopReason = event.delta.stop_reason;
      } else if (event.type === "error") {
        throw new Error(claudeErrorMessage(0, JSON.stringify(event)));
      }
    }).then(function () {
      return stopReason;
    });
  }

  var GEMINI_BLOCKED = /^(SAFETY|BLOCKLIST|PROHIBITED_CONTENT|SPII|RECITATION|OTHER)$/;

  function streamGemini(apiKey, model, prompt, signal, onText) {
    var body = {
      systemInstruction: { parts: [{ text: prompt.system }] },
      contents: [{ role: "user", parts: [{ text: prompt.user }] }],
      generationConfig: { maxOutputTokens: 8192, thinkingConfig: { thinkingLevel: "LOW" } },
    };
    var headers = { "content-type": "application/json", "x-goog-api-key": apiKey };
    var url = GEMINI_URL + encodeURIComponent(model) + ":streamGenerateContent?alt=sse";

    var stopReason = null;
    return postStream(url, headers, body, signal, geminiErrorMessage, function (event) {
      if (event.error) throw new Error(geminiErrorMessage(event.error.code || 0, JSON.stringify(event)));
      if (event.promptFeedback && event.promptFeedback.blockReason) stopReason = "refusal";
      var candidate = event.candidates && event.candidates[0];
      if (!candidate) return;
      var parts = (candidate.content && candidate.content.parts) || [];
      parts.forEach(function (part) {
        if (part.text && !part.thought) onText(part.text);
      });
      if (candidate.finishReason === "MAX_TOKENS") stopReason = "max_tokens";
      else if (GEMINI_BLOCKED.test(candidate.finishReason || "")) stopReason = "refusal";
      else if (candidate.finishReason) stopReason = candidate.finishReason;
    }).then(function () {
      return stopReason;
    });
  }

  function setWritingUi(on) {
    if (on) els.writeBtn.textContent = "Stop";
    else updateWriteLabel();
    els.writeBtn.classList.toggle("stop", on);
    els.draftBtn.disabled = on;
  }

  function onWrite() {
    if (writing) {
      writing.abort();
      return;
    }
    if (!requireEmail()) return;
    var which = provider();
    var name = PROVIDER_NAMES[which];
    var apiKey = readApiKey(which);
    if (!apiKey) {
      openSettings();
      setStatus("Add your " + name + " API key first, then click Save.", "warn");
      (which === "gemini" ? els.setGeminiKey : els.setApiKey).focus();
      return;
    }

    var ctx = promptContext();
    var all = replyMode() === "replyAll";
    var prompt = T.buildReplyPrompt({
      myName: email.myName,
      myEmail: email.myEmail,
      fromMe: email.fromMe,
      replyAll: all,
      tone: ctx.tone,
      language: ctx.language,
      instruction: ctx.instruction,
      styleNotes: settings.styleNotes,
      subject: email.subject,
      senderLine: email.senderLine,
      to: email.to,
      cc: email.cc,
      date: email.date,
      body: ctx.cleaned.text,
    });

    var controller = new AbortController();
    writing = controller;
    var started = false;
    els.reply.value = "";
    setWritingUi(true);
    setStatus(name + " is reading the thread…", "info");

    function onText(text) {
      if (!started) {
        started = true;
        text = text.replace(/^\s+/, "");
        setStatus("Writing…", "info");
      }
      els.reply.value += text;
      els.reply.scrollTop = els.reply.scrollHeight;
    }

    var run = which === "gemini"
      ? streamGemini(apiKey, GEMINI_MODELS[settings.geminiModel] ? settings.geminiModel : DEFAULTS.geminiModel, prompt, controller.signal, onText)
      : streamClaude(apiKey, MODELS[settings.model] ? settings.model : DEFAULTS.model, prompt, controller.signal, onText);

    run
      .then(function (stopReason) {
        els.reply.value = els.reply.value.trim();
        var note = ctx.cleaned.truncated ? " (long thread shortened to fit the max length)" : "";
        if (stopReason === "refusal") {
          setStatus(name + " declined to draft this one. Try changing the instruction, the AI in Settings, or use ChatGPT instead.", "warn");
        } else if (stopReason === "max_tokens") {
          setStatus("The reply was cut off because it got too long. Check the end before using it.", "warn");
        } else if (!els.reply.value) {
          setStatus(name + " didn't return any text. Try again.", "warn");
        } else {
          setStatus("Reply ready ✓ Edit it if you like, then click Create reply draft" + note + ".", "ok");
        }
      })
      .catch(function (err) {
        if (err && err.name === "AbortError") {
          setStatus("Stopped.", "info");
        } else if (err instanceof TypeError) {
          setStatus("Couldn't reach " + name + ". Check your internet connection and try again.", "err");
        } else {
          setStatus(err && err.message ? err.message : "Something went wrong while writing.", "err");
        }
      })
      .then(function () {
        if (writing === controller) writing = null;
        setWritingUi(false);
      });
  }

  // ---------- ChatGPT (manual copy/paste) ----------

  function buildChatGPTPrompt() {
    var ctx = promptContext();
    var text = T.fillTemplate(settings.template, {
      tone: ctx.tone,
      instruction: ctx.instruction || "none",
      language: ctx.language,
      senderName: email.senderName,
      senderEmail: email.senderEmail,
      subject: email.subject,
      date: email.date,
      to: email.to,
      cc: email.cc,
      body: ctx.cleaned.text,
    });
    return { text: text, truncated: ctx.cleaned.truncated };
  }

  // execCommand runs synchronously, so it stays inside the click (which WebKit
  // in Outlook for Mac requires); the async Clipboard API is the second try.
  function execCopy(text) {
    var ta = document.createElement("textarea");
    ta.value = text;
    ta.setAttribute("readonly", "");
    ta.style.position = "fixed";
    ta.style.top = "-1000px";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    var ok = false;
    try {
      ta.select();
      ta.setSelectionRange(0, text.length);
      ok = document.execCommand("copy");
    } catch (err) {
      ok = false;
    }
    document.body.removeChild(ta);
    return ok;
  }

  function copyText(text) {
    if (execCopy(text)) return Promise.resolve(true);
    if (navigator.clipboard && navigator.clipboard.writeText) {
      try {
        return navigator.clipboard.writeText(text).then(
          function () { return true; },
          function () { return false; }
        );
      } catch (err) {
        /* fall through */
      }
    }
    return Promise.resolve(false);
  }

  function openChatGPT() {
    try {
      if (Office.context.requirements.isSetSupported("OpenBrowserWindowApi", "1.1")) {
        Office.context.ui.openBrowserWindow(CHATGPT_URL);
        return;
      }
    } catch (err) {
      /* fall back to window.open */
    }
    window.open(CHATGPT_URL, "_blank", "noopener");
  }

  function onCopyPrompt() {
    if (!requireEmail()) return;
    var prompt = buildChatGPTPrompt();
    els.promptText.value = prompt.text;
    var note = prompt.truncated ? " (long email shortened)" : "";

    copyText(prompt.text).then(function (copied) {
      if (!copied) {
        els.promptText.focus();
        els.promptText.select();
        setStatus("Couldn't copy automatically. Click \"Select all & copy\" or press ⌘C, then open ChatGPT.", "warn");
        return;
      }
      if (settings.useMacApp) {
        setStatus("Prompt copied ✓ Paste it into ChatGPT (⌘V)" + note, "ok");
      } else {
        openChatGPT();
        setStatus("Prompt copied ✓ Paste it into ChatGPT (⌘V), then paste its answer below" + note, "ok");
      }
    });
  }

  function onSelectAndCopy() {
    if (!els.promptText.value) {
      if (!requireEmail()) return;
      els.promptText.value = buildChatGPTPrompt().text;
    }
    els.promptText.focus();
    els.promptText.select();
    var ok = false;
    try {
      ok = document.execCommand("copy");
    } catch (err) {
      ok = false;
    }
    setStatus(ok ? "Prompt copied ✓" : "Text selected. Press ⌘C to copy it.", ok ? "ok" : "warn");
  }

  function onPaste() {
    if (!navigator.clipboard || !navigator.clipboard.readText) {
      els.reply.focus();
      setStatus("Click in the reply box and press ⌘V to paste.", "info");
      return;
    }
    var p;
    try {
      p = navigator.clipboard.readText();
    } catch (err) {
      p = Promise.reject(err);
    }
    p.then(
      function (text) {
        if (!text || !text.trim()) {
          setStatus("The clipboard is empty. Copy ChatGPT's answer first.", "warn");
          return;
        }
        els.reply.value = text;
        setStatus("Pasted ✓ Check it, then create the reply draft.", "ok");
      },
      function () {
        els.reply.focus();
        setStatus("Pasting is blocked here. Click in the reply box and press ⌘V.", "info");
      }
    );
  }

  // ---------- reply draft ----------

  function onCreateDraft() {
    var text = els.reply.value;
    if (!text.trim()) {
      els.reply.focus();
      setStatus("Write the reply first (or paste one into the box).", "warn");
      return;
    }
    var item = currentItem();
    if (!item) {
      setStatus("No email selected. Select the email you're replying to.", "warn");
      return;
    }
    var html = T.textToHtml(text);
    if (html.length > 32000) {
      setStatus("That reply is too long for Outlook (about 32,000 characters max). Shorten it and try again.", "warn");
      return;
    }
    var all = replyMode() === "replyAll";
    try {
      if (all) item.displayReplyAllForm({ htmlBody: html });
      else item.displayReplyForm({ htmlBody: html });
      setStatus((all ? "Reply All" : "Reply") + " draft opened ✓ Review it and send it yourself.", "ok");
    } catch (err) {
      setStatus("Couldn't open the reply draft: " + err.message, "err");
    }
  }

  // ---------- theme ----------

  function hexLuminance(hex) {
    var m = /^#?([0-9a-f]{6})$/i.exec(String(hex || "").trim());
    if (!m) return null;
    var n = parseInt(m[1], 16);
    return (0.299 * ((n >> 16) & 255) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255;
  }

  // Follows the Office theme where Outlook reports it; otherwise the CSS follows macOS.
  function applyOfficeTheme(theme) {
    try {
      theme = theme || (Office.context && Office.context.officeTheme);
      var lum = theme && hexLuminance(theme.bodyBackgroundColor);
      if (lum === null || lum === undefined) return;
      document.documentElement.setAttribute("data-theme", lum < 0.5 ? "dark" : "light");
    } catch (err) {
      /* keep the macOS appearance */
    }
  }

  // ---------- start ----------

  function bind() {
    [
      "main", "mainView", "settingsView", "emailFrom", "emailSubject", "tone", "language",
      "instruction", "writeBtn", "copyPromptBtn", "promptText", "selectCopyBtn", "openChatBtn",
      "reply", "pasteBtn", "draftBtn", "status", "settingsBtn", "setApiKey", "setModel",
      "setProvider", "geminiSettings", "claudeSettings", "setGeminiKey", "setGeminiModel",
      "setStyleNotes", "setFullThread", "setTemplate", "resetTemplateBtn", "setTone",
      "setReplyMode", "setMaxBody", "setMacApp", "saveSettingsBtn", "cancelSettingsBtn",
    ].forEach(function (id) {
      els[id] = $(id);
    });

    els.writeBtn.addEventListener("click", safe(onWrite));
    els.copyPromptBtn.addEventListener("click", safe(onCopyPrompt));
    els.selectCopyBtn.addEventListener("click", safe(onSelectAndCopy));
    els.openChatBtn.addEventListener("click", safe(function () {
      openChatGPT();
      setStatus("ChatGPT opened in your browser.", "info");
    }));
    els.pasteBtn.addEventListener("click", safe(onPaste));
    els.draftBtn.addEventListener("click", safe(onCreateDraft));
    els.settingsBtn.addEventListener("click", safe(function () {
      if (els.settingsView.hidden) openSettings();
      else closeSettings();
    }));
    els.setProvider.addEventListener("change", safe(showProviderSettings));
    els.cancelSettingsBtn.addEventListener("click", safe(closeSettings));
    els.saveSettingsBtn.addEventListener("click", safe(onSaveSettings));
    els.resetTemplateBtn.addEventListener("click", safe(function () {
      els.setTemplate.value = T.DEFAULT_TEMPLATE;
      setStatus("Template reset. Click Save to keep it.", "info");
    }));
  }

  function start() {
    bind();

    if (!window.Office || !Office.onReady) {
      setStatus("Office.js didn't load. Check your internet connection, then reopen the panel.", "err");
      return;
    }

    Office.onReady(function (info) {
      try {
        if (info.host !== Office.HostType.Outlook || !mailbox()) {
          applyDefaultsToForm();
          setStatus("Open this panel from Outlook to use it.", "warn");
          return;
        }
        loadSettings();
        applyDefaultsToForm();
        applyOfficeTheme();
        loadEmail();

        mailbox().addHandlerAsync(Office.EventType.ItemChanged, safe(onItemChanged), function () {});
        if (Office.EventType.OfficeThemeChanged) {
          mailbox().addHandlerAsync(Office.EventType.OfficeThemeChanged, function (e) {
            applyOfficeTheme(e && e.officeTheme);
          }, function () {});
        }
      } catch (err) {
        console.error(err);
        setStatus("Couldn't start: " + err.message, "err");
      }
    });
  }

  window.addEventListener("error", function (e) {
    setStatus("Something went wrong: " + (e.message || "unknown error"), "err");
  });
  window.addEventListener("unhandledrejection", function (e) {
    e.preventDefault();
    setStatus("Something went wrong: " + ((e.reason && e.reason.message) || "unknown error"), "err");
  });

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start);
  else start();
})();
