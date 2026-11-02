/*
 * Reply with ChatGPT: task pane logic.
 * Everything runs inside this panel. The only things it opens are chatgpt.com
 * (in your browser, for you) and Outlook's own reply form. It never sends email.
 */
(function () {
  "use strict";

  var T = window.RWCText;
  var CHATGPT_URL = "https://chatgpt.com/";
  var SETTINGS_KEY = "rwcSettings";
  var DEFAULTS = {
    template: T.DEFAULT_TEMPLATE,
    tone: "Professional",
    replyMode: "reply",
    maxBody: 6000,
    useMacApp: false,
  };

  var settings = Object.assign({}, DEFAULTS);
  var email = null; // the email as read when the panel opened or the item changed
  var emailState = "loading"; // loading | ready | none | error
  var loadToken = 0;
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

  function fillToneSelect(select, value) {
    select.innerHTML = "";
    Object.keys(T.TONES).forEach(function (name) {
      var opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      select.appendChild(opt);
    });
    select.value = T.TONES[value] ? value : DEFAULTS.tone;
  }

  // ---------- settings (roamingSettings) ----------

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
  }

  function openSettings() {
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
    if (!(max >= 500 && max <= 50000)) {
      setStatus("Max email length must be between 500 and 50000.", "warn");
      return;
    }
    var template = els.setTemplate.value.trim() ? els.setTemplate.value : T.DEFAULT_TEMPLATE;
    if (template.indexOf("{body}") === -1) {
      setStatus("The template needs {body} so ChatGPT sees the email.", "warn");
      return;
    }
    var next = {
      template: template,
      tone: els.setTone.value,
      replyMode: els.setReplyMode.value,
      maxBody: max,
      useMacApp: els.setMacApp.checked,
    };
    els.saveSettingsBtn.disabled = true;
    saveSettings(next, function (ok, message) {
      els.saveSettingsBtn.disabled = false;
      if (ok) {
        applyDefaultsToForm();
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
    els.emailFrom.textContent = formatPerson(from) || "(unknown sender)";
    els.emailFrom.title = formatPerson(from);
    els.emailSubject.textContent = item.subject || "(no subject)";
    els.emailSubject.title = item.subject || "";
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
          to: (item.to || []).map(formatPerson).join(", "),
          cc: (item.cc || []).map(formatPerson).join(", "),
          subject: item.subject || "",
          date: item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString() : "",
          rawBody: result.value || "",
        };
        emailState = "ready";
        if (!email.rawBody.trim()) setStatus("This email has no text body, but you can still ask ChatGPT.", "warn");
        else setStatus("Ready. Choose a tone, then copy the prompt.", "info");
      });
    } catch (err) {
      emailState = "error";
      setStatus("Couldn't read this email: " + err.message, "err");
    }
  }

  function onItemChanged() {
    // Pinned panel: a different email was selected. Clear the old reply so it can't go to the wrong person.
    els.reply.value = "";
    els.promptText.value = "";
    els.promptDetails.open = false;
    loadEmail();
  }

  // ---------- step 1: prompt ----------

  function buildPrompt() {
    var cleaned = T.cleanEmailBody(email.rawBody, settings.maxBody);
    var instruction = els.instruction.value.trim();
    var prompt = T.fillTemplate(settings.template, {
      tone: T.TONES[els.tone.value] || els.tone.value,
      instruction: instruction || "none",
      language: els.language.value || "the same language as the email",
      senderName: email.senderName,
      senderEmail: email.senderEmail,
      subject: email.subject,
      date: email.date,
      to: email.to,
      cc: email.cc,
      body: cleaned.text,
    });
    return { text: prompt, truncated: cleaned.truncated };
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

  function showManualCopy(message) {
    els.promptDetails.open = true;
    els.promptText.focus();
    els.promptText.select();
    setStatus(message, "warn");
  }

  function onCopyPrompt() {
    if (emailState === "loading") {
      setStatus("Still reading the email, try again in a moment.", "info");
      return;
    }
    if (emailState !== "ready" || !email) {
      setStatus("No readable email selected. Select an email first.", "warn");
      return;
    }
    var prompt = buildPrompt();
    els.promptText.value = prompt.text;
    var note = prompt.truncated ? " (long email shortened)" : "";

    copyText(prompt.text).then(function (copied) {
      if (!copied) {
        showManualCopy("Couldn't copy automatically. Click \"Select all & copy\" or press ⌘C, then open ChatGPT.");
        return;
      }
      if (settings.useMacApp) {
        setStatus("Prompt copied ✓ Paste it into ChatGPT (⌘V)" + note, "ok");
      } else {
        openChatGPT();
        setStatus("Prompt copied ✓ Paste it into ChatGPT (⌘V), then copy its answer back here" + note, "ok");
      }
    });
  }

  function onSelectAndCopy() {
    if (!els.promptText.value) {
      if (emailState === "ready" && email) els.promptText.value = buildPrompt().text;
      else {
        setStatus("No readable email selected.", "warn");
        return;
      }
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

  // ---------- step 2: reply draft ----------

  function onPaste() {
    if (!navigator.clipboard || !navigator.clipboard.readText) {
      els.reply.focus();
      setStatus("Click in the box and press ⌘V to paste.", "info");
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
        setStatus("Pasting is blocked here. Click in the box and press ⌘V.", "info");
      }
    );
  }

  function onCreateDraft() {
    var text = els.reply.value;
    if (!text.trim()) {
      els.reply.focus();
      setStatus("Paste ChatGPT's reply into the box first.", "warn");
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
      "instruction", "copyPromptBtn", "promptDetails", "promptText", "selectCopyBtn", "openChatBtn",
      "reply", "pasteBtn", "draftBtn", "status", "settingsBtn", "setTemplate", "resetTemplateBtn",
      "setTone", "setReplyMode", "setMaxBody", "setMacApp", "saveSettingsBtn", "cancelSettingsBtn",
    ].forEach(function (id) {
      els[id] = $(id);
    });

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
