/*
 * Pure text helpers for Reply with ChatGPT (no Office.js, no DOM).
 * Loaded by taskpane.html as window.RWCText and by the Node tests via require().
 */
(function (root) {
  "use strict";

  var DEFAULT_TEMPLATE = [
    "Draft a reply to the email below.",
    "Tone: {tone}.",
    "Extra instruction: {instruction}.",
    "Language: {language}.",
    "Output only the reply body: no subject line, no signature, no commentary.",
    "---",
    "From: {senderName} <{senderEmail}>",
    "Subject: {subject}",
    "",
    "{body}",
  ].join("\n");

  // What each Tone choice tells ChatGPT.
  var TONES = {
    Professional: "professional and clear",
    Friendly: "friendly and warm",
    Brief: "brief, a few short sentences at most",
    "Polite decline": "polite, graciously declining the request",
    "Follow-up": "a polite follow-up that moves things forward",
  };

  function normalizeNewlines(text) {
    return String(text == null ? "" : text).replace(/\r\n?/g, "\n");
  }

  var HEADER_LINE = /^\s*\*?(Sent|Date|To|Cc|Subject)\s*:/i;

  // Index of the first line that starts quoted history, or -1.
  function findQuoteStart(lines) {
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      if (/^\s*-{2,}\s*Original Message\s*-{2,}\s*$/i.test(line)) return i;
      // Outlook's separator line above a quoted header block.
      if (/^\s*_{10,}\s*$/.test(line) && nextHasHeader(lines, i + 1)) return i;
      // "On Mon, 3 Nov 2026 at 10:00, Name <a@b.c> wrote:", sometimes wrapped onto two lines.
      if (/^\s*On\s.+/i.test(line)) {
        if (/wrote:\s*$/i.test(line)) return i;
        if (i + 1 < lines.length && /wrote:\s*$/i.test(lines[i + 1]) && line.length < 200) return i;
      }
      // "From: ..." followed by Sent:/To:/Subject: lines.
      if (/^\s*\*?From\s*:/i.test(line) && nextHasHeader(lines, i + 1)) return i;
    }
    return -1;
  }

  function nextHasHeader(lines, start) {
    for (var j = start; j < Math.min(lines.length, start + 5); j++) {
      if (HEADER_LINE.test(lines[j])) return true;
    }
    return false;
  }

  // Removes old quoted history (when it isn't the whole email) and tidies blank lines.
  function stripQuotedHistory(text) {
    var lines = normalizeNewlines(text).split("\n");
    var cut = findQuoteStart(lines);
    if (cut > 0) {
      var kept = lines.slice(0, cut);
      if (kept.join("").trim()) lines = kept;
    }
    return lines
      .join("\n")
      .replace(/[ \t ]+\n/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  // Shortens text to maxChars without splitting an emoji or other surrogate pair.
  function limitLength(text, maxChars) {
    var max = Math.max(200, Number(maxChars) || 6000);
    if (text.length <= max) return { text: text, truncated: false };
    var cut = text.slice(0, max);
    var last = cut.charCodeAt(cut.length - 1);
    if (last >= 0xd800 && last <= 0xdbff) cut = cut.slice(0, -1);
    return { text: cut.trimEnd() + "\n[… email shortened]", truncated: true };
  }

  // keepHistory: send the whole thread (only tidy blank lines) instead of just the newest message.
  function cleanEmailBody(text, maxChars, keepHistory) {
    var body = keepHistory
      ? normalizeNewlines(text).replace(/[ \t ]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim()
      : stripQuotedHistory(text);
    return limitLength(body, maxChars);
  }

  var CLAUDE_SYSTEM = [
    "You draft email replies on behalf of {myName} <{myEmail}>. Write in the first person as {myName}.",
    "Read the whole thread to understand who said what, what was asked and what is still open, then reply to the latest message.",
    "Answer every question addressed to {myName}. Don't invent facts, dates, prices or commitments that aren't in the thread; where you need something you don't know, put a short placeholder in square brackets, such as [confirm time].",
    "Output only the reply body as plain text: no subject line, no signature block, no commentary before or after it. A greeting and a closing line such as \"Best regards,\" are fine, but don't write {myName}'s name under it, because Outlook adds the signature.",
  ].join("\n");

  // Builds the system prompt and user message for Claude from the email context.
  function buildClaudePrompt(c) {
    var me = { myName: c.myName || "the user", myEmail: c.myEmail || "" };
    var lines = [
      "Draft my reply to the email thread below.",
      "Tone: " + c.tone + ".",
      "Language: " + c.language + ".",
    ];
    if (c.instruction) lines.push("What I want to say: " + c.instruction);
    if (c.fromMe) {
      lines.push(
        "The latest message was sent by me, so write a follow-up from me to the other people on the thread, not a reply to myself."
      );
    }
    lines.push(
      c.replyAll
        ? "This goes to everyone on the thread (Reply All): " + [c.senderLine, c.to, c.cc].filter(Boolean).join(", ") + "."
        : "This goes only to " + c.senderLine + "."
    );
    if (c.styleNotes) lines.push("About me and how I write: " + c.styleNotes);
    lines.push(
      "",
      "<thread>",
      "Subject: " + c.subject,
      "Latest message from: " + c.senderLine,
      "To: " + (c.to || "-"),
      "Cc: " + (c.cc || "-"),
      "Date: " + (c.date || "-"),
      "",
      c.body,
      "</thread>"
    );
    return { system: fillTemplate(CLAUDE_SYSTEM, me), user: lines.join("\n") };
  }

  // Replaces {name} placeholders; unknown placeholders are left as they are.
  function fillTemplate(template, values) {
    return String(template || DEFAULT_TEMPLATE).replace(/\{(\w+)\}/g, function (match, key) {
      return Object.prototype.hasOwnProperty.call(values, key) ? String(values[key]) : match;
    });
  }

  function escapeHtml(text) {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  // Tidies text copied from ChatGPT: drops a stray "Subject:" line and code fences.
  function tidyReplyText(text) {
    var t = normalizeNewlines(text).trim();
    t = t.replace(/^```[^\n]*\n([\s\S]*?)\n```$/, "$1").trim();
    t = t.replace(/^\s*\**Subject\**\s*:.*\n+/i, "");
    return t.trim();
  }

  // Plain text → simple HTML: blank lines make paragraphs, single newlines make <br>.
  // **bold** (as ChatGPT's copy button gives it) becomes <strong>.
  function textToHtml(text) {
    var paragraphs = tidyReplyText(text).split(/\n[ \t]*\n+/);
    return paragraphs
      .filter(function (p) {
        return p.trim();
      })
      .map(function (p) {
        var html = escapeHtml(p.trim())
          .replace(/\*\*(?=\S)([^*\n]+?)\*\*/g, "<strong>$1</strong>")
          .replace(/\n/g, "<br>");
        return '<p style="margin:0 0 12px 0">' + html + "</p>";
      })
      .join("");
  }

  var api = {
    DEFAULT_TEMPLATE: DEFAULT_TEMPLATE,
    TONES: TONES,
    stripQuotedHistory: stripQuotedHistory,
    limitLength: limitLength,
    cleanEmailBody: cleanEmailBody,
    buildClaudePrompt: buildClaudePrompt,
    fillTemplate: fillTemplate,
    escapeHtml: escapeHtml,
    tidyReplyText: tidyReplyText,
    textToHtml: textToHtml,
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.RWCText = api;
})(this);
