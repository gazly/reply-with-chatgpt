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

  function cleanEmailBody(text, maxChars) {
    return limitLength(stripQuotedHistory(text), maxChars);
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
    fillTemplate: fillTemplate,
    escapeHtml: escapeHtml,
    tidyReplyText: tidyReplyText,
    textToHtml: textToHtml,
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.RWCText = api;
})(this);
