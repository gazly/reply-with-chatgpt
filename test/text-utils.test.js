"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const T = require("../docs/text-utils.js");

test("cuts Outlook-style quoted history", () => {
  const body = "Thanks, see you then.\n\n________________________________\nFrom: Ana <ana@x.com>\nSent: Monday\nTo: Me\nSubject: Re: Meet\n\nOld text";
  assert.equal(T.stripQuotedHistory(body), "Thanks, see you then.");
});

test("cuts 'On ... wrote:' including when wrapped over two lines", () => {
  assert.equal(T.stripQuotedHistory("Sure!\n\nOn Mon, 2 Nov 2026 at 09:00, Bob <b@x.com> wrote:\n> hi"), "Sure!");
  assert.equal(T.stripQuotedHistory("Sure!\nOn Mon, 2 Nov 2026 at 09:00, Bob\n<b@x.com> wrote:\n> hi"), "Sure!");
});

test("cuts -----Original Message-----", () => {
  assert.equal(T.stripQuotedHistory("New\r\n-----Original Message-----\r\nFrom: a\r\nOld"), "New");
});

test("keeps a 'From:' line that isn't a quoted header", () => {
  const body = "From: the team, thanks for everything.\nBest";
  assert.equal(T.stripQuotedHistory(body), body);
});

test("never cuts away the whole email", () => {
  const body = "From: Ana\nSent: Monday\nSubject: X\n\nOnly a forwarded message";
  assert.equal(T.stripQuotedHistory(body), body);
});

test("limits length without splitting an emoji", () => {
  const text = "a".repeat(499) + "😀" + "b".repeat(100);
  const out = T.limitLength(text, 500);
  assert.equal(out.truncated, true);
  assert.ok(out.text.startsWith("a".repeat(499) + "\n[…"));
});

test("keeps Sinhala and Tamil text intact", () => {
  const body = "ආයුබෝවන්, ස්තූතියි.\nவணக்கம், நன்றி.";
  assert.equal(T.cleanEmailBody(body, 6000).text, body);
});

test("fills known placeholders and leaves unknown ones", () => {
  assert.equal(T.fillTemplate("{tone} / {x}", { tone: "Brief" }), "Brief / {x}");
});

test("textToHtml escapes and builds paragraphs and line breaks", () => {
  const html = T.textToHtml("Hi <Ana> & co,\nthanks.\n\n**Yes**, Thursday works.");
  assert.equal(
    html,
    '<p style="margin:0 0 12px 0">Hi &lt;Ana&gt; &amp; co,<br>thanks.</p>' +
      '<p style="margin:0 0 12px 0"><strong>Yes</strong>, Thursday works.</p>'
  );
});

test("drops a stray Subject line and code fence from ChatGPT's answer", () => {
  assert.equal(T.tidyReplyText("Subject: Re: Meet\n\nHello"), "Hello");
  assert.equal(T.tidyReplyText("```\nHello\n```"), "Hello");
});

test("keeps the whole thread when asked", () => {
  const body = "Sounds good.\n\nOn Mon, 2 Nov 2026 at 09:00, Bob <b@x.com> wrote:\n> Can we meet?";
  assert.equal(T.cleanEmailBody(body, 6000, true).text, body);
  assert.equal(T.cleanEmailBody(body, 6000, false).text, "Sounds good.");
});

test("AI prompt writes as me and handles my own latest message", () => {
  const p = T.buildReplyPrompt({
    myName: "Muhammed Gazzaly", myEmail: "g@x.com", fromMe: true, replyAll: true,
    tone: "professional and clear", language: "the same language as the email",
    instruction: "Confirm Thursday", styleNotes: "UK spelling",
    subject: "Demo", senderLine: "Muhammed Gazzaly <g@x.com>", to: "Harshit <h@y.com>", cc: "",
    date: "today", body: "Thread text",
  });
  assert.match(p.system, /on behalf of Muhammed Gazzaly <g@x\.com>/);
  assert.match(p.user, /latest message was sent by me/);
  assert.match(p.user, /Reply All\): Muhammed Gazzaly <g@x\.com>, Harshit <h@y\.com>\./);
  assert.match(p.user, /What I want to say: Confirm Thursday/);
  assert.match(p.user, /<thread>[\s\S]*Thread text\n<\/thread>$/);
});
