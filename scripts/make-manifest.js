// Writes manifest.prod.xml with your hosting URL in place of https://localhost:3000.
// Usage: npm run manifest -- https://yourname.github.io/reply-with-chatgpt
"use strict";

const fs = require("fs");
const path = require("path");

const base = (process.argv[2] || "").trim().replace(/\/+$/, "");
if (!/^https:\/\/[^\s"<>]+$/.test(base)) {
  console.error("Give your hosting URL (must start with https://), for example:");
  console.error("  npm run manifest -- https://yourname.github.io/reply-with-chatgpt");
  process.exit(1);
}

const src = path.join(__dirname, "..", "manifest.xml");
const out = path.join(__dirname, "..", "manifest.prod.xml");
const xml = fs.readFileSync(src, "utf8").split("https://localhost:3000").join(base);
fs.writeFileSync(out, xml);
console.log(`Wrote manifest.prod.xml for ${base}`);
console.log(`Check it loads: ${base}/taskpane.html`);
