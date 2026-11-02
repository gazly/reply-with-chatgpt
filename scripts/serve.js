// Local HTTPS server for testing the add-in at https://localhost:3000.
// Serves the static files in docs/ using the Office dev certificate
// (installed on first run; macOS may ask for your password once).
"use strict";

const https = require("https");
const fs = require("fs");
const path = require("path");
const devCerts = require("office-addin-dev-certs");

const PORT = 3000;
const ROOT = path.join(__dirname, "..", "docs");
const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".json": "application/json; charset=utf-8",
  ".ico": "image/x-icon",
};

function handle(req, res) {
  let urlPath;
  try {
    urlPath = decodeURIComponent(new URL(req.url, "https://localhost").pathname);
  } catch {
    res.writeHead(400).end("Bad request");
    return;
  }
  if (urlPath.endsWith("/")) urlPath += "index.html";
  const file = path.normalize(path.join(ROOT, urlPath));
  if (!file.startsWith(ROOT + path.sep)) {
    res.writeHead(403).end("Forbidden");
    return;
  }
  fs.readFile(file, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain" }).end("Not found");
      console.log(`404 ${urlPath}`);
      return;
    }
    res.writeHead(200, {
      "Content-Type": TYPES[path.extname(file)] || "application/octet-stream",
      "Cache-Control": "no-store",
    });
    res.end(data);
  });
}

devCerts
  .getHttpsServerOptions()
  .then((options) => {
    https.createServer(options, handle).listen(PORT, () => {
      console.log(`Serving docs/ at https://localhost:${PORT}`);
      console.log(`Check it in a browser: https://localhost:${PORT}/taskpane.html`);
      console.log("Press Ctrl+C to stop.");
    });
  })
  .catch((err) => {
    console.error("Could not set up the HTTPS certificate:", err.message || err);
    console.error("Try: npm run certs");
    process.exit(1);
  });
