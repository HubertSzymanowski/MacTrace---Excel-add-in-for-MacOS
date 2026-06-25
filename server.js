/*
 * MacTrace dev server
 * -------------------
 * Serves the add-in's static files over HTTPS on https://localhost:3000.
 * Office add-ins require HTTPS even locally, so we use office-addin-dev-certs
 * to create (and trust) a localhost certificate the first time it runs.
 *
 *   node server.js            -> serves the project root
 */
const https = require("https");
const fs = require("fs");
const path = require("path");
const devCerts = require("office-addin-dev-certs");

const PORT = 3100;
const ROOT = __dirname;

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".xml": "text/xml; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon"
};

function send(res, status, body, type) {
  res.writeHead(status, {
    "Content-Type": type || "text/plain; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Cache-Control": "no-store"
  });
  res.end(body);
}

(async () => {
  let options;
  try {
    // Installs/trusts a localhost CA on first run (may prompt for your password).
    options = await devCerts.getHttpsServerOptions();
  } catch (e) {
    console.error("Could not get HTTPS certificates:", e.message || e);
    console.error("Try:  npx office-addin-dev-certs install");
    process.exit(1);
  }

  https
    .createServer(options, (req, res) => {
      let urlPath = decodeURIComponent(req.url.split("?")[0]);
      console.log(new Date().toISOString(), req.method, urlPath); // request log (diagnostics)
      if (urlPath === "/") urlPath = "/src/taskpane.html";

      // Resolve safely inside ROOT (block path traversal).
      const filePath = path.normalize(path.join(ROOT, urlPath));
      if (!filePath.startsWith(ROOT)) return send(res, 403, "Forbidden");

      fs.readFile(filePath, (err, data) => {
        if (err) return send(res, 404, "Not found: " + urlPath);
        send(res, 200, data, MIME[path.extname(filePath)] || "application/octet-stream");
      });
    })
    .listen(PORT, () => {
      console.log("MacTrace is serving at https://localhost:" + PORT);
      console.log("  Task pane: https://localhost:" + PORT + "/src/taskpane.html");
      console.log("Leave this running, then open Excel.");
    });
})();
