/* Plain-HTTP static server for the mock preview only (no certs needed). */
const http = require("http");
const fs = require("fs");
const path = require("path");

const ROOT = path.join(__dirname, ".."); // MacTrace/
const PORT = 4123;
const MIME = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".png": "image/png" };

http
  .createServer((req, res) => {
    let urlPath = decodeURIComponent(req.url.split("?")[0]);
    if (urlPath === "/") urlPath = "/preview/preview.html";
    const filePath = path.normalize(path.join(ROOT, urlPath));
    if (!filePath.startsWith(ROOT)) { res.writeHead(403); return res.end("Forbidden"); }
    fs.readFile(filePath, (err, data) => {
      if (err) { res.writeHead(404); return res.end("Not found"); }
      res.writeHead(200, { "Content-Type": MIME[path.extname(filePath)] || "application/octet-stream" });
      res.end(data);
    });
  })
  .listen(PORT, () => console.log("Preview at http://localhost:" + PORT));
