const http = require("http");
const fs = require("fs");
const path = require("path");

const publicDir = path.join(__dirname, "public");

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".png": "image/png",
  ".xml": "application/xml; charset=utf-8",
};

const server = http.createServer((req, res) => {
  let filePath = req.url === "/" ? "/taskpane.html" : req.url;
  filePath = filePath.split("?")[0];

  const fullPath = path.join(publicDir, filePath);

  if (!fullPath.startsWith(publicDir)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  fs.readFile(fullPath, (err, content) => {
    if (err) {
      res.writeHead(404);
      res.end("Not found");
      return;
    }

    const ext = path.extname(fullPath);
    res.writeHead(200, {
      "Content-Type": mimeTypes[ext] || "text/plain; charset=utf-8",
    });
    res.end(content);
  });
});

server.listen(3000, () => {
  console.log("Local server running on http://localhost:3000");
});