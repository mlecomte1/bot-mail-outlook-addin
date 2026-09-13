const http = require("http");
const fs = require("fs");
const path = require("path");
const generateReply = require("./api/generate-reply");
const { applyCors } = require("./api/security");

const publicDir = path.resolve(__dirname, "public");
const port = Number(process.env.PORT) || 3000;
const MAX_BODY_BYTES = 100_000;

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".png": "image/png",
  ".xml": "application/xml; charset=utf-8",
};

const securityHeaders = {
  "X-Content-Type-Options": "nosniff",
  "Referrer-Policy": "no-referrer",
  "Cache-Control": "no-store",
  "Content-Security-Policy":
    "default-src 'self'; script-src 'self' https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com; style-src 'self'; connect-src 'self' https://www.google.com https://www.googleapis.com https://speech.googleapis.com; img-src 'self'; media-src 'self'; object-src 'none'; base-uri 'self'; form-action 'none'",
};

function responseHeaders(res) {
  const headers = {
    ...securityHeaders,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };
  const origin = res.getHeader("Access-Control-Allow-Origin");

  if (origin) {
    headers["Access-Control-Allow-Origin"] = origin;
    headers.Vary = "Origin";
  }

  return headers;
}

function sendJson(req, res, statusCode, body) {
  applyCors(req, res);
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    ...responseHeaders(res),
  });
  res.end(JSON.stringify(body));
}

function createApiResponse(req, res) {
  applyCors(req, res);

  const apiResponse = {
    setHeader(name, value) {
      res.setHeader(name, value);
      return apiResponse;
    },
    status(statusCode) {
      return {
        json(responseBody) {
          sendJson(req, res, statusCode, responseBody);
        },
        send(responseBody) {
          applyCors(req, res);
          res.writeHead(statusCode, responseHeaders(res));
          res.end(responseBody ?? "");
        },
        end() {
          applyCors(req, res);
          res.writeHead(statusCode, responseHeaders(res));
          res.end();
        },
      };
    },
  };

  return apiResponse;
}

function isInsidePublic(fullPath) {
  const resolved = path.resolve(fullPath);
  return resolved === publicDir || resolved.startsWith(publicDir + path.sep);
}

const server = http.createServer((req, res) => {
  let urlPath;

  try {
    urlPath = decodeURIComponent((req.url || "/").split("?")[0]);
  } catch {
    res.writeHead(400, securityHeaders);
    res.end("Bad request");
    return;
  }

  if (urlPath === "/api/generate-reply") {
    if (req.method === "OPTIONS") {
      applyCors(req, res);
      res.writeHead(204, responseHeaders(res));
      res.end();
      return;
    }

    let requestBody = "";
    let tooLarge = false;

    req.on("data", (chunk) => {
      if (tooLarge) {
        return;
      }

      requestBody += chunk;

      if (Buffer.byteLength(requestBody) > MAX_BODY_BYTES) {
        tooLarge = true;
        req.destroy();
        sendJson(req, res, 413, { error: "Requête trop volumineuse." });
      }
    });

    req.on("end", async () => {
      if (tooLarge) {
        return;
      }

      let body = {};

      if (requestBody) {
        try {
          body = JSON.parse(requestBody);
        } catch {
          sendJson(req, res, 400, { error: "JSON invalide." });
          return;
        }
      }

      try {
        await generateReply(
          { method: req.method, body, headers: req.headers },
          createApiResponse(req, res)
        );
      } catch (error) {
        console.error("ERREUR API LOCALE :", error?.message || "unknown");
        sendJson(req, res, 500, { error: "Erreur serveur pendant la génération." });
      }
    });

    return;
  }

  if (req.method !== "GET" && req.method !== "HEAD") {
    res.writeHead(405, securityHeaders);
    res.end("Method not allowed");
    return;
  }

  const filePath = urlPath === "/" ? "/taskpane.html" : urlPath;
  const fullPath = path.resolve(publicDir, `.${filePath}`);

  if (!isInsidePublic(fullPath)) {
    res.writeHead(403, securityHeaders);
    res.end("Forbidden");
    return;
  }

  fs.readFile(fullPath, (err, content) => {
    if (err) {
      res.writeHead(404, securityHeaders);
      res.end("Not found");
      return;
    }

    const ext = path.extname(fullPath);
    res.writeHead(200, {
      "Content-Type": mimeTypes[ext] || "text/plain; charset=utf-8",
      ...securityHeaders,
    });
    res.end(content);
  });
});

server.listen(port, "127.0.0.1", () => {
  console.log(`Local server running on http://127.0.0.1:${port}`);
});
