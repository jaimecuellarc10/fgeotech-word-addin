const https = require("https");
const http = require("http");
const fs = require("fs");
const path = require("path");
const PORT = 3000;
const USE_HTTPS = process.argv.includes("--https");

const MIME_TYPES = {
  ".html": "text/html",
  ".js": "application/javascript",
  ".css": "text/css",
  ".json": "application/json",
  ".png": "image/png",
  ".ico": "image/x-icon",
  ".xml": "application/xml",
};

function serveStatic(req, res) {
  const parsed = new URL(req.url, "https://localhost");
  const pathname = parsed.pathname;

  // Proxy endpoint: GET /proxy/projects?number=1234 with header X-Synergy-Key
  if (pathname === "/proxy/projects") {
    const projectNumber = parsed.searchParams.get("number");
    const apiKey = req.headers["x-synergy-key"];
    const orgSlug = req.headers["x-synergy-org"] || "actgeotechnicalengineers";

    if (!projectNumber || !apiKey) {
      res.writeHead(400, { "Content-Type": "application/json" });
      return res.end(JSON.stringify({ error: "Missing project number or API key" }));
    }

    const apiOptions = {
      hostname: "api.totalsynergy.com",
      path: `/api/v2/Organisation/${encodeURIComponent(orgSlug)}/Projects?criteria.projectNumber=${encodeURIComponent(projectNumber)}`,
      method: "GET",
      headers: {
        "access-token": apiKey,
        Accept: "application/json",
      },
    };

    const proxyReq = https.request(apiOptions, (proxyRes) => {
      let data = "";
      proxyRes.on("data", (chunk) => (data += chunk));
      proxyRes.on("end", () => {
        console.log(`[Synergy API] ${apiOptions.path} → ${proxyRes.statusCode}`);
        fs.writeFileSync("synergy_response.json", data);
        res.writeHead(proxyRes.statusCode, {
          "Content-Type": "application/json",
          "Access-Control-Allow-Origin": "*",
        });
        res.end(data);
      });
    });

    proxyReq.on("error", (err) => {
      res.writeHead(502, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ error: "Failed to reach Total Synergy API", detail: err.message }));
    });

    return proxyReq.end();
  }

  // Static file serving
  let filePath = path.join(__dirname, pathname === "/" ? "/taskpane.html" : pathname);
  const ext = path.extname(filePath);
  const contentType = MIME_TYPES[ext] || "text/plain";

  fs.readFile(filePath, (err, content) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain" });
      return res.end("Not found");
    }
    res.writeHead(200, {
      "Content-Type": contentType,
      "Access-Control-Allow-Origin": "*",
    });
    res.end(content);
  });
}

if (USE_HTTPS) {
  // For HTTPS, generate certs first:
  // npx office-addin-dev-certs install
  // Then run: node server.js --https
  try {
    const certDir = path.join(require("os").homedir(), ".office-addin-dev-certs");
    const options = {
      key: fs.readFileSync(path.join(certDir, "localhost.key")),
      cert: fs.readFileSync(path.join(certDir, "localhost.crt")),
    };
    https.createServer(options, serveStatic).listen(PORT, () => {
      console.log(`Synergy Panel server running at https://localhost:${PORT}`);
      console.log("Open taskpane.html via the Synergy Panel button in Word Home ribbon.");
    });
  } catch {
    console.error("HTTPS certs not found. Run: npx office-addin-dev-certs install");
    console.error("Or start without --https flag for HTTP (requires manifest update).");
    process.exit(1);
  }
} else {
  http.createServer(serveStatic).listen(PORT, () => {
    console.log(`Synergy Panel server running at http://localhost:${PORT}`);
    console.log("If Word rejects HTTP, run with --https after installing dev certs.");
  });
}
