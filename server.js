// ppMath 用の静的ファイルサーバー
//   node server.js          … https://localhost:3000 （PowerPoint 用、開発証明書が必要）
//   node server.js --http   … http://localhost:3000  （ブラウザで UI だけ確認する用）
const fs = require("fs");
const path = require("path");
const http = require("http");
const https = require("https");

const PORT = Number(process.env.PORT || 3000);
const ROOT = path.join(__dirname, "docs");
const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".woff2": "font/woff2",
  ".woff": "font/woff",
  ".ttf": "font/ttf",
  ".xml": "application/xml",
};

function resolve(urlPath) {
  return safeJoin(ROOT, urlPath === "/" ? "taskpane.html" : urlPath.slice(1));
}

function safeJoin(base, rel) {
  const p = path.normalize(path.join(base, rel));
  return p.startsWith(base) ? p : null;
}

function handler(req, res) {
  const urlPath = decodeURIComponent(new URL(req.url, "http://x").pathname);
  const file = resolve(urlPath);
  if (!file || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
    res.writeHead(404);
    return res.end("Not found");
  }
  res.writeHead(200, {
    "Content-Type": TYPES[path.extname(file).toLowerCase()] || "application/octet-stream",
    "Cache-Control": "no-cache",
    "Access-Control-Allow-Origin": "*",
  });
  fs.createReadStream(file).pipe(res);
}

async function main() {
  if (process.argv.includes("--http")) {
    http.createServer(handler).listen(PORT, () => console.log(`http://localhost:${PORT}/taskpane.html`));
    return;
  }
  const devCerts = require("office-addin-dev-certs");
  const options = await devCerts.getHttpsServerOptions();
  https.createServer(options, handler).listen(PORT, () => console.log(`https://localhost:${PORT}/taskpane.html`));
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
