// Servidor HTTPS local mínimo para servir el complemento (solo módulos nativos de Node).
const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");

const PORT = 3000;
const ROOT = path.join(__dirname, "public");
const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".json": "application/json",
};

let options;
try {
  options = {
    key: fs.readFileSync(path.join(CERT_DIR, "localhost.key")),
    cert: fs.readFileSync(path.join(CERT_DIR, "localhost.crt")),
  };
} catch (e) {
  console.error("No se encontraron los certificados. Ejecuta primero:\n  npx office-addin-dev-certs install");
  process.exit(1);
}

https
  .createServer(options, (req, res) => {
    let urlPath;
    try {
      urlPath = decodeURIComponent(req.url.split("?")[0]);
    } catch (e) {
      res.writeHead(400);
      return res.end("Solicitud inválida");
    }
    if (urlPath === "/") urlPath = "/taskpane.html";
    const file = path.normalize(path.join(ROOT, urlPath));
    if (!file.startsWith(ROOT + path.sep)) {
      res.writeHead(403);
      return res.end("Prohibido");
    }
    fs.readFile(file, (err, data) => {
      if (err) {
        res.writeHead(404);
        return res.end("No encontrado");
      }
      res.writeHead(200, {
        "Content-Type": TYPES[path.extname(file)] || "application/octet-stream",
        "Cache-Control": "no-store",
      });
      res.end(data);
    });
  })
  .on("error", (e) => {
    if (e.code === "EADDRINUSE") {
      console.error(`El puerto ${PORT} ya está en uso. Cierra el otro programa (o el otro "npm start") y vuelve a intentar.`);
    } else {
      console.error(e.message);
    }
    process.exit(1);
  })
  .listen(PORT, "localhost", () => console.log(`Complemento sirviendo en https://localhost:${PORT}/taskpane.html`));
