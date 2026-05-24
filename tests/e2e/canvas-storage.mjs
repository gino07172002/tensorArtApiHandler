import { chromium } from "playwright";
import assert from "node:assert/strict";
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..", "..");

const MIME = {
  ".html": "text/html",
  ".js": "application/javascript",
  ".css": "text/css",
};

const server = http.createServer((req, res) => {
  const url = req.url === "/" ? "/canvas.html" : req.url.split("?")[0];
  const file = path.join(root, url);
  fs.readFile(file, (err, data) => {
    if (err) {
      res.statusCode = 404;
      res.end();
      return;
    }
    res.setHeader("Content-Type", MIME[path.extname(file)] || "application/octet-stream");
    res.end(data);
  });
});

const port = 4569;
await new Promise((resolve) => server.listen(port, resolve));

const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });

const inpaintPowerShell = `Invoke-WebRequest -Uri "https://api.tensor.art/works/v1/works/task_by_image" \`
  -Method "POST" \`
  -Headers @{
    "authorization"="Bearer canvas-token"
    "origin"="https://tensor.art"
    "x-request-sign"="canvas-sig"
  } \`
  -ContentType "application/json" \`
  -Body '{"params":{"baseModel":{"modelId":"111","modelFileId":"222"},"prompt":"saved prompt","negativePrompt":"saved negative","height":768,"width":1024,"steps":25,"cfgScale":5,"seed":"-1","denoisingStrength":0.75,"images":["https://example.test/base.jpg"],"inpaint":{"maskImage":"https://example.test/mask.png","maskBlur":4}},"taskType":"IMAGE_TO_INPAINT","imageUrl":"https://example.test/base.jpg","imageId":"image-9"}'`;

const editedBody = JSON.stringify({
  taskType: "IMAGE_TO_INPAINT",
  imageUrl: "https://example.test/base.jpg",
  imageId: "image-9",
  params: {
    prompt: "manual body survives reload",
    images: ["https://example.test/base.jpg"],
    inpaint: { maskImage: "https://example.test/manual-mask.png" },
  },
}, null, 2);

await page.goto(`http://localhost:${port}/canvas.html`);
await page.waitForSelector("#canvas-stage");
await page.locator("#canvas-inpaint-source-summary").click();
await page.locator("#canvas-inpaint-powershell").fill(inpaintPowerShell);
await page.locator("#canvas-inpaint-parse").click();
await page.waitForFunction(() => document.querySelector("#canvas-body")?.value.includes("saved prompt"));
await page.locator("#canvas-body").fill(editedBody);

await page.reload();
await page.waitForSelector("#canvas-stage");

assert.equal(await page.locator("#canvas-inpaint-powershell").inputValue(), inpaintPowerShell);
assert.deepEqual(JSON.parse(await page.locator("#canvas-inpaint-headers").inputValue()), {
  authorization: "Bearer canvas-token",
  "x-request-sign": "canvas-sig",
});
assert.equal(await page.locator("#canvas-body").inputValue(), editedBody);

await page.locator("#canvas-inpaint-source-summary").click();
await page.locator("#canvas-inpaint-powershell").fill(inpaintPowerShell.replace("saved prompt", "reparsed prompt"));
await page.locator("#canvas-inpaint-parse").click();
await page.waitForFunction(() => document.querySelector("#canvas-body")?.value.includes("reparsed prompt"));
assert.match(await page.locator("#canvas-body").inputValue(), /reparsed prompt/);

await browser.close();
server.close();
console.log("canvas localStorage persistence test passed");
