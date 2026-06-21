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

async function revealInpaintExperiment() {
  await page.locator("#canvas-inpaint-source-fold").evaluate((element) => {
    element.open = true;
  });
  await page.locator("#canvas-inpaint-source-fold").scrollIntoViewIfNeeded();
}

async function canvasClientPoint(x, y) {
  const metrics = await page.locator("#canvas-stage").evaluate((canvas) => {
    const rect = canvas.getBoundingClientRect();
    return {
      left: rect.left,
      top: rect.top,
      width: rect.width,
      height: rect.height,
      canvasWidth: canvas.width,
      canvasHeight: canvas.height,
    };
  });
  return {
    x: metrics.left + (x * metrics.width / metrics.canvasWidth),
    y: metrics.top + (y * metrics.height / metrics.canvasHeight),
  };
}

await page.goto(`http://localhost:${port}/canvas.html`);
await page.waitForSelector("#canvas-stage");
await revealInpaintExperiment();
await page.locator("#canvas-inpaint-powershell").fill(inpaintPowerShell);
await page.locator("#canvas-inpaint-parse").click();
await page.waitForFunction(() => document.querySelector("#canvas-body")?.value.includes("saved prompt"));
await page.locator("#canvas-body").fill(editedBody);

await page.locator('[data-canvas-layer="mask"]').click();
await page.locator('[data-canvas-tool="brush"]').click();
await page.locator("#canvas-stage").scrollIntoViewIfNeeded();
const brushStart = await canvasClientPoint(120, 120);
const brushEnd = await canvasClientPoint(220, 120);
await page.mouse.move(brushStart.x, brushStart.y);
await page.mouse.down();
await page.mouse.move(brushEnd.x, brushEnd.y);
await page.mouse.up();
await page.locator('[data-canvas-tool="eraser"]').click();
const eraseStart = await canvasClientPoint(170, 120);
const eraseEnd = await canvasClientPoint(172, 120);
await page.mouse.move(eraseStart.x, eraseStart.y);
await page.mouse.down();
await page.mouse.move(eraseEnd.x, eraseEnd.y);
await page.mouse.up();

await page.waitForFunction(() => window.__canvasEditorForTest?.maskCtx?.getImageData(120, 120, 1, 1).data[3] > 0);
const beforeReloadMask = await page.evaluate(() => ({
  edgeAlpha: window.__canvasEditorForTest.maskCtx.getImageData(120, 120, 1, 1).data[3],
  erasedAlpha: window.__canvasEditorForTest.maskCtx.getImageData(170, 120, 1, 1).data[3],
  layer: window.__canvasEditorForTest.layer,
  tool: window.__canvasEditorForTest.tool,
}));
assert.ok(beforeReloadMask.edgeAlpha > 0);
assert.equal(beforeReloadMask.erasedAlpha, 0);
assert.equal(beforeReloadMask.layer, "mask");
assert.equal(beforeReloadMask.tool, "eraser");

await page.reload();
await page.waitForSelector("#canvas-stage");
await revealInpaintExperiment();
await page.waitForFunction(() => window.__canvasEditorForTest?.maskCtx?.getImageData(120, 120, 1, 1).data[3] > 0);

assert.equal(await page.locator("#canvas-inpaint-powershell").inputValue(), inpaintPowerShell);
assert.deepEqual(JSON.parse(await page.locator("#canvas-inpaint-headers").inputValue()), {
  authorization: "Bearer canvas-token",
  "x-request-sign": "canvas-sig",
});
assert.equal(await page.locator("#canvas-body").inputValue(), editedBody);
const afterReloadMask = await page.evaluate(() => ({
  edgeAlpha: window.__canvasEditorForTest.maskCtx.getImageData(120, 120, 1, 1).data[3],
  erasedAlpha: window.__canvasEditorForTest.maskCtx.getImageData(170, 120, 1, 1).data[3],
  layer: window.__canvasEditorForTest.layer,
  tool: window.__canvasEditorForTest.tool,
}));
assert.ok(afterReloadMask.edgeAlpha > 0);
assert.equal(afterReloadMask.erasedAlpha, 0);
assert.equal(afterReloadMask.layer, "mask");
assert.equal(afterReloadMask.tool, "eraser");

await revealInpaintExperiment();
await page.locator("#canvas-inpaint-powershell").fill(inpaintPowerShell.replace("saved prompt", "reparsed prompt"));
await page.locator("#canvas-inpaint-parse").click();
await page.waitForFunction(() => document.querySelector("#canvas-body")?.value.includes("reparsed prompt"));
assert.match(await page.locator("#canvas-body").inputValue(), /reparsed prompt/);

await browser.close();
server.close();
console.log("canvas localStorage persistence test passed");
