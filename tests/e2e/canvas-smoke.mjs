import { chromium } from "playwright";
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
  ".json": "application/json",
  ".png": "image/png",
};

const server = http.createServer((req, res) => {
  const url = req.url === "/" ? "/canvas.html" : req.url.split("?")[0];
  const file = path.join(root, url);
  fs.readFile(file, (err, data) => {
    if (err) { res.statusCode = 404; res.end(); return; }
    res.setHeader("Content-Type", MIME[path.extname(file)] || "application/octet-stream");
    res.end(data);
  });
});

const port = 4567;
server.listen(port);

const issues = [];
function note(label, detail) {
  issues.push({ label, detail });
  console.log(`[issue] ${label}: ${detail}`);
}

const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
page.on("console", (msg) => { if (msg.type() === "error") console.log("PAGE ERR:", msg.text()); });
page.on("pageerror", (err) => console.log("PAGE PAGEERR:", err.message));

await page.goto(`http://localhost:${port}/canvas.html`);
await page.waitForSelector("#canvas-stage");
await page.waitForTimeout(200);

// Test 1: change API to portrait 1024x1536 - canvas should resize
await page.locator("#canvas-height").fill("1536");
await page.waitForTimeout(150);
const portrait = await page.locator("#canvas-stage").evaluate((el) => ({
  width: el.width, height: el.height,
  styleW: parseFloat(el.style.width), styleH: parseFloat(el.style.height),
}));
console.log("Portrait sync:", portrait);
if (portrait.height !== 1536) note("canvas.height not synced to 1536", `got ${portrait.height}`);

// Test 2: brush popover positioned to LEFT of toolbar (no canvas overlap)
await page.locator('[data-canvas-tool="brush"]').click();
await page.waitForTimeout(150);
const popoverGeom = await page.locator("#canvas-brush-popover").boundingBox();
const canvasGeom = await page.locator("#canvas-stage").boundingBox();
if (popoverGeom && canvasGeom && popoverGeom.x + popoverGeom.width > canvasGeom.x) {
  note("popover overlaps canvas horizontally", `pop right=${popoverGeom.x + popoverGeom.width} canvas left=${canvasGeom.x}`);
}

// Close popover
await page.locator("h2").first().click();
await page.waitForTimeout(150);

// Test 3: switch to mask mode - paint-only UI hides
await page.locator('[data-canvas-layer="mask"]').click();
await page.waitForTimeout(200);
const maskHidden = await page.evaluate(() => ({
  pick: getComputedStyle(document.querySelector("#canvas-pick")).display,
  layerPanel: getComputedStyle(document.querySelector(".layer-panel.paint-only")).display,
  moveBtn: getComputedStyle(document.querySelector('[data-canvas-tool="move"]')).display,
  clearPaint: getComputedStyle(document.querySelector("#canvas-clear-paint")).display,
  clearMask: getComputedStyle(document.querySelector("#canvas-clear-mask")).display,
  transformMini: getComputedStyle(document.querySelector(".transform-mini")).display,
}));
if (maskHidden.pick !== "none") note("pick still visible in mask mode", maskHidden.pick);
if (maskHidden.layerPanel !== "none") note("paint layer panel still visible in mask mode", maskHidden.layerPanel);
if (maskHidden.moveBtn !== "none") note("move tool still visible in mask mode", maskHidden.moveBtn);
if (maskHidden.clearPaint !== "none") note("clear-paint still visible in mask mode", maskHidden.clearPaint);
if (maskHidden.clearMask === "none") note("clear-mask hidden in mask mode", maskHidden.clearMask);
if (maskHidden.transformMini !== "none") note("transform-mini visible in mask mode", maskHidden.transformMini);

// Back to paint
await page.locator('[data-canvas-layer="paint"]').click();
await page.waitForTimeout(200);

// Test 4: Simulate canvasImport with explicit width/height -> canvas resizes to imported dims
await page.evaluate(() => {
  const stored = JSON.parse(localStorage.getItem("tensor-api-qa-console-v2") || "{}");
  stored.canvasImport = {
    mode: "img2img",
    imageUrl: "",
    width: 832,
    height: 1216,
    prompt: "test prompt",
    seed: "-1",
  };
  localStorage.setItem("tensor-api-qa-console-v2", JSON.stringify(stored));
});
await page.reload();
await page.waitForSelector("#canvas-stage");
await page.waitForTimeout(300);
const afterImport = await page.locator("#canvas-stage").evaluate((el) => ({
  width: el.width, height: el.height,
  inputW: document.querySelector("#canvas-width").value,
  inputH: document.querySelector("#canvas-height").value,
  badge: document.querySelector("#canvas-stage-size").textContent,
}));
console.log("After import (832x1216):", afterImport);
if (afterImport.width !== 832) note("import width not applied to canvas", `got ${afterImport.width}, input=${afterImport.inputW}`);
if (afterImport.height !== 1216) note("import height not applied to canvas", `got ${afterImport.height}, input=${afterImport.inputH}`);

// Test 5: Loading a local image file with no API dims should adopt image dims
// Reset state
await page.evaluate(() => localStorage.removeItem("tensor-api-qa-console-v2"));
await page.reload();
await page.waitForSelector("#canvas-stage");
await page.waitForTimeout(200);

// Use a known small PNG (1x1 pixel data URL would work but let's craft 600x400)
const pngBuffer = await page.evaluate(async () => {
  const c = document.createElement("canvas");
  c.width = 600;
  c.height = 400;
  const ctx = c.getContext("2d");
  ctx.fillStyle = "#abcdef";
  ctx.fillRect(0, 0, 600, 400);
  return new Promise((resolve) => {
    c.toBlob(async (blob) => {
      const buf = await blob.arrayBuffer();
      resolve(Array.from(new Uint8Array(buf)));
    }, "image/png");
  });
});

const tmpPath = path.join(root, "tests/e2e/tmp-600x400.png");
fs.writeFileSync(tmpPath, Buffer.from(pngBuffer));
await page.locator("#canvas-file").setInputFiles(tmpPath);
await page.waitForTimeout(400);
const afterLocal = await page.locator("#canvas-stage").evaluate((el) => ({
  width: el.width, height: el.height,
  inputW: document.querySelector("#canvas-width").value,
  inputH: document.querySelector("#canvas-height").value,
}));
console.log("After local 600x400 load:", afterLocal);
if (afterLocal.width !== 600) note("local image width not adopted to canvas", `got ${afterLocal.width}, input=${afterLocal.inputW}`);
if (afterLocal.height !== 400) note("local image height not adopted to canvas", `got ${afterLocal.height}, input=${afterLocal.inputH}`);
fs.unlinkSync(tmpPath);

await page.screenshot({ path: path.join(root, "tests/e2e/canvas-shot-final.png"), fullPage: false });

await browser.close();
server.close();

console.log("\n--- ISSUES ---");
for (const i of issues) console.log(`- ${i.label}: ${i.detail}`);
console.log(`Total: ${issues.length}`);
process.exit(issues.length > 0 ? 1 : 0);
