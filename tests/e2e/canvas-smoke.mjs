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

// Test 1: change API to portrait 1024x1536 - canvas should resize and fit larger
await page.locator("#canvas-height").fill("1536");
await page.waitForTimeout(150);
const portrait = await page.locator("#canvas-stage").evaluate((el) => ({
  width: el.width, height: el.height,
  styleW: parseFloat(el.style.width), styleH: parseFloat(el.style.height),
}));
console.log("Portrait sync:", portrait, "badge:", await page.locator("#canvas-stage-size").textContent());
if (portrait.height !== 1536) note("canvas.height not synced to 1536", `got ${portrait.height}`);
if (portrait.styleH < portrait.styleW * 1.2) note("portrait display not tall enough", JSON.stringify(portrait));

// Test 2: brush popover doesn't cover the canvas
await page.locator('[data-canvas-tool="brush"]').click();
await page.waitForTimeout(150);
const popoverGeom = await page.locator("#canvas-brush-popover").boundingBox();
const canvasGeom = await page.locator("#canvas-stage").boundingBox();
console.log("Popover:", popoverGeom, "Canvas:", canvasGeom);
if (popoverGeom && canvasGeom) {
  const xOverlap = popoverGeom.x < canvasGeom.x + canvasGeom.width && popoverGeom.x + popoverGeom.width > canvasGeom.x;
  const yOverlap = popoverGeom.y < canvasGeom.y + canvasGeom.height && popoverGeom.y + popoverGeom.height > canvasGeom.y;
  if (xOverlap && yOverlap) note("popover still overlaps canvas", `popover x=${popoverGeom.x} canvas x=${canvasGeom.x}`);
}

// Close popover by clicking elsewhere
await page.locator("h2").first().click();
await page.waitForTimeout(150);

// Test 3: switch to mask mode - paint-only UI should hide
await page.locator('[data-canvas-layer="mask"]').click();
await page.waitForTimeout(200);
const maskHidden = await page.evaluate(() => {
  const pick = document.querySelector("#canvas-pick");
  const colorRow = document.querySelector(".popover-row.paint-only");
  const layerPanel = document.querySelector(".layer-panel.paint-only");
  const moveBtn = document.querySelector('[data-canvas-tool="move"]');
  const clearPaint = document.querySelector("#canvas-clear-paint");
  const clearMask = document.querySelector("#canvas-clear-mask");
  const transformMini = document.querySelector(".transform-mini");
  return {
    pick: pick ? getComputedStyle(pick).display : "missing",
    layerPanel: layerPanel ? getComputedStyle(layerPanel).display : "missing",
    moveBtn: moveBtn ? getComputedStyle(moveBtn).display : "missing",
    clearPaint: clearPaint ? getComputedStyle(clearPaint).display : "missing",
    clearMask: clearMask ? getComputedStyle(clearMask).display : "missing",
    transformMini: transformMini ? getComputedStyle(transformMini).display : "missing",
  };
});
console.log("Mask mode visibility:", maskHidden);
if (maskHidden.pick !== "none") note("pick still visible in mask mode", maskHidden.pick);
if (maskHidden.layerPanel !== "none") note("paint layer panel still visible in mask mode", maskHidden.layerPanel);
if (maskHidden.moveBtn !== "none") note("move tool still visible in mask mode", maskHidden.moveBtn);
if (maskHidden.clearPaint !== "none") note("clear-paint still visible in mask mode", maskHidden.clearPaint);
if (maskHidden.clearMask === "none") note("clear-mask hidden in mask mode", maskHidden.clearMask);
if (maskHidden.transformMini !== "none") note("transform-mini visible in mask mode", maskHidden.transformMini);

// Test 4: switch back to paint - paint UI returns
await page.locator('[data-canvas-layer="paint"]').click();
await page.waitForTimeout(200);
const paintReturn = await page.evaluate(() => ({
  pick: getComputedStyle(document.querySelector("#canvas-pick")).display,
  moveBtn: getComputedStyle(document.querySelector('[data-canvas-tool="move"]')).display,
  layerPanel: getComputedStyle(document.querySelector(".layer-panel.paint-only")).display,
  clearMask: getComputedStyle(document.querySelector("#canvas-clear-mask")).display,
}));
console.log("Paint mode visibility:", paintReturn);
if (paintReturn.pick === "none") note("pick hidden in paint mode", paintReturn.pick);
if (paintReturn.moveBtn === "none") note("move hidden in paint mode", paintReturn.moveBtn);
if (paintReturn.layerPanel === "none") note("layer panel hidden in paint mode", paintReturn.layerPanel);
if (paintReturn.clearMask !== "none") note("clear-mask visible in paint mode", paintReturn.clearMask);

// Test 5: dimmed stage frame visible around canvas
const frameContrast = await page.evaluate(() => {
  const frame = document.querySelector(".stage-frame");
  const host = document.querySelector(".editor-canvas-host");
  return {
    framePadding: getComputedStyle(frame).padding,
    frameBg: getComputedStyle(frame).backgroundColor,
    frameRect: frame.getBoundingClientRect().toJSON(),
    hostRect: host?.getBoundingClientRect().toJSON(),
  };
});
console.log("Stage frame:", frameContrast);

// Test 6: layer add/remove in paint mode
const layersBefore = await page.locator("#canvas-layer-list li").count();
await page.locator("#canvas-layer-add").click();
await page.waitForTimeout(100);
const layersAfter = await page.locator("#canvas-layer-list li").count();
console.log("Layers before/after add:", layersBefore, layersAfter);
if (layersAfter !== layersBefore + 1) note("add layer failed", `${layersBefore} -> ${layersAfter}`);

await page.screenshot({ path: path.join(root, "tests/e2e/canvas-shot-paint.png"), fullPage: false });

// Test 7: switch to mask + screenshot
await page.locator('[data-canvas-layer="mask"]').click();
await page.waitForTimeout(200);
await page.screenshot({ path: path.join(root, "tests/e2e/canvas-shot-mask.png"), fullPage: false });

await browser.close();
server.close();

console.log("\n--- ISSUES ---");
for (const i of issues) console.log(`- ${i.label}: ${i.detail}`);
console.log(`Total: ${issues.length}`);
process.exit(issues.length > 0 ? 1 : 0);
