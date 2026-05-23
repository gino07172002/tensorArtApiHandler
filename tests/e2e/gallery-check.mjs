import { chromium } from "playwright";
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..", "..");
const MIME = { ".html": "text/html", ".js": "application/javascript", ".css": "text/css" };

const server = http.createServer((req, res) => {
  const url = req.url === "/" ? "/index.html" : req.url.split("?")[0];
  const file = path.join(root, url);
  fs.readFile(file, (err, data) => {
    if (err) { res.statusCode = 404; res.end(); return; }
    res.setHeader("Content-Type", MIME[path.extname(file)] || "text/plain");
    res.end(data);
  });
});
const port = 4568;
server.listen(port);

const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
page.on("console", (m) => { if (m.type() === "error") console.log("ERR:", m.text()); });
page.on("pageerror", (e) => console.log("PAGE ERR:", e.message));

// Inject a fake gallery item via localStorage
await page.goto(`http://localhost:${port}/gallery.html`);
await page.evaluate(() => {
  const stored = JSON.parse(localStorage.getItem("tensor-api-qa-console-v2") || "{}");
  stored.galleryItems = [{
    taskId: "fake-task-1",
    status: "SUCCESS",
    url: "https://placehold.co/200x200.png",
    generationImageId: "fake-img-1",
    metadata: {
      prompt: "test prompt",
      negativePrompt: "test neg",
      width: 1024,
      height: 768,
      seed: "12345",
      size: "1024x768",
      imageUrl: "https://placehold.co/200x200.png",
      imageId: "fake-img-1",
    },
  }];
  localStorage.setItem("tensor-api-qa-console-v2", JSON.stringify(stored));
});
await page.reload();
await page.waitForSelector(".gallery-card", { timeout: 5000 }).catch(() => {});
await page.waitForTimeout(300);

const cards = await page.locator(".gallery-card").count();
console.log("gallery cards:", cards);
const remixButtons = await page.locator(".remix-button").count();
console.log("remix buttons:", remixButtons);
const buttonText = await page.locator(".remix-button").allTextContents();
console.log("remix labels:", buttonText);

const html = await page.locator(".gallery-card").first().innerHTML().catch(() => "");
console.log("first card html (first 500ch):", html.slice(0, 500));

await page.locator(".gallery-card").first().scrollIntoViewIfNeeded();
await page.waitForTimeout(200);

const remixState = await page.evaluate(() => {
  const btns = document.querySelectorAll(".remix-button");
  return Array.from(btns).map((b) => {
    const cs = getComputedStyle(b);
    const r = b.getBoundingClientRect();
    return {
      text: b.textContent,
      display: cs.display,
      visibility: cs.visibility,
      opacity: cs.opacity,
      color: cs.color,
      bg: cs.backgroundColor,
      width: r.width,
      height: r.height,
      x: r.x,
      y: r.y,
      parentClass: b.parentElement?.className,
      parentDisplay: b.parentElement ? getComputedStyle(b.parentElement).display : "",
      parentRect: b.parentElement?.getBoundingClientRect().toJSON(),
    };
  });
});
console.log("remix state:", JSON.stringify(remixState, null, 2));

await page.screenshot({ path: path.join(root, "tests/e2e/gallery-shot.png"), fullPage: true });

await browser.close();
server.close();
