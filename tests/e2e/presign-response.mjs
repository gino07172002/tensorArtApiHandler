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

const port = 4568;
await new Promise((resolve) => server.listen(port, resolve));

const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });

const responseBody = {
  data: {
    uploadUrl: "https://upload.example.test/full-mask-body.jpeg",
    displayUrl: "https://display.example.test/full-mask-body.jpeg",
    token: "line-check",
  },
};

await page.route("https://api.tensor.art/community-web/v1/cloudflare/upload/pre_sign", async (route) => {
  await route.fulfill({
    status: 200,
    headers: {
      "access-control-allow-origin": `http://localhost:${port}`,
      "access-control-allow-credentials": "true",
      "content-type": "application/json",
    },
    body: JSON.stringify(responseBody),
  });
});

await page.goto(`http://localhost:${port}/canvas.html`);
await page.waitForSelector("#presign-request");

await page.locator("#presign-source-summary").click();
await page.locator("#presign-url").fill("https://api.tensor.art/community-web/v1/cloudflare/upload/pre_sign");
await page.locator("#presign-method").fill("POST");
await page.locator("#presign-headers").fill(JSON.stringify({ authorization: "Bearer test" }, null, 2));
await page.locator("#presign-body").fill(JSON.stringify({ scene: "IMAGE_TO_IMAGE", fileNameSuffix: "jpeg" }, null, 2));
await page.locator("#presign-request").click();
await page.waitForFunction(() => document.querySelector("#presign-response")?.textContent.includes("full-mask-body.jpeg"));

const isOpen = await page.locator("#presign-response-fold").evaluate((el) => el.open);
const rendered = await page.locator("#presign-response").textContent();
const initialView = await page.locator("#presign-response-fold").getAttribute("data-response-view");
const initialBox = await page.locator("#presign-response").boundingBox();

assert.equal(isOpen, true);
assert.equal(initialView, "preview");
assert.match(rendered, /HTTP 200/);
assert.match(rendered, /full-mask-body\.jpeg/);
assert.match(rendered, /--- pre_sign upload fields ---/);
assert.match(rendered, /Use PUT with image\/jpeg mask blob/);
assert.ok(initialBox.height < 140, `preview should stay near five lines, got ${initialBox.height}`);

await page.locator('[data-response-action="expand"]').click();
const expandedView = await page.locator("#presign-response-fold").getAttribute("data-response-view");
const expandedBox = await page.locator("#presign-response").boundingBox();
assert.equal(expandedView, "full");
assert.ok(expandedBox.height > initialBox.height, `expanded body should be taller than preview: ${expandedBox.height} <= ${initialBox.height}`);

await page.locator('[data-response-action="collapse"]').click();
const collapsedView = await page.locator("#presign-response-fold").getAttribute("data-response-view");
const collapsedBox = await page.locator("#presign-response").boundingBox();
assert.equal(collapsedView, "preview");
assert.ok(collapsedBox.height <= initialBox.height + 2, `collapsed body should return to five-line preview: ${collapsedBox.height} > ${initialBox.height}`);

await page.locator('[data-response-action="expand"]').click();
assert.equal(await page.locator("#presign-response-fold").getAttribute("data-response-view"), "full");
await page.locator("#presign-request").click();
await page.waitForFunction(() => document.querySelector("#presign-response")?.textContent.includes("full-mask-body.jpeg"));
const afterRequestView = await page.locator("#presign-response-fold").getAttribute("data-response-view");
const afterRequestBox = await page.locator("#presign-response").boundingBox();
assert.equal(afterRequestView, "preview");
assert.ok(afterRequestBox.height <= initialBox.height + 2, `request should reset response to five-line preview: ${afterRequestBox.height} > ${initialBox.height}`);

await browser.close();
server.close();
console.log("presign expanded response body test passed");
