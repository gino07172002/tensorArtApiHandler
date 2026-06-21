import { chromium } from 'playwright';

const browser = await chromium.launch({ headless: true });
const ctx = await browser.newContext();
const page = await ctx.newPage();

const consoleErrors = [];
page.on('console', msg => {
  if (msg.type() === 'error') consoleErrors.push(msg.text());
  if (msg.type() === 'warn') console.log('[WARN]', msg.text());
});
page.on('pageerror', err => consoleErrors.push('PAGE ERROR: ' + err.message));

// 1. Open canvas page directly
await page.goto('https://gino07172002.github.io/tensorArtApiHandler/canvas.html', { waitUntil: 'networkidle' });
await page.screenshot({ path: 'd:/tensorArt/__ss1-canvas.png', fullPage: true });
console.log('=== URL:', page.url());

// 2. Find all buttons and selects
const elements = await page.evaluate(() => {
  const result = [];
  document.querySelectorAll('button, select, input[type="radio"], [data-mode]').forEach(el => {
    result.push({
      tag: el.tagName,
      id: el.id,
      text: el.innerText?.trim().substring(0, 60),
      value: el.value,
      dataMode: el.dataset?.mode,
      name: el.name,
      type: el.type,
    });
  });
  return result;
});
console.log('=== INTERACTIVE ELEMENTS:');
elements.forEach(e => console.log(JSON.stringify(e)));

// 3. Check for mode selector (inpaint mode)
const modeSelect = await page.locator('select[id*="mode"], select[name*="mode"], [data-mode]').first();
const modeExists = await modeSelect.count();
console.log('=== MODE ELEMENT EXISTS:', modeExists);

if (modeExists) {
  const modeVal = await modeSelect.evaluate(el => el.value || el.dataset.mode);
  console.log('=== CURRENT MODE:', modeVal);
}

// 4. Look for inpaint-related UI
const inpaintSection = await page.evaluate(() => {
  const els = [];
  document.querySelectorAll('[class*="inpaint"], [id*="inpaint"], [data-mode="inpaint"]').forEach(el => {
    els.push({ tag: el.tagName, id: el.id, class: el.className, text: el.innerText?.substring(0,100) });
  });
  return els;
});
console.log('=== INPAINT ELEMENTS:', JSON.stringify(inpaintSection, null, 2));

// 5. Find Send button
const sendButtons = await page.evaluate(() => {
  const btns = [];
  document.querySelectorAll('button').forEach(b => {
    if (b.innerText.match(/send|送出|Submit/i)) {
      btns.push({ text: b.innerText.trim(), id: b.id, class: b.className, dataset: JSON.stringify(b.dataset) });
    }
  });
  return btns;
});
console.log('=== SEND BUTTONS:', JSON.stringify(sendButtons));

// 6. Print HTML structure around the main controls
const mainHtml = await page.evaluate(() => {
  const main = document.querySelector('main, .canvas-shell, .app-shell, body');
  return main ? main.innerHTML.substring(0, 5000) : 'not found';
});
console.log('=== MAIN HTML (5000):', mainHtml);

await browser.close();
console.log('=== CONSOLE ERRORS:', JSON.stringify(consoleErrors, null, 2));
