import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";

const source = fs.readFileSync(new URL("../app.js", import.meta.url), "utf8");

const localStorageStore = new Map();
const context = {
  Blob,
  URL: { createObjectURL: () => "blob:test", revokeObjectURL: () => {} },
  document: {
    body: { dataset: { page: "" }, append() {} },
    createElement: () => ({ click() {}, remove() {} }),
    querySelector: () => null,
    querySelectorAll: () => [],
  },
  localStorage: {
    getItem: (key) => localStorageStore.get(key) ?? null,
    setItem: (key, value) => localStorageStore.set(key, String(value)),
    removeItem: (key) => localStorageStore.delete(key),
  },
  location: { reload() {} },
  console,
};

vm.createContext(context);
vm.runInContext(`${source}
globalThis.__testExports = {
  applyGenerationMetadataToBody,
  copyGenerationToSendBody,
  getState: () => state,
  parsePowerShellRequest,
  sanitizeHeaders,
};
`, context);

const {
  applyGenerationMetadataToBody,
  copyGenerationToSendBody,
  getState,
  parsePowerShellRequest,
  sanitizeHeaders,
} = context.__testExports;

const plain = (value) => JSON.parse(JSON.stringify(value));

assert.deepEqual(plain(sanitizeHeaders({
  accept: "application/json",
  cookie: "secret",
  "content-length": "999",
  "x-request-sign": "sig-1",
})), {
  accept: "application/json",
  "x-request-sign": "sig-1",
});

const parsedRequest = parsePowerShellRequest(`
Invoke-WebRequest -Uri "https://api.tensor.art/works/v1/works/tasks/query" \`
  -Method "POST" \`
  -Headers @{
    "accept"="application/json"
    "cookie"="secret"
  } \`
  -ContentType "application/json" \`
  -Body '{"size":20}'
`);

assert.equal(parsedRequest.url, "https://api.tensor.art/works/v1/works/tasks/query");
assert.equal(parsedRequest.method, "POST");
assert.deepEqual(plain(parsedRequest.headers), { accept: "application/json" });
assert.equal(parsedRequest.bodyText, '{"size":20}');

const sourceBody = JSON.stringify({
  params: {
    prompt: "old prompt",
    seed: "12345",
    steps: 12,
    keepMe: true,
  },
  projectId: "project-1",
});

const metadata = {
  prompt: "new prompt",
  negativePrompt: "low quality",
  seed: "987654321",
  steps: "30",
  cfgScale: "7.5",
  guidance: "3.5",
  clipSkip: "2",
  width: "768",
  height: "1024",
  vae: "Automatic",
  sampler: "Euler a",
  schedule: "Karras",
  denoisingStrength: "0.45",
  modelId: 111,
  modelFileId: 222,
};

const copiedWithoutSeed = JSON.parse(applyGenerationMetadataToBody(sourceBody, metadata, { includeSeed: false }));

assert.equal(copiedWithoutSeed.projectId, "project-1");
assert.equal(copiedWithoutSeed.params.keepMe, true);
assert.equal(copiedWithoutSeed.params.prompt, "new prompt");
assert.equal(copiedWithoutSeed.params.negativePrompt, "low quality");
assert.equal(copiedWithoutSeed.params.seed, "-1");
assert.equal(copiedWithoutSeed.params.steps, 30);
assert.equal(copiedWithoutSeed.params.cfgScale, 7.5);
assert.equal(copiedWithoutSeed.params.guidance, 3.5);
assert.equal(copiedWithoutSeed.params.clipSkip, 2);
assert.equal(copiedWithoutSeed.params.width, 768);
assert.equal(copiedWithoutSeed.params.height, 1024);
assert.equal(copiedWithoutSeed.params.sdVae, "Automatic");
assert.equal(copiedWithoutSeed.params.ksamplerName, "Euler a");
assert.equal(copiedWithoutSeed.params.schedule, "Karras");
assert.equal(copiedWithoutSeed.params.denoisingStrength, 0.45);
assert.deepEqual(copiedWithoutSeed.params.baseModel, {
  modelId: "111",
  modelFileId: "222",
});

const copiedWithSeed = JSON.parse(applyGenerationMetadataToBody(sourceBody, metadata, { includeSeed: true }));
assert.equal(copiedWithSeed.params.seed, "987654321");

const state = getState();
const dom = { stats: { textContent: "" }, sendSection: null };
state.send.bodyText = sourceBody;
state.queryResultCopySeed = false;
copyGenerationToSendBody({ metadata }, dom);
assert.equal(JSON.parse(state.send.bodyText).params.seed, "-1");
assert.match(dom.stats.textContent, /seed 已設為 -1/);

state.send.bodyText = sourceBody;
state.queryResultCopySeed = true;
copyGenerationToSendBody({ metadata }, dom);
assert.equal(JSON.parse(state.send.bodyText).params.seed, "987654321");
assert.match(dom.stats.textContent, /包含 seed/);

console.log("app helper tests passed");
