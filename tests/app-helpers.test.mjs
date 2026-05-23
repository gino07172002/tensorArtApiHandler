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
  buildCanvasEditorRequestDraft,
  createCanvasImageLoadAttempts,
  getCanvasStrokeMode,
  readCanvasColorAtPoint,
  applyParsedCanvasRequestToForm,
  storeCanvasParsedRequest,
  buildCanvasImportFromGalleryEntry,
  copyGenerationToSendBody,
  getState: () => state,
  parsePowerShellRequest,
  renderGallery,
  sanitizeHeaders,
};
`, context);

const {
  applyGenerationMetadataToBody,
  buildCanvasEditorRequestDraft,
  createCanvasImageLoadAttempts,
  getCanvasStrokeMode,
  readCanvasColorAtPoint,
  applyParsedCanvasRequestToForm,
  storeCanvasParsedRequest,
  buildCanvasImportFromGalleryEntry,
  copyGenerationToSendBody,
  getState,
  parsePowerShellRequest,
  renderGallery,
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

const parsedLooseRequest = parsePowerShellRequest(`
Invoke-WebRequest -UseBasicParsing \`
  -Uri 'https://api.tensor.art/works/v1/works/task_by_image' \`
  -WebSession $session \`
  -Headers @{
    "authorization"="Bearer token"
    "origin"="https://tensor.art"
    "x-request-sign"="sig-2"
  } \`
  -Method POST \`
  -ContentType "application/json" \`
  -Body '{"params":{"prompt":"ok"}}'
`);

assert.equal(parsedLooseRequest.url, "https://api.tensor.art/works/v1/works/task_by_image");
assert.equal(parsedLooseRequest.method, "POST");
assert.deepEqual(plain(parsedLooseRequest.headers), {
  authorization: "Bearer token",
  "x-request-sign": "sig-2",
});
assert.equal(parsedLooseRequest.bodyText, '{"params":{"prompt":"ok"}}');

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
assert.match(dom.stats.textContent, /seed/);

state.send.bodyText = sourceBody;
state.queryResultCopySeed = true;
copyGenerationToSendBody({ metadata }, dom);
assert.equal(JSON.parse(state.send.bodyText).params.seed, "987654321");
assert.match(dom.stats.textContent, /seed/);

state.galleryItems = [{
  taskId: "task-1",
  status: "FINISH",
  url: "https://example.test/image.png",
  generationImageId: "image-1",
  metadata,
}];
state.selectedImageIds = [];

const galleryRoot = {
  innerHTML: "",
  querySelectorAll: () => [],
};
renderGallery({
  stats: { textContent: "" },
  root: galleryRoot,
  selectedCount: [{ textContent: "" }],
});

const galleryHead = galleryRoot.innerHTML.match(/<div class="gallery-head">([\s\S]*?)<\/div>/)?.[1] || "";
const gallerySelection = galleryRoot.innerHTML.match(/<div class="gallery-selection">([\s\S]*?)<\/div>/)?.[1] || "";
assert.match(galleryHead, /Task task-1/);
assert.doesNotMatch(galleryHead, /data-action="copy-to-send"/);
assert.match(gallerySelection, /type="checkbox"/);
assert.match(gallerySelection, /data-image-id="image-1"/);
assert.match(gallerySelection, /加入 ID/);
assert.match(gallerySelection, /<details class="gallery-action-menu">/);
assert.match(gallerySelection, /<summary[^>]*>Actions<\/summary>/);
assert.match(gallerySelection, /data-action="copy-to-send"/);
assert.match(gallerySelection, /data-action="copy-to-send-seed"/);
assert.match(gallerySelection, /data-action="remix-inpaint"/);
assert.match(gallerySelection, /data-action="remix-img2img"/);
assert.match(gallerySelection, />Inpaint</);
assert.match(gallerySelection, />Img2Img</);
assert.match(gallerySelection, /class="gallery-open-button"/);
assert.match(gallerySelection, /data-action="open-original"/);
assert.match(gallerySelection, />Open</);
assert.doesNotMatch(gallerySelection, /remax/);
assert.doesNotMatch(galleryRoot.innerHTML, /複製到 API 1 Body/);

const canvasImport = buildCanvasImportFromGalleryEntry(state.galleryItems[0], "img2img");
assert.equal(canvasImport.mode, "img2img");
assert.equal(canvasImport.imageUrl, "https://example.test/image.png");
assert.equal(canvasImport.imageId, "image-1");
assert.equal(canvasImport.taskId, "task-1");
assert.equal(canvasImport.prompt, metadata.prompt);
assert.equal(canvasImport.negativePrompt, metadata.negativePrompt);
assert.equal(canvasImport.modelId, String(metadata.modelId));
assert.equal(canvasImport.modelFileId, String(metadata.modelFileId));

assert.deepEqual(plain(createCanvasImageLoadAttempts("https://image.tensorartassets.com/example.png")), [
  { url: "https://image.tensorartassets.com/example.png", crossOrigin: "anonymous" },
  { url: "https://image.tensorartassets.com/example.png", crossOrigin: "" },
]);
assert.equal(getCanvasStrokeMode({ layer: "paint", tool: "brush" }), "paint-add");
assert.equal(getCanvasStrokeMode({ layer: "paint", tool: "eraser" }), "paint-remove");
assert.equal(getCanvasStrokeMode({ layer: "mask", tool: "brush" }), "mask-add");
assert.equal(getCanvasStrokeMode({ layer: "mask", tool: "eraser" }), "mask-remove");
assert.equal(getCanvasStrokeMode({ layer: "mask", tool: "move" }), "none");
assert.equal(getCanvasStrokeMode({ layer: "mask", tool: "eyedropper" }), "none");
assert.deepEqual(plain(readCanvasColorAtPoint({
  getImageData: () => ({ data: [1, 35, 255, 255] }),
}, { x: 10, y: 20 })), { color: "#0123ff", error: "" });
const blockedColorRead = readCanvasColorAtPoint({
  getImageData: () => {
    const error = new Error("blocked");
    error.name = "SecurityError";
    throw error;
  },
}, { x: 10, y: 20 });
assert.equal(blockedColorRead.color, "");
assert.match(blockedColorRead.error, /CORS/);

const canvasFormDom = {
  mode: { value: "" },
  prompt: { value: "" },
  negativePrompt: { value: "" },
  width: { value: "" },
  height: { value: "" },
  steps: { value: "" },
  cfgScale: { value: "" },
  seed: { value: "" },
  denoisingStrength: { value: "" },
  modelId: { value: "" },
  modelFileId: { value: "" },
  body: { value: "" },
};
const canvasEditor = {
  imageDataUrl: "",
  sourceImageUrl: "",
  imageId: "",
  taskId: "",
  ksamplerName: "",
  schedule: "",
};
const parsedInpaint = parsePowerShellRequest(`
Invoke-WebRequest -Uri "https://api.tensor.art/works/v1/works/task_by_image" \`
  -Method "POST" \`
  -Headers @{
    "authorization"="Bearer token"
    "origin"="https://tensor.art"
    "x-request-sign"="sig"
  } \`
  -ContentType "application/json" \`
  -Body '{"params":{"baseModel":{"modelId":"111","modelFileId":"222"},"prompt":"p","negativePrompt":"n","height":1536,"width":1024,"steps":25,"cfgScale":5,"seed":"-1","denoisingStrength":0.75,"ksamplerName":"euler_ancestral","schedule":"normal","images":["https://example.test/base.jpg"],"inpaint":{"maskImage":"https://example.test/mask.png","maskBlur":4}},"taskType":"IMAGE_TO_INPAINT","imageUrl":"https://example.test/base.jpg","imageId":"image-9"}'
`);
applyParsedCanvasRequestToForm(parsedInpaint, canvasEditor, canvasFormDom);
assert.equal(parsedInpaint.headers.origin, undefined);
assert.equal(canvasFormDom.mode.value, "inpaint");
assert.equal(canvasFormDom.prompt.value, "p");
assert.equal(canvasFormDom.negativePrompt.value, "n");
assert.equal(canvasFormDom.width.value, "1024");
assert.equal(canvasFormDom.height.value, "1536");
assert.equal(canvasFormDom.modelId.value, "111");
assert.equal(canvasFormDom.modelFileId.value, "222");
assert.equal(canvasEditor.sourceImageUrl, "https://example.test/base.jpg");
assert.equal(canvasEditor.imageId, "image-9");
assert.equal(JSON.parse(canvasFormDom.body.value).params.inpaint.maskImage, "https://example.test/mask.png");
state.send.bodyText = '{"api1":true}';
state.send.headers = { "x-api-1": "keep" };
storeCanvasParsedRequest(parsedInpaint);
assert.equal(state.send.bodyText, '{"api1":true}');
assert.deepEqual(plain(state.send.headers), { "x-api-1": "keep" });
assert.equal(state.canvasRequest.url, "https://api.tensor.art/works/v1/works/task_by_image");
assert.equal(JSON.parse(state.canvasRequest.bodyText).taskType, "IMAGE_TO_INPAINT");

const canvasBase = {
  prompt: "portrait study",
  negativePrompt: "low quality",
  imageDataUrl: "data:image/png;base64,base",
  maskDataUrl: "data:image/png;base64,mask",
  width: 768,
  height: 1024,
  steps: 25,
  cfgScale: 5,
  seed: "-1",
  denoisingStrength: 0.51,
  guidance: 3.5,
  clipSkip: 2,
  modelId: "990778216270553015",
  modelFileId: "990778216270553016",
  ksamplerName: "euler_ancestral",
  schedule: "normal",
};

const inpaintDraft = buildCanvasEditorRequestDraft({ ...canvasBase, mode: "inpaint" });
assert.equal(inpaintDraft.url, "https://api.tensor.art/works/v1/works/task_by_image");
assert.equal(inpaintDraft.method, "POST");
assert.equal(inpaintDraft.contentType, "application/json");
assert.equal(inpaintDraft.body.taskType, "IMAGE_TO_INPAINT");
assert.equal(inpaintDraft.body.params.inpaint.maskImage, canvasBase.maskDataUrl);
assert.deepEqual(plain(inpaintDraft.body.params.images), [canvasBase.imageDataUrl]);

const img2imgDraft = buildCanvasEditorRequestDraft({ ...canvasBase, mode: "img2img" });
assert.equal(img2imgDraft.url, "https://api.tensor.art/works/v1/works/task");
assert.equal(img2imgDraft.contentType, "text/plain;charset=UTF-8");
assert.equal(img2imgDraft.body.taskType, "IMG2IMG");
assert.equal(img2imgDraft.body.isRemix, true);
assert.equal(img2imgDraft.body.captchaType, "CLOUDFLARE_TURNSTILE");
assert.equal(img2imgDraft.body.params.inpaint, undefined);

console.log("app helper tests passed");
