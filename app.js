const STORAGE_KEY = "tensor-api-qa-console-v2";
const PNG_SIGNATURE = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
const JPEG_SOI = [0xff, 0xd8];
const FORBIDDEN_HEADERS = new Set([
  "accept-encoding",
  "authority",
  "connection",
  "content-length",
  "cookie",
  "host",
  "method",
  "origin",
  "path",
  "priority",
  "referer",
  "scheme",
  "sec-ch-ua",
  "sec-ch-ua-mobile",
  "sec-ch-ua-platform",
  "sec-fetch-dest",
  "sec-fetch-mode",
  "sec-fetch-site",
]);

const state = loadState();
const page = document.body.dataset.page;

bindCommonControls();

if (page === "send") {
  initSendPage();
}

if (page === "dashboard") {
  initDashboardPage();
}

if (page === "gallery") {
  initGalleryPage();
}

if (page === "metadata") {
  initMetadataPage();
}

if (page === "canvas") {
  initCanvasPage();
}

function blankRequestState() {
  return {
    powershell: "",
    url: "",
    method: "POST",
    headers: {},
    bodyText: "",
    responseText: "",
    clearOnSubmit: false,
    presetId: null,
  };
}

function loadState() {
  const base = {
    send: blankRequestState(),
    query: blankRequestState(),
    post: blankRequestState(),
    selectedImageIds: [],
    galleryItems: [],
    queryResultCopySeed: false,
    canvasImport: null,
    importedMetadata: null,
    savedSettings: { send: [], query: [], post: [] },
  };

  const saved = localStorage.getItem(STORAGE_KEY);
  if (!saved) return base;

  try {
    const parsed = JSON.parse(saved);
    return {
      send: { ...blankRequestState(), ...(parsed.send || {}) },
      query: { ...blankRequestState(), ...(parsed.query || {}) },
      post: { ...blankRequestState(), ...(parsed.post || {}) },
      selectedImageIds: Array.isArray(parsed.selectedImageIds) ? parsed.selectedImageIds : [],
      galleryItems: Array.isArray(parsed.galleryItems) ? parsed.galleryItems : [],
      queryResultCopySeed: Boolean(parsed.queryResultCopySeed),
      canvasImport: parsed.canvasImport ?? null,
      importedMetadata: parsed.importedMetadata ?? null,
      savedSettings: parsed.savedSettings || { send: [], query: [], post: [] },
    };
  } catch {
    return base;
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function cloneRequestDraft(request) {
  return { ...blankRequestState(), ...JSON.parse(JSON.stringify(request || {})) };
}

function normalizeSavedSettings(savedSettings) {
  const source = savedSettings || {};
  return {
    send: Array.isArray(source.send) ? source.send : [],
    query: Array.isArray(source.query) ? source.query : [],
    post: Array.isArray(source.post) ? source.post : [],
  };
}

function buildSnapshotData(source) {
  return {
    send: cloneRequestDraft(source.send),
    query: cloneRequestDraft(source.query),
    post: cloneRequestDraft(source.post),
    selectedImageIds: Array.isArray(source.selectedImageIds) ? [...source.selectedImageIds] : [],
    galleryItems: Array.isArray(source.galleryItems) ? JSON.parse(JSON.stringify(source.galleryItems)) : [],
    queryResultCopySeed: Boolean(source.queryResultCopySeed),
    canvasImport: source.canvasImport ?? null,
    importedMetadata: source.importedMetadata ?? null,
    savedSettings: normalizeSavedSettings(source.savedSettings),
  };
}

function bindCommonControls() {
  const exportButton = document.querySelector("#export-storage");
  const importButton = document.querySelector("#import-storage");
  const importFile = document.querySelector("#import-storage-file");
  const clearButton = document.querySelector("#clear-storage");

  exportButton?.addEventListener("click", exportStorageSnapshot);
  importButton?.addEventListener("click", () => importFile?.click());
  importFile?.addEventListener("change", importStorageSnapshot);
  clearButton?.addEventListener("click", clearStorage);
}

function exportStorageSnapshot() {
  const data = buildSnapshotData(state);
  const payload = {
    exportedAt: new Date().toISOString(),
    storageKey: STORAGE_KEY,
    snapshotVersion: 2,
    data,
  };
  triggerDownload(new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" }), `tensor-api-qa-${timestamp()}.json`);
}

async function importStorageSnapshot(event) {
  const [file] = event.target.files || [];
  if (!file) return;

  try {
    const parsed = JSON.parse(await file.text());
    const incoming = buildSnapshotData(parsed.data ?? parsed);
    Object.assign(state.send, incoming.send);
    Object.assign(state.query, incoming.query);
    Object.assign(state.post, incoming.post);
    state.selectedImageIds = incoming.selectedImageIds;
    state.galleryItems = incoming.galleryItems;
    state.queryResultCopySeed = incoming.queryResultCopySeed;
    state.canvasImport = incoming.canvasImport;
    state.importedMetadata = incoming.importedMetadata;
    state.savedSettings = incoming.savedSettings;
    saveState();
    location.reload();
  } finally {
    event.target.value = "";
  }
}

function clearStorage() {
  localStorage.removeItem(STORAGE_KEY);
  location.reload();
}

function initSendPage() {
  const section = bindRequestSection("send");
  renderRequestSection("send", section);
  initializeRequestSectionVisibility("send", section);

  section.parseButton.addEventListener("click", () => {
    try {
      Object.assign(state.send, parsePowerShellRequest(section.powershell.value));
      state.send.bodyText = formatJsonString(state.send.bodyText);
      saveState();
      renderRequestSection("send", section);
      collapseRequestSection(section);
    } catch (error) {
      setResponse("send", `解析失敗: ${error.message}`);
      renderRequestSection("send", section);
    }
  });

  section.requestButton.addEventListener("click", async () => {
    await submitRequestSection("send", section, false);
  });

  section.formatButton.addEventListener("click", () => {
    state.send.bodyText = formatJsonString(section.body.value);
    saveState();
    renderRequestSection("send", section);
  });

  bindSectionInputs("send", section);
  bindPresetControls("send", section);
}

function initDashboardPage() {
  const send = bindRequestSection("send");
  const query = bindRequestSection("query");
  const post = bindRequestSection("post");
  const gallery = {
    stats: document.querySelector("#gallery-stats"),
    root: document.querySelector("#gallery"),
    selectedCount: document.querySelectorAll(".selected-count"),
    postClearIds: document.querySelectorAll(".post-clear-ids"),
    copySeed: document.querySelector("#gallery-copy-seed"),
    sendSection: send,
  };

  bindGalleryOptions(gallery);
  renderRequestSection("send", send);
  renderRequestSection("query", query);
  renderRequestSection("post", post);
  renderGallery(gallery);
  initializeRequestSectionVisibility("send", send);
  initializeRequestSectionVisibility("query", query);
  initializeRequestSectionVisibility("post", post);

  bindSectionInputs("send", send);
  bindSectionInputs("query", query);
  bindSectionInputs("post", post);
  
  bindPresetControls("send", send);
  bindPresetControls("query", query);
  bindPresetControls("post", post);

  bindParseAction("send", send, () => {});
  bindParseAction("query", query, () => {});
  bindParseAction("post", post, () => {
    syncGenerationImageIds();
    renderRequestSection("post", post);
    renderGallery(gallery);
  });

  send.requestButton.addEventListener("click", async () => {
    await submitRequestSection("send", send, false);
  });
  query.requestButton.addEventListener("click", async () => {
    await submitRequestSection("query", query, true);
    renderGallery(gallery);
  });
  post.requestButton.addEventListener("click", async () => {
    syncGenerationImageIds();
    renderRequestSection("post", post);
    const status = await submitRequestSection("post", post, false);
    if (state.post.clearOnSubmit && status === 200) {
      state.selectedImageIds = [];
      syncGenerationImageIds();
      saveState();
      renderRequestSection("post", post);
      renderGallery(gallery);
    }
  });

  send.formatButton.addEventListener("click", () => formatBody("send", send));
  query.formatButton.addEventListener("click", () => formatBody("query", query));
  post.formatButton.addEventListener("click", () => formatBody("post", post));

  document.querySelectorAll(".post-clear-ids").forEach(btn => {
    btn.addEventListener("click", () => {
      state.selectedImageIds = [];
      syncGenerationImageIds();
      saveState();
      renderRequestSection("post", post);
      renderGallery(gallery);
    });
  });
}

function initGalleryPage() {
  const query = bindRequestSection("query");
  const post = bindRequestSection("post");
  const gallery = {
    stats: document.querySelector("#gallery-stats"),
    root: document.querySelector("#gallery"),
    selectedCount: document.querySelectorAll(".selected-count"),
    postClearIds: document.querySelectorAll(".post-clear-ids"),
    copySeed: document.querySelector("#gallery-copy-seed"),
    sendSection: null,
  };

  bindGalleryOptions(gallery);
  renderRequestSection("query", query);
  renderRequestSection("post", post);
  renderGallery(gallery);
  initializeRequestSectionVisibility("query", query);
  initializeRequestSectionVisibility("post", post);

  query.parseButton.addEventListener("click", () => {
    try {
      Object.assign(state.query, parsePowerShellRequest(query.powershell.value));
      state.query.bodyText = formatJsonString(state.query.bodyText);
      saveState();
      renderRequestSection("query", query);
      collapseRequestSection(query);
    } catch (error) {
      setResponse("query", `解析失敗: ${error.message}`);
      renderRequestSection("query", query);
    }
  });

  query.requestButton.addEventListener("click", async () => {
    await submitRequestSection("query", query, true);
    renderGallery(gallery);
  });

  query.formatButton.addEventListener("click", () => {
    state.query.bodyText = formatJsonString(query.body.value);
    saveState();
    renderRequestSection("query", query);
  });

  post.parseButton.addEventListener("click", () => {
    try {
      Object.assign(state.post, parsePowerShellRequest(post.powershell.value));
      state.post.bodyText = formatJsonString(state.post.bodyText);
      syncGenerationImageIds();
      saveState();
      renderRequestSection("post", post);
      renderGallery(gallery);
      collapseRequestSection(post);
    } catch (error) {
      setResponse("post", `解析失敗: ${error.message}`);
      renderRequestSection("post", post);
    }
  });

  post.requestButton.addEventListener("click", async () => {
    syncGenerationImageIds();
    renderRequestSection("post", post);
    const status = await submitRequestSection("post", post, false);
    if (state.post.clearOnSubmit && status === 200) {
      state.selectedImageIds = [];
      syncGenerationImageIds();
      saveState();
      renderRequestSection("post", post);
      renderGallery(gallery);
    }
  });

  post.formatButton.addEventListener("click", () => {
    state.post.bodyText = formatJsonString(post.body.value);
    saveState();
    renderRequestSection("post", post);
  });

  gallery.postSync?.addEventListener("click", () => {
    syncGenerationImageIds();
    saveState();
    renderRequestSection("post", post);
    renderGallery(gallery);
  });

  document.querySelectorAll(".post-clear-ids").forEach(btn => {
    btn.addEventListener("click", () => {
      state.selectedImageIds = [];
      syncGenerationImageIds();
      saveState();
      renderRequestSection("post", post);
      renderGallery(gallery);
    });
  });

  bindSectionInputs("query", query);
  bindSectionInputs("post", post);
  
  bindPresetControls("query", query);
  bindPresetControls("post", post);
}

function bindParseAction(key, section, afterParse) {
  section.parseButton.addEventListener("click", () => {
    try {
      Object.assign(state[key], parsePowerShellRequest(section.powershell.value));
      state[key].bodyText = formatJsonString(state[key].bodyText);
      saveState();
      renderRequestSection(key, section);
      collapseRequestSection(section);
      afterParse();
    } catch (error) {
      setResponse(key, `解析失敗: ${error.message}`);
      renderRequestSection(key, section);
    }
  });
}

function formatBody(key, section) {
  state[key].bodyText = formatJsonString(section.body.value);
  saveState();
  renderRequestSection(key, section);
}

function initMetadataPage() {
  const fileInput = document.querySelector("#metadata-file");
  const output = document.querySelector("#metadata-output");
  const applyButton = document.querySelector("#metadata-apply");

  output.textContent = state.importedMetadata
    ? JSON.stringify(state.importedMetadata, null, 2)
    : "尚未載入 metadata。";

  fileInput.addEventListener("change", async (event) => {
    const [file] = event.target.files || [];
    if (!file) return;

    try {
      state.importedMetadata = await readMetadataFromImage(file);
    } catch (error) {
      state.importedMetadata = { error: error.message };
    }

    saveState();
    output.textContent = JSON.stringify(state.importedMetadata, null, 2);
  });

  applyButton.addEventListener("click", () => {
    if (!state.importedMetadata || state.importedMetadata.error) return;

    try {
      state.send.bodyText = applyGenerationMetadataToBody(
        state.send.bodyText,
        state.importedMetadata,
        { includeSeed: true },
      );
      saveState();
      output.textContent = `${JSON.stringify(state.importedMetadata, null, 2)}\n\n已套用到 Send Request Body。`;
    } catch (error) {
      output.textContent = `套用失敗: ${error.message}`;
    }
  });
}

function initCanvasPage() {
  const dom = {
    canvas: document.querySelector("#canvas-stage"),
    file: document.querySelector("#canvas-file"),
    mode: document.querySelector("#canvas-mode"),
    toolButtons: document.querySelectorAll("[data-canvas-tool]"),
    layerButtons: document.querySelectorAll("[data-canvas-layer]"),
    brushSize: document.querySelector("#canvas-brush-size"),
    brushOpacity: document.querySelector("#canvas-brush-opacity"),
    color: document.querySelector("#canvas-color"),
    x: document.querySelector("#canvas-x"),
    y: document.querySelector("#canvas-y"),
    scale: document.querySelector("#canvas-scale"),
    rotation: document.querySelector("#canvas-rotation"),
    prompt: document.querySelector("#canvas-prompt"),
    negativePrompt: document.querySelector("#canvas-negative"),
    width: document.querySelector("#canvas-width"),
    height: document.querySelector("#canvas-height"),
    steps: document.querySelector("#canvas-steps"),
    cfgScale: document.querySelector("#canvas-cfg"),
    seed: document.querySelector("#canvas-seed"),
    denoisingStrength: document.querySelector("#canvas-denoise"),
    modelId: document.querySelector("#canvas-model-id"),
    modelFileId: document.querySelector("#canvas-model-file-id"),
    body: document.querySelector("#canvas-body"),
    stats: document.querySelector("#canvas-stats"),
    build: document.querySelector("#canvas-build"),
    applyToSend: document.querySelector("#canvas-apply-send"),
    send: document.querySelector("#canvas-send"),
    response: document.querySelector("#canvas-response"),
    responseFold: document.querySelector("#canvas-response-fold"),
    clearMask: document.querySelector("#canvas-clear-mask"),
    clearPaint: document.querySelector("#canvas-clear-paint"),
    stageSize: document.querySelector("#canvas-stage-size"),
    layerList: document.querySelector("#canvas-layer-list"),
    layerAdd: document.querySelector("#canvas-layer-add"),
    layerRemove: document.querySelector("#canvas-layer-remove"),
    layerUp: document.querySelector("#canvas-layer-up"),
    layerDown: document.querySelector("#canvas-layer-down"),
    activeLayerName: document.querySelector("#canvas-active-layer-name"),
    pick: document.querySelector("#canvas-pick"),
    brushPopover: document.querySelector("#canvas-brush-popover"),
    brushSizeVal: document.querySelector("#canvas-brush-size-val"),
    brushOpacityVal: document.querySelector("#canvas-brush-opacity-val"),
  };

  if (!dom.canvas) return;

  const editor = createCanvasEditor(dom.canvas, dom);
  renderCanvasEditor(editor);
  applyCanvasImport(editor, dom);
  updateCanvasBodyPreview(editor, dom);

  dom.file.addEventListener("change", (event) => loadCanvasBaseImage(event, editor, dom));
  dom.mode.addEventListener("change", () => updateCanvasBodyPreview(editor, dom));
  dom.build.addEventListener("click", () => updateCanvasBodyPreview(editor, dom));
  dom.applyToSend.addEventListener("click", () => {
    const draft = buildCanvasEditorDraftFromDom(editor, dom);
    state.send.url = draft.url;
    state.send.method = draft.method;
    state.send.headers = {
      ...getCanvasInheritedHeaders(),
      "content-type": draft.contentType,
    };
    state.send.bodyText = JSON.stringify(draft.body, null, 2);
    saveState();
    dom.stats.textContent = "已套用到 Dashboard 的 API 1 Body。";
  });
  dom.send.addEventListener("click", () => sendCanvasRequest(editor, dom));
  dom.clearMask.addEventListener("click", () => {
    editor.maskCtx.clearRect(0, 0, editor.maskCanvas.width, editor.maskCanvas.height);
    renderCanvasEditor(editor);
    updateCanvasBodyPreview(editor, dom);
  });
  dom.clearPaint.addEventListener("click", () => {
    const active = getActiveLayer(editor);
    active.ctx.clearRect(0, 0, active.canvas.width, active.canvas.height);
    renderCanvasEditor(editor);
    updateCanvasBodyPreview(editor, dom);
  });

  dom.toolButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const tool = button.dataset.canvasTool;
      const alreadyActive = editor.tool === tool;
      editor.tool = tool;
      dom.toolButtons.forEach((item) => item.classList.toggle("is-active", item === button));
      if (dom.brushPopover) {
        if (tool === "brush") {
          dom.brushPopover.hidden = alreadyActive ? !dom.brushPopover.hidden : false;
        } else {
          dom.brushPopover.hidden = true;
        }
      }
      renderCanvasEditor(editor);
    });
  });

  if (dom.brushPopover) {
    document.addEventListener("pointerdown", (event) => {
      if (dom.brushPopover.hidden) return;
      if (dom.brushPopover.contains(event.target)) return;
      if (event.target.closest('[data-canvas-tool="brush"]')) return;
      dom.brushPopover.hidden = true;
    });
  }

  if (dom.brushSize && dom.brushSizeVal) {
    const updateSize = () => { dom.brushSizeVal.textContent = dom.brushSize.value; };
    dom.brushSize.addEventListener("input", updateSize);
    updateSize();
  }
  if (dom.brushOpacity && dom.brushOpacityVal) {
    const updateOpacity = () => { dom.brushOpacityVal.textContent = dom.brushOpacity.value; };
    dom.brushOpacity.addEventListener("input", updateOpacity);
    updateOpacity();
  }

  if (dom.pick) {
    dom.pick.addEventListener("click", () => runColorPicker(dom));
  }

  const applyLayerMode = () => {
    const toolbarEl = document.querySelector(".canvas-toolbar");
    const inspectorEl = document.querySelector(".canvas-inspector");
    if (toolbarEl) toolbarEl.dataset.activeLayer = editor.layer;
    if (inspectorEl) inspectorEl.classList.toggle("is-mask-mode", editor.layer === "mask");

    if (editor.layer === "mask") {
      if (editor.tool === "move" || editor.tool === "eyedropper") {
        editor.tool = "brush";
        dom.toolButtons.forEach((item) => item.classList.toggle("is-active", item.dataset.canvasTool === "brush"));
      }
    }
  };

  dom.layerButtons.forEach((button) => {
    button.addEventListener("click", () => {
      editor.layer = button.dataset.canvasLayer;
      dom.layerButtons.forEach((item) => item.classList.toggle("is-active", item === button));
      applyLayerMode();
      renderCanvasEditor(editor);
    });
  });
  applyLayerMode();

  if (dom.layerList) {
    renderLayerPanel(editor, dom);
    dom.layerAdd?.addEventListener("click", () => {
      editor.layerCounter += 1;
      const next = createPaintLayer(`Layer ${editor.layerCounter}`, editor.canvas.width, editor.canvas.height);
      const activeIndex = editor.layers.findIndex((l) => l.id === editor.activeLayerId);
      editor.layers.splice(activeIndex + 1, 0, next);
      editor.activeLayerId = next.id;
      renderLayerPanel(editor, dom);
      renderCanvasEditor(editor);
      updateCanvasBodyPreview(editor, dom);
    });
    dom.layerRemove?.addEventListener("click", () => {
      if (editor.layers.length <= 1) {
        dom.stats.textContent = "至少要保留一個圖層。";
        return;
      }
      const idx = editor.layers.findIndex((l) => l.id === editor.activeLayerId);
      editor.layers.splice(idx, 1);
      const nextIdx = Math.min(idx, editor.layers.length - 1);
      editor.activeLayerId = editor.layers[nextIdx].id;
      renderLayerPanel(editor, dom);
      renderCanvasEditor(editor);
      updateCanvasBodyPreview(editor, dom);
    });
    dom.layerUp?.addEventListener("click", () => moveActiveLayer(editor, dom, 1));
    dom.layerDown?.addEventListener("click", () => moveActiveLayer(editor, dom, -1));
  }

  [dom.x, dom.y, dom.scale, dom.rotation].forEach((input) => {
    input.addEventListener("input", () => {
      editor.transform.x = Number(dom.x.value);
      editor.transform.y = Number(dom.y.value);
      editor.transform.scale = Number(dom.scale.value) / 100;
      editor.transform.rotation = Number(dom.rotation.value);
      renderCanvasEditor(editor);
    });
  });

  [
    dom.prompt,
    dom.negativePrompt,
    dom.steps,
    dom.cfgScale,
    dom.seed,
    dom.denoisingStrength,
    dom.modelId,
    dom.modelFileId,
  ].forEach((input) => {
    input.addEventListener("input", () => updateCanvasBodyPreview(editor, dom));
  });

  [dom.width, dom.height].forEach((input) => {
    input.addEventListener("input", () => {
      resizeCanvasEditor(editor, dom);
      updateCanvasBodyPreview(editor, dom);
    });
  });

  if (dom.stageSize) {
    dom.stageSize.textContent = `${editor.canvas.width}×${editor.canvas.height}`;
  }

  requestAnimationFrame(() => fitCanvasDisplay(editor));
  window.addEventListener("resize", () => fitCanvasDisplay(editor));

  dom.canvas.addEventListener("pointerdown", (event) => handleCanvasPointerDown(event, editor, dom));
  dom.canvas.addEventListener("pointermove", (event) => handleCanvasPointerMove(event, editor, dom));
  const endPointer = () => {
    if (editor.pointer && editor.pointer.mode !== "stroke") {
      updateCanvasBodyPreview(editor, dom);
    }
    editor.pointer = null;
    dom.canvas.style.cursor = "";
  };
  dom.canvas.addEventListener("pointerup", endPointer);
  dom.canvas.addEventListener("pointercancel", endPointer);
  dom.canvas.addEventListener("pointerleave", () => {
    if (!editor.pointer) dom.canvas.style.cursor = "";
  });
}

function bindRequestSection(prefix) {
  return {
    sourceFold: document.querySelector(`#${prefix}-source-fold`),
    sourceSummary: document.querySelector(`#${prefix}-source-summary`),
    responseFold: document.querySelector(`#${prefix}-response-fold`),
    responseSummary: document.querySelector(`#${prefix}-response-summary`),
    powershell: document.querySelector(`#${prefix}-powershell`),
    url: document.querySelector(`#${prefix}-url`),
    method: document.querySelector(`#${prefix}-method`),
    headers: document.querySelector(`#${prefix}-headers`),
    body: document.querySelector(`#${prefix}-body`),
    response: document.querySelector(`#${prefix}-response`),
    parseButton: document.querySelector(`#${prefix}-parse`),
    requestButton: document.querySelector(`#${prefix}-request`),
    formatButton: document.querySelector(`#${prefix}-format`),
    clearOnSubmit: document.querySelector(`#${prefix}-clear-on-submit`),
    presetSelect: document.querySelector(`#${prefix}-preset-select`),
    presetSave: document.querySelector(`#${prefix}-preset-save`),
    presetSaveAs: document.querySelector(`#${prefix}-preset-saveas`),
    presetDelete: document.querySelector(`#${prefix}-preset-delete`),
  };
}

function bindSectionInputs(key, section) {
  section.powershell.addEventListener("input", (event) => {
    state[key].powershell = event.target.value;
    saveState();
  });
  section.url.addEventListener("input", (event) => {
    state[key].url = event.target.value;
    saveState();
  });
  section.method.addEventListener("input", (event) => {
    state[key].method = event.target.value;
    saveState();
  });
  section.headers.addEventListener("input", (event) => {
    try {
      state[key].headers = JSON.parse(event.target.value || "{}");
    } catch {
      state[key].headers = {};
    }
    saveState();
  });
  section.body.addEventListener("input", (event) => {
    state[key].bodyText = event.target.value;
    saveState();
  });
  if (section.clearOnSubmit) {
    section.clearOnSubmit.addEventListener("change", (event) => {
      state[key].clearOnSubmit = event.target.checked;
      saveState();
    });
  }
}

function generateId() {
  return Math.random().toString(36).substring(2, 9);
}

function bindPresetControls(key, section) {
  if (!section.presetSelect) return;

  renderPresetOptions(key, section);

  section.presetSelect.addEventListener("change", (e) => {
    const selectedId = e.target.value;
    if (!selectedId) {
      state[key].presetId = null;
    } else {
      const preset = state.savedSettings[key].find(p => p.id === selectedId);
      if (preset) {
        state[key] = { ...blankRequestState(), ...JSON.parse(JSON.stringify(preset.request)), presetId: preset.id };
        saveState();
        renderRequestSection(key, section);
      }
    }
  });

  section.presetSave.addEventListener("click", () => {
    if (!state[key].presetId) {
      section.presetSaveAs.click();
      return;
    }
    const idx = state.savedSettings[key].findIndex(p => p.id === state[key].presetId);
    if (idx !== -1) {
      state.savedSettings[key][idx].request = extractRequestForPreset(key);
      saveState();
      alert("設定已更新並儲存");
    }
  });

  section.presetSaveAs.addEventListener("click", () => {
    const name = prompt("請輸入新設定名稱：");
    if (!name) return;
    const newPreset = {
      id: generateId(),
      name,
      request: extractRequestForPreset(key)
    };
    state.savedSettings[key].push(newPreset);
    state[key].presetId = newPreset.id;
    saveState();
    renderPresetOptions(key, section);
    alert("另存新設定成功");
  });

  section.presetDelete.addEventListener("click", () => {
    if (!state[key].presetId) return;
    if (!confirm("確定要刪除此設定嗎？")) return;
    state.savedSettings[key] = state.savedSettings[key].filter(p => p.id !== state[key].presetId);
    state[key].presetId = null;
    saveState();
    renderPresetOptions(key, section);
  });
}

function extractRequestForPreset(key) {
  const req = state[key];
  return {
    powershell: req.powershell,
    url: req.url,
    method: req.method,
    headers: JSON.parse(JSON.stringify(req.headers)),
    bodyText: req.bodyText,
    clearOnSubmit: req.clearOnSubmit || false
  };
}

function renderPresetOptions(key, section) {
  if (!section.presetSelect) return;
  const presets = state.savedSettings[key] || [];
  
  const options = [`<option value="">-- 未儲存的草稿 --</option>`];
  presets.forEach(p => {
    options.push(`<option value="${escapeHtml(p.id)}">${escapeHtml(p.name)}</option>`);
  });
  section.presetSelect.innerHTML = options.join("");
  section.presetSelect.value = state[key].presetId || "";
  
  if (section.presetSave) section.presetSave.disabled = !state[key].presetId;
  if (section.presetDelete) section.presetDelete.disabled = !state[key].presetId;
}

function renderRequestSection(key, section) {
  section.powershell.value = state[key].powershell;
  section.url.value = state[key].url;
  section.method.value = state[key].method;
  section.headers.value = JSON.stringify(state[key].headers, null, 2);
  section.body.value = state[key].bodyText;
  section.response.textContent = buildResponseBodyPreview(state[key].responseText);
  section.sourceSummary.textContent = buildSourcePreview(state[key]);
  section.responseSummary.textContent = buildResponsePreview(state[key].responseText);
  
  if (section.clearOnSubmit) {
    section.clearOnSubmit.checked = state[key].clearOnSubmit || false;
  }
}

function collapseRequestSection(section) {
  section.sourceFold.open = false;
  section.responseFold.open = false;
}

function initializeRequestSectionVisibility(key, section) {
  const request = state[key];
  const hasAPIData = Boolean(
    request.url.trim()
    || request.method.trim()
    || Object.keys(request.headers).length
  );

  section.sourceFold.open = !hasAPIData && !request.powershell.trim();
  section.responseFold.open = false;
}

async function submitRequestSection(key, section, updateGallery) {
  try {
    const request = buildFetchRequest(key);
    setResponse(key, "送出中...");
    renderRequestSection(key, section);
    section.responseFold.open = false;

    const response = await fetch(request.url, request.options);
    const text = await response.text();
    setResponse(key, `HTTP ${response.status}\n${formatResponse(text)}`);
    renderRequestSection(key, section);
    section.responseFold.open = false;

    if (updateGallery) {
      const json = JSON.parse(text);
      state.galleryItems = flattenTasks(json?.data?.tasks ?? []);
      saveState();
    }
    return response.status;
  } catch (error) {
    setResponse(key, `Request 失敗: ${error.message}`);
    renderRequestSection(key, section);
    section.responseFold.open = false;
    return null;
  }
}

function buildFetchRequest(key) {
  const request = state[key];
  const headers = sanitizeHeaders(request.headers);
  const options = {
    method: (request.method || "POST").toUpperCase(),
    headers,
    mode: "cors",
    credentials: "include",
  };

  if (request.bodyText.trim()) {
    options.body = JSON.stringify(JSON.parse(request.bodyText));
  }

  return { url: request.url, options };
}

function setResponse(key, text) {
  state[key].responseText = text;
  saveState();
}

function syncGenerationImageIds() {
  if (!state.post.bodyText.trim()) return;

  try {
    const parsed = JSON.parse(state.post.bodyText);
    parsed.generationImageIds = [...new Set(state.selectedImageIds)];
    state.post.bodyText = JSON.stringify(parsed, null, 2);
  } catch {
    // keep user input unchanged when invalid JSON
  }
}

function bindGalleryOptions(dom) {
  if (!dom.copySeed) return;

  dom.copySeed.checked = state.queryResultCopySeed;
  dom.copySeed.addEventListener("change", (event) => {
    state.queryResultCopySeed = event.target.checked;
    saveState();
  });
}

function renderGallery(dom) {
  dom.selectedCount.forEach(el => { el.textContent = String(state.selectedImageIds.length); });

  if (!state.galleryItems.length) {
    dom.stats.textContent = "目前沒有查詢結果。";
    dom.root.innerHTML = "";
    return;
  }

  dom.stats.textContent = `共 ${state.galleryItems.length} 張圖片，可直接送到 Canvas Editor 做 Inpaint 或 Img2Img。`;
  dom.root.innerHTML = state.galleryItems.map((entry, index) => {
    return `
      <article class="gallery-card">
        <img src="${escapeHtml(entry.url)}" alt="Task ${escapeHtml(entry.taskId)}">
        <div class="gallery-body">
          <div class="gallery-head">
            <h3 class="gallery-title">Task ${escapeHtml(entry.taskId)}</h3>
            <span class="pill">${escapeHtml(entry.status || "UNKNOWN")}</span>
          </div>
          <div class="gallery-selection">
            <button type="button" class="remix-button" data-action="remix-inpaint" data-index="${index}">Inpaint</button>
            <button type="button" class="remix-button" data-action="remix-img2img" data-index="${index}">Img2Img</button>
          </div>
          <div class="gallery-meta">
            <span class="pill">Seed ${escapeHtml(String(entry.metadata.seed || "-"))}</span>
            <span class="pill">${escapeHtml(entry.metadata.size || "-")}</span>
          </div>
          <pre>${escapeHtml(JSON.stringify(entry.metadata, null, 2))}</pre>
          <div class="gallery-actions">
            <button type="button" data-action="open-original" data-index="${index}">開啟原圖</button>
          </div>
        </div>
      </article>
    `;
  }).join("");

  dom.root.querySelectorAll("button[data-action]").forEach((button) => {
    button.addEventListener("click", () => handleGalleryAction(button.dataset.action, Number(button.dataset.index), dom));
  });
}

function toggleSelectedImageId(imageId, checked) {
  if (checked) {
    state.selectedImageIds = [...new Set([...state.selectedImageIds, imageId])];
  } else {
    state.selectedImageIds = state.selectedImageIds.filter((id) => id !== imageId);
  }
  syncGenerationImageIds();
  saveState();
  const postBody = document.querySelector("#post-body");
  if (postBody) postBody.value = state.post.bodyText;
}

function convertBlobToJpeg(blob) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = img.width;
      canvas.height = img.height;
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0);
      canvas.toBlob((jpegBlob) => {
        if (jpegBlob) resolve(jpegBlob);
        else reject(new Error("無法轉換為 JPG"));
      }, "image/jpeg", 0.95);
      URL.revokeObjectURL(img.src);
    };
    img.onerror = () => reject(new Error("圖片載入失敗，無法轉換"));
    img.src = URL.createObjectURL(blob);
  });
}

function extractBlobFromImageElement(imgElement, targetMimeType) {
  return new Promise((resolve, reject) => {
    try {
      if (!imgElement || !imgElement.src) {
        reject(new Error("圖片尚未載入或無法存取"));
        return;
      }
      // 建立新的 Image 物件並設定 crossorigin，避免汙染 gallery 中的原始 img
      const corsImg = new Image();
      corsImg.crossOrigin = "anonymous";
      corsImg.onload = () => {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = corsImg.naturalWidth || corsImg.width;
          canvas.height = corsImg.naturalHeight || corsImg.height;
          const ctx = canvas.getContext("2d");
          if (targetMimeType === "image/jpeg" || targetMimeType === "image/jpg") {
            ctx.fillStyle = "#ffffff";
            ctx.fillRect(0, 0, canvas.width, canvas.height);
          }
          ctx.drawImage(corsImg, 0, 0);
          canvas.toBlob((blob) => {
            if (blob) resolve(blob);
            else reject(new Error("無法從 Canvas 取得影像資料"));
          }, targetMimeType, 0.95);
        } catch (e) {
          reject(e);
        }
      };
      corsImg.onerror = () => reject(new Error("以 CORS 模式載入圖片失敗"));
      corsImg.src = imgElement.src;
    } catch (e) {
      reject(e);
    }
  });
}

async function handleGalleryAction(action, index, dom) {
  const entry = state.galleryItems[index];
  if (!entry) return;

  if (action === "remix-inpaint" || action === "remix-img2img") {
    const mode = action === "remix-inpaint" ? "inpaint" : "img2img";
    state.canvasImport = buildCanvasImportFromGalleryEntry(entry, mode);
    saveState();
    window.location.href = "./canvas.html";
    return;
  }

  if (action === "copy-to-send") {
    copyGenerationToSendBody(entry, dom);
    return;
  }

  const signedExpiry = getSignedUrlExpiry(entry.url);
  if (signedExpiry && new Date() > signedExpiry) {
    dom.stats.textContent = `無法下載：圖片連結已在 ${signedExpiry.toLocaleTimeString("zh-TW", { hour12: false })} 過期，請重新查詢一次 API。`;
    return;
  }

  const baseName = (entry.metadata.downloadFileName || `${entry.taskId}`).replace(/\.[^.]+$/, "");
  const originalFileName = `${baseName}.${resolveExtension(entry.metadata.mimeType || "image/png")}`;

  if (action === "open-original") {
    window.open(entry.url, "_blank", "noopener");
    return;
  }

  try {
    const suggestedFileName = action === "download-jpg" ? `${baseName}.jpg` : originalFileName;
    const mimeType = action === "download-jpg" ? "image/jpeg" : (entry.metadata.mimeType || "image/png");
    
    let fileHandle = null;
    let usePicker = "showSaveFilePicker" in window;
    
    if (usePicker) {
      try {
        fileHandle = await window.showSaveFilePicker({
          suggestedName: sanitizeFileName(suggestedFileName),
          types: buildPickerTypes(mimeType),
        });
      } catch (e) {
        if (e.name === "AbortError") return;
        usePicker = false;
      }
    }
    
    dom.stats.textContent = "正在從快取轉換圖片，請稍候...";
    
    let blob;
    try {
      // 直接從畫面上已經載入的 img 標籤提取並轉檔，不重新 fetch()
      const imgElement = dom.root.querySelectorAll('.gallery-card img')[index];
      const targetMimeType = action === "download-jpg" ? "image/jpeg" : mimeType;
      blob = await extractBlobFromImageElement(imgElement, targetMimeType);
    } catch (e) {
      // 降級方案：再次發送請求
      dom.stats.textContent = "無法直接從快取讀取，正嘗試重新下載圖片...";
      blob = await fetchBlobForSave(entry.url, mimeType);
      if (action === "download-jpg" && blob.type !== "image/jpeg" && blob.type !== "image/jpg") {
        blob = await convertBlobToJpeg(blob);
      }
    }
    
    if (fileHandle) {
      const writable = await fileHandle.createWritable();
      await writable.write(blob);
      await writable.close();
    } else {
      triggerDownload(blob, suggestedFileName);
    }
    
    dom.stats.textContent = `已完成另存新檔：${suggestedFileName}`;
  } catch (error) {
    if (error.name === "AbortError") return;
    dom.stats.textContent = `無法下載：${error.message} (建議重新 Query API 取得新連結)`;
  }
}

function copyGenerationToSendBody(entry, dom) {
  try {
    state.send.bodyText = applyGenerationMetadataToBody(
      state.send.bodyText,
      entry.metadata,
      { includeSeed: state.queryResultCopySeed },
    );
    saveState();

    if (dom.sendSection) {
      renderRequestSection("send", dom.sendSection);
    }

    dom.stats.textContent = state.queryResultCopySeed
      ? "已複製生成資料到 API 1 Body（包含 seed）。"
      : "已複製生成資料到 API 1 Body（seed 已設為 -1）。";
  } catch (error) {
    dom.stats.textContent = `無法複製生成資料：${error.message}`;
  }
}

function buildCanvasImportFromGalleryEntry(entry, mode) {
  const metadata = entry.metadata || {};
  const imageUrl = entry.url || metadata.imageUrl || "";
  return {
    mode,
    imageUrl,
    imageId: entry.generationImageId || metadata.imageId || "",
    taskId: entry.taskId || "",
    prompt: metadata.prompt || "",
    negativePrompt: metadata.negativePrompt || "",
    width: castNumber(metadata.width, ""),
    height: castNumber(metadata.height, ""),
    steps: castNumber(metadata.steps, ""),
    cfgScale: castNumber(metadata.cfgScale, ""),
    seed: String(metadata.seed || "-1"),
    denoisingStrength: castNumber(metadata.denoisingStrength, ""),
    modelId: metadata.modelId ? String(metadata.modelId) : "",
    modelFileId: metadata.modelFileId ? String(metadata.modelFileId) : "",
    ksamplerName: metadata.kSampler || metadata.sampler || "euler_ancestral",
    schedule: metadata.schedule || "normal",
  };
}

function parsePowerShellRequest(text) {
  if (!text.trim()) {
    throw new Error("請先貼上 PowerShell 內容");
  }

  const url = capture(text, /-Uri\s+"([\s\S]*?)"\s*`?\s*-Method/i);
  const method = capture(text, /-Method\s+"([^"]+)"/i);
  const headersBlock = captureOptional(text, /-Headers\s+@\{([\s\S]*?)\}\s*`?\s*-ContentType/i);
  const bodyLiteral = extractPowerShellBodyLiteral(text);

  return {
    powershell: text,
    url: decodePowerShellString(url),
    method: decodePowerShellString(method).toUpperCase(),
    headers: sanitizeHeaders(parseHeaderBlock(headersBlock || "")),
    bodyText: decodePowerShellQuotedString(bodyLiteral.value, bodyLiteral.quote),
  };
}

function parseHeaderBlock(block) {
  const headers = {};
  const pattern = /"([^"]+)"\s*=\s*"([\s\S]*?)"/g;
  let match;

  while ((match = pattern.exec(block)) !== null) {
    headers[match[1]] = decodePowerShellString(match[2]);
  }

  return headers;
}

function sanitizeHeaders(headers) {
  return Object.fromEntries(
    Object.entries(headers).filter(([key]) => !FORBIDDEN_HEADERS.has(key.toLowerCase())),
  );
}

function capture(text, pattern) {
  const match = text.match(pattern);
  if (!match) {
    throw new Error(`找不到必要欄位: ${pattern}`);
  }
  return match[1];
}

function captureOptional(text, pattern) {
  const match = text.match(pattern);
  return match ? match[1] : "";
}

function extractPowerShellBodyLiteral(text) {
  const bodyMatch = /-Body\b/i.exec(text);
  if (!bodyMatch) {
    return { value: "", quote: "\"" };
  }

  const segment = text.slice(bodyMatch.index + bodyMatch[0].length);
  const literal = findFirstPowerShellStringLiteral(segment);
  return literal || { value: "", quote: "\"" };
}

function findFirstPowerShellStringLiteral(text) {
  for (let index = 0; index < text.length; index += 1) {
    const quote = text[index];
    if (quote !== "\"" && quote !== "'") {
      continue;
    }

    const literal = quote === "\""
      ? readDoubleQuotedPowerShellString(text, index)
      : readSingleQuotedPowerShellString(text, index);

    if (literal) {
      return literal;
    }
  }

  return null;
}

function readDoubleQuotedPowerShellString(text, startIndex) {
  let index = startIndex + 1;

  while (index < text.length) {
    const char = text[index];
    if (char === "`") {
      index += 2;
      continue;
    }

    if (char === "\"") {
      return {
        value: text.slice(startIndex + 1, index),
        quote: "\"",
      };
    }

    index += 1;
  }

  return null;
}

function readSingleQuotedPowerShellString(text, startIndex) {
  let index = startIndex + 1;

  while (index < text.length) {
    if (text[index] === "'" && text[index + 1] === "'") {
      index += 2;
      continue;
    }

    if (text[index] === "'") {
      return {
        value: text.slice(startIndex + 1, index),
        quote: "'",
      };
    }

    index += 1;
  }

  return null;
}

function decodePowerShellString(value) {
  return value
    .replace(/`"/g, "\"")
    .replace(/``/g, "`")
    .replace(/`r/g, "\r")
    .replace(/`n/g, "\n")
    .replace(/`\$/g, "$");
}

function decodePowerShellQuotedString(value, quote = "\"") {
  if (quote === "'") {
    return value.replace(/''/g, "'");
  }

  return decodePowerShellString(value);
}

function formatJsonString(text) {
  if (!text.trim()) return "";
  try {
    return JSON.stringify(JSON.parse(text), null, 2);
  } catch {
    return text;
  }
}

function formatResponse(text) {
  try {
    return JSON.stringify(JSON.parse(text), null, 2);
  } catch {
    return text;
  }
}

function buildResponsePreview(text) {
  if (!text) return "回應預覽";
  const lines = text.split(/\r?\n/);
  const preview = lines.slice(0, 5).join(" ");
  return lines.length > 5 ? `${preview} ...` : preview;
}

function buildResponseBodyPreview(text) {
  if (!text) return "";
  const lines = text.split(/\r?\n/);
  const preview = lines.slice(0, 5).join("\n");
  return lines.length > 5 ? `${preview}\n...` : preview;
}

function buildSourcePreview(request) {
  if (!request.url.trim()) {
    return "API 詳細設定已就緒 / 解析等待中";
  }

  const method = (request.method || "POST").toUpperCase();
  const url = request.url.trim();
  if (!url) {
    return "PowerShell 已輸入";
  }

  const preview = `${method} ${url}`;
  return preview.length > 72 ? `${preview.slice(0, 69)}...` : preview;
}

function flattenTasks(tasks) {
  return tasks.flatMap((task) => {
    const metadata = extractTaskMetadata(task);
    return (task.items || []).map((item) => ({
      taskId: task.taskId,
      status: item.status || task.status,
      url: item.url,
      generationImageId: item.generationImageId || item.imageId,
      metadata: {
        ...metadata,
        width: item.width || metadata.width,
        height: item.height || metadata.height,
        mimeType: item.mimeType || "image/png",
        downloadFileName: item.downloadFileName || `${task.taskId}.png`,
      },
    }));
  });
}

function extractTaskMetadata(task) {
  const visualMap = Object.fromEntries((task.visualParameters || []).map((entry) => [entry.name, entry.value]));
  const size = visualMap.Size || [task.items?.[0]?.width, task.items?.[0]?.height].filter(Boolean).join("x");
  const [width, height] = String(size || "x").split("x");

  return {
    prompt: task.inputData?.prompt || visualMap.Prompt || "",
    negativePrompt: visualMap["Negative prompt"] || "",
    model: visualMap.Model || task.baseModel?.name || "",
    seed: visualMap.Seed || task.items?.[0]?.seed || "",
    steps: visualMap.Steps || "",
    cfgScale: visualMap["CFG scale"] || "",
    sampler: visualMap.Sampler || "",
    kSampler: visualMap.KSampler || "",
    schedule: visualMap.Schedule || "",
    guidance: visualMap.Guidance || "",
    vae: visualMap.VAE || "",
    clipSkip: visualMap["Clip skip"] || "",
    denoisingStrength: visualMap["Denoising strength"] || "",
    size: size || "",
    width: width || "",
    height: height || "",
    taskType: task.taskType || "",
    modelId: task.baseModel?.modelId || "",
    modelFileId: task.baseModel?.modelFileId || "",
  };
}

function applyGenerationMetadataToBody(bodyText, metadata, options = {}) {
  const payload = JSON.parse(bodyText || "{}");
  const params = payload.params || {};
  const includeSeed = Boolean(options.includeSeed);

  params.prompt = metadata.prompt ?? params.prompt ?? "";
  params.negativePrompt = metadata.negativePrompt ?? params.negativePrompt ?? "";
  params.steps = castNumber(metadata.steps, params.steps);
  params.cfgScale = castNumber(metadata.cfgScale, params.cfgScale);
  params.guidance = castNumber(metadata.guidance, params.guidance);
  params.clipSkip = castNumber(metadata.clipSkip, params.clipSkip);
  params.seed = includeSeed ? String(metadata.seed ?? params.seed ?? "-1") : "-1";
  params.sdVae = metadata.vae ?? params.sdVae ?? "Automatic";
  params.ksamplerName = metadata.kSampler || metadata.sampler || params.ksamplerName || "";
  params.schedule = metadata.schedule ?? params.schedule ?? "";
  params.width = castNumber(metadata.width, params.width);
  params.height = castNumber(metadata.height, params.height);
  params.denoisingStrength = castNumber(metadata.denoisingStrength, params.denoisingStrength);

  if (metadata.modelId && metadata.modelFileId) {
    params.baseModel = {
      ...(params.baseModel || {}),
      modelId: String(metadata.modelId),
      modelFileId: String(metadata.modelFileId),
    };
  }

  payload.params = params;
  return JSON.stringify(payload, null, 2);
}

function buildCanvasEditorRequestDraft(input) {
  const params = {
    baseModel: {
      modelId: String(input.modelId || ""),
      modelFileId: String(input.modelFileId || ""),
    },
    sdxl: { refiner: false },
    models: [],
    embeddingModels: [],
    sdVae: input.sdVae || "Automatic",
    prompt: input.prompt || "",
    negativePrompt: input.negativePrompt || "",
    height: castNumber(input.height, 1024),
    width: castNumber(input.width, 1024),
    imageCount: castNumber(input.imageCount, 1),
    steps: castNumber(input.steps, 25),
    images: input.imageDataUrl ? [input.imageDataUrl] : [],
    cfgScale: castNumber(input.cfgScale, 5),
    seed: String(input.seed || "-1"),
    clipSkip: castNumber(input.clipSkip, 2),
    etaNoiseSeedDelta: castNumber(input.etaNoiseSeedDelta, 31327),
    v1Clip: true,
    enablePix2pix: false,
    guidance: castNumber(input.guidance, 3.5),
    useFirstLastFrame: false,
    ksamplerName: input.ksamplerName || "euler_ancestral",
    schedule: input.schedule || "normal",
    denoisingStrength: castNumber(input.denoisingStrength, 0.51),
  };

  if (input.mode === "inpaint") {
    params.inpaint = {
      maskImage: input.maskDataUrl || "",
      maskBlur: castNumber(input.maskBlur, 4),
    };

    return {
      url: "https://api.tensor.art/works/v1/works/task_by_image",
      method: "POST",
      contentType: "application/json",
      body: {
        params,
        credits: castNumber(input.credits, 3.66),
        taskType: "IMAGE_TO_INPAINT",
        imageUrl: input.imageDataUrl || "",
        imageId: input.imageId || "",
      },
    };
  }

  return {
    url: "https://api.tensor.art/works/v1/works/task",
    method: "POST",
    contentType: "text/plain;charset=UTF-8",
    body: {
      params,
      credits: castNumber(input.credits, 3.66),
      taskType: "IMG2IMG",
      isRemix: true,
      remixPostId: input.remixPostId || "",
      remixPostImageId: input.remixPostImageId || "",
      captchaType: "CLOUDFLARE_TURNSTILE",
    },
  };
}

function getCanvasInheritedHeaders() {
  const candidates = [state.send.headers, state.query.headers, state.post.headers];
  const inherited = candidates.find((headers) => headers && Object.keys(headers).length) || {};
  return sanitizeHeaders(inherited);
}

async function sendCanvasRequest(editor, dom) {
  const draft = buildCanvasEditorDraftFromDom(editor, dom);
  const headers = {
    ...getCanvasInheritedHeaders(),
    "content-type": draft.contentType,
  };

  if (!Object.keys(getCanvasInheritedHeaders()).length) {
    dom.stats.textContent = "找不到可用的 headers。請先在 Dashboard 貼一次 PowerShell 來繼承授權 cookie。";
    dom.responseFold.open = true;
    dom.response.textContent = "尚未送出：缺少 headers (authorization / cookie 等)。";
    return;
  }

  dom.send.disabled = true;
  dom.responseFold.open = true;
  dom.response.textContent = "送出中...";
  dom.stats.textContent = "送出中...";

  try {
    const response = await fetch(draft.url, {
      method: draft.method,
      headers,
      mode: "cors",
      credentials: "include",
      body: JSON.stringify(draft.body),
    });
    const text = await response.text();
    dom.response.textContent = `HTTP ${response.status}\n${formatResponse(text)}`;
    dom.stats.textContent = `已送出 (HTTP ${response.status})`;
  } catch (error) {
    dom.response.textContent = `Request 失敗: ${error.message}`;
    dom.stats.textContent = `送出失敗: ${error.message}`;
  } finally {
    dom.send.disabled = false;
  }
}

function clampCanvasDim(value) {
  const n = Math.round(Number(value) || 0);
  return Math.max(64, Math.min(4096, n));
}

function fitCanvasDisplay(editor) {
  const canvas = editor.canvas;
  const host = canvas.parentElement;
  const frame = host?.parentElement;
  if (!frame) return;
  const aw = canvas.width;
  const ah = canvas.height;
  if (!aw || !ah) return;
  const style = getComputedStyle(frame);
  const padX = parseFloat(style.paddingLeft) + parseFloat(style.paddingRight);
  const padY = parseFloat(style.paddingTop) + parseFloat(style.paddingBottom);
  const fw = Math.max(1, (frame.clientWidth || aw) - padX);
  const fh = Math.max(1, (frame.clientHeight || ah) - padY);
  const scale = Math.min(fw / aw, fh / ah);
  const displayW = Math.max(1, Math.floor(aw * scale));
  const displayH = Math.max(1, Math.floor(ah * scale));
  canvas.style.width = `${displayW}px`;
  canvas.style.height = `${displayH}px`;
}

function resizeCanvasEditor(editor, dom) {
  const nextW = clampCanvasDim(Number(dom.width.value) || editor.canvas.width);
  const nextH = clampCanvasDim(Number(dom.height.value) || editor.canvas.height);
  if (nextW === editor.canvas.width && nextH === editor.canvas.height) return;

  editor.layers = editor.layers.map((layer) => {
    const next = document.createElement("canvas");
    next.width = nextW;
    next.height = nextH;
    next.getContext("2d").drawImage(layer.canvas, 0, 0);
    return { ...layer, canvas: next, ctx: next.getContext("2d") };
  });

  const newMask = document.createElement("canvas");
  newMask.width = nextW;
  newMask.height = nextH;
  newMask.getContext("2d").drawImage(editor.maskCanvas, 0, 0);

  editor.canvas.width = nextW;
  editor.canvas.height = nextH;
  editor.maskCanvas = newMask;
  editor.maskCtx = newMask.getContext("2d");

  editor.transform.x = Math.round(nextW / 2);
  editor.transform.y = Math.round(nextH / 2);
  dom.x.value = String(editor.transform.x);
  dom.y.value = String(editor.transform.y);

  if (dom.stageSize) {
    dom.stageSize.textContent = `${nextW}×${nextH}`;
  }

  fitCanvasDisplay(editor);
  renderCanvasEditor(editor);
}

function createCanvasEditor(canvas, dom) {
  const maskCanvas = document.createElement("canvas");
  const width = clampCanvasDim(Number(dom.width.value) || 1024);
  const height = clampCanvasDim(Number(dom.height.value) || 768);

  canvas.width = width;
  canvas.height = height;
  maskCanvas.width = width;
  maskCanvas.height = height;

  const firstLayer = createPaintLayer("Layer 1", width, height);

  return {
    canvas,
    ctx: canvas.getContext("2d"),
    layers: [firstLayer],
    activeLayerId: firstLayer.id,
    layerCounter: 1,
    maskCanvas,
    maskCtx: maskCanvas.getContext("2d"),
    baseImage: null,
    imageDataUrl: "",
    sourceImageUrl: "",
    imageId: "",
    taskId: "",
    ksamplerName: "euler_ancestral",
    schedule: "normal",
    layer: "paint",
    tool: "move",
    pointer: null,
    transform: {
      x: Number(dom.x.value),
      y: Number(dom.y.value),
      scale: Number(dom.scale.value) / 100,
      rotation: Number(dom.rotation.value),
    },
  };
}

function createPaintLayer(name, width, height) {
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  return {
    id: `layer-${Math.random().toString(36).slice(2, 9)}`,
    name,
    visible: true,
    canvas,
    ctx: canvas.getContext("2d"),
  };
}

function getActiveLayer(editor) {
  return editor.layers.find((layer) => layer.id === editor.activeLayerId) || editor.layers[0];
}

function moveActiveLayer(editor, dom, delta) {
  const idx = editor.layers.findIndex((l) => l.id === editor.activeLayerId);
  const target = idx + delta;
  if (idx < 0 || target < 0 || target >= editor.layers.length) return;
  const [moved] = editor.layers.splice(idx, 1);
  editor.layers.splice(target, 0, moved);
  renderLayerPanel(editor, dom);
  renderCanvasEditor(editor);
  updateCanvasBodyPreview(editor, dom);
}

function renderLayerPanel(editor, dom) {
  if (!dom.layerList) return;
  dom.layerList.innerHTML = "";

  for (let i = editor.layers.length - 1; i >= 0; i -= 1) {
    const layer = editor.layers[i];
    const li = document.createElement("li");
    li.className = "layer-item" + (layer.id === editor.activeLayerId ? " is-active" : "");
    li.dataset.layerId = layer.id;

    const vis = document.createElement("button");
    vis.type = "button";
    vis.className = "layer-visibility" + (layer.visible ? "" : " is-hidden");
    vis.title = layer.visible ? "Hide layer" : "Show layer";
    vis.textContent = layer.visible ? "👁" : "🚫";
    vis.addEventListener("click", (event) => {
      event.stopPropagation();
      layer.visible = !layer.visible;
      renderLayerPanel(editor, dom);
      renderCanvasEditor(editor);
      updateCanvasBodyPreview(editor, dom);
    });

    const name = document.createElement("input");
    name.type = "text";
    name.className = "layer-name";
    name.value = layer.name;
    name.readOnly = true;
    name.addEventListener("dblclick", () => {
      name.readOnly = false;
      name.focus();
      name.select();
    });
    name.addEventListener("blur", () => {
      layer.name = name.value.trim() || layer.name;
      name.value = layer.name;
      name.readOnly = true;
      if (layer.id === editor.activeLayerId && dom.activeLayerName) {
        dom.activeLayerName.textContent = layer.name;
      }
    });
    name.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        name.blur();
      } else if (event.key === "Escape") {
        name.value = layer.name;
        name.blur();
      }
    });

    li.addEventListener("click", () => {
      editor.activeLayerId = layer.id;
      renderLayerPanel(editor, dom);
    });

    li.append(vis, name);
    dom.layerList.append(li);
  }

  if (dom.activeLayerName) {
    dom.activeLayerName.textContent = getActiveLayer(editor).name;
  }
}

function loadCanvasBaseImage(event, editor, dom) {
  const [file] = event.target.files || [];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = () => {
    const image = new Image();
    image.onload = () => {
      editor.baseImage = image;
      editor.imageDataUrl = String(reader.result || "");
      editor.sourceImageUrl = "";
      editor.imageId = "";
      editor.taskId = "";
      adoptCanvasImageDimensions(editor, dom, image, { matchSize: true });
      renderCanvasEditor(editor);
      updateCanvasBodyPreview(editor, dom);
      dom.stats.textContent = `已載入本地圖片：${file.name}（畫布同步為 ${image.width}×${image.height}）`;
    };
    image.src = String(reader.result || "");
  };
  reader.readAsDataURL(file);
}

function adoptCanvasImageDimensions(editor, dom, image, { matchSize }) {
  if (matchSize && image.width && image.height) {
    dom.width.value = String(image.width);
    dom.height.value = String(image.height);
    resizeCanvasEditor(editor, dom);
  }
  const fitScale = Math.min(editor.canvas.width / image.width, editor.canvas.height / image.height) * 0.9;
  editor.transform = {
    x: Math.round(editor.canvas.width / 2),
    y: Math.round(editor.canvas.height / 2),
    scale: Number.isFinite(fitScale) ? fitScale : 1,
    rotation: 0,
  };
  dom.x.value = String(editor.transform.x);
  dom.y.value = String(editor.transform.y);
  dom.scale.value = String(Math.round(editor.transform.scale * 100));
  dom.rotation.value = "0";
}

function applyCanvasImport(editor, dom) {
  const incoming = state.canvasImport;
  if (!incoming) return;

  dom.mode.value = incoming.mode || "img2img";
  dom.prompt.value = incoming.prompt || "";
  dom.negativePrompt.value = incoming.negativePrompt || "";
  if (incoming.width) dom.width.value = String(incoming.width);
  if (incoming.height) dom.height.value = String(incoming.height);
  if (incoming.steps) dom.steps.value = String(incoming.steps);
  if (incoming.cfgScale) dom.cfgScale.value = String(incoming.cfgScale);
  dom.seed.value = String(incoming.seed || "-1");
  if (incoming.denoisingStrength) dom.denoisingStrength.value = String(incoming.denoisingStrength);
  dom.modelId.value = incoming.modelId || "";
  dom.modelFileId.value = incoming.modelFileId || "";

  editor.importHadDims = Boolean(incoming.width && incoming.height);
  if (editor.importHadDims) {
    resizeCanvasEditor(editor, dom);
  }

  editor.imageDataUrl = incoming.imageUrl || "";
  editor.sourceImageUrl = incoming.imageUrl || "";
  editor.imageId = incoming.imageId || "";
  editor.taskId = incoming.taskId || "";
  editor.ksamplerName = incoming.ksamplerName || editor.ksamplerName;
  editor.schedule = incoming.schedule || editor.schedule;

  if (incoming.imageUrl) {
    loadCanvasImageUrlWithFallback(incoming.imageUrl, editor, dom);
  }

  dom.stats.textContent = incoming.mode === "inpaint"
    ? "已從 Query Results 帶入 Inpaint。"
    : "已從 Query Results 帶入 Img2Img。";
}

function loadCanvasImageUrlWithFallback(url, editor, dom) {
  const attempts = createCanvasImageLoadAttempts(url);
  let index = 0;

  const tryLoad = () => {
    const attempt = attempts[index];
    if (!attempt) {
      dom.stats.textContent = "Image could not be loaded. The API body will keep the original image URL.";
      updateCanvasBodyPreview(editor, dom);
      return;
    }

    const image = new Image();
    if (attempt.crossOrigin) {
      image.crossOrigin = attempt.crossOrigin;
    }

    image.onload = () => {
      editor.baseImage = image;
      adoptCanvasImageDimensions(editor, dom, image, { matchSize: !editor.importHadDims });
      renderCanvasEditor(editor);
      updateCanvasBodyPreview(editor, dom);
      if (!attempt.crossOrigin) {
        dom.stats.textContent = "Image loaded for editing preview. Export will keep the original image URL because CORS is blocked.";
      }
    };

    image.onerror = () => {
      index += 1;
      tryLoad();
    };

    image.src = attempt.url;
  };

  tryLoad();
}

function createCanvasImageLoadAttempts(url) {
  return [
    { url, crossOrigin: "anonymous" },
    { url, crossOrigin: "" },
  ];
}

function loadCanvasImageUrl(url, editor, dom) {
  const image = new Image();
  image.crossOrigin = "anonymous";
  image.onload = () => {
    editor.baseImage = image;
    adoptCanvasImageDimensions(editor, dom, image, { matchSize: !editor.importHadDims });
    renderCanvasEditor(editor);
    updateCanvasBodyPreview(editor, dom);
  };
  image.onerror = () => {
    dom.stats.textContent = "圖片連結無法載入到畫布，Body 仍會使用原圖 URL。";
    updateCanvasBodyPreview(editor, dom);
  };
  image.src = url;
}

function runColorPicker(dom) {
  if ("EyeDropper" in window) {
    dom.stats.textContent = "Pick a color from the screen.";
    new window.EyeDropper().open()
      .then((result) => {
        dom.color.value = result.sRGBHex;
        dom.stats.textContent = `Picked ${result.sRGBHex}`;
      })
      .catch(() => {
        dom.stats.textContent = "Color pick cancelled.";
      });
  } else {
    dom.color?.click();
  }
}

function handleCanvasPointerDown(event, editor, dom) {
  const point = canvasEventPoint(event, editor.canvas);
  editor.canvas.setPointerCapture?.(event.pointerId);

  if (editor.tool === "move" && editor.baseImage) {
    const handle = hitTestCanvasHandle(point, editor);
    if (handle) {
      editor.pointer = {
        mode: handle,
        startX: point.x,
        startY: point.y,
        transformX: editor.transform.x,
        transformY: editor.transform.y,
        scale: editor.transform.scale,
        rotation: editor.transform.rotation,
      };
      return;
    }
  }

  editor.pointer = {
    mode: editor.tool === "move" ? "pan" : "stroke",
    x: point.x,
    y: point.y,
    startX: point.x,
    startY: point.y,
    transformX: editor.transform.x,
    transformY: editor.transform.y,
    scale: editor.transform.scale,
    rotation: editor.transform.rotation,
  };

  if (editor.tool !== "move") {
    drawCanvasStroke(point, point, editor, dom);
    renderCanvasEditor(editor);
    updateCanvasBodyPreview(editor, dom);
  }
}

function handleCanvasPointerMove(event, editor, dom) {
  const point = canvasEventPoint(event, editor.canvas);

  if (!editor.pointer) {
    if (editor.tool === "move" && editor.baseImage) {
      const handle = hitTestCanvasHandle(point, editor);
      editor.canvas.style.cursor = canvasCursorForHandle(handle);
    } else {
      editor.canvas.style.cursor = "";
    }
    return;
  }

  const p = editor.pointer;

  if (p.mode === "pan") {
    editor.transform.x = Math.round(p.transformX + point.x - p.startX);
    editor.transform.y = Math.round(p.transformY + point.y - p.startY);
    syncCanvasTransformInputs(editor, dom);
    renderCanvasEditor(editor);
    return;
  }

  if (p.mode === "rotate") {
    const a0 = Math.atan2(p.startY - p.transformY, p.startX - p.transformX);
    const a1 = Math.atan2(point.y - p.transformY, point.x - p.transformX);
    let next = p.rotation + ((a1 - a0) * 180) / Math.PI;
    next = ((next + 180) % 360 + 360) % 360 - 180;
    editor.transform.rotation = Math.round(next);
    syncCanvasTransformInputs(editor, dom);
    renderCanvasEditor(editor);
    return;
  }

  if (p.mode && p.mode.startsWith("scale-")) {
    const startDist = Math.hypot(p.startX - p.transformX, p.startY - p.transformY);
    const curDist = Math.hypot(point.x - p.transformX, point.y - p.transformY);
    if (startDist > 0) {
      const next = Math.max(0.05, Math.min(4, (p.scale * curDist) / startDist));
      editor.transform.scale = next;
      syncCanvasTransformInputs(editor, dom);
      renderCanvasEditor(editor);
    }
    return;
  }

  if (p.mode === "stroke") {
    drawCanvasStroke({ x: p.x, y: p.y }, point, editor, dom);
    editor.pointer = { ...p, x: point.x, y: point.y };
    renderCanvasEditor(editor);
    updateCanvasBodyPreview(editor, dom);
  }
}

function syncCanvasTransformInputs(editor, dom) {
  dom.x.value = String(editor.transform.x);
  dom.y.value = String(editor.transform.y);
  dom.scale.value = String(Math.round(editor.transform.scale * 100));
  dom.rotation.value = String(editor.transform.rotation);
}

function getCanvasImageCorners(editor) {
  const img = editor.baseImage;
  if (!img) return null;
  const { x, y, scale, rotation } = editor.transform;
  const hw = (img.width * scale) / 2;
  const hh = (img.height * scale) / 2;
  const rad = (rotation * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const local = [
    { lx: -hw, ly: -hh },
    { lx: hw, ly: -hh },
    { lx: hw, ly: hh },
    { lx: -hw, ly: hh },
  ];
  const corners = local.map(({ lx, ly }) => ({
    x: x + lx * cos - ly * sin,
    y: y + lx * sin + ly * cos,
  }));
  const rotateHandle = {
    x: x + 0 * cos - (-hh - 36) * sin,
    y: y + 0 * sin + (-hh - 36) * cos,
  };
  return { corners, rotateHandle, hw, hh, cos, sin };
}

function hitTestCanvasHandle(point, editor) {
  const info = getCanvasImageCorners(editor);
  if (!info) return null;
  const tolerance = 14;
  if (Math.hypot(point.x - info.rotateHandle.x, point.y - info.rotateHandle.y) <= tolerance) {
    return "rotate";
  }
  for (let i = 0; i < info.corners.length; i += 1) {
    const c = info.corners[i];
    if (Math.hypot(point.x - c.x, point.y - c.y) <= tolerance) {
      return `scale-${i}`;
    }
  }
  const { x, y } = editor.transform;
  const dx = point.x - x;
  const dy = point.y - y;
  const lx = dx * info.cos + dy * info.sin;
  const ly = -dx * info.sin + dy * info.cos;
  if (Math.abs(lx) <= info.hw && Math.abs(ly) <= info.hh) {
    return null;
  }
  return null;
}

function canvasCursorForHandle(handle) {
  if (!handle) return "";
  if (handle === "rotate") return "grab";
  return "nwse-resize";
}

function drawCanvasStroke(from, to, editor, dom) {
  const size = Number(dom.brushSize.value || 20);
  const opacity = Number(dom.brushOpacity.value || 70) / 100;
  const strokeMode = getCanvasStrokeMode(editor);
  if (strokeMode === "none") return;
  let target;
  if (editor.layer === "mask") {
    target = editor.maskCtx;
  } else {
    const active = getActiveLayer(editor);
    if (!active.visible) {
      active.visible = true;
      renderLayerPanel(editor, dom);
    }
    target = active.ctx;
  }

  target.save();
  target.lineCap = "round";
  target.lineJoin = "round";
  target.lineWidth = size;

  if (strokeMode === "paint-remove" || strokeMode === "mask-remove") {
    target.globalCompositeOperation = "destination-out";
    target.strokeStyle = "rgba(0,0,0,1)";
  } else if (strokeMode === "mask-add") {
    target.globalAlpha = opacity;
    target.strokeStyle = "#ffffff";
  } else {
    target.globalAlpha = opacity;
    target.strokeStyle = dom.color.value || "#d9480f";
  }

  target.beginPath();
  target.moveTo(from.x, from.y);
  target.lineTo(to.x, to.y);
  target.stroke();
  target.restore();
}

function getCanvasStrokeMode(editor) {
  if (editor.tool !== "brush" && editor.tool !== "eraser") return "none";
  const layer = editor.layer === "mask" ? "mask" : "paint";
  return `${layer}-${editor.tool === "eraser" ? "remove" : "add"}`;
}

function renderCanvasEditor(editor) {
  const { ctx, canvas } = editor;
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  drawCanvasGrid(ctx, canvas.width, canvas.height);

  if (editor.baseImage) {
    ctx.save();
    ctx.translate(editor.transform.x, editor.transform.y);
    ctx.rotate((editor.transform.rotation * Math.PI) / 180);
    ctx.scale(editor.transform.scale, editor.transform.scale);
    ctx.drawImage(editor.baseImage, -editor.baseImage.width / 2, -editor.baseImage.height / 2);
    ctx.restore();
  }

  editor.layers.forEach((layer) => {
    if (layer.visible) ctx.drawImage(layer.canvas, 0, 0);
  });

  if (editor.layer === "mask") {
    ctx.save();
    ctx.globalAlpha = 0.72;
    ctx.globalCompositeOperation = "source-over";
    ctx.drawImage(tintCanvas(editor.maskCanvas, "#111111"), 0, 0);
    ctx.restore();
  }

  if (editor.tool === "move" && editor.baseImage) {
    drawCanvasTransformHandles(ctx, editor);
  }
}

function drawCanvasTransformHandles(ctx, editor) {
  const info = getCanvasImageCorners(editor);
  if (!info) return;
  const { corners, rotateHandle } = info;

  ctx.save();
  ctx.strokeStyle = "#c44f2b";
  ctx.lineWidth = 1.5;
  ctx.setLineDash([6, 4]);
  ctx.beginPath();
  ctx.moveTo(corners[0].x, corners[0].y);
  for (let i = 1; i < corners.length; i += 1) ctx.lineTo(corners[i].x, corners[i].y);
  ctx.closePath();
  ctx.stroke();
  ctx.setLineDash([]);

  const topMid = {
    x: (corners[0].x + corners[1].x) / 2,
    y: (corners[0].y + corners[1].y) / 2,
  };
  ctx.beginPath();
  ctx.moveTo(topMid.x, topMid.y);
  ctx.lineTo(rotateHandle.x, rotateHandle.y);
  ctx.stroke();

  corners.forEach((c) => {
    ctx.fillStyle = "#fff7f1";
    ctx.beginPath();
    ctx.arc(c.x, c.y, 6, 0, Math.PI * 2);
    ctx.fill();
    ctx.stroke();
  });

  ctx.fillStyle = "#fff7f1";
  ctx.beginPath();
  ctx.arc(rotateHandle.x, rotateHandle.y, 7, 0, Math.PI * 2);
  ctx.fill();
  ctx.stroke();
  ctx.restore();
}

function drawCanvasGrid(ctx, width, height) {
  ctx.fillStyle = "#f9f5ed";
  ctx.fillRect(0, 0, width, height);
  ctx.strokeStyle = "rgba(36, 22, 15, 0.06)";
  ctx.lineWidth = 1;
  for (let x = 0; x <= width; x += 32) {
    ctx.beginPath();
    ctx.moveTo(x, 0);
    ctx.lineTo(x, height);
    ctx.stroke();
  }
  for (let y = 0; y <= height; y += 32) {
    ctx.beginPath();
    ctx.moveTo(0, y);
    ctx.lineTo(width, y);
    ctx.stroke();
  }
}

function tintCanvas(source, color) {
  const tinted = document.createElement("canvas");
  tinted.width = source.width;
  tinted.height = source.height;
  const ctx = tinted.getContext("2d");
  ctx.drawImage(source, 0, 0);
  ctx.globalCompositeOperation = "source-in";
  ctx.fillStyle = color;
  ctx.fillRect(0, 0, tinted.width, tinted.height);
  return tinted;
}

function canvasEventPoint(event, canvas) {
  const rect = canvas.getBoundingClientRect();
  return {
    x: Math.round((event.clientX - rect.left) * (canvas.width / rect.width)),
    y: Math.round((event.clientY - rect.top) * (canvas.height / rect.height)),
  };
}

function buildCanvasEditorDraftFromDom(editor, dom) {
  return buildCanvasEditorRequestDraft({
    mode: dom.mode.value,
    prompt: dom.prompt.value,
    negativePrompt: dom.negativePrompt.value,
    imageDataUrl: exportCanvasImage(editor),
    maskDataUrl: editor.maskCanvas.toDataURL("image/png"),
    imageId: editor.imageId,
    remixPostId: editor.taskId,
    remixPostImageId: editor.imageId,
    width: dom.width.value,
    height: dom.height.value,
    steps: dom.steps.value,
    cfgScale: dom.cfgScale.value,
    seed: dom.seed.value,
    denoisingStrength: dom.denoisingStrength.value,
    modelId: dom.modelId.value,
    modelFileId: dom.modelFileId.value,
    guidance: 3.5,
    clipSkip: 2,
    ksamplerName: editor.ksamplerName || "euler_ancestral",
    schedule: editor.schedule || "normal",
  });
}

function exportCanvasImage(editor) {
  if (!editor.baseImage && editor.sourceImageUrl) {
    return editor.sourceImageUrl;
  }

  const output = document.createElement("canvas");
  output.width = editor.canvas.width;
  output.height = editor.canvas.height;
  const ctx = output.getContext("2d");
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, output.width, output.height);

  if (editor.baseImage) {
    ctx.save();
    ctx.translate(editor.transform.x, editor.transform.y);
    ctx.rotate((editor.transform.rotation * Math.PI) / 180);
    ctx.scale(editor.transform.scale, editor.transform.scale);
    ctx.drawImage(editor.baseImage, -editor.baseImage.width / 2, -editor.baseImage.height / 2);
    ctx.restore();
  }

  editor.layers.forEach((layer) => {
    if (layer.visible) ctx.drawImage(layer.canvas, 0, 0);
  });
  try {
    return output.toDataURL("image/png");
  } catch {
    return editor.sourceImageUrl || editor.imageDataUrl || "";
  }
}

function updateCanvasBodyPreview(editor, dom) {
  const draft = buildCanvasEditorDraftFromDom(editor, dom);
  dom.body.value = JSON.stringify(draft.body, null, 2);
}

function readCanvasColorAtPoint(ctx, point) {
  try {
    const pixel = ctx.getImageData(point.x, point.y, 1, 1).data;
    return { color: rgbToHex(pixel[0], pixel[1], pixel[2]), error: "" };
  } catch (error) {
    return {
      color: "",
      error: "Cannot pick this pixel because the image blocks CORS. Use a local image or a CORS-enabled source.",
    };
  }
}

function rgbToHex(r, g, b) {
  return `#${[r, g, b].map((value) => value.toString(16).padStart(2, "0")).join("")}`;
}

function castNumber(value, fallback) {
  if (value === undefined || value === null || value === "") return fallback;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : fallback;
}

function timestamp() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  const hour = String(now.getHours()).padStart(2, "0");
  const minute = String(now.getMinutes()).padStart(2, "0");
  const second = String(now.getSeconds()).padStart(2, "0");
  return `${year}${month}${day}-${hour}${minute}${second}`;
}

async function embedMetadataInOriginal(blob, metadata) {
  if (blob.type === "image/png") return embedMetadataInPng(blob, metadata);
  if (blob.type === "image/jpeg" || blob.type === "image/jpg") return embedMetadataInJpeg(blob, metadata);
  return blob;
}

async function embedMetadataInPng(blob, metadata) {
  const original = new Uint8Array(await blob.arrayBuffer());
  if (!matchesSignature(original, PNG_SIGNATURE)) {
    throw new Error("檔案不是有效 PNG");
  }

  const payload = new TextEncoder().encode(`AITestMetadata\0${JSON.stringify(metadata)}`);
  const chunk = createPngChunk("tEXt", payload);
  const iendOffset = findPngIendOffset(original);
  const merged = new Uint8Array(original.length + chunk.length);
  merged.set(original.slice(0, iendOffset), 0);
  merged.set(chunk, iendOffset);
  merged.set(original.slice(iendOffset), iendOffset + chunk.length);
  return new Blob([merged], { type: "image/png" });
}

function createPngChunk(type, data) {
  const typeBytes = new TextEncoder().encode(type);
  const chunk = new Uint8Array(4 + 4 + data.length + 4);
  chunk.set(uint32ToBytes(data.length), 0);
  chunk.set(typeBytes, 4);
  chunk.set(data, 8);
  const crcSource = new Uint8Array(typeBytes.length + data.length);
  crcSource.set(typeBytes, 0);
  crcSource.set(data, typeBytes.length);
  chunk.set(uint32ToBytes(crc32(crcSource)), 8 + data.length);
  return chunk;
}

function findPngIendOffset(bytes) {
  let offset = PNG_SIGNATURE.length;
  while (offset < bytes.length) {
    const length = bytesToUint32(bytes.slice(offset, offset + 4));
    const type = new TextDecoder().decode(bytes.slice(offset + 4, offset + 8));
    if (type === "IEND") return offset;
    offset += 12 + length;
  }
  throw new Error("找不到 PNG IEND chunk");
}

async function embedMetadataInJpeg(blob, metadata) {
  const original = new Uint8Array(await blob.arrayBuffer());
  if (original[0] !== JPEG_SOI[0] || original[1] !== JPEG_SOI[1]) {
    throw new Error("檔案不是有效 JPEG");
  }

  const comment = new TextEncoder().encode(`AITestMetadata:${JSON.stringify(metadata)}`);
  const segment = new Uint8Array(comment.length + 4);
  segment[0] = 0xff;
  segment[1] = 0xfe;
  segment[2] = ((comment.length + 2) >> 8) & 0xff;
  segment[3] = (comment.length + 2) & 0xff;
  segment.set(comment, 4);

  const merged = new Uint8Array(original.length + segment.length);
  merged.set(original.slice(0, 2), 0);
  merged.set(segment, 2);
  merged.set(original.slice(2), 2 + segment.length);
  return new Blob([merged], { type: "image/jpeg" });
}

async function readMetadataFromImage(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  if (matchesSignature(bytes, PNG_SIGNATURE)) return readMetadataFromPng(bytes);
  if (bytes[0] === JPEG_SOI[0] && bytes[1] === JPEG_SOI[1]) return readMetadataFromJpeg(bytes);
  throw new Error("只支援 PNG 或 JPEG");
}

function readMetadataFromPng(bytes) {
  let offset = PNG_SIGNATURE.length;
  while (offset < bytes.length) {
    const length = bytesToUint32(bytes.slice(offset, offset + 4));
    const type = new TextDecoder().decode(bytes.slice(offset + 4, offset + 8));
    if (type === "tEXt") {
      const payload = new TextDecoder().decode(bytes.slice(offset + 8, offset + 8 + length));
      if (payload.startsWith("AITestMetadata\0")) {
        return JSON.parse(payload.slice("AITestMetadata\0".length));
      }
    }
    offset += 12 + length;
  }
  throw new Error("圖片內找不到 AITestMetadata");
}

function readMetadataFromJpeg(bytes) {
  let offset = 2;
  while (offset < bytes.length) {
    if (bytes[offset] !== 0xff) {
      offset += 1;
      continue;
    }
    const marker = bytes[offset + 1];
    if (marker === 0xd9 || marker === 0xda) break;
    const length = (bytes[offset + 2] << 8) + bytes[offset + 3];
    if (marker === 0xfe) {
      const payload = new TextDecoder().decode(bytes.slice(offset + 4, offset + 2 + length));
      if (payload.startsWith("AITestMetadata:")) {
        return JSON.parse(payload.slice("AITestMetadata:".length));
      }
    }
    offset += 2 + length;
  }
  throw new Error("圖片內找不到 AITestMetadata");
}

function matchesSignature(bytes, signature) {
  return signature.every((value, index) => bytes[index] === value);
}

function bytesToUint32(bytes) {
  return (((bytes[0] << 24) >>> 0) + (bytes[1] << 16) + (bytes[2] << 8) + bytes[3]) >>> 0;
}

function uint32ToBytes(value) {
  return new Uint8Array([
    (value >>> 24) & 0xff,
    (value >>> 16) & 0xff,
    (value >>> 8) & 0xff,
    value & 0xff,
  ]);
}

function crc32(bytes) {
  let crc = 0xffffffff;
  for (let index = 0; index < bytes.length; index += 1) {
    crc ^= bytes[index];
    for (let bit = 0; bit < 8; bit += 1) {
      const mask = -(crc & 1);
      crc = (crc >>> 1) ^ (0xedb88320 & mask);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function classifyImageAccessFailure(url, error) {
  const signedExpiry = getSignedUrlExpiry(url);
  const now = new Date();

  if (signedExpiry && now > signedExpiry) {
    return `圖片連結已過期（${signedExpiry.toLocaleString("zh-TW", { hour12: false })}）`;
  }

  if (String(error?.message || "").includes("404")) {
    if (signedExpiry) {
      return `圖片回傳 404，且連結可能已過期（${signedExpiry.toLocaleString("zh-TW", { hour12: false })}）`;
    }
    return "圖片回傳 404，請重新查詢取得新連結";
  }

  return "圖片下載失敗，可能是 CORS 或連結失效";
}

function getSignedUrlExpiry(url) {
  try {
    const parsed = new URL(url);
    const signedAt = parsed.searchParams.get("X-Amz-Date");
    const expires = parsed.searchParams.get("X-Amz-Expires");
    if (!signedAt || !expires) return null;
    const match = signedAt.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z$/);
    if (!match) return null;
    const [, year, month, day, hour, minute, second] = match;
    const issuedAtUtc = Date.UTC(
      Number(year),
      Number(month) - 1,
      Number(day),
      Number(hour),
      Number(minute),
      Number(second),
    );
    return new Date(issuedAtUtc + Number(expires) * 1000);
  } catch {
    return null;
  }
}

function resolveExtension(mimeType) {
  if (mimeType === "image/png") return "png";
  if (mimeType === "image/jpeg" || mimeType === "image/jpg") return "jpg";
  return "bin";
}

function triggerDownload(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

function directBrowserDownload(url, fileName) {
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.target = "_blank";
  anchor.rel = "noopener";
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
}

async function fetchBlobForSave(url, mimeTypeHint) {
  let response;
  try {
    response = await fetch(url, { mode: "cors", credentials: "omit", cache: "reload" });
  } catch (error) {
    throw new Error(classifyImageAccessFailure(url, error));
  }

  if (!response.ok) {
    throw new Error(`下載失敗: HTTP ${response.status}`);
  }

  const blob = await response.blob();
  if (blob.type) return blob;
  return new Blob([await blob.arrayBuffer()], { type: mimeTypeHint || "application/octet-stream" });
}

async function saveBlobAsFile(blob, suggestedFileName, mimeTypeHint) {
  const fileName = sanitizeFileName(suggestedFileName);
  const mimeType = blob.type || mimeTypeHint || guessMimeFromFileName(fileName);

  if ("showSaveFilePicker" in window) {
    const handle = await window.showSaveFilePicker({
      suggestedName: fileName,
      types: buildPickerTypes(mimeType),
    });
    const writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();
    return;
  }

  triggerDownload(blob, fileName);
}

function sanitizeFileName(fileName) {
  const cleaned = String(fileName || "").trim().replace(/[\\/:*?"<>|]/g, "_");
  return cleaned || "download.bin";
}

function buildPickerTypes(mimeType) {
  if (!mimeType || mimeType === "application/octet-stream") {
    return [{
      description: "All files",
      accept: { "application/octet-stream": [".bin"] },
    }];
  }

  const ext = resolveExtension(mimeType);
  return [{
    description: `${mimeType} file`,
    accept: { [mimeType]: [`.${ext}`] },
  }];
}

function guessMimeFromFileName(fileName) {
  const lower = fileName.toLowerCase();
  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  if (lower.endsWith(".webp")) return "image/webp";
  return "application/octet-stream";
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&#39;");
}


