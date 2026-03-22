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
  const payload = {
    exportedAt: new Date().toISOString(),
    storageKey: STORAGE_KEY,
    data: state,
  };
  triggerDownload(new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" }), `tensor-api-qa-${timestamp()}.json`);
}

async function importStorageSnapshot(event) {
  const [file] = event.target.files || [];
  if (!file) return;

  try {
    const parsed = JSON.parse(await file.text());
    const incoming = parsed.data ?? parsed;
    Object.assign(state.send, { ...blankRequestState(), ...(incoming.send || {}) });
    Object.assign(state.query, { ...blankRequestState(), ...(incoming.query || {}) });
    Object.assign(state.post, { ...blankRequestState(), ...(incoming.post || {}) });
    state.selectedImageIds = Array.isArray(incoming.selectedImageIds) ? incoming.selectedImageIds : [];
    state.galleryItems = Array.isArray(incoming.galleryItems) ? incoming.galleryItems : [];
    state.importedMetadata = incoming.importedMetadata ?? null;
    state.savedSettings = incoming.savedSettings || { send: [], query: [], post: [] };
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
    selectedCount: document.querySelector("#selected-count"),
    postClearIds: document.querySelector("#post-clear-ids"),
  };

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
    if (state.post.clearOnSubmit) {
      state.selectedImageIds = [];
      syncGenerationImageIds();
      renderGallery(gallery);
    }
    renderRequestSection("post", post);
    await submitRequestSection("post", post, false);
  });

  send.formatButton.addEventListener("click", () => formatBody("send", send));
  query.formatButton.addEventListener("click", () => formatBody("query", query));
  post.formatButton.addEventListener("click", () => formatBody("post", post));

  gallery.postClearIds?.addEventListener("click", () => {
    state.selectedImageIds = [];
    syncGenerationImageIds();
    saveState();
    renderRequestSection("post", post);
    renderGallery(gallery);
  });
}

function initGalleryPage() {
  const query = bindRequestSection("query");
  const post = bindRequestSection("post");
  const gallery = {
    stats: document.querySelector("#gallery-stats"),
    root: document.querySelector("#gallery"),
    selectedCount: document.querySelector("#selected-count"),
    postClearIds: document.querySelector("#post-clear-ids"),
  };

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
    if (state.post.clearOnSubmit) {
      state.selectedImageIds = [];
      syncGenerationImageIds();
      renderGallery(gallery);
    }
    renderRequestSection("post", post);
    await submitRequestSection("post", post, false);
  });

  post.formatButton.addEventListener("click", () => {
    state.post.bodyText = formatJsonString(post.body.value);
    saveState();
    renderRequestSection("post", post);
  });

  gallery.postSync.addEventListener("click", () => {
    syncGenerationImageIds();
    saveState();
    renderRequestSection("post", post);
    renderGallery(gallery);
  });

  gallery.postClearIds?.addEventListener("click", () => {
    state.selectedImageIds = [];
    syncGenerationImageIds();
    saveState();
    renderRequestSection("post", post);
    renderGallery(gallery);
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
      const payload = JSON.parse(state.send.bodyText || "{}");
      const params = payload.params || {};
      const meta = state.importedMetadata;
      params.prompt = meta.prompt ?? params.prompt ?? "";
      params.negativePrompt = meta.negativePrompt ?? params.negativePrompt ?? "";
      params.steps = castNumber(meta.steps, params.steps);
      params.cfgScale = castNumber(meta.cfgScale, params.cfgScale);
      params.guidance = castNumber(meta.guidance, params.guidance);
      params.clipSkip = castNumber(meta.clipSkip, params.clipSkip);
      params.seed = String(meta.seed ?? params.seed ?? "-1");
      params.sdVae = meta.vae ?? params.sdVae ?? "Automatic";
      params.ksamplerName = meta.kSampler ?? params.ksamplerName ?? "";
      params.schedule = meta.schedule ?? params.schedule ?? "";
      params.width = castNumber(meta.width, params.width);
      params.height = castNumber(meta.height, params.height);

      if (meta.modelId && meta.modelFileId) {
        params.baseModel = {
          ...(params.baseModel || {}),
          modelId: String(meta.modelId),
          modelFileId: String(meta.modelFileId),
        };
      }

      payload.params = params;
      state.send.bodyText = JSON.stringify(payload, null, 2);
      saveState();
      output.textContent = `${JSON.stringify(state.importedMetadata, null, 2)}\n\n已套用到 Send Request Body。`;
    } catch (error) {
      output.textContent = `套用失敗: ${error.message}`;
    }
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
  } catch (error) {
    setResponse(key, `Request 失敗: ${error.message}`);
    renderRequestSection(key, section);
    section.responseFold.open = false;
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

function renderGallery(dom) {
  dom.selectedCount.textContent = String(state.selectedImageIds.length);

  if (!state.galleryItems.length) {
    dom.stats.textContent = "目前沒有查詢結果。";
    dom.root.innerHTML = "";
    return;
  }

  dom.stats.textContent = `共 ${state.galleryItems.length} 張圖片，可勾選後同步到 post 的 generationImageIds。`;
  dom.root.innerHTML = state.galleryItems.map((entry, index) => {
    const checked = state.selectedImageIds.includes(entry.generationImageId) ? "checked" : "";
    return `
      <article class="gallery-card">
        <img src="${escapeHtml(entry.url)}" alt="Task ${escapeHtml(entry.taskId)}">
        <div class="gallery-body">
          <div class="gallery-head">
            <h3 class="gallery-title">Task ${escapeHtml(entry.taskId)}</h3>
            <span class="pill">${escapeHtml(entry.status || "UNKNOWN")}</span>
          </div>
          <div class="gallery-selection">
            <input type="checkbox" id="pick-${index}" data-image-id="${escapeHtml(entry.generationImageId)}" ${checked}>
            <label for="pick-${index}">加入 generationImageIds</label>
          </div>
          <div class="gallery-meta">
            <span class="pill">Seed ${escapeHtml(String(entry.metadata.seed || "-"))}</span>
            <span class="pill">${escapeHtml(entry.metadata.size || "-")}</span>
          </div>
          <pre>${escapeHtml(JSON.stringify(entry.metadata, null, 2))}</pre>
          <div class="gallery-actions">
            <button type="button" data-action="download-original" data-index="${index}">另存原圖</button>
            <button type="button" data-action="download-jpg" data-index="${index}">另存為 JPG</button>
            <button type="button" data-action="open-original" data-index="${index}">開啟原圖</button>
          </div>
        </div>
      </article>
    `;
  }).join("");

  dom.root.querySelectorAll('input[type="checkbox"]').forEach((checkbox) => {
    checkbox.addEventListener("change", (event) => {
      toggleSelectedImageId(event.target.dataset.imageId, event.target.checked);
      renderGallery(dom);
    });
  });

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
      if (!imgElement || !imgElement.complete || imgElement.naturalHeight === 0) {
        reject(new Error("圖片尚未載入或無法存取"));
        return;
      }
      const canvas = document.createElement("canvas");
      canvas.width = imgElement.naturalWidth || imgElement.width;
      canvas.height = imgElement.naturalHeight || imgElement.height;
      const ctx = canvas.getContext("2d");
      
      if (targetMimeType === "image/jpeg" || targetMimeType === "image/jpg") {
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, canvas.width, canvas.height);
      }
      ctx.drawImage(imgElement, 0, 0);
      
      canvas.toBlob((blob) => {
        if (blob) resolve(blob);
        else reject(new Error("無法從 Canvas 取得影像資料"));
      }, targetMimeType, 0.95);
    } catch (e) {
      reject(e);
    }
  });
}

async function handleGalleryAction(action, index, dom) {
  const entry = state.galleryItems[index];
  if (!entry) return;

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
    
    // 如果是 HTTP 狀態碼錯誤或連結過期，切勿使用直接下載以免載下 XML 錯誤頁面
    if (error.message.includes("HTTP") || error.message.includes("已過期") || error.message.includes("403") || error.message.includes("404")) {
      dom.stats.textContent = `無法下載：${error.message} (建議重新 Query API 取得新連結)`;
    } else {
      directBrowserDownload(entry.url, originalFileName);
      dom.stats.textContent = `轉檔/存檔失敗，已改用瀏覽器直接下載：${error.message}`;
    }
  }
}
function parsePowerShellRequest(text) {
  if (!text.trim()) {
    throw new Error("請先貼上 PowerShell 內容");
  }

  const url = capture(text, /-Uri\s+"([\s\S]*?)"\s*`?\s*-Method/i);
  const method = capture(text, /-Method\s+"([^"]+)"/i);
  const bodyRaw = captureOptional(text, /-Body\s+"([\s\S]*?)"\s*$/i);
  const headersBlock = captureOptional(text, /-Headers\s+@\{([\s\S]*?)\}\s*`?\s*-ContentType/i);

  return {
    powershell: text,
    url: decodePowerShellString(url),
    method: decodePowerShellString(method).toUpperCase(),
    headers: sanitizeHeaders(parseHeaderBlock(headersBlock || "")),
    bodyText: decodePowerShellString(bodyRaw || ""),
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

function decodePowerShellString(value) {
  return value
    .replace(/`"/g, "\"")
    .replace(/``/g, "`")
    .replace(/`r/g, "\r")
    .replace(/`n/g, "\n")
    .replace(/`\$/g, "$");
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


