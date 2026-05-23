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
const SHARED_HEADER_NAMES = new Set([
  "x-echoing-env",
  "x-request-lang",
  "x-request-package-id",
  "x-request-package-sign-version",
  "x-request-sign",
  "x-request-sign-type",
  "x-request-sign-version",
  "x-request-timestamp",
]);
const REQUEST_KEYS = ["send", "query", "post"];
const API_URL_MATCHERS = {
  send: /\/works\/v1\/works\/task$/i,
  query: /\/works\/v1\/works\/tasks\/query$/i,
  post: /\/community-web\/v1\/post\/create$/i,
};
const DEFAULT_REQUESTS = {
  send: {
    url: "https://api.tensor.art/works/v1/works/task",
    method: "POST",
    headers: {
      accept: "*/*",
      "accept-language": "zh-TW,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6,ja;q=0.5",
    },
    body: {
      params: {
        baseModel: {
          modelId: "990778216270553015",
          modelFileId: "990778216270553016",
        },
        sdxl: {
          refiner: false,
        },
        models: [],
        embeddingModels: [],
        sdVae: "Automatic",
        prompt: "adult elf woman, sideless fantasy outfit, fantasy, elegant, beautiful, detailed fantasy illustration, non-explicit",
        negativePrompt: "bad quality,worst quality,worst detail,sketch,censor,bad body proportions ,",
        height: 1536,
        width: 1024,
        imageCount: 1,
        steps: 25,
        images: [],
        cfgScale: 5,
        seed: "-1",
        clipSkip: 2,
        etaNoiseSeedDelta: 31327,
        v1Clip: true,
        enablePix2pix: false,
        guidance: 3.5,
        useFirstLastFrame: false,
        samplerName: "Euler a",
      },
      credits: 1.31,
      taskType: "TXT2IMG",
      isRemix: false,
      captchaType: "CLOUDFLARE_TURNSTILE",
    },
  },
  query: {
    url: "https://api.tensor.art/works/v1/works/tasks/query",
    method: "POST",
    headers: {
      accept: "*/*",
      "accept-language": "zh-TW,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6,ja;q=0.5",
    },
    body: {
      size: 30,
      cursor: "0",
      returnAllTask: true,
      isSortAsc: false,
      startedAt: "0",
      endedAt: "0",
    },
  },
  post: {
    url: "https://api.tensor.art/community-web/v1/post/create",
    method: "POST",
    headers: {
      accept: "*/*",
      "accept-language": "zh-TW,zh;q=0.9",
    },
    body: {
      content: "",
      tags: [],
      title: "",
      channelId: "107",
      generationImageIds: [],
    },
  },
};

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

function getDefaultRequestState(key) {
  const defaults = DEFAULT_REQUESTS[key];
  const blank = blankRequestState();

  if (!defaults) return blank;

  return {
    ...blank,
    url: defaults.url,
    method: defaults.method,
    headers: JSON.parse(JSON.stringify(defaults.headers)),
    bodyText: JSON.stringify(defaults.body, null, 2),
  };
}

function blankCommonState() {
  return {
    powershell: "",
    headers: {},
  };
}

function loadState() {
  const base = {
    common: blankCommonState(),
    send: getDefaultRequestState("send"),
    query: getDefaultRequestState("query"),
    post: getDefaultRequestState("post"),
    selectedImageIds: [],
    galleryItems: [],
    importedMetadata: null,
    savedSettings: { send: [], query: [], post: [] },
  };

  const saved = localStorage.getItem(STORAGE_KEY);
  if (!saved) return base;

  try {
    const parsed = JSON.parse(saved);
    const loaded = {
      common: { ...blankCommonState(), ...(parsed.common || {}) },
      send: { ...getDefaultRequestState("send"), ...(parsed.send || {}) },
      query: { ...getDefaultRequestState("query"), ...(parsed.query || {}) },
      post: { ...getDefaultRequestState("post"), ...(parsed.post || {}) },
      selectedImageIds: Array.isArray(parsed.selectedImageIds) ? parsed.selectedImageIds : [],
      galleryItems: Array.isArray(parsed.galleryItems) ? parsed.galleryItems : [],
      importedMetadata: parsed.importedMetadata ?? null,
      savedSettings: parsed.savedSettings || { send: [], query: [], post: [] },
    };

    migrateSharedHeaders(loaded);
    return loaded;
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

function cloneCommonDraft(common) {
  return { ...blankCommonState(), ...JSON.parse(JSON.stringify(common || {})) };
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
    common: cloneCommonDraft(source.common),
    send: cloneRequestDraft(source.send),
    query: cloneRequestDraft(source.query),
    post: cloneRequestDraft(source.post),
    selectedImageIds: Array.isArray(source.selectedImageIds) ? [...source.selectedImageIds] : [],
    galleryItems: Array.isArray(source.galleryItems) ? JSON.parse(JSON.stringify(source.galleryItems)) : [],
    importedMetadata: source.importedMetadata ?? null,
    savedSettings: normalizeSavedSettings(source.savedSettings),
  };
}

function extractSharedHeaders(headers) {
  return Object.fromEntries(
    Object.entries(sanitizeHeaders(headers || {})).filter(([key, value]) => (
      SHARED_HEADER_NAMES.has(key.toLowerCase())
      && (key.toLowerCase() === "x-echoing-env" || String(value ?? "").trim())
    )),
  );
}

function getRequestSpecificHeaders(headers) {
  return Object.fromEntries(
    Object.entries(sanitizeHeaders(headers || {})).filter(([key]) => (
      !SHARED_HEADER_NAMES.has(key.toLowerCase())
      && key.toLowerCase() !== "authorization"
    )),
  );
}

function buildEffectiveHeaders(requestHeaders, sharedHeaders) {
  return sanitizeHeaders({
    ...getRequestSpecificHeaders(requestHeaders),
    ...extractSharedHeaders(sharedHeaders),
  });
}

function migrateSharedHeaders(loaded) {
  const existingSharedHeaders = extractSharedHeaders(loaded.common?.headers || {});
  const inferredHeaders = REQUEST_KEYS.reduce((headers, key) => ({
    ...headers,
    ...extractSharedHeaders(loaded[key]?.headers || {}),
  }), {});

  loaded.common = {
    ...blankCommonState(),
    ...(loaded.common || {}),
    headers: {
      ...inferredHeaders,
      ...existingSharedHeaders,
    },
  };

  if (!loaded.common.powershell) {
    const source = REQUEST_KEYS.map((key) => loaded[key]?.powershell).find(Boolean);
    loaded.common.powershell = source || "";
  }

  REQUEST_KEYS.forEach((key) => {
    loaded[key] = {
      ...blankRequestState(),
      ...(loaded[key] || {}),
      headers: getRequestSpecificHeaders(loaded[key]?.headers || {}),
    };
  });
}

function findRequestKeyForUrl(url) {
  return Object.entries(API_URL_MATCHERS).find(([, pattern]) => pattern.test(url || ""))?.[0] || null;
}

function applyParsedPowerShellToRequestState(request, common, parsed) {
  const sharedHeaders = extractSharedHeaders(parsed.headers);
  common.powershell = parsed.powershell;
  common.headers = {
    ...(common.headers || {}),
    ...sharedHeaders,
  };

  Object.assign(request, {
    powershell: parsed.powershell,
    url: parsed.url,
    method: parsed.method,
    headers: getRequestSpecificHeaders(parsed.headers),
    bodyText: parsed.bodyText,
  });
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
  bindCommonSignatureControls();
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

function bindCommonSignatureControls() {
  const dom = {
    powershell: document.querySelector("#common-powershell"),
    parseButton: document.querySelector("#common-parse"),
    headers: document.querySelector("#common-headers"),
    summary: document.querySelector("#common-summary"),
    status: document.querySelector("#common-status"),
  };

  if (!dom.powershell && !dom.headers) return;

  renderCommonSignatureControls(dom);

  dom.powershell?.addEventListener("input", (event) => {
    state.common.powershell = event.target.value;
    saveState();
  });

  dom.headers?.addEventListener("input", (event) => {
    try {
      state.common.headers = extractSharedHeaders(JSON.parse(event.target.value || "{}"));
    } catch {
      state.common.headers = {};
    }
    saveState();
    renderCommonSignatureControls(dom);
  });

  dom.parseButton?.addEventListener("click", () => {
    try {
      const parsed = parsePowerShellRequest(dom.powershell.value);
      const matchedKey = findRequestKeyForUrl(parsed.url);
      const sharedHeaders = extractSharedHeaders(parsed.headers);

      if (!Object.keys(sharedHeaders).length) {
        throw new Error("No TensorArt x-request signature headers found.");
      }

      state.common.powershell = parsed.powershell;
      state.common.headers = {
        ...state.common.headers,
        ...sharedHeaders,
      };

      if (matchedKey) {
        applyParsedPowerShellToRequestState(state[matchedKey], state.common, parsed);
        state[matchedKey].bodyText = formatJsonString(state[matchedKey].bodyText);
      }

      saveState();
      renderCommonSignatureControls(dom);
      if (matchedKey && page !== "metadata") {
        const section = bindRequestSection(matchedKey);
        if (section.body) renderRequestSection(matchedKey, section);
      }
      setCommonStatus(dom, matchedKey
        ? `Updated shared signature and ${matchedKey} request defaults.`
        : "Updated shared signature headers.");
    } catch (error) {
      setCommonStatus(dom, `Parse failed: ${error.message}`);
    }
  });
}

function renderCommonSignatureControls(dom) {
  if (dom.powershell) dom.powershell.value = state.common.powershell || "";
  if (dom.headers) dom.headers.value = JSON.stringify(extractSharedHeaders(state.common.headers), null, 2);
  if (dom.summary) {
    const count = Object.keys(extractSharedHeaders(state.common.headers)).length;
    dom.summary.textContent = count
      ? `Shared TensorArt signature headers: ${count} ready`
      : "Shared TensorArt signature headers: not set";
  }
}

function setCommonStatus(dom, text) {
  if (dom.status) dom.status.textContent = text;
}

function initSendPage() {
  const section = bindRequestSection("send");
  renderRequestSection("send", section);
  initializeRequestSectionVisibility("send", section);

  section.parseButton.addEventListener("click", () => {
    try {
      applyParsedPowerShellToRequestState(state.send, state.common, parsePowerShellRequest(section.powershell.value));
      state.send.bodyText = formatJsonString(state.send.bodyText);
      saveState();
      renderRequestSection("send", section);
      collapseRequestSection(section);
    } catch (error) {
      setResponse("send", `閫??憭望?: ${error.message}`);
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
  };

  renderRequestSection("query", query);
  renderRequestSection("post", post);
  renderGallery(gallery);
  initializeRequestSectionVisibility("query", query);
  initializeRequestSectionVisibility("post", post);

  query.parseButton.addEventListener("click", () => {
    try {
      applyParsedPowerShellToRequestState(state.query, state.common, parsePowerShellRequest(query.powershell.value));
      state.query.bodyText = formatJsonString(state.query.bodyText);
      saveState();
      renderRequestSection("query", query);
      collapseRequestSection(query);
    } catch (error) {
      setResponse("query", `閫??憭望?: ${error.message}`);
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
      applyParsedPowerShellToRequestState(state.post, state.common, parsePowerShellRequest(post.powershell.value));
      state.post.bodyText = formatJsonString(state.post.bodyText);
      syncGenerationImageIds();
      saveState();
      renderRequestSection("post", post);
      renderGallery(gallery);
      collapseRequestSection(post);
    } catch (error) {
      setResponse("post", `閫??憭望?: ${error.message}`);
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

  gallery.postSync.addEventListener("click", () => {
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
      applyParsedPowerShellToRequestState(state[key], state.common, parsePowerShellRequest(section.powershell.value));
      state[key].bodyText = formatJsonString(state[key].bodyText);
      saveState();
      renderRequestSection(key, section);
      collapseRequestSection(section);
      afterParse();
    } catch (error) {
      setResponse(key, `閫??憭望?: ${error.message}`);
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
    : "No metadata loaded";

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
      applyMetadataToSendBody(state.importedMetadata, { keepSeed: true });
      output.textContent = `${JSON.stringify(state.importedMetadata, null, 2)}\n\nApplied metadata to Send Request Body.`;
    } catch (error) {
      output.textContent = `Apply failed: ${error.message}`;
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
      alert("Preset updated.");
    }
  });

  section.presetSaveAs.addEventListener("click", () => {
    const name = prompt("Preset name");
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
    alert("Preset saved.");
  });

  section.presetDelete.addEventListener("click", () => {
    if (!state[key].presetId) return;
    if (!confirm("確定要刪除此預設嗎?")) return;
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
    setResponse(key, "送出中..");
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
    setResponse(key, formatRequestError(error, globalThis.location?.href || ""));
    renderRequestSection(key, section);
    section.responseFold.open = false;
    return null;
  }
}

function formatRequestError(error, pageUrl = "") {
  const message = error?.message || String(error);
  const hints = [];
  const isLocalPage = /^https?:\/\/(127\.0\.0\.1|localhost)(:\d+)?\//i.test(pageUrl);

  if (/failed to fetch/i.test(message)) {
    if (isLocalPage) {
      hints.push("代理伺服器似乎沒在跑。請在專案目錄執行 `node server.js`,然後從 http://localhost:8787/index.html 開啟此頁面。");
    } else {
      hints.push("瀏覽器無法直接呼叫 api.tensor.art(CORS 政策)。請在本機跑 `node server.js`,透過 http://localhost:8787 開啟頁面以使用內建代理。");
    }
    hints.push("另一個可能:PowerShell 簽章已過期,請重新從 TensorArt 複製一份 PowerShell 並重新解析。");
  }

  return [`Request failed: ${message}`, ...hints].join("\n");
}

function buildFetchRequest(key) {
  const request = state[key];
  const headers = buildEffectiveHeaders(request.headers, state.common.headers);
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
  dom.selectedCount.forEach(el => { el.textContent = String(state.selectedImageIds.length); });

  if (!state.galleryItems.length) {
    dom.stats.textContent = "No gallery images yet.";
    dom.root.innerHTML = "";
    return;
  }

  dom.stats.textContent = `Loaded ${state.galleryItems.length} images. Selected images will be written to post generationImageIds.`;
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
            <button type="button" data-action="open-original" data-index="${index}">開啟原圖</button>
            <button type="button" data-action="apply-to-send" data-index="${index}">套用到 Send Body</button>
            <label style="display:inline-flex; align-items:center; gap:0.25rem; font-size:0.875rem;">
              <input type="checkbox" class="keep-seed" data-index="${index}" checked>
              沿用 seed
            </label>
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
        else reject(new Error("Cannot convert to JPG"));
      }, "image/jpeg", 0.95);
      URL.revokeObjectURL(img.src);
    };
    img.onerror = () => reject(new Error("Image load failed"));
    img.src = URL.createObjectURL(blob);
  });
}

function extractBlobFromImageElement(imgElement, targetMimeType) {
  return new Promise((resolve, reject) => {
    try {
      if (!imgElement || !imgElement.src) {
        reject(new Error("Image element is missing a source"));
        return;
      }
      // 建立新的 Image 物件並設定 crossorigin,避免重用 gallery 中既有的 img
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
            else reject(new Error("Cannot export canvas image"));
          }, targetMimeType, 0.95);
        } catch (e) {
          reject(e);
        }
      };
      corsImg.onerror = () => reject(new Error("Image failed CORS loading"));
      corsImg.src = imgElement.src;
    } catch (e) {
      reject(e);
    }
  });
}

async function handleGalleryAction(action, index, dom) {
  const entry = state.galleryItems[index];
  if (!entry) return;

  if (action === "apply-to-send") {
    const keepSeedInput = dom.root.querySelector(`input.keep-seed[data-index="${index}"]`);
    const keepSeed = keepSeedInput ? keepSeedInput.checked : true;
    try {
      applyMetadataToSendBody(entry.metadata, { keepSeed });
      const seedLabel = keepSeed ? `seed=${entry.metadata.seed || "-"}` : "seed=-1 (random)";
      dom.stats.textContent = `已套用 Task ${entry.taskId} 的設定到 Send Body (${seedLabel})。`;
    } catch (error) {
      dom.stats.textContent = `套用失敗: ${error.message}`;
    }
    return;
  }

  const signedExpiry = getSignedUrlExpiry(entry.url);
  if (signedExpiry && new Date() > signedExpiry) {
    dom.stats.textContent = `Signed image URL expired at ${signedExpiry.toLocaleTimeString("zh-TW", { hour12: false })}. Refresh the query API.`;
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
    
    dom.stats.textContent = "正在處理圖片,請稍候..";
    
    let blob;
    try {
      // 優先嘗試從已載入的 img 物件中匯出位元組,避免重新 fetch()
      const imgElement = dom.root.querySelectorAll('.gallery-card img')[index];
      const targetMimeType = action === "download-jpg" ? "image/jpeg" : mimeType;
      blob = await extractBlobFromImageElement(imgElement, targetMimeType);
    } catch (e) {
      // 備援方案:用 fetch 重新下載
      dom.stats.textContent = "無法從圖片元素匯出,改用 fetch 重新下載..";
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
    
    dom.stats.textContent = `已儲存為 ${suggestedFileName}`;
  } catch (error) {
    if (error.name === "AbortError") return;
    dom.stats.textContent = `下載失敗:${error.message} (建議重新呼叫 Query API 取得新網址)`;
  }
}
function parsePowerShellRequest(text) {
  if (!text.trim()) {
    throw new Error("請貼上 PowerShell 內容");
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
    throw new Error(`找不到符合的欄位: ${pattern}`);
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
  if (!text) return "無內容";
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
    return "API not configured";
  }

  const method = (request.method || "POST").toUpperCase();
  const url = request.url.trim();
  if (!url) {
    return "PowerShell not parsed";
  }

  const preview = `${method} ${url}`;
  return preview.length > 72 ? `${preview.slice(0, 69)}...` : preview;
}

function applyMetadataToSendBody(meta, { keepSeed = true } = {}) {
  const payload = JSON.parse(state.send.bodyText || "{}");
  const params = payload.params || {};
  params.prompt = meta.prompt ?? params.prompt ?? "";
  params.negativePrompt = meta.negativePrompt ?? params.negativePrompt ?? "";
  params.steps = castNumber(meta.steps, params.steps);
  params.cfgScale = castNumber(meta.cfgScale, params.cfgScale);
  params.guidance = castNumber(meta.guidance, params.guidance);
  params.clipSkip = castNumber(meta.clipSkip, params.clipSkip);
  params.seed = keepSeed ? String(meta.seed ?? params.seed ?? "-1") : "-1";
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

  const sendBody = document.querySelector("#send-body");
  if (sendBody) sendBody.value = state.send.bodyText;
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
    throw new Error("檔案不是 PNG");
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
  throw new Error("?曆???PNG IEND chunk");
}

async function embedMetadataInJpeg(blob, metadata) {
  const original = new Uint8Array(await blob.arrayBuffer());
  if (original[0] !== JPEG_SOI[0] || original[1] !== JPEG_SOI[1]) {
    throw new Error("檔案不是 JPEG");
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
  throw new Error("不支援的格式,僅支援 PNG 或 JPEG");
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
  throw new Error("找不到 AITestMetadata");
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
  throw new Error("找不到 AITestMetadata");
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
    return `Signed image URL expired at ${signedExpiry.toLocaleString("zh-TW", { hour12: false })}.`;
  }

  if (String(error?.message || "").includes("404")) {
    if (signedExpiry) {
      return `Image URL returned 404. Signed URL may have expired at ${signedExpiry.toLocaleString("zh-TW", { hour12: false })}.`;
    }
    return "Image URL returned 404. Refresh the query API.";
  }

  return "Image access failed. It may be CORS or an expired signed URL.";
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


