import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";

const source = fs.readFileSync(new URL("../app.js", import.meta.url), "utf8");

const localStorageStore = new Map();
const context = {
  Blob,
  URL: { createObjectURL: () => "blob:test", revokeObjectURL: () => {} },
  document: {
    body: { dataset: { page: "" } },
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
  extractSharedHeaders,
  getRequestSpecificHeaders,
  buildEffectiveHeaders,
  applyParsedPowerShellToRequestState,
  getDefaultRequestState,
  formatRequestError,
};
`, context);

const {
  extractSharedHeaders,
  getRequestSpecificHeaders,
  buildEffectiveHeaders,
  applyParsedPowerShellToRequestState,
  getDefaultRequestState,
  formatRequestError,
} = context.__testExports;

const plain = (value) => JSON.parse(JSON.stringify(value));

const mixedHeaders = {
  accept: "application/json",
  authorization: "Bearer secret",
  cookie: "SESSIONID=secret",
  "x-echoing-env": "production",
  "x-request-sign": "sig-1",
  "x-request-timestamp": "123",
  "x-request-lang": "en",
  "x-request-package-id": "pkg",
  "x-request-sign-version": "v1",
  "x-request-sign-type": "hmac",
  "x-request-package-sign-version": "p1",
};

assert.deepEqual(plain(extractSharedHeaders(mixedHeaders)), {
  "x-echoing-env": "production",
  "x-request-sign": "sig-1",
  "x-request-timestamp": "123",
  "x-request-lang": "en",
  "x-request-package-id": "pkg",
  "x-request-sign-version": "v1",
  "x-request-sign-type": "hmac",
  "x-request-package-sign-version": "p1",
});

assert.deepEqual(plain(extractSharedHeaders({ "x-echoing-env": "" })), {
  "x-echoing-env": "",
});

assert.deepEqual(plain(getRequestSpecificHeaders(mixedHeaders)), {
  accept: "application/json",
});

assert.deepEqual(
  plain(
  buildEffectiveHeaders(
    { accept: "application/json", "x-request-sign": "old" },
    { "x-request-sign": "fresh", "x-request-timestamp": "456" },
  ),
  ),
  {
    accept: "application/json",
    "x-request-sign": "fresh",
    "x-request-timestamp": "456",
  },
);

const request = {
  powershell: "",
  url: "",
  method: "POST",
  headers: {},
  bodyText: "",
};
const common = { powershell: "", headers: {} };
applyParsedPowerShellToRequestState(
  request,
  common,
  {
    powershell: "Invoke-WebRequest ...",
    url: "https://api.tensor.art/works/v1/works/tasks/query",
    method: "POST",
    headers: mixedHeaders,
    bodyText: '{"size":20}',
  },
);

assert.equal(request.url, "https://api.tensor.art/works/v1/works/tasks/query");
assert.equal(request.bodyText, '{"size":20}');
assert.deepEqual(plain(request.headers), { accept: "application/json" });
assert.equal(common.headers["x-request-sign"], "sig-1");
assert.equal(common.headers.authorization, undefined);

const defaultSend = getDefaultRequestState("send");
assert.equal(defaultSend.url, "https://api.tensor.art/works/v1/works/task");
assert.equal(defaultSend.method, "POST");
assert.equal(defaultSend.headers.accept, "*/*");
assert.match(defaultSend.bodyText, /adult elf woman/);
assert.doesNotMatch(defaultSend.bodyText, /authorization|x-request-sign|cookie/i);

const defaultQuery = getDefaultRequestState("query");
assert.equal(defaultQuery.url, "https://api.tensor.art/works/v1/works/tasks/query");
assert.match(defaultQuery.bodyText, /"size": 30/);

assert.match(
  formatRequestError(new TypeError("Failed to fetch"), "http://127.0.0.1:8788/index.html"),
  /GitHub Pages/,
);

console.log("app helper tests passed");
