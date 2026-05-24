# Canvas pre_sign Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a Canvas-top pre_sign experiment panel that parses Tensor PowerShell, sends the pre_sign request, and displays the response.

**Architecture:** Reuse the existing request-section abstraction in `app.js` for the new `presign` key. Store `presign` alongside existing request drafts in localStorage snapshots, and render the new controls in `canvas.html` above the editor workbench.

**Tech Stack:** Static HTML, CSS, vanilla JavaScript modules, Node test runner.

---

### Task 1: Lock the State and Parser Contract

**Files:**
- Modify: `tests/app-helpers.test.mjs`
- Modify: `app.js`

- [ ] Write a failing test asserting `state.presign` exists, snapshots preserve it, and a pre_sign PowerShell request parses into sanitized headers and JSON body text.
- [ ] Run `npm test` and confirm the new test fails because `state.presign` and snapshot support are missing.
- [ ] Add `presign: blankRequestState()` to base state, saved-state merge, snapshot export, and snapshot import.
- [ ] Run `npm test` and confirm the state/parser test passes.

### Task 2: Add Canvas UI and Bindings

**Files:**
- Modify: `canvas.html`
- Modify: `app.js`
- Modify: `styles.css`

- [ ] Add a `pre_sign` panel above `.canvas-workbench` with ids matching `bindRequestSection("presign")`.
- [ ] In `initCanvasPage`, bind `presign` with parse, send, format, inputs, and initial render/visibility.
- [ ] Add compact CSS so the panel sits above Canvas without crowding the editor.
- [ ] Run `npm test`.

### Task 3: Verify Locally, Commit, Push, and Verify GitHub Pages

**Files:**
- No additional source files.

- [ ] Serve the static app locally and verify Canvas shows the pre_sign panel.
- [ ] Run `npm test`.
- [ ] Commit the changes.
- [ ] Push `main` to `origin`.
- [ ] Verify the deployed GitHub Pages URL loads the updated UI.
