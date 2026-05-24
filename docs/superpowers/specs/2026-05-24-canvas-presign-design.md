# Canvas pre_sign Design

## Goal

Add a pre_sign experiment panel above the Canvas workspace so the parsed Tensor API request can be tested before mask upload work is connected.

## Design

The Canvas page gets a dedicated `pre_sign` request section above the editor. It reuses the existing PowerShell parser, header sanitization, JSON body editor, request sender, and response preview flow already used by Dashboard request panels.

The feature stores its own draft in `state.presign`, separate from Dashboard API 1/2/3 and separate from the Canvas task request. The parsed response remains visible in the panel so the user can inspect whether Tensor returns `data.uploadUrl` and `data.displayUrl`.

## Scope

- Add `state.presign` to localStorage snapshots and imports.
- Add a `pre_sign` panel to `canvas.html` above `.canvas-workbench`.
- Bind parse, format, and send actions from `initCanvasPage`.
- Keep actual mask PUT upload as a later step. This change verifies whether the parsed API credentials and request body can obtain an upload URL.

## Testing

Node tests cover the new `presign` state shape and the PowerShell parser behavior for a pre_sign request that contains escaped JSON body text and browser-only headers that must be sanitized.
