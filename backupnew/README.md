# tensorArtApiHandler

A pure frontend TensorArt API helper.

## Pure Frontend Flow

The app stores editable Send / Query / Post request settings in localStorage.
Normal use only needs one fresh TensorArt PowerShell request:

1. Sign in to TensorArt in the same browser profile.
2. Refresh TensorArt / reCAPTCHA when needed.
3. Copy any working TensorArt API request as PowerShell.
4. Paste it into the "共用 API 簽章" section.
5. Click "更新共用簽章".

The shared parser keeps only the headers that are actually needed for TensorArt
request signing:

```text
x-echoing-env
x-request-lang
x-request-package-id
x-request-package-sign-version
x-request-sign
x-request-sign-type
x-request-sign-version
x-request-timestamp
```

Cookies are not stored. Browser login state still comes from
`credentials: "include"` in the signed-in browser profile.

Use the GitHub Pages URL for real requests. TensorArt currently allows requests
from `https://gino07172002.github.io/tensorArtApiHandler/`, but the same request
from `http://127.0.0.1` / `localhost` can fail at the browser CORS layer with
`Failed to fetch`.

Each Send / Query / Post panel still has its own PowerShell parser for advanced
updates. Parsing a per-request PowerShell command updates that request's URL,
method, body, and non-sensitive request-specific headers, while moving the
TensorArt signature headers into the shared signature store.

## Defaults

If localStorage is empty, the app starts with editable defaults for:

- `https://api.tensor.art/works/v1/works/task`
- `https://api.tensor.art/works/v1/works/tasks/query`
- `https://api.tensor.art/community-web/v1/post/create`

Existing localStorage values always win over these built-in defaults.
