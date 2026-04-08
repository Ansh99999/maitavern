# MaiTavern

A minimalist, mobile-friendly tavern UI for OpenAI-compatible APIs with early support for characters, providers, presets, assets, and recent chats.

## Current features

- OLED-style welcome screen and minimalist black UI
- Home page with Characters, Presets, Providers, and Chats sections
- SVG icon-based interface
- Character library with:
  - search
  - import SillyTavern-compatible JSON
  - create/edit inside the UI
  - export character JSON
  - optional asset assignment for avatars
- Asset upload area for images and JSON files
- Provider manager with:
  - search
  - import/export JSON
  - create new provider entries
  - apply a provider as the active endpoint
- Chat view with:
  - bottom-sticky composer
  - send button only behavior
  - Enter key does not send
  - impersonation dropdown
  - hamburger drawer for homepage/presets/providers
  - editable chat title
  - current preset/provider status in the drawer
- Custom endpoint URL
- Custom API key
- Custom auth header name and prefix
- Extra headers and extra JSON body fields
- Multiple model names
- Saves state in local storage

## Usage

Open `index.html` in a browser or serve the folder with any static file server.

## Endpoint URL behavior

You can enter either:
- a base URL like `https://example.com/v1`
- or the full URL like `https://example.com/v1/chat/completions`

If you enter only the base URL, the app automatically appends:

```text
/chat/completions
```

## Notes

- This app sends your API key directly from the browser, so use it only in trusted environments.
- Some custom providers expect headers other than `Authorization: Bearer ...`, so the UI lets you customize the header name/prefix.
- If your gateway requires extra request fields, use the extra body JSON box.
- Presets are still early-stage and currently expose a default applied preset state.
- For production use, a small backend proxy is safer.
