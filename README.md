# MaiTavern

> A mobile-first AI roleplay frontend — *"SillyTavern for your phone"* — shipped as an Android **APK** (with a PWA/desktop build from the same codebase).

MaiTavern is a local-first, deeply customizable AI roleplay chat app. It talks **directly** to LLM provider APIs from your device (BYOK — bring your own key), so there is **no mandatory backend**. The surface is a minimalist, relaxed "lounge"; underneath sits a feature-rich engine: SillyTavern-compatible characters, lorebooks and presets, a full memory system, group chats, and total UI customization.

> **Status: Phase 1 (MVP chat) built.** Streaming single-character chat, BYOK provider profile (OpenAI-compatible + Anthropic), SillyTavern card import, sampler presets, macro engine + prompt assembler, request-log viewer, 4 chat styles and theme presets. Phase 2 (lorebooks, swipes, personas) is next — see [`docs/ROADMAP.md`](docs/ROADMAP.md).

---

## Highlights

- **4 chat display styles** — bubble · document · discord · novel — with never-cropped, aspect-aware avatars and a draggable/zoomable avatar viewer.
- **A "lounge" homepage** — logo, auto-rotating carousel of recent chats, time-based greeting, live Characters/Chats previews, and Provider/Preset status chips.
- **Character Studio** — an AI-guided, Chrome-tabbed character editor (Overview · Personality · Trackers · Lorebook · System) with **Mai**, an agentic BYOK build assistant.
- **Three-tier memory** — Manual, Automated (summarization), and **Agentic** (a background agent that updates lorebooks, character details, and live stats, with a review-first gate).
- **Copy-on-write everywhere** — in-chat edits fork to chat scope and never mutate your canonical characters/lorebooks; good discoveries graduate via "promote."
- **Providers & Models** — simple **Provider** and advanced **Router** connections (multi-key rotation, fallbacks, per-model pricing & cost tracking, param include/exclude), plus image & voice models and a central **Model Roles** map.
- **SillyTavern compatibility** — character cards (V2/V3 PNG/JSON), lorebooks, presets, `.jsonl` chats; plus AgnAI and other imports.
- **Deep theming** — preset themes + a full Custom Theme Editor (colors, backgrounds, bubble styles, fonts incl. `.ttf`/Google Fonts, custom CSS).
- **Global → Character → Chat** override scope for nearly every setting.

See [`docs/`](docs/) for the complete design.

## Planned tech stack

| Layer | Choice |
|---|---|
| UI | React 18 + TypeScript + Vite |
| Mobile shell | Capacitor (Android APK) + PWA/desktop from the same build |
| State | Zustand |
| Styling | Tailwind CSS + CSS variables (runtime theming) |
| Storage | Dexie / IndexedDB (local-first) |
| Rendering | react-markdown + DOMPurify, virtualized message list |
| Networking | direct streaming fetch + a native Kotlin OkHttp/SSE plugin for custom endpoints |
| Build/CI | GitHub Actions → debug APK + PWA artifact |

Rationale and locked decisions: [`docs/00-vision-and-architecture.md`](docs/00-vision-and-architecture.md).

## Repository layout

```
maitavern/
├── docs/                 # full design documentation
│   ├── 00-vision-and-architecture.md
│   ├── 01-chat-interface.md
│   ├── 02-homepage-and-library.md
│   ├── 03-new-chat-flow.md
│   ├── 04-character-studio.md
│   ├── 05-settings.md
│   ├── 06-memory-system.md
│   ├── 07-providers-and-models.md
│   ├── 08-presets.md
│   ├── 09-data-model.md
│   ├── 10-prompt-engine.md
│   ├── 11-lorebook-editor.md
│   ├── ROADMAP.md
│   └── DECISIONS.md
├── README.md
├── LICENSE               # MIT
└── .gitignore            # ready for the Vite + Capacitor app
```

## Roadmap (short)

- **Phase 0** — scaffold (Vite+React+TS+Capacitor), CI emitting a debug APK + PWA, storage + theming shell.
- **Phase 1 (MVP)** — single-char streaming chat, one provider, ST card import + basic creator, sampler preset, macro engine + prompt assembler, log viewer, 4 chat styles + theming.
- **Phase 2** — lorebooks, author's note, scenario injection, swipes/branching, personas, full provider trio.
- **Phase 3** — memory (summarization, agentic editing), group chats, regex scripts, guided generations.
- **Phase 4** — AgnAI + other imports, backup/restore, prompt-manager presets, onboarding, signed release.
- **Phase 5** — RAG/vector memory, TTS/STT, image hooks, cloud sync.

Full detail: [`docs/ROADMAP.md`](docs/ROADMAP.md).

## License

[MIT](LICENSE) © 2026 MaiTavern contributors.
