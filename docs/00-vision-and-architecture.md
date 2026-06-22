# 00 — Vision & Architecture

## Vision

MaiTavern is a **mobile-first AI roleplay frontend** shipped as an Android **APK** (plus a PWA/desktop build from the same codebase). The feeling on the surface is a calm, minimalist **lounge**; underneath is a deeply feature-rich, customizable engine comparable to SillyTavern, designed for phones first.

**North star:** easy, beautiful mobile chatting on top; serious power underneath. **Local-first, no mandatory backend** — the app calls LLM provider APIs directly from the device (BYOK).

## Locked stack

| Layer | Choice | Why |
|---|---|---|
| UI | React 18 + TypeScript + Vite | Web tech maximizes SillyTavern format/UX compatibility + dev speed |
| Mobile shell | **Capacitor** → Android APK; PWA/desktop from same build | HTML/CSS-in-message, custom fonts, theming, card PNG parsing are all native to web |
| State | Zustand | Lightweight, fits the override-layer model |
| Styling | Tailwind CSS + CSS variables | Runtime theming, custom backgrounds/fonts |
| Storage | Dexie / IndexedDB; Capacitor Filesystem for import/export | Local-first; SQLite optional for heavy logs |
| Rendering | react-markdown + remark/rehype + DOMPurify; react-virtuoso | Safe HTML, long-chat performance |
| Networking | Direct streaming `fetch` for big-3 + **native Kotlin OkHttp/SSE plugin** | Bypass CORS for arbitrary custom endpoints; reliable streaming |
| Build/CI | GitHub Actions | Debug APK + PWA artifact; signed release later |

### Platform & data decisions

- **Android-first**, plus a **PWA/desktop** build (manifest + service worker) from the same codebase. No iOS for now; code kept portable.
- **Local zip export/import first; cloud sync (WebDAV/Drive) deferred to a later phase** — but the data model is sync-ready from the start.
- **API keys** live encrypted on device (a client app can never truly hide keys).

## The two foundational principles

1. **One prompt pipeline.** Everything — system prompt, persona, character def, scenario, world info (before/after), example messages, chat history, summaries, author's note, trackers, RAG recall — flows through a single **ordered prompt assembler** modeled on SillyTavern's Prompt Manager. This makes the **Log viewer** able to show the exact assembled request, block by block, with token counts.

2. **Global → Character → Chat override scope.** Almost every setting has a global default (Settings), overridable per-character (Character Studio), overridable again per-chat (Quick Settings); lower wins. **In-chat edits fork to chat scope** (copy-on-write) and never mutate canonical assets — good discoveries graduate via an explicit "promote/canonize" action.

## Key technical risks (verify at build time)

1. **Streaming + CORS over WebView.** Default: direct fetch for the big-3 (Anthropic needs `anthropic-dangerous-direct-browser-access: true`); a **native Kotlin OkHttp+SSE plugin** for arbitrary endpoints lacking CORS. Don't rely on CapacitorHttp for SSE.
2. **Cleartext/LAN endpoints** (KoboldCpp/Ollama/oobabooga) need `usesCleartextTraffic` / network-security-config.
3. **Tokenizers** — per-provider token counting (tiktoken for OpenAI; approximate for others) for the budget visualizer.
4. **APK signing in CI** — keystore as a base64 GitHub secret; never logged.
5. Confirm current `actions/setup-java`, Gradle, and Capacitor major versions when scaffolding.

## Reference projects (formats/UX to mine)

- **SillyTavern** — `world-info.js`, `PromptManager.js`/`openai.js`, `macros.js`, `group-chats.js`, instruct/context templates, regex + summarize extensions, `.jsonl` chats, settings/preset JSON.
- **Character Card V2/V3 spec** — `TavernCardV2` / `chara_card_v3`; PNG tEXt `chara`/`ccv3`.
- **RisuAI** — mobile UX, CCv3, lorebook scripting.
- **AgnAI** — character/chat/memory-book export schema.
- Chub / Risu / Pygmalion / TavernAI legacy formats.
