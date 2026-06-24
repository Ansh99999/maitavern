# Roadmap

We are **not** one-shotting this. Phased build; each phase is shippable.

## Phase 0 — Skeleton
Repo, Vite + React + TypeScript + Capacitor, Tailwind + theming shell, storage layer (Dexie), and a **GitHub Actions** pipeline that emits a **debug APK and a deployable PWA** (manifest + service worker). Proves both output targets end-to-end.

## Phase 1 — MVP chat ("Standard MVP")
Single-character streaming chat · one provider profile · character import (ST PNG/JSON) + basic creator · sampler preset · macro engine + ordered prompt assembler · basic log viewer · the 4 chat styles + core theming.

## Phase 2 — Depth
Lorebook / World Info · author's note · scenario injection + Prompt Snippets library · swipes / branching · persona system · full provider trio (OpenAI / Anthropic / Gemini) + native streaming plugin.

## Phase 3 — Memory & groups
Auto-summarization · manual memory · **agentic lorebook editing** (review-first) · live trackers · group chats · regex scripts · guided generations.

## Phase 4 — Compatibility & polish
AgnAI + other imports · full backup/restore · prompt-manager preset editing · token visualizer · onboarding · signed release pipeline.

## Phase 5 — Extras
RAG / vector memory (semantic recall + Documents) · TTS/STT · image-gen hooks · cloud sync (WebDAV/Drive) · Custom Theme Editor depth · Character Wiki.

---

### Design status

| Area | Doc | Status |
|---|---|---|
| Vision & architecture | [00](00-vision-and-architecture.md) | ✅ locked |
| Chat interface | [01](01-chat-interface.md) | ✅ locked |
| Homepage & Library | [02](02-homepage-and-library.md) | ✅ locked |
| New chat flow | [03](03-new-chat-flow.md) | ✅ locked |
| Character Studio | [04](04-character-studio.md) | ✅ locked |
| Settings | [05](05-settings.md) | ✅ locked |
| Memory system | [06](06-memory-system.md) | ✅ locked |
| Providers & Models | [07](07-providers-and-models.md) | ✅ locked |
| Presets / Prompt Manager | [08](08-presets.md) | ✅ locked |
| Data model | [09](09-data-model.md) | ✅ locked |
| Prompt engine | [10](10-prompt-engine.md) | ✅ locked |
| Lorebook editor (deep) | — | ⏳ next |
| Logs viewer · Persona manager · Theme editor · Memory panel UI | — | ⏳ todo |
| Import / compatibility pipeline | — | ⏳ todo |
| Character Wiki | — | 🔒 reserved (awaiting spec) |

### Build status

| Step | Status |
|---|---|
| Repo + docs skeleton | ✅ done |
| Foundation specs (data model + prompt engine) | ✅ done |
| Phase 0 scaffold (Vite+React+TS+Tailwind+Capacitor+CI) | ⏳ next |
