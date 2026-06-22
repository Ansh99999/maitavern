# Decision Log

Locked decisions, newest grouped by area. These are the choices the user explicitly confirmed during design.

## Stack & platform
- **Shell:** Capacitor + React + TypeScript (not React Native / Flutter / Tauri).
- **Targets:** Android-first **+ PWA/desktop** from the same codebase. No iOS for now.
- **Sync:** local zip export/import first; cloud sync deferred to Phase 5 (data model kept sync-ready).
- **License:** MIT.
- **Phase 1 MVP:** "Standard MVP" (single-char streaming chat, one provider, ST import + basic creator, sampler preset, macros + assembler, log viewer, 4 chat styles + theming). Lorebook is Phase 2.

## Chat interface
- **Four** display styles: bubble · document · discord · novel.
- Avatars: show/hide in **every** mode (per-style memory; novel included); **never cropped** (native aspect ratio); sizing **Auto (accommodate aspect)** or **Custom**; text **wrap-around or below**, every mode.
- Send button **morphs into Stop** during generation.
- Composer `⋮` = **combined sheet**: generation actions on top + expand into full Quick Settings.

## Homepage
- Layout: **"Lounge with live previews"** (Characters shelf, live Chats previews, Provider/Preset status chips).

## New chat flow
- **Edit scope = per-chat override** (global edits via Character Studio); Chat stores an override layer over the base Character.
- **Scenario default = the character's own scenario** (overridable per chat).

## Character Studio
- Editor named **"Character Studio"**; 5 Chrome-style tabs; agentic **Mai** assistant.
- **Trackers = per-stat choice** (fixed value or live AI-updated tracker).
- **Mai = dedicated assistant model.**

## Settings
- **Full Global → Character → Chat** override hierarchy.
- Theme = **presets + full Custom Theme Editor** (incl. custom CSS).

## Memory system
- **Lorebook forks = one chat-lorebook per chat + version history.**
- **Agentic apply = review-first, with per-chat auto opt-in.**
- **Memory-agent model + cadence = user-configurable** (+ cost cap).

## Providers & Models
- **Model Roles = both** (central map + inline shortcuts).
- **Cost tracking** available to plain Providers via optional manual pricing (Router has it built in).
