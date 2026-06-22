# 05 — Settings

A **searchable, categorized hub** (search indexes every leaf setting) → tap a category → detail page.

```
┌──────────────────────────────────────────────┐
│ ‹  Settings    [ 🔍 Search settings…       ]  │
├──────────────────────────────────────────────┤
│ 🎨 Appearance & Theme   💬 Chat Display        │
│ 🔤 Fonts & Text         🔌 Providers & Models  │
│ 🎚 Presets & Generation 🧠 Memory & Context    │
│ 🎭 Personas             🏠 Homepage            │
│ 💾 Data, Backup & Import 🔒 Privacy & Security  │
│ 📜 Logs & Advanced      ℹ️ About & Updates     │
└──────────────────────────────────────────────┘
```

## Core principle — scope

Settings is the **global default layer**:

```
Global (Settings)  →  Character (Studio)  →  Chat (Quick Settings)   ← lower wins
```

**Confirmed: full Global → Character → Chat hierarchy.** Override-capable pages note "global default; characters/chats can override"; changing a control inside a chat shows a scope toggle (This chat / This character / Global). The data model stores override layers at character and chat level.

> Providers, Presets, and Library are their own top-level areas (home tiles too); Settings links to the **same** screens and adds app-wide defaults.

## Categories

- **🎨 Appearance & Theme** — **confirmed: presets + a full Custom Theme Editor** (background color/gradient/image + blur/opacity, bubble colors + styles, text/theme/dialogue colors, fonts, **custom CSS**), exportable/shareable; app theme (Light/Dark/System/Custom), accent, per-character theme overrides.
- **💬 Chat Display** — default style (bubble/document/discord/novel), avatars (show-hide per style, sizing auto/custom, text wrap-vs-below, rounding), name toggle, inline actions, timestamps, swipes, markdown/HTML, composer (send-on-enter, streaming, token counter).
- **🔤 Fonts & Text** — family (system / Google Fonts / upload `.ttf`), size, line height, spacing, UI-vs-chat fonts, novel justification.
- **🔌 Providers & Models** — see [07-providers-and-models.md](07-providers-and-models.md).
- **🎚 Presets & Generation** — sampler presets (temp, top_p/k/a, min_p, typical, penalties — each with a **disable toggle**, sampler order, max context/response), prompt-manager presets, instruct/context templates, default preset.
- **🧠 Memory & Context** — see [06-memory-system.md](06-memory-system.md).
- **🎭 Personas** — manage personas, default, per-character default.
- **🏠 Homepage** — carousel, greeting, section visibility/order, global search toggle.
- **💾 Data, Backup & Import** — full zip backup/restore, per-entity export (PNG/JSON/jsonl), compatibility import (ST / AgnAI / Chub / Risu / Pygmalion / TavernAI), cloud sync (later), storage, factory reset.
- **🔒 Privacy & Security** — key encryption status, app lock (PIN/biometric), incognito (ephemeral chats), no telemetry.
- **📜 Logs & Advanced** — request/response logging + global log viewer, custom CSS injection, regex/find-replace scripts (ST-compatible), macros reference + custom macros, quick-replies manager, experimental flags, reset.
- **ℹ️ About & Updates** — version/build, check-for-updates (sideloaded → GitHub Releases / Obtainium), changelog, licenses, credits.
