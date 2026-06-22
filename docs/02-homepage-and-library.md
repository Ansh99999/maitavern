# 02 — Homepage & Library

## Homepage — "the lounge"

A calm, personal landing space. Confirmed layout: **"Lounge with live previews."**

```
┌──────────────────────────────────────────────┐
│  ◆ MaiTavern                            🕐    │  logo (L) · clock (R) → most recent chat
├──────────────────────────────────────────────┤
│  ╭────────────────────────────────────────╮   │
│  │         [ carousel slide image ]        │  │  auto-slide every 3–4s
│  │                            ●  ○  ○  ○    │  │  recent 3–4 chats' avatars / custom imgs
│  ╰────────────────────────────────────────╯   │  tap a slide → open that chat
│  Good evening 🌙                               │  time-based greeting (+ optional subtitle)
│                                                │
│  Characters                          see all › │  horizontal shelf of faces
│  ⟨ (A) (B) (C) (D)   ＋ ⟩                       │
│  Chats                               see all › │  live recent-message previews
│  • Aria — "you made it…"            2m ago     │
│  ┌──────────┐ ┌──────────┐                     │
│  │ Provider │ │ Preset   │  ← active chips     │
│  └──────────┘ └──────────┘                     │
│  ┌──────────┐ ┌──────────┐                     │
│  │ Library  │ │ Settings │                     │
│  └──────────┘ └──────────┘                     │
└──────────────────────────────────────────────┘
```

- **Top bar:** logo (left), **clock** (right → resume most recent chat in full chat interface). Minimal — only these two.
- **Carousel:** auto-slides every 3–4s through recent chats' avatars or custom images; **tap a slide → open that chat**. Settings: enable/disable, interval, source (recent / custom / both), transition, dots, pause-on-touch, tap behavior.
- **Greeting:** time-based (morning/afternoon/evening/night), optionally personalized with persona name + rotating subtitle. Toggleable.
- **Characters:** shelf of favorites/recent + `＋` create + "see all ›" → full gallery.
- **Chats:** recent-chats preview (last message + timestamp) + "see all ›" → full list.
- **Provider / Preset tiles:** double as **active-status chips** (show current config, tap to change).
- **Library / Settings tiles:** open their screens.

> Optional: a global search field under the greeting (characters + chats + lorebooks).

### Section destinations

Characters → gallery/management (grid, tags, search, favorites, create). Chats → full chat list. Provider → connection profiles. Preset → presets manager. Library → dashboard below. Settings → app settings.

## Library — the dashboard

```
┌──────────────────────────────────────────────┐
│  ‹  Library                                    │
├──────────────────────────────────────────────┤
│  ┌── 📖 Lorebooks ──┐  ┌── ✎ Prompt Snippets ─┐ │
│  ┌── 🖼 Gallery ────┐  ┌── 📄 Documents (RAG) ─┐ │
│  ┌── 🗂 Character Wiki (reserved) ───────────┐  │
└──────────────────────────────────────────────┘
```

- **📖 Lorebooks** — World Info books: create / import (ST, AgnAI) / edit / assign (global or per-chat). Two kinds: **normal** and **chat** lorebooks (see Memory doc).
- **✎ Prompt Snippets** — reusable prompt blocks. Includes **scenario snippets** (feed new-chat scenario injection) plus system/main-prompt, jailbreak/post-history, author's-note templates, and style snippets. Tagged & searchable.
- **🖼 Gallery** — central media store (backgrounds, avatars, custom images, later generated); reused for backgrounds, avatars, and the homepage carousel.
- **📄 Documents (RAG)** — upload txt/md/pdf → chunk + embed → retrieve into context ("data bank"); assignable per character/chat. The retrieval half of the memory system.
- **🗂 Character Wiki** — **reserved**; spec to come.
