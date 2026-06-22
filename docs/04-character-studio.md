# 04 — Character Studio

The character editor, reached from the character quick menu / new-chat flow. **Feel: a brainstorming canvas, AI-guided — not a form.** Big inviting input areas, inline AI help, organized by **Chrome-style tabs**. Persistent **Save** in the header + a `⋯` menu (export V2/V3 PNG card / JSON, import).

```
┌──────────────────────────────────────────────────┐
│ ‹  Character Studio                 ⋯   [ Save ]  │
│ ╭ Overview ╮ Personality  Trackers  Lore  System  │  chrome-style tabs
├──────────────────────────────────────────────────┤
│                  (active tab content)              │
│                                          ( ✦ Mai ) │  floating assistant
└──────────────────────────────────────────────────┘
```

## Tabs

1. **Overview** (identity + presentation) — display picture, cover picture (opt), chat background (opt), chat banner, name (+ nickname/tags/creator/version), **theme color** + **dialogue color**. Feeds the chat per-character theme. Images preserve aspect ratio.
2. **Personality** (definition canvas) — personality · character description · dialogue/message examples (multiple blocks) · greeting/first message + alternate greetings · creator notes.
3. **Trackers** (opt-in; off by default) — user-defined stats. Per stat: type (number / bar w/ min-max / text / tag), initial value, visible-in-chat?, AI-updates-live?
4. **Lorebook** — the character's **embedded** World Info (V2/V3 `character_book`): create/edit/import entries; travels with the card on export.
5. **System** — system prompt (snippet-pullable) · scenario prompt (supports scenario snippets; per-chat overridable) · post-history/jailbreak · character gallery (photos; V3 `assets`) · **✓ Finalize character** (commit + export card/JSON).

## ✦ Mai — the agentic build assistant

A floating button on every tab. Conversational: the user describes intent → Mai drafts/refines fields. **Agentic & multi-field** (can flesh out a whole character across tabs, ask clarifying questions). **Context-aware** quick actions per tab. **Non-destructive** — every change is an Accept/Edit/Reject diff; the user stays the author.

## Confirmed decisions

- **Trackers = per-stat choice:** each stat is **either a fixed sheet value or a live tracker** the AI updates during chat (shown in a chat stats panel, optionally injected into context). Live trackers tie into the memory/state system.
- **Mai = a dedicated assistant model** (separately configured BYOK provider/model, kept apart from the roleplay model). Lives in the Model Roles map (see Providers doc).

## Card spec mapping (compatibility)

Fields map to **Character Card V2/V3**: `name`, `description`, `personality`, `scenario`, `first_mes`, `mes_example`, `alternate_greetings`, `system_prompt`, `post_history_instructions`, `character_book`, `tags`, `creator`, `character_version`, plus V3 `assets` (gallery). Trackers and theme are stored in `extensions`.
