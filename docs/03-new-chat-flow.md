# 03 — New Chat Flow

Character-first, with a fast default path. Everything chosen here is editable mid-chat later.

## Entry: the character quick menu

In the **Characters** region, tap a character → a small popover:

```
┌──────────────────────────┐
│  ✎  Edit in Character     │   → full character editor
│     Studio                │
│  ✦  New chat              │   → setup screen below
│  ☰  Chat list             │   → all chats with this character
└──────────────────────────┘
```

(Extras like favorite/duplicate/export can hang off long-press.)

## "New chat" → the setup screen

A single scrollable config screen (not a wizard). Every field defaults sensibly, so the fast path is **tap character → New chat → Launch** (two taps).

```
┌────────────────────────────────────────────┐
│  ‹  New chat                                 │
├────────────────────────────────────────────┤
│  ┌────┐  Aria        edit definition ›       │
│  └────┘                                       │
│  Preset       [ Default (global)        ▾ ]  │
│  Provider     [ Default (global)        ▾ ]  │
│  Chat style   [ Bubble ▾ ]      ⚙ customize  │
│  Persona      [ (P) Aanya ▾ ]        ✎ edit  │
│  Scenario     [ Character's own ▾ ]  ＋ library│
│  ┌────────────────────────────────────────┐ │
│  │ ＋  Make this a group chat              │ │
│  └────────────────────────────────────────┘ │
│             [   🚀  Launch chat   ]          │
└────────────────────────────────────────────┘
```

- **Preset** — defaults to the global default preset.
- **Provider** — defaults to the global default connection.
- **Chat style** — the 4 styles; defaults to global. **"⚙ customize"** opens display settings, saveable as a **per-character override or globally**.
- **Persona** — defaults to your active persona; edit/create inline.
- **Scenario** — **defaults to the character's own scenario**, overridable per-chat with a Library scenario snippet or custom text.
- **＋ Make this a group chat** — add other bots (turn-order/members editable later). A complementary **"＋ New group chat"** entry in the Characters region starts a group from scratch.
- **edit definition ›** — inline quick-edit (description/personality/scenario/first message) + link to full **Character Studio**.
- **🚀 Launch chat** → opens the chat interface.

## Confirmed decisions

- **Edit scope = per-chat override.** Editing the character definition from New Chat or inside a chat applies **only to that chat**; the original stays clean. Global edits go through **Character Studio**.
  - *Data model:* a Chat stores an **override layer** (description/personality/scenario/first-message/etc.) over the base Character; the prompt assembler merges override-over-base at build time.
- **Scenario default = the character's own scenario**, overridable per chat.
