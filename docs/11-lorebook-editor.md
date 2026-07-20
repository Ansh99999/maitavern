# 11 — Lorebook Editor (deep design)

The World Info authoring surface. **One shared editor component**, mounted in two places: **Library → 📖 Lorebooks** (standalone books) and **Character Studio → Lore tab** (the character's embedded `character_book`). Every field maps 1:1 to `Lorebook` / `LorebookEntry` in [09-data-model.md](09-data-model.md); runtime matching is the scan algorithm in [10-prompt-engine.md](10-prompt-engine.md) — the editor never invents behavior the engine doesn't have.

## Lorebooks list (Library → 📖)

```
┌──────────────────────────────────────────────┐
│ ‹  Lorebooks                    [Import] [＋] │
│ 🔍 search…      ( All │ Normal │ Chat │ Emb ) │
├──────────────────────────────────────────────┤
│ 📖 Eldoria Setting            24 entries   ⋮ │
│    ↳ 💬 Eldoria Setting — Aria #1  (chat)  ⋮ │  fork, collapsed under parent
│ 📖 Aria (embedded)             6 entries   ⋮ │  from character card
└──────────────────────────────────────────────┘
```

- Rows: name · entry count · kind badge (**normal / chat / embedded**) · assignment chips (global / character / chat) · `⋮` (assign · export ST JSON · duplicate · rename · delete).
- **Chat forks nest under their parent** (`parentLorebookId`), collapsed by default; embedded books show their owning character and deep-link into the Studio Lore tab.
- `[Import]` accepts SillyTavern World Info JSON and `character_book` extracted from V2/V3 cards (AgnAI in Phase 4).

## Book editor

```
┌──────────────────────────────────────────────┐
│ ‹  Eldoria Setting                 ⋯  [Save] │
│ 💬 forked from “Eldoria Setting” for Aria #1 │  chat books only → [History] [Promote]
│ Scan depth 4 · Budget 25% · Recursion ✓   ✎  │  book settings (blank = engine defaults)
│ 🔍 filter entries…        sort ▾   [＋ Entry] │
├──────────────────────────────────────────────┤
│ ⋮⋮ ● The Silver Court    keys: court, silver │  collapsed row
│ ⋮⋮ ● Magic system        📌 constant         │
│ ⋮⋮ ○ Old war (disabled)  keys: war  🎲 50%   │
│ ── tap a row → expands into the entry editor │
└──────────────────────────────────────────────┘
```

- **Header:** name (inline-editable) · `⋯` (export · import-entries-into-book · duplicate · assign · delete) · persistent **Save** (Studio-style).
- **Book settings row:** `scanDepth` · `tokenBudget` · `recursiveScanning`. Unset = inherit engine defaults (4 / 25% / on, limit 2); the row shows inherited values dimmed.
- **Entry rows (collapsed):** drag handle (`⋮⋮`, writes `order`) · enable dot (`●`/`○`) · title (`comment`) · key preview · badges: `📌` constant · `.*` regex keys · `🎲 n%` probability < 100 · `↳ d` at-depth. Long-press = multi-select (bulk enable/disable · move/copy to book · delete).
- **Sort:** insertion order (default, = `order`) · alphabetical · recently edited. Only insertion-order view allows dragging.

## Entry editor (expanded row, accordion — no separate page)

```
│ ▼ The Silver Court                        ⋮  │
│   Title      [The Silver Court           ]   │
│   Keys       (court ×) (silver ×) (＋)       │
│   ◉ Keyed   ○ Constant (always on)           │
│   Selective ▢ → secondary keys + logic       │
│       (queen ×) (＋)   [AND ANY ▾]           │
│   Content    ┌───────────────────────────┐   │
│              │ big macro-enabled canvas… │   │
│              └───────────────────────────┘   │
│   Insertion  [Before char ▾]  depth [–]      │
│   Order [100]  Probability [────────● 100%]  │
│   ▢ Case-sensitive   ▢ Regex keys            │
└──────────────────────────────────────────────┘
```

- Fields map exactly to `LorebookEntry`: `comment`, `keys`, `constant`, `selective` + `secondaryKeys` + `selectiveLogic` (AND ANY / AND ALL / NOT ANY / NOT ALL), `content`, `position` (Before char · After char · At depth `n` · A/N top · A/N bottom), `order`, `probability`, `caseSensitive`, `useRegex`, `enabled`.
- **Content is a big inviting canvas** (Studio feel), macro-enabled, with a live token estimate.
- Entry `⋮`: duplicate · move/copy to another book · save as snippet · delete.

## 🔬 Scan tester (dry run)

A drawer at the bottom of the book editor. Input: pasted sample text **or** "last N messages of chat …". Output: the entries that would fire, each with its **reason** (key hit highlighted in the sample / constant / secondary-logic result / probability roll shown as rolled-vs-needed), recursion passes visualized, and total tokens vs the book budget. **Runs the real engine scan function** — same code path as generation, so the tester can never lie.

## Chat forks, version history, promote

- A **chat book** shows the fork banner (`parentLorebookId` → origin, owning chat) with **[History]** and **[Promote]**.
- **History sheet:** the book's `LorebookVersion` snapshots (auto-pushed per edit session, per [06](06-memory-system.md)) — tap to preview, **Rollback** restores.
- **Promote (canonize):** entry-by-entry diff against the parent (added / changed / removed), pick what graduates, apply as a merge to the parent book. Never a silent overwrite. Mirrors Mai's accept/reject diffs.

## Assignment

**Assign sheet** (from list or book `⋯`): activate the book **globally**, **per character**, or **per chat** (`Chat.lorebookIds`). Embedded books are always active with their character; chat forks are auto-assigned to their chat and **shadow their parent** there (engine rule, doc 10). Chips on each row show where the book is live.

## ✦ Mai in the editor

Same floating assistant as the Studio (04), available in both mounts. Lore-specific quick actions: **draft entries from a description**, **suggest keys for an entry's content**, **split an oversized entry**, **find gaps** (world topics with no entry). All output lands as Accept/Edit/Reject diffs — the user stays the author.

## Component map

`LorebookListScreen` → `LorebookRow` · `LorebookEditor` (shared Library/Studio) = `BookHeader` + `ForkBanner` + `BookSettingsRow` + `EntryList` → `EntryRow` ⇄ `EntryEditor` (accordion) · drawers/sheets: `ScanTesterDrawer`, `VersionHistorySheet`, `PromoteDiffSheet`, `AssignSheet`, `BulkActionBar`.

## Defaults chosen (vetoable)

- **New-entry defaults:** enabled · keyed (not constant) · selective off · position **Before char** · order **100** · probability **100** · case-insensitive · non-regex keys (matches ST expectations for imported users).
- **Accordion editing** (expand-in-place) instead of a per-entry page — fewer navigation levels on mobile; the content canvas grows to full height when focused.
- **Drag order writes `order` in steps of 10** so manual values can be slotted between rows.
- **Book settings blank-inherit** engine defaults rather than copying them (so tuning a global default flows through).
- **Promote = reviewed merge** into the parent (by entry id), never replace-all.
