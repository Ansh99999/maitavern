# 08 — Presets

The generation engine's user-facing editor. Chrome-style tabs (like Character Studio). The **Prompt Blocks tab is the visible face of the one prompt assembler** ([10](10-prompt-engine.md)) and feeds the Log viewer. Full **SillyTavern** prompt-manager + macro compatibility throughout.

```
┌──────────────────────────────────────────────────┐
│ ‹  Preset editor              ⋯   [ Save ]        │
│ ╭ General ╮ Prompt Blocks  Parameters  Regex      │
└──────────────────────────────────────────────────┘
```

## Tab 1 — General

Name · **icon + banner + theme color** (from in-app Gallery or phone files; aspect preserved → flows into the Preset picker and the homepage Preset chip) · **default provider** for this preset (a *suggestion*, overridable per chat) · context size · temperature · max response tokens · streaming · notes/tags.

> Temperature & context are the **same values** as the Parameters tab (single source of truth, surfaced in both).

## Tab 2 — Prompt Blocks (SillyTavern Prompt Manager)

Ordered, **drag-to-reorder** blocks; reordering literally changes the assembled prompt (and the Log viewer). Two kinds:

- **Markers** — dynamic injection points (chat history, character description, personality, scenario, persona, dialogue examples, world-info before/after).
- **Custom** — your static, macro-enabled text.

Standard ST blocks are present out of the box; **importing an ST preset** populates this exactly. Tap a block → editor window: Title · Role (system/user/assistant) · Position (relative order or in-chat at depth) · Content (macros) · Enabled · Forbid-override · live token count.

Default order is the SillyTavern-compatible sequence in [10-prompt-engine.md](10-prompt-engine.md).

## Tab 3 — Parameters (samplers)

Every parameter has an **enable toggle** (off = not sent; respects the connection's include/exclude valve). Set by input or slider. Grouped so it isn't a wall:

- **Core** — temperature, top_p, top_k, max tokens
- **Penalties** — repetition, frequency, presence, no-repeat-ngram
- **Advanced** — min_p, typical_p, TFS, top_a, mirostat (τ/η), DynaTemp, smoothing, sampler order
- **Misc** — seed, stop sequences, logit bias, n, grammar/JSON schema

Params unsupported by the selected method are flagged/greyed. (Note: "euler" etc. are *image* samplers — those live with image models in Providers, not here.)

## Tab 4 — Regex (SillyTavern-style scripts)

Ordered find/replace scripts. Per script: name · find `/…/flags` · replace (capture groups + macros) · **Affects** (user input / AI output / display-only / prompt-only / slash-commands) · min/max depth · run-on-edit · enabled · a **🧪 test box** (paste sample → live transformed result).

**Scope: regex stacks across Global / Character / Preset / Chat** (matches the override hierarchy). Preset-scoped scripts travel with the preset.

## Macros — app-wide SillyTavern compatibility

A shared core service resolves ST macros everywhere text is entered. Full list and evaluation rules in [10-prompt-engine.md](10-prompt-engine.md).

## Productive features

Live context-budget meter (per-block token counts) · dry-run preview (assemble against a sample character/chat — a mini Log viewer — without sending) · import/export ST presets (JSON, both ways; visual identity travels) · clone + version history · macro lint.

## Confirmed decisions

- **Preset ↔ provider = default suggestion** (overridable per chat; one preset works across providers).
- **Regex = stacks across Global / Character / Preset / Chat** scopes.
