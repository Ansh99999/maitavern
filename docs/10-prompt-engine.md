# 10 — Prompt Engine

The runtime core. One deterministic pipeline turns the data model ([09](09-data-model.md)) into a provider request, and **every stage is inspectable in the Log viewer**. This is also exactly what the Preset → Prompt Blocks tab visualizes.

## Pipeline (per generation)

```
1. Resolve scope      merge Global → Character → Chat overrides → effective config
2. Gather sources     character (merged w/ per-chat overrides), persona, scenario,
                      author's note, summary, trackers, assigned lorebooks, RAG docs
3. World-info scan     match lorebook entries against recent context (+ recursion)
4. Assemble blocks     order the Prompt Blocks; resolve markers → concrete content
5. Macros              expand {{...}} across all text (after assembly)
6. Regex (prompt)      apply prompt-affecting regex (stacked scopes) to user input / blocks
7. Budget & truncate   count tokens; summarize-then-drop oldest history to fit
8. Adapt to method     build the method-specific request (Anthropic/OpenAI/Gemini/Responses)
9. Param filter        drop disabled / excluded sampler params
10. Send (stream)      direct fetch or native OkHttp/SSE plugin; apply fallback chain
11. Post men           regex (display) on output; write swipe; update tokens/cost
12. Log                persist assembled blocks + final payload + response + usage + ms
13. Agentic pass       (async, cadence/trigger) separate memory-agent request → review queue
```

## Default block order (SillyTavern-compatible)

```
1  Main / System prompt
2  World Info (before)
3  Persona description
4  Character description
5  Character personality
6  Scenario
7  World Info (after)
8  Dialogue examples
9  Chat history   ← author's note + summary injected at depth here
10 Post-history instructions / jailbreak
```

Blocks are fully reorderable per preset; **markers** (chatHistory, charDescription, worldInfoBefore, …) resolve to live content, **custom** blocks carry static macro-enabled text. Reordering in the editor literally changes this sequence — and the Log viewer shows the result.

## World-Info scanning

```
- Build scan text = last `scanDepth` messages (default 4) of the active branch, plus
  optionally persona/char fields.
- For each enabled entry across assigned + embedded books (chat fork shadows its parent):
    constant            → always include
    keyed               → include if a primary key matches scan text
    selective           → also require secondary keys per selectiveLogic (andAny/andAll/notAny/notAll)
    probability < 100   → roll; include on success
- Recursion: included entries' content is re-scanned for more keys, up to
  `recursionLimit` (default 2) passes; respects `tokenBudget` (default 25% of context).
- Placement: position → worldInfoBefore / worldInfoAfter blocks, or at_depth in chat,
  or as author's-note-adjacent. order breaks ties.
```

## Macro engine

A shared service expanding **SillyTavern-compatible** macros everywhere text is authored (blocks, regex replace, scenario, author's note, character fields). Evaluated **after** assembly so positional macros see final context; non-deterministic ones (`{{roll}}`, `{{random}}`, `{{pick}}`) evaluate **per generation**.

- Identity/content: `{{char}} {{user}} {{persona}} {{description}} {{personality}} {{scenario}} {{mesExamples}}`
- Context: `{{lastMessage}} {{lastUserMessage}} {{lastCharMessage}} {{input}}`
- Time/random: `{{time}} {{date}} {{weekday}} {{isotime}} {{random:a,b,c}} {{roll:dN}} {{pick:...}}`
- Formatting/control: `{{newline}} {{trim}} {{noop}} {{// comment}}`
- Variables: chat-scoped `{{setvar::k::v}} {{getvar::k}} {{addvar}} {{incvar}}` (stored on the chat) **and** global `{{setglobalvar}} {{getglobalvar}}` (in `Settings.globalVars`).
- MaiTavern extensions: `{{tracker::<name>}}` (live tracker value), memory hooks.

## Token budget & truncation (summarize-then-drop)

```
budget = effective.contextSize − maxResponseTokens − safetyMargin (default 256)
spent  = Σ tokens(all assembled blocks)            // approx tokenizer; exact for OpenAI-family
while spent > budget:
    take the oldest not-yet-summarized run of history messages
    → memory-summarize them into the rolling Summary (layered/hierarchical)
    → drop those raw messages from context, keep the Summary block
    recount
```

If summarization is disabled, fall back to **drop-oldest**; never silently exceed the window. `hiddenFromContext` messages are excluded up front. The live **context-budget meter** in the Preset editor reflects this math.

## Memory injection points

- **Summary** → injected within/above chat history at a configurable depth.
- **Manual memories (pinned)** → high-priority block near the top of system context.
- **RAG recall** → top-k relevant `docChunks` (+ optional semantic recall of old messages) inserted as a labeled block; k and threshold configurable.
- **Trackers** → a compact state block (only stats with `injectIntoContext`).
- **Author's note** → injected at `authorsNote.depth` in history.

## Provider adapter (method mapping)

One assembled, role-tagged message list → method-specific request:

| Method | System | Messages | Stream |
|---|---|---|---|
| `anthropic_messages` | top-level `system` | `messages[]` (user/assistant) | SSE deltas; header `anthropic-dangerous-direct-browser-access: true` |
| `openai_chat` | `messages[0]` system | `messages[]` | SSE `data:` chunks |
| `openai_responses` | `instructions` | `input[]` | SSE events |
| `gemini` | `systemInstruction` | `contents[]` (roles map to user/model) | `streamGenerateContent` |

- **Param filter:** disabled params (Parameters tab) and connection `paramFilter` include/exclude are applied here; method-incompatible params are dropped (e.g. `top_k` for Anthropic is allowed, but penalties unsupported by a method are stripped).
- **Transport:** direct streaming `fetch` for CORS-friendly endpoints; the **native Kotlin OkHttp/SSE plugin** for arbitrary/custom endpoints. **Fallback chain** retries down the list on error/timeout/rate-limit; **multi-key rotation** picks a healthy key.

## Agentic memory pass (async)

Runs on its **own request** off the main reply, on the user-configured **model + cadence** (every N messages / triggers / manual), capped by a per-session cost limit:

```
build a memory-agent prompt (recent history + current lorebook/char/trackers state)
→ tools: lorebookAdd/Update, charFieldEdit, trackerSet, rename, writeMemory
→ proposed changes → AgentAction[] (status: proposed)
→ review-first by default; if chat.agentAutonomous → apply directly (still logged, undoable)
→ writes target CHAT SCOPE (chat-lorebook fork / per-chat character override) — never canonical
→ opt-in batched notification → tap to review/approve/revert
```

## Logging

Every generation persists a `LogEntry`: the ordered assembled blocks (title/role/content/tokens), the final request payload (keys redacted), raw response, token usage + cost, and latency. The per-message **"view request log"** and the global Logs viewer read these.

## Defaults chosen (vetoable)

- **scanDepth = 4** messages; **recursionLimit = 2**; world-info **tokenBudget = 25%** of context.
- **safetyMargin = 256** tokens; tokenizer approximate (chars/4) + exact for OpenAI-family.
- **Default shipped preset** ("RP — Balanced"): temperature **0.9**, top_p **0.95** enabled; top_k / penalties **disabled**; maxResponseTokens **600**; streaming **on**; SillyTavern default block order above.
- **Group chat turn order (provisional, finalized when we design groups):** list/round-robin order, with per-member talkativeness as a later weight.
- **Macro var scope:** both chat + global supported.
