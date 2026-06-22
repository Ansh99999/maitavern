# 06 — Memory System

Three types of memory, all of which become **blocks in the prompt assembler** (and therefore visible in the Log viewer).

## 1. Manual (user-authored)

Self-made **lorebook** entries · **scenario snippets** · **summary states** (hand-written/edited). Plus the per-message "Save to memory/lorebook" action and highlight → "remember this." Pinnable, scope-able (chat or character).

## 2. Automated (summarization)

A **Summary section** you can **tap to summarize now**. Triggers: **every 20 / every X messages**, and **summarize the last X messages**. The result is a rolling, editable summary injected at a configurable depth.

## 3. Agentic (the autonomous one)

A **background agent** running in a hidden / semi-hidden "terminal" — a **separate request** from the main reply (not a JSON-schema bolt-on). Tools let it **create/update lorebook entries, edit specific points of the character description, update Stats (trackers), even rename the character** via commands + self-text-insertion. It decides when to act; **opt-in notifications** keep you informed.

## Copy-on-write (the unifying philosophy)

- An **in-chat lorebook edit forks** the lorebook into a **chat lorebook** named `[Lorebook Name] [Chat Name] #[copy number]`; the original is untouched. The Lorebook section distinguishes **normal** vs **chat** lorebooks.
- Agentic edits to character description / stats / name land in the chat's **per-chat override layer** (same mechanism as the New Chat flow), never the canonical character.
- So the agent is **safe by construction** — canonical assets change only when *you* **promote/canonize**.

## Productive additions

1. **Review & undo gate + Agent Activity log** — every agentic change logged with one-tap undo; auto-apply vs review-first mode; notification → approve/revert. Mirrors Mai's accept/reject diffs.
2. **Unified "🧠 Memory" panel in chat** (tabs: Summary · Manual · Lorebook-fired · Trackers · Agent log) — the home for the tap-to-summarize section.
3. **Smarter summarization** — trigger "when context > Y%" + **hierarchical summary-of-summaries** for very long chats.
4. **Semantic recall** — RAG over past messages (reuse the Library embedding pipeline) to pull relevant old messages back into context.
5. **Dedicated cheaper "memory model" + cadence + cost cap** — agent runs async every N messages / on triggers, not every turn (BYOK = the user pays).
6. **Canonize/promote** — graduate chat-lorebook discoveries / summary / agentic edits into the canonical character/lorebook.
7. **Batched notifications** + memory blocks shown inside the request Log.

**Risks:** name changes touch `{{char}}` references everywhere (update refs, keep chat-scoped); background-agent latency/cost (mitigated by a cheaper model + cadence).

## Confirmed decisions

- **Lorebook fork granularity = one fork per chat + version history.** First in-chat edit forks the lorebook once for that chat; later edits update that same chat-lorebook but keep internal version snapshots (rollback). `#[copy number]` only for rare manual duplicates.
- **Agentic apply mode = review-first, with per-chat auto opt-in.** Proposals wait in a review queue / notification; a chat can be flipped to fully autonomous when trusted.
- **Memory-agent model + cadence = user-configurable.** Choose the model (a dedicated cheaper memory model, Mai's assistant model, or the main chat model) and the cadence (every N messages / on triggers / manual), plus a per-session cost cap. No model is forced.
