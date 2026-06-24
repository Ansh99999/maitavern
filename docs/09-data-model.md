# 09 — Data Model

The storage foundation. Everything we designed plugs into these entities. Local-first via **Dexie/IndexedDB**; large/append-heavy stores (logs, message archives) may move to **Capacitor SQLite** later without changing the shapes below.

## Conventions

- **IDs:** UUID v4 strings (stable across export/sync). Never reuse.
- **Timestamps:** `createdAt` / `updatedAt` as epoch-ms numbers. Bump `updatedAt` on every write (sync needs it).
- **Soft delete:** `deletedAt?: number` tombstones instead of hard deletes, so future cloud sync can reconcile. UI hides tombstoned rows; a periodic compaction purges old ones.
- **Secrets:** API keys are stored **encrypted** (Capacitor secure storage / EncryptedSharedPreferences-backed key wrapping the value); never written to logs or exports unless the user explicitly opts in.
- **Schema version:** a top-level `schemaVersion` drives Dexie migrations.

## Dexie tables

```
characters, lorebooks, lorebookEntries, personas, presets,
connections, modelRoles, chats, messages, branches,
memories, summaries, documents, docChunks,
themes, galleryAssets, regexScripts, logs, settings, kv
```

> `kv` holds singletons (global settings, global macro variables, app state). `messages`/`branches` are split out of `chats` so long chats stay light to query.

---

## Core entities

### Character (canonical / base)

Mirrors the **Character Card V2/V3** spec so import/export is loss-light.

```ts
interface Character {
  id: string;
  spec: 'chara_card_v2' | 'chara_card_v3';
  name: string;
  nickname?: string;                 // v3
  description: string;
  personality: string;
  scenario: string;
  firstMes: string;                  // greeting
  alternateGreetings: string[];
  groupOnlyGreetings?: string[];     // v3
  mesExample: string;                // dialogue examples
  systemPrompt?: string;
  postHistoryInstructions?: string;  // jailbreak
  creatorNotes?: string;
  tags: string[];
  creator?: string;
  characterVersion?: string;
  embeddedLorebookId?: string;       // -> Lorebook (character_book)
  assets: GalleryAssetRef[];         // v3 assets: avatar, expressions, gallery
  avatarAssetId?: string;
  // MaiTavern extensions (live in card `extensions` on export):
  trackers: TrackerDef[];            // see below
  theme?: ThemeOverride;             // per-character theme (colors, banner, bg)
  defaultPersonaId?: string;
  createdAt: number; updatedAt: number; deletedAt?: number;
}

interface TrackerDef {
  id: string; name: string;
  type: 'number' | 'bar' | 'text' | 'tag';
  initial: number | string;
  min?: number; max?: number;        // for bar
  visibleInChat: boolean;
  aiUpdatesLive: boolean;            // static sheet value vs live tracker
  injectIntoContext: boolean;
}
```

### Lorebook & entries (normal + chat)

```ts
interface Lorebook {
  id: string; name: string;
  kind: 'normal' | 'chat';           // chat = copy-on-write fork
  parentLorebookId?: string;         // chat books link to their origin
  chatId?: string;                   // owning chat (kind === 'chat')
  scanDepth?: number; tokenBudget?: number; recursiveScanning?: boolean;
  versions: LorebookVersion[];       // chat books keep rollback snapshots
  createdAt: number; updatedAt: number; deletedAt?: number;
}
interface LorebookVersion { id: string; at: number; note?: string; entryIds: string[]; }

interface LorebookEntry {
  id: string; lorebookId: string;
  keys: string[]; secondaryKeys: string[];
  comment: string;                   // title/name
  content: string;
  constant: boolean;                 // always on vs keyed
  selective: boolean;                // require secondary logic
  selectiveLogic: 'andAny' | 'andAll' | 'notAny' | 'notAll';
  position: 'before_char' | 'after_char' | 'at_depth' | 'an_top' | 'an_bottom';
  depth?: number; order: number; probability: number;  // 0..100
  caseSensitive?: boolean; useRegex?: boolean; enabled: boolean;
  createdAt: number; updatedAt: number;
}
```

**Copy-on-write:** the first in-chat edit of a `normal` book clones it into one `chat` book (`[Name] [Chat] #n`), linked by `parentLorebookId`; later edits update that same chat book and push a `LorebookVersion` snapshot (rollback). See [10-prompt-engine.md](10-prompt-engine.md) for scan/insertion.

### Persona

```ts
interface Persona { id: string; name: string; description: string; avatarAssetId?: string;
  createdAt: number; updatedAt: number; deletedAt?: number; }
```

### Preset

```ts
interface Preset {
  id: string; name: string;
  // General tab / identity:
  iconAssetId?: string; bannerAssetId?: string; themeColor?: string;
  notes?: string; tags: string[];
  defaultConnectionId?: string;      // SUGGESTION, overridable per chat
  contextSize: number; maxResponseTokens: number; streaming: boolean;
  // Prompt Blocks tab:
  promptBlocks: PromptBlock[];       // ordered; markers + custom
  // Parameters tab:
  parameters: Record<string, ParamValue>;   // keyed by param id (temperature, top_p, …)
  // Regex tab (preset-scoped; stacks with other scopes):
  regexScriptIds: string[];
  createdAt: number; updatedAt: number; deletedAt?: number;
}
interface PromptBlock {
  id: string; title: string;
  kind: 'marker' | 'custom';
  marker?: MarkerType;               // 'chatHistory' | 'charDescription' | 'personality' |
                                     // 'scenario' | 'persona' | 'dialogueExamples' |
                                     // 'worldInfoBefore' | 'worldInfoAfter' | …
  role: 'system' | 'user' | 'assistant';
  content?: string;                  // custom blocks (macro-enabled)
  position: 'relative' | 'at_depth'; depth?: number;
  enabled: boolean; forbidOverride?: boolean;
}
interface ParamValue { enabled: boolean; value: number | string | number[]; }
```

### RegexScript (multi-scope, stacking)

```ts
interface RegexScript {
  id: string; name: string;
  find: string; flags: string; replace: string;     // capture groups + macros
  affects: { userInput: boolean; aiOutput: boolean; displayOnly: boolean;
             promptOnly: boolean; slashCommands: boolean; };
  minDepth?: number; maxDepth?: number; runOnEdit: boolean; enabled: boolean; order: number;
  scope: 'global' | 'character' | 'preset' | 'chat';
  ownerId?: string;                  // character/preset/chat id when not global
}
```

### Connections (Provider / Router) + Model Roles

```ts
interface Connection {
  id: string; name: string; kind: 'provider' | 'router';
  category: 'text' | 'image' | 'voice';
  method: 'anthropic_messages' | 'openai_chat' | 'openai_responses' | 'gemini' |
          'image' | 'tts' | 'stt';
  baseUrl: string;
  apiKeys: EncryptedKey[];           // 1..n
  rotation?: 'round_robin' | 'random' | 'failover' | 'weighted';
  model?: string; headers?: Record<string,string>; contextWindow?: number;
  // Router extras:
  pricing?: Record<string, { inputPerM: number; outputPerM: number }>;  // by model
  trackCost?: boolean;               // also allowed on providers via manual pricing
  fallback?: { connectionId: string; model: string }[];
  customParams?: Record<string, unknown>;
  paramFilter?: { include?: string[]; exclude?: string[] };
  // image/voice specifics:
  imageDefaults?: Record<string, unknown>; voice?: { ttsVoice?: string; sttModel?: string };
  createdAt: number; updatedAt: number; deletedAt?: number;
}
interface EncryptedKey { id: string; cipher: string; health: 'ok' | 'rate_limited' | 'bad'; }

interface ModelRoles {                // singleton in `kv`
  roleplay?: ModelRef; mai?: ModelRef; memoryAgent?: ModelRef;
  embeddings?: ModelRef; image?: ModelRef; voice?: ModelRef;
}
interface ModelRef { connectionId: string; model: string; }
```

### Chat + Messages + Branches

```ts
interface Chat {
  id: string;
  title: string;
  characterIds: string[];            // 1 = solo, >1 = group
  personaId?: string;
  presetId?: string; connectionId?: string; model?: string;  // resolved at send if unset
  chatStyle?: ChatStyle;             // per-chat override
  // Per-chat OVERRIDE LAYER (sparse — only set keys override):
  characterOverrides: Record<string, Partial<Character>>;    // by characterId
  settingOverrides: Record<string, unknown>;                 // display/theme/etc.
  scenarioOverride?: string;
  authorsNote?: { text: string; depth: number; enabled: boolean };
  lorebookIds: string[];             // assigned books (incl. chat forks)
  regexScriptIds: string[];          // chat-scoped scripts
  // Memory:
  summaryId?: string;                // rolling summary
  memoryIds: string[];               // manual pinned memories
  trackerState: Record<string, number | string>;             // live tracker values (by trackerDef id)
  agentLog: AgentAction[];           // agentic memory activity + review queue
  agentAutonomous: boolean;          // per-chat auto opt-in (default false = review-first)
  // Branching:
  rootBranchId: string; activeBranchId: string;
  incognito?: boolean;               // ephemeral, not persisted to history
  createdAt: number; updatedAt: number; deletedAt?: number;
}

interface Branch {                   // the tree: lightweight
  id: string; chatId: string; parentBranchId?: string; forkedFromMessageId?: string;
  label?: string; createdAt: number;
}

interface Message {
  id: string; chatId: string; branchId: string;
  role: 'user' | 'assistant' | 'system';
  characterId?: string;              // which bot (group chats)
  swipes: Swipe[]; activeSwipe: number;   // alternate generations
  editHistory: { at: number; content: string }[];  // version history
  hiddenFromContext: boolean;        // "mute" but keep visible
  tokens?: number; createdAt: number; updatedAt: number; deletedAt?: number;
}
interface Swipe { content: string; at: number; meta?: { model?: string; logId?: string } }
```

**Branching = both:** the `Branch` tree gives lightweight in-chat branches (sharing history up to `forkedFromMessageId`); a **split** operation deep-copies a branch's message chain into a brand-new `Chat`.

### Memory: summaries, manual memories, RAG documents

```ts
interface Summary { id: string; chatId: string; text: string;
  coversUpToMessageId?: string; layers: { text: string; at: number }[]; // hierarchical
  editedByUser: boolean; updatedAt: number; }

interface Memory { id: string; scope: 'chat' | 'character'; ownerId: string;
  text: string; pinned: boolean; source: 'manual' | 'agent'; createdAt: number; }

interface AgentAction { id: string; at: number;
  kind: 'lorebook_add' | 'lorebook_update' | 'char_field' | 'tracker' | 'rename' | 'memory';
  summary: string; payload: unknown; status: 'proposed' | 'applied' | 'reverted'; }

interface Document { id: string; name: string; mime: string; assignedTo: string[]; createdAt: number; }
interface DocChunk { id: string; documentId: string; text: string; embedding: Float32Array; }
```

### Media, themes, logs, settings

```ts
interface GalleryAsset { id: string; kind: 'avatar'|'background'|'banner'|'image';
  blobKey: string; width: number; height: number; mime: string; createdAt: number; }
interface GalleryAssetRef { assetId: string; role?: string }

interface Theme { id: string; name: string; isBuiltin: boolean;
  vars: Record<string, string>;      // CSS-variable contract (colors, fonts, bubbles)
  customCss?: string; createdAt: number; updatedAt: number; }
interface ThemeOverride { themeId?: string; vars?: Record<string,string>;
  bannerAssetId?: string; backgroundAssetId?: string; }

interface LogEntry { id: string; chatId: string; messageId?: string; at: number;
  assembledBlocks: { title: string; role: string; content: string; tokens: number }[];
  requestPayload: unknown; responseRaw?: unknown;
  usage?: { in: number; out: number; cost?: number }; ms?: number; }

interface Settings { /* singleton in kv */ schemaVersion: number;
  globalDefaults: Record<string, unknown>;  // the Global layer of the override hierarchy
  globalVars: Record<string, unknown>;       // {{setglobalvar}} store
  appLock?: { enabled: boolean; method: 'pin' | 'biometric' };
  crashLog: { enabled: boolean }; }          // opt-in, off by default
}
```

---

## Override-layer mechanics (Global → Character → Chat)

Resolution is a **sparse merge**, lowest-wins:

```
effective(setting) = merge(
  Settings.globalDefaults[setting],          // Global
  Character.<field or settingOverride>,      // Character
  Chat.settingOverrides[setting]             // Chat
)
```

- Only overridden keys are stored at each level (sparse), so changing a global default still flows to anything that didn't override it.
- The **scope toggle** in the chat UI writes to the appropriate level (chat / character / global).
- Character *definition* edits from a chat write to `Chat.characterOverrides[charId]` (per-chat override); the prompt engine merges override-over-base when building character blocks.

## Backup / sync format (sync-ready from day one)

- Export = a `.zip`: `manifest.json` (schemaVersion, exportedAt) + one NDJSON per table + a `blobs/` dir for gallery assets. Keys excluded unless explicitly opted in.
- Every row carries `id` + `updatedAt` (+ `deletedAt` tombstones) → a future WebDAV/Drive sync is a last-writer-wins merge with no schema change.
- Per-entity export reuses the same row shapes (a character exports as a V2/V3 PNG/JSON; a chat as `.jsonl`).

## Defaults chosen (vetoable)

- **IDs** UUID v4; **timestamps** epoch-ms; **soft-delete** tombstones for sync-readiness.
- **Tokenizer:** approximate (chars/4 heuristic) by default, with exact tiktoken for OpenAI-family where feasible; budget uses the estimate + a safety margin (see doc 10).
- **Macro variables:** support **both** chat-scoped (`{{setvar}}`, stored on the chat) and global (`{{setglobalvar}}`, in `Settings.globalVars`).
- **Live tracker values** live on `Chat.trackerState` (per-chat), seeded from `Character.trackers[].initial`.
- **Agent default = review-first** (`agentAutonomous: false`).
- **Incognito chats** keep messages in memory only (not persisted), honoring the privacy setting.
