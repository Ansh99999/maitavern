import { db } from './db';
import { now, uuid } from '@/lib/id';
import type {
  Character,
  Chat,
  Connection,
  Message,
  Preset,
  PromptBlock,
  Settings,
} from '@/types';

const SETTINGS_KEY = 'settings';

export const DEFAULT_SETTINGS: Settings = {
  schemaVersion: 2,
  userName: 'You',
  themeId: 'lounge-dark',
  chatStyle: 'bubble',
  showAvatars: { bubble: true, document: true, discord: true, novel: false },
};

export async function getSettings(): Promise<Settings> {
  const row = await db.kv.get(SETTINGS_KEY);
  return { ...DEFAULT_SETTINGS, ...((row?.value as Partial<Settings>) ?? {}) };
}

export async function patchSettings(patch: Partial<Settings>): Promise<Settings> {
  const next = { ...(await getSettings()), ...patch };
  await db.kv.put({ key: SETTINGS_KEY, value: next });
  return next;
}

/** SillyTavern-compatible default block order (docs/10). */
export function defaultPromptBlocks(): PromptBlock[] {
  const marker = (title: string, m: PromptBlock['marker']): PromptBlock => ({
    id: uuid(),
    title,
    kind: 'marker',
    marker: m,
    role: 'system',
    position: 'relative',
    enabled: true,
  });
  return [
    {
      id: uuid(),
      title: 'Main prompt',
      kind: 'custom',
      role: 'system',
      content:
        "Write {{char}}'s next reply in a fictional roleplay chat between {{char}} and {{user}}. Stay in character.",
      position: 'relative',
      enabled: true,
    },
    marker('Persona', 'persona'),
    marker('Character description', 'charDescription'),
    marker('Character personality', 'personality'),
    marker('Scenario', 'scenario'),
    marker('Dialogue examples', 'dialogueExamples'),
    marker('Chat history', 'chatHistory'),
  ];
}

/** Default shipped preset "RP — Balanced" (docs/10 defaults). */
export function makeDefaultPreset(): Preset {
  const t = now();
  return {
    id: uuid(),
    name: 'RP — Balanced',
    tags: ['built-in'],
    contextSize: 8192,
    maxResponseTokens: 600,
    streaming: true,
    promptBlocks: defaultPromptBlocks(),
    parameters: {
      temperature: { enabled: true, value: 0.9 },
      top_p: { enabled: true, value: 0.95 },
      top_k: { enabled: false, value: 40 },
      frequency_penalty: { enabled: false, value: 0 },
      presence_penalty: { enabled: false, value: 0 },
    },
    createdAt: t,
    updatedAt: t,
  };
}

/** Idempotent first-run seeding: settings singleton + one default preset. */
export async function seed(): Promise<void> {
  await db.transaction('rw', db.kv, db.presets, async () => {
    const existing = await db.kv.get(SETTINGS_KEY);
    if (existing) return;
    const preset = makeDefaultPreset();
    await db.presets.add(preset);
    await db.kv.put({
      key: SETTINGS_KEY,
      value: { ...DEFAULT_SETTINGS, activePresetId: preset.id },
    });
  });
}

export function touch<T extends { updatedAt: number }>(row: T): T {
  row.updatedAt = now();
  return row;
}

// ---- characters ----

export function makeCharacter(partial: Partial<Character> = {}): Character {
  const t = now();
  return {
    id: uuid(),
    spec: 'chara_card_v2',
    name: 'New character',
    description: '',
    personality: '',
    scenario: '',
    firstMes: '',
    alternateGreetings: [],
    mesExample: '',
    tags: [],
    createdAt: t,
    updatedAt: t,
    ...partial,
  };
}

export async function saveCharacter(c: Character): Promise<void> {
  await db.characters.put(touch(c));
}

// ---- chats & messages ----

/** Create a chat with its root branch and the character's greeting (if any). */
export async function createChat(character: Character): Promise<Chat> {
  const t = now();
  const branchId = uuid();
  const chat: Chat = {
    id: uuid(),
    title: character.name,
    characterIds: [character.id],
    settingOverrides: {},
    rootBranchId: branchId,
    activeBranchId: branchId,
    createdAt: t,
    updatedAt: t,
  };
  await db.transaction('rw', db.chats, db.branches, db.messages, async () => {
    await db.chats.add(chat);
    await db.branches.add({ id: branchId, chatId: chat.id, createdAt: t });
    if (character.firstMes.trim()) {
      await db.messages.add(
        makeMessage(chat, 'assistant', character.firstMes, { characterId: character.id }),
      );
    }
  });
  return chat;
}

export function makeMessage(
  chat: Chat,
  role: Message['role'],
  content: string,
  extra: Partial<Message> = {},
): Message {
  const t = now();
  return {
    id: uuid(),
    chatId: chat.id,
    branchId: chat.activeBranchId,
    role,
    swipes: [{ content, at: t }],
    activeSwipe: 0,
    editHistory: [],
    hiddenFromContext: false,
    createdAt: t,
    updatedAt: t,
    ...extra,
  };
}

export async function branchMessages(chat: Chat): Promise<Message[]> {
  const rows = await db.messages
    .where('chatId')
    .equals(chat.id)
    .and((m) => m.branchId === chat.activeBranchId && !m.deletedAt)
    .sortBy('createdAt');
  return rows;
}

export async function touchChat(chatId: string): Promise<void> {
  await db.chats.update(chatId, { updatedAt: now() });
}

// ---- connections ----

export function makeConnection(partial: Partial<Connection> = {}): Connection {
  const t = now();
  return {
    id: uuid(),
    name: 'My provider',
    kind: 'provider',
    category: 'text',
    method: 'openai_chat',
    baseUrl: 'https://api.openai.com/v1',
    apiKey: '',
    createdAt: t,
    updatedAt: t,
    ...partial,
  };
}
