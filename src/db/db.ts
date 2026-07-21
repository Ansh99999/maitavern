import Dexie, { type Table } from 'dexie';
import type {
  Branch,
  Character,
  Chat,
  Connection,
  GalleryAsset,
  LogEntry,
  Message,
  Preset,
} from '@/types';

/*
 * Local-first storage. Schema mirrors docs/09-data-model.md.
 * Only indexed fields are listed in the store strings; full row shapes are the
 * TypeScript interfaces in src/types. v2 types the Phase 1 tables and adds
 * logs; later phases add documents/docChunks etc. (logs may move to SQLite).
 */
export interface BaseRow {
  id: string;
  createdAt: number;
  updatedAt: number;
  deletedAt?: number;
}

export class MaiTavernDB extends Dexie {
  characters!: Table<Character, string>;
  lorebooks!: Table<BaseRow, string>;
  lorebookEntries!: Table<BaseRow, string>;
  personas!: Table<BaseRow, string>;
  presets!: Table<Preset, string>;
  connections!: Table<Connection, string>;
  chats!: Table<Chat, string>;
  messages!: Table<Message, string>;
  branches!: Table<Branch, string>;
  memories!: Table<BaseRow, string>;
  summaries!: Table<BaseRow, string>;
  themes!: Table<BaseRow, string>;
  galleryAssets!: Table<GalleryAsset, string>;
  regexScripts!: Table<BaseRow, string>;
  logs!: Table<LogEntry, string>;
  kv!: Table<{ key: string; value: unknown }, string>;

  constructor() {
    super('maitavern');
    this.version(1).stores({
      characters: 'id, name, updatedAt, deletedAt',
      lorebooks: 'id, kind, parentLorebookId, chatId, updatedAt, deletedAt',
      lorebookEntries: 'id, lorebookId, updatedAt',
      personas: 'id, name, updatedAt, deletedAt',
      presets: 'id, name, updatedAt, deletedAt',
      connections: 'id, category, kind, updatedAt, deletedAt',
      chats: 'id, title, updatedAt, deletedAt, *characterIds',
      messages: 'id, chatId, branchId, createdAt, deletedAt',
      branches: 'id, chatId, parentBranchId',
      memories: 'id, scope, ownerId, updatedAt',
      summaries: 'id, chatId, updatedAt',
      themes: 'id, name, updatedAt',
      galleryAssets: 'id, kind, createdAt',
      regexScripts: 'id, scope, ownerId, order',
      kv: 'key',
    });
    this.version(2).stores({
      logs: 'id, chatId, messageId, at',
    });
  }
}

export const db = new MaiTavernDB();
