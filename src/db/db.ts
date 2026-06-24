import Dexie, { type Table } from 'dexie';

/*
 * Local-first storage. Schema mirrors docs/09-data-model.md.
 * Only indexed fields are listed in the store strings; full row shapes are the
 * TypeScript interfaces in src/types. Phase 0 stands up the core tables;
 * later phases add documents/docChunks, logs (may move to SQLite), etc.
 */
export interface BaseRow {
  id: string;
  createdAt: number;
  updatedAt: number;
  deletedAt?: number;
}

export class MaiTavernDB extends Dexie {
  characters!: Table<BaseRow, string>;
  lorebooks!: Table<BaseRow, string>;
  lorebookEntries!: Table<BaseRow, string>;
  personas!: Table<BaseRow, string>;
  presets!: Table<BaseRow, string>;
  connections!: Table<BaseRow, string>;
  chats!: Table<BaseRow, string>;
  messages!: Table<BaseRow & { chatId: string; branchId: string }, string>;
  branches!: Table<BaseRow & { chatId: string }, string>;
  memories!: Table<BaseRow, string>;
  summaries!: Table<BaseRow, string>;
  themes!: Table<BaseRow, string>;
  galleryAssets!: Table<BaseRow, string>;
  regexScripts!: Table<BaseRow, string>;
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
  }
}

export const db = new MaiTavernDB();
