/*
 * Phase 1 subset of the data model (docs/09-data-model.md).
 * Shapes match the doc exactly so later phases only ADD fields.
 */

export type ChatStyle = 'bubble' | 'document' | 'discord' | 'novel';

export interface Character {
  id: string;
  spec: 'chara_card_v2' | 'chara_card_v3';
  name: string;
  nickname?: string;
  description: string;
  personality: string;
  scenario: string;
  firstMes: string;
  alternateGreetings: string[];
  mesExample: string;
  systemPrompt?: string;
  postHistoryInstructions?: string;
  creatorNotes?: string;
  tags: string[];
  creator?: string;
  characterVersion?: string;
  avatarAssetId?: string;
  createdAt: number;
  updatedAt: number;
  deletedAt?: number;
}

export interface Preset {
  id: string;
  name: string;
  notes?: string;
  tags: string[];
  defaultConnectionId?: string;
  contextSize: number;
  maxResponseTokens: number;
  streaming: boolean;
  promptBlocks: PromptBlock[];
  parameters: Record<string, ParamValue>;
  createdAt: number;
  updatedAt: number;
  deletedAt?: number;
}

export type MarkerType =
  | 'chatHistory'
  | 'charDescription'
  | 'personality'
  | 'scenario'
  | 'persona'
  | 'dialogueExamples';

export interface PromptBlock {
  id: string;
  title: string;
  kind: 'marker' | 'custom';
  marker?: MarkerType;
  role: 'system' | 'user' | 'assistant';
  content?: string;
  position: 'relative' | 'at_depth';
  depth?: number;
  enabled: boolean;
}

export interface ParamValue {
  enabled: boolean;
  value: number | string | number[];
}

export type ConnectionMethod = 'anthropic_messages' | 'openai_chat';

export interface Connection {
  id: string;
  name: string;
  kind: 'provider';
  category: 'text';
  method: ConnectionMethod;
  baseUrl: string;
  apiKey: string; // Phase 1: plain local storage; encrypted-at-rest wrapper lands with native secure storage
  model?: string;
  headers?: Record<string, string>;
  contextWindow?: number;
  createdAt: number;
  updatedAt: number;
  deletedAt?: number;
}

export interface Chat {
  id: string;
  title: string;
  characterIds: string[];
  presetId?: string;
  connectionId?: string;
  model?: string;
  chatStyle?: ChatStyle;
  settingOverrides: Record<string, unknown>;
  scenarioOverride?: string;
  rootBranchId: string;
  activeBranchId: string;
  createdAt: number;
  updatedAt: number;
  deletedAt?: number;
}

export interface Branch {
  id: string;
  chatId: string;
  parentBranchId?: string;
  forkedFromMessageId?: string;
  label?: string;
  createdAt: number;
}

export interface Message {
  id: string;
  chatId: string;
  branchId: string;
  role: 'user' | 'assistant' | 'system';
  characterId?: string;
  swipes: Swipe[];
  activeSwipe: number;
  editHistory: { at: number; content: string }[];
  hiddenFromContext: boolean;
  tokens?: number;
  createdAt: number;
  updatedAt: number;
  deletedAt?: number;
}

export interface Swipe {
  content: string;
  at: number;
  meta?: { model?: string; logId?: string };
}

export interface GalleryAsset {
  id: string;
  kind: 'avatar' | 'background' | 'banner' | 'image';
  blob: Blob;
  width: number;
  height: number;
  mime: string;
  createdAt: number;
}

export interface LogEntry {
  id: string;
  chatId: string;
  messageId?: string;
  at: number;
  assembledBlocks: { title: string; role: string; content: string; tokens: number }[];
  requestPayload: unknown;
  responseRaw?: unknown;
  usage?: { in: number; out: number };
  ms?: number;
}

export interface Settings {
  schemaVersion: number;
  userName: string; // {{user}} until the Persona system (Phase 2)
  themeId: string;
  chatStyle: ChatStyle;
  showAvatars: Partial<Record<ChatStyle, boolean>>;
  activePresetId?: string;
  activeConnectionId?: string;
}

/** Convenient helper for the message content the UI shows/sends. */
export function activeContent(m: Message): string {
  return m.swipes[m.activeSwipe]?.content ?? '';
}
