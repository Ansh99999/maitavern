import { db } from '@/db/db';
import { branchMessages, getSettings, touchChat } from '@/db/repo';
import { now, uuid } from '@/lib/id';
import type { Character, Chat, Connection, LogEntry, Message, Preset } from '@/types';
import { assemblePrompt } from './assemble';
import { adaptRequest } from './adapters';
import { streamCompletion } from './stream';

/*
 * Generation orchestrator (docs/10 pipeline). Resolves scope, assembles,
 * adapts, streams, then persists the result + a LogEntry. The caller owns the
 * target Message row (already inserted) and receives deltas for live display.
 */

export class GenerateConfigError extends Error {}

export interface ResolvedGenConfig {
  preset: Preset;
  connection: Connection;
}

/** Chat overrides win over global settings (docs/09 override mechanics). */
export async function resolveGenConfig(chat: Chat): Promise<ResolvedGenConfig> {
  const settings = await getSettings();
  const presetId = chat.presetId ?? settings.activePresetId;
  const preset = presetId ? await db.presets.get(presetId) : undefined;
  if (!preset) throw new GenerateConfigError('No preset selected — pick one in Settings.');
  const connectionId =
    chat.connectionId ?? settings.activeConnectionId ?? preset.defaultConnectionId;
  const connection = connectionId ? await db.connections.get(connectionId) : undefined;
  if (!connection) {
    throw new GenerateConfigError('No provider connected — add one in Settings → Provider.');
  }
  if (!connection.apiKey) throw new GenerateConfigError(`Provider "${connection.name}" has no API key.`);
  if (!connection.model) throw new GenerateConfigError(`Provider "${connection.name}" has no model set.`);
  return { preset, connection };
}

export interface GenerateArgs {
  chat: Chat;
  character: Character;
  /** message row receiving the reply (already in db, may have empty content) */
  target: Message;
  /** history to prompt with (oldest→newest, EXCLUDING the target row) */
  history: Message[];
  onDelta: (text: string) => void;
  signal: AbortSignal;
}

/** Run one generation; persists the reply text into `target` + a LogEntry. Returns final text. */
export async function generateReply(args: GenerateArgs): Promise<string> {
  const { chat, character, target } = args;
  const settings = await getSettings();
  const { preset, connection } = await resolveGenConfig(chat);

  const assembled = assemblePrompt({
    character,
    settings,
    preset,
    history: args.history,
    scenarioOverride: chat.scenarioOverride,
  });
  const req = adaptRequest(connection, preset, assembled.messages, preset.streaming);

  const started = now();
  let finalText = '';
  let aborted = false;
  let usage: { in: number; out: number } | undefined;
  try {
    const result = await streamCompletion(
      connection,
      req,
      preset.streaming,
      (text) => {
        finalText = text;
        args.onDelta(text);
      },
      args.signal,
    );
    finalText = result.text || finalText;
    usage = result.usage;
  } catch (err) {
    // A user Stop keeps the partial text; anything else propagates after logging.
    if ((err as Error).name === 'AbortError') aborted = true;
    else {
      await writeLog(chat, target, assembled.assembledBlocks, req.body, started, usage, String(err));
      throw err;
    }
  }

  const logId = await writeLog(chat, target, assembled.assembledBlocks, req.body, started, usage);
  const t = now();
  const swipes = [...target.swipes];
  swipes[target.activeSwipe] = {
    content: finalText || (aborted ? '' : finalText),
    at: t,
    meta: { model: connection.model, logId },
  };
  await db.messages.update(target.id, { swipes, updatedAt: t });
  await touchChat(chat.id);
  return finalText;
}

async function writeLog(
  chat: Chat,
  target: Message,
  assembledBlocks: LogEntry['assembledBlocks'],
  requestPayload: unknown,
  started: number,
  usage?: { in: number; out: number },
  error?: string,
): Promise<string> {
  const id = uuid();
  const entry: LogEntry = {
    id,
    chatId: chat.id,
    messageId: target.id,
    at: started,
    assembledBlocks,
    requestPayload, // headers (and the API key) are never included
    responseRaw: error ? { error } : undefined,
    usage,
    ms: now() - started,
  };
  await db.logs.add(entry);
  return id;
}

/** Reload the active branch history for a chat, excluding one message id. */
export async function historyExcluding(chat: Chat, excludeId: string): Promise<Message[]> {
  return (await branchMessages(chat)).filter((m) => m.id !== excludeId);
}
