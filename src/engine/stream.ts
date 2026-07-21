import type { Connection } from '@/types';
import type { AdaptedRequest } from './adapters';

/*
 * Streaming transport (docs/10 step 10): direct streaming fetch + SSE parsing
 * for both methods. The native OkHttp/SSE plugin for arbitrary endpoints
 * arrives in Phase 2; CORS-friendly endpoints work today.
 */

export interface StreamResult {
  text: string;
  usage?: { in: number; out: number };
  raw?: unknown;
}

export async function streamCompletion(
  connection: Connection,
  req: AdaptedRequest,
  stream: boolean,
  onDelta: (fullText: string) => void,
  signal: AbortSignal,
): Promise<StreamResult> {
  const res = await fetch(req.url, {
    method: 'POST',
    headers: req.headers,
    body: JSON.stringify(req.body),
    signal,
  });
  if (!res.ok) {
    const detail = await res.text().catch(() => '');
    throw new Error(`${res.status} ${res.statusText}${detail ? ` — ${detail.slice(0, 400)}` : ''}`);
  }

  if (!stream || !res.body) {
    const json = await res.json();
    return { text: extractFullText(connection, json), usage: extractUsage(connection, json), raw: json };
  }

  let text = '';
  let usage: StreamResult['usage'];
  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    // SSE frames are separated by blank lines; keep the trailing partial frame.
    const frames = buffer.split(/\r?\n\r?\n/);
    buffer = frames.pop() ?? '';
    for (const frame of frames) {
      for (const line of frame.split(/\r?\n/)) {
        if (!line.startsWith('data:')) continue;
        const payload = line.slice(5).trim();
        if (!payload || payload === '[DONE]') continue;
        let evt: unknown;
        try {
          evt = JSON.parse(payload);
        } catch {
          continue;
        }
        const delta = extractDelta(connection, evt);
        if (delta) {
          text += delta;
          onDelta(text);
        }
        usage = extractStreamUsage(connection, evt, usage);
      }
    }
  }
  return { text, usage };
}

/* eslint-disable @typescript-eslint/no-explicit-any */

function extractDelta(c: Connection, evt: any): string {
  if (c.method === 'anthropic_messages') {
    return evt?.type === 'content_block_delta' ? (evt.delta?.text ?? '') : '';
  }
  return evt?.choices?.[0]?.delta?.content ?? '';
}

function extractStreamUsage(
  c: Connection,
  evt: any,
  prev: StreamResult['usage'],
): StreamResult['usage'] {
  if (c.method === 'anthropic_messages') {
    if (evt?.type === 'message_start') {
      return { in: evt.message?.usage?.input_tokens ?? 0, out: prev?.out ?? 0 };
    }
    if (evt?.type === 'message_delta' && evt.usage) {
      return { in: prev?.in ?? 0, out: evt.usage.output_tokens ?? 0 };
    }
    return prev;
  }
  if (evt?.usage) return { in: evt.usage.prompt_tokens ?? 0, out: evt.usage.completion_tokens ?? 0 };
  return prev;
}

function extractFullText(c: Connection, json: any): string {
  if (c.method === 'anthropic_messages') {
    return (json?.content ?? [])
      .filter((b: any) => b.type === 'text')
      .map((b: any) => b.text)
      .join('');
  }
  return json?.choices?.[0]?.message?.content ?? '';
}

function extractUsage(c: Connection, json: any): StreamResult['usage'] {
  const u = json?.usage;
  if (!u) return undefined;
  if (c.method === 'anthropic_messages') return { in: u.input_tokens ?? 0, out: u.output_tokens ?? 0 };
  return { in: u.prompt_tokens ?? 0, out: u.completion_tokens ?? 0 };
}
