import type { Connection, ParamValue, Preset } from '@/types';
import type { RoleMessage } from './assemble';

/*
 * Provider adapter (docs/10 method mapping). One role-tagged message list →
 * a method-specific request. The param filter drops disabled params and
 * anything the method doesn't support.
 */

export interface AdaptedRequest {
  url: string;
  headers: Record<string, string>;
  body: unknown;
}

const METHOD_PARAMS: Record<Connection['method'], string[]> = {
  anthropic_messages: ['temperature', 'top_p', 'top_k'],
  openai_chat: ['temperature', 'top_p', 'frequency_penalty', 'presence_penalty'],
};

function enabledParams(preset: Preset, method: Connection['method']): Record<string, ParamValue['value']> {
  const allowed = METHOD_PARAMS[method];
  const out: Record<string, ParamValue['value']> = {};
  for (const [key, pv] of Object.entries(preset.parameters)) {
    if (pv.enabled && allowed.includes(key)) out[key] = pv.value;
  }
  return out;
}

/** Merge consecutive same-role messages (both APIs dislike system runs mid-list). */
function coalesce(messages: RoleMessage[]): RoleMessage[] {
  const out: RoleMessage[] = [];
  for (const m of messages) {
    const prev = out[out.length - 1];
    if (prev && prev.role === m.role) prev.content += `\n\n${m.content}`;
    else out.push({ ...m });
  }
  return out;
}

export function adaptRequest(
  connection: Connection,
  preset: Preset,
  messages: RoleMessage[],
  stream: boolean,
): AdaptedRequest {
  const base = connection.baseUrl.replace(/\/+$/, '');
  const params = enabledParams(preset, connection.method);
  const model = connection.model ?? '';

  if (connection.method === 'anthropic_messages') {
    // System blocks → top-level `system`; the rest must alternate user/assistant.
    const system = messages.filter((m) => m.role === 'system').map((m) => m.content).join('\n\n');
    let convo = coalesce(messages.filter((m) => m.role !== 'system'));
    if (!convo.length || convo[0].role !== 'user') {
      convo = [{ role: 'user', content: '(begin)' }, ...convo];
      convo = coalesce(convo);
    }
    return {
      url: `${base}/messages`,
      headers: {
        'content-type': 'application/json',
        'x-api-key': connection.apiKey,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
        ...connection.headers,
      },
      body: {
        model,
        system: system || undefined,
        messages: convo,
        max_tokens: preset.maxResponseTokens,
        stream,
        ...params,
      },
    };
  }

  // openai_chat
  return {
    url: `${base}/chat/completions`,
    headers: {
      'content-type': 'application/json',
      authorization: `Bearer ${connection.apiKey}`,
      ...connection.headers,
    },
    body: {
      model,
      messages: coalesce(messages),
      max_tokens: preset.maxResponseTokens,
      stream,
      ...params,
    },
  };
}
