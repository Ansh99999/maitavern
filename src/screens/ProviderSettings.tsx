import { useEffect, useState } from 'react';
import { Link } from 'react-router-dom';
import { useLiveQuery } from 'dexie-react-hooks';
import { db } from '@/db/db';
import Icon from '@/components/Icon';
import { getSettings, makeConnection, patchSettings, touch } from '@/db/repo';
import type { Connection, ConnectionMethod } from '@/types';

/*
 * Provider settings (docs/07, "simple Provider" tier — Phase 1 scope is one
 * provider profile; Router, key rotation and the full trio land in Phase 2).
 * Presets for the two supported methods prefill base URL + a sane model.
 */

const METHOD_DEFAULTS: Record<ConnectionMethod, { baseUrl: string; model: string; label: string }> = {
  openai_chat: {
    label: 'OpenAI-compatible (chat completions)',
    baseUrl: 'https://api.openai.com/v1',
    model: 'gpt-4o-mini',
  },
  anthropic_messages: {
    label: 'Anthropic (messages)',
    baseUrl: 'https://api.anthropic.com/v1',
    model: 'claude-sonnet-5',
  },
};

export default function ProviderSettings() {
  const connections = useLiveQuery(
    () => db.connections.filter((c) => !c.deletedAt).sortBy('createdAt'),
    [],
  );
  const [activeId, setActiveId] = useState<string>();
  const [draft, setDraft] = useState<Connection>();
  const [savedAt, setSavedAt] = useState(0);

  useEffect(() => {
    getSettings().then((s) => setActiveId(s.activeConnectionId));
  }, []);

  // Edit the first (usually only) connection, or a fresh draft.
  useEffect(() => {
    if (!draft && connections) {
      setDraft(connections[0] ? { ...connections[0] } : makeConnection());
    }
  }, [connections, draft]);

  if (!draft) return null;

  const set = <K extends keyof Connection>(key: K, value: Connection[K]) =>
    setDraft({ ...draft, [key]: value });

  function setMethod(method: ConnectionMethod) {
    if (!draft) return;
    const d = METHOD_DEFAULTS[method];
    const untouched =
      draft.baseUrl === METHOD_DEFAULTS[draft.method].baseUrl || draft.baseUrl.trim() === '';
    setDraft({
      ...draft,
      method,
      baseUrl: untouched ? d.baseUrl : draft.baseUrl,
      model: draft.model || d.model,
    });
  }

  async function onSave() {
    if (!draft) return;
    await db.connections.put(touch({ ...draft }));
    await patchSettings({ activeConnectionId: draft.id });
    setActiveId(draft.id);
    setSavedAt(Date.now());
  }

  const input =
    'w-full px-3 py-2 rounded-xl bg-surface border border-border outline-none focus:border-accent';

  return (
    <div className="min-h-full bg-bg text-text font-ui">
      <header className="flex items-center gap-3 px-5 pt-4 pb-3 border-b border-border">
        <Link to="/settings" aria-label="Back" className="text-muted">
          <Icon name="chevronLeft" size={22} />
        </Link>
        <h1 className="text-lg font-semibold flex-1">Provider</h1>
        <button
          onClick={onSave}
          className="text-sm px-4 py-1.5 rounded-lg bg-accent text-bg font-medium"
        >
          {Date.now() - savedAt < 2000 ? 'Saved' : 'Save & activate'}
        </button>
      </header>

      <main className="px-5 py-4 flex flex-col gap-4 max-w-xl mx-auto pb-16">
        {activeId === draft.id && (
          <p className="text-xs text-accent">This is the active global provider.</p>
        )}
        <label className="block">
          <div className="text-sm font-medium mb-1">Name</div>
          <input className={input} value={draft.name} onChange={(e) => set('name', e.target.value)} />
        </label>
        <label className="block">
          <div className="text-sm font-medium mb-1">API format</div>
          <select
            className={input}
            value={draft.method}
            onChange={(e) => setMethod(e.target.value as ConnectionMethod)}
          >
            {Object.entries(METHOD_DEFAULTS).map(([value, d]) => (
              <option key={value} value={value}>{d.label}</option>
            ))}
          </select>
          <div className="text-xs text-muted mt-1">
            OpenAI-compatible covers OpenRouter, llama.cpp, LM Studio, vLLM, etc. — just change the
            base URL.
          </div>
        </label>
        <label className="block">
          <div className="text-sm font-medium mb-1">Base URL</div>
          <input
            className={input}
            value={draft.baseUrl}
            onChange={(e) => set('baseUrl', e.target.value)}
            placeholder={METHOD_DEFAULTS[draft.method].baseUrl}
            inputMode="url"
          />
        </label>
        <label className="block">
          <div className="text-sm font-medium mb-1">API key</div>
          <input
            className={input}
            type="password"
            value={draft.apiKey}
            onChange={(e) => set('apiKey', e.target.value)}
            placeholder="sk-…"
            autoComplete="off"
          />
          <div className="text-xs text-muted mt-1">
            Stored only on this device (BYOK). Never included in exports or logs.
          </div>
        </label>
        <label className="block">
          <div className="text-sm font-medium mb-1">Model</div>
          <input
            className={input}
            value={draft.model ?? ''}
            onChange={(e) => set('model', e.target.value || undefined)}
            placeholder={METHOD_DEFAULTS[draft.method].model}
          />
        </label>
      </main>
    </div>
  );
}
