import { useEffect, useState } from 'react';
import { Link } from 'react-router-dom';
import { useLiveQuery } from 'dexie-react-hooks';
import { db } from '@/db/db';
import Icon from '@/components/Icon';
import { getSettings, makeDefaultPreset, patchSettings, touch } from '@/db/repo';
import type { ParamValue, Preset } from '@/types';

/*
 * Sampler preset editor (docs/08, Phase 1 slice): General (context size, max
 * response, streaming) + Parameters with per-param enable toggles. The Prompt
 * Blocks reorder UI and Regex tab arrive with the full Prompt Manager.
 */

const PARAM_META: Record<string, { label: string; min: number; max: number; step: number }> = {
  temperature: { label: 'Temperature', min: 0, max: 2, step: 0.05 },
  top_p: { label: 'Top P', min: 0, max: 1, step: 0.01 },
  top_k: { label: 'Top K', min: 1, max: 200, step: 1 },
  frequency_penalty: { label: 'Frequency penalty', min: -2, max: 2, step: 0.05 },
  presence_penalty: { label: 'Presence penalty', min: -2, max: 2, step: 0.05 },
};

export default function PresetSettings() {
  const presets = useLiveQuery(() => db.presets.filter((p) => !p.deletedAt).sortBy('name'), []);
  const [activeId, setActiveId] = useState<string>();
  const [draft, setDraft] = useState<Preset>();
  const [savedAt, setSavedAt] = useState(0);

  useEffect(() => {
    getSettings().then((s) => setActiveId(s.activePresetId));
  }, []);

  useEffect(() => {
    if (!draft && presets?.length && activeId !== undefined) {
      const active = presets.find((p) => p.id === activeId) ?? presets[0];
      setDraft({ ...active });
    }
  }, [presets, draft, activeId]);

  if (!draft) return null;

  async function onSave() {
    if (!draft) return;
    await db.presets.put(touch({ ...draft }));
    await patchSettings({ activePresetId: draft.id });
    setActiveId(draft.id);
    setSavedAt(Date.now());
  }

  async function onDuplicate() {
    if (!draft) return;
    const copy = { ...makeDefaultPreset(), ...draft, id: crypto.randomUUID(), name: `${draft.name} (copy)`, tags: [] };
    await db.presets.add(copy);
    setDraft(copy);
  }

  function setParam(key: string, patch: Partial<ParamValue>) {
    if (!draft) return;
    const current = draft.parameters[key] ?? { enabled: false, value: 0 };
    setDraft({
      ...draft,
      parameters: { ...draft.parameters, [key]: { ...current, ...patch } },
    });
  }

  const input =
    'w-full px-3 py-2 rounded-xl bg-surface border border-border outline-none focus:border-accent';

  return (
    <div className="min-h-full bg-bg text-text font-ui">
      <header className="flex items-center gap-3 px-5 pt-4 pb-3 border-b border-border">
        <Link to="/settings" aria-label="Back" className="text-muted">
          <Icon name="chevronLeft" size={22} />
        </Link>
        <h1 className="text-lg font-semibold flex-1">Preset</h1>
        <button onClick={onDuplicate} className="text-sm px-3 py-1.5 rounded-lg bg-surface border border-border">
          Duplicate
        </button>
        <button onClick={onSave} className="text-sm px-4 py-1.5 rounded-lg bg-accent text-bg font-medium">
          {Date.now() - savedAt < 2000 ? 'Saved' : 'Save & activate'}
        </button>
      </header>

      <main className="px-5 py-4 flex flex-col gap-4 max-w-xl mx-auto pb-16">
        <label className="block">
          <div className="text-sm font-medium mb-1">Preset</div>
          <select
            className={input}
            value={draft.id}
            onChange={(e) => {
              const p = presets?.find((x) => x.id === e.target.value);
              if (p) setDraft({ ...p });
            }}
          >
            {(presets ?? []).map((p) => (
              <option key={p.id} value={p.id}>
                {p.name}{p.id === activeId ? ' · active' : ''}
              </option>
            ))}
          </select>
        </label>

        <label className="block">
          <div className="text-sm font-medium mb-1">Name</div>
          <input className={input} value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} />
        </label>

        <div className="grid grid-cols-2 gap-3">
          <label className="block">
            <div className="text-sm font-medium mb-1">Context size</div>
            <input
              className={input}
              type="number"
              min={512}
              value={draft.contextSize}
              onChange={(e) => setDraft({ ...draft, contextSize: Number(e.target.value) || 0 })}
            />
          </label>
          <label className="block">
            <div className="text-sm font-medium mb-1">Max response tokens</div>
            <input
              className={input}
              type="number"
              min={16}
              value={draft.maxResponseTokens}
              onChange={(e) => setDraft({ ...draft, maxResponseTokens: Number(e.target.value) || 0 })}
            />
          </label>
        </div>

        <label className="flex items-center gap-2 text-sm">
          <input
            type="checkbox"
            checked={draft.streaming}
            onChange={(e) => setDraft({ ...draft, streaming: e.target.checked })}
          />
          Stream responses
        </label>

        <section>
          <h2 className="text-sm font-semibold text-muted uppercase tracking-wide mb-2">Sampler parameters</h2>
          <p className="text-xs text-muted mb-3">
            Disabled parameters are omitted from the request; parameters a provider doesn't support
            are dropped automatically.
          </p>
          {Object.entries(PARAM_META).map(([key, meta]) => {
            const pv = draft.parameters[key] ?? { enabled: false, value: meta.min };
            return (
              <div key={key} className="flex items-center gap-3 mb-3">
                <input
                  type="checkbox"
                  checked={pv.enabled}
                  onChange={(e) => setParam(key, { enabled: e.target.checked })}
                />
                <div className="flex-1">
                  <div className="flex justify-between text-sm">
                    <span className={pv.enabled ? '' : 'text-muted'}>{meta.label}</span>
                    <span className="text-muted">{pv.value}</span>
                  </div>
                  <input
                    type="range"
                    className="w-full accent-[rgb(var(--mt-accent))]"
                    min={meta.min}
                    max={meta.max}
                    step={meta.step}
                    value={Number(pv.value)}
                    disabled={!pv.enabled}
                    onChange={(e) => setParam(key, { value: Number(e.target.value) })}
                  />
                </div>
              </div>
            );
          })}
        </section>
      </main>
    </div>
  );
}
