import { useEffect, useState } from 'react';
import { Link, useNavigate, useParams } from 'react-router-dom';
import { db } from '@/db/db';
import { makeCharacter, saveCharacter } from '@/db/repo';
import { saveAvatarBlob } from '@/lib/cardImport';
import Avatar from '@/components/Avatar';
import type { Character } from '@/types';

/*
 * Basic character creator (Phase 1 MVP). A single scrolling form covering the
 * core card fields; the full tabbed Character Studio (docs/04) replaces this
 * shell later — field names already match the Character shape 1:1.
 */
export default function CharacterEditor() {
  const { characterId } = useParams();
  const navigate = useNavigate();
  const isNew = characterId === 'new';
  const [c, setC] = useState<Character | null>(isNew ? makeCharacter({ name: '' }) : null);
  const [saved, setSaved] = useState(false);

  useEffect(() => {
    if (!isNew && characterId) db.characters.get(characterId).then((row) => setC(row ?? null));
  }, [characterId, isNew]);

  if (!c) {
    return <div className="min-h-full bg-bg text-text grid place-items-center text-muted">Loading…</div>;
  }

  const set = <K extends keyof Character>(key: K, value: Character[K]) => {
    setC({ ...c, [key]: value });
    setSaved(false);
  };

  async function onSave() {
    if (!c || !c.name.trim()) return;
    await saveCharacter(c);
    setSaved(true);
    if (isNew) navigate(`/characters/${c.id}`, { replace: true });
  }

  async function onAvatarPick(files: FileList | null) {
    const file = files?.[0];
    if (!file || !c) return;
    set('avatarAssetId', await saveAvatarBlob(file));
  }

  return (
    <div className="min-h-full bg-bg text-text font-ui">
      <header className="flex items-center gap-3 px-5 pt-4 pb-3 border-b border-border sticky top-0 bg-bg z-10">
        <Link to="/characters" aria-label="Back" className="text-muted text-lg">
          ‹
        </Link>
        <h1 className="text-lg font-semibold flex-1 truncate">{isNew ? 'New character' : c.name}</h1>
        <button
          onClick={onSave}
          disabled={!c.name.trim()}
          className="text-sm px-4 py-1.5 rounded-lg bg-accent text-bg font-medium disabled:opacity-40"
        >
          {saved ? 'Saved ✓' : 'Save'}
        </button>
      </header>

      <main className="px-5 py-4 pb-16 flex flex-col gap-4 max-w-xl mx-auto">
        <label className="flex items-center gap-4">
          <Avatar assetId={c.avatarAssetId} name={c.name || '?'} className="w-20 h-20" />
          <div>
            <div className="text-sm font-medium mb-1">Display picture</div>
            <input type="file" accept="image/*" onChange={(e) => onAvatarPick(e.target.files)} className="text-xs text-muted" />
          </div>
        </label>

        <Field label="Name" value={c.name} onChange={(v) => set('name', v)} placeholder="Aria" />
        <Area label="Description" value={c.description} onChange={(v) => set('description', v)} rows={5}
          hint="Who they are. Supports {{char}} / {{user}} macros." />
        <Area label="Personality" value={c.personality} onChange={(v) => set('personality', v)} rows={3} />
        <Area label="Scenario" value={c.scenario} onChange={(v) => set('scenario', v)} rows={3}
          hint="Default setting for new chats (overridable per chat)." />
        <Area label="First message (greeting)" value={c.firstMes} onChange={(v) => set('firstMes', v)} rows={4} />
        <Area label="Dialogue examples" value={c.mesExample} onChange={(v) => set('mesExample', v)} rows={5}
          hint={'Separate examples with <START> lines, e.g.\n<START>\n{{user}}: hi\n{{char}}: hey'} />
        <Area label="System prompt (optional)" value={c.systemPrompt ?? ''} onChange={(v) => set('systemPrompt', v || undefined)} rows={3} />
        <Area label="Post-history instructions (optional)" value={c.postHistoryInstructions ?? ''} onChange={(v) => set('postHistoryInstructions', v || undefined)} rows={3} />
        <Field label="Tags (comma-separated)" value={c.tags.join(', ')}
          onChange={(v) => set('tags', v.split(',').map((s) => s.trim()).filter(Boolean))} placeholder="fantasy, sci-fi" />
      </main>
    </div>
  );
}

function Field({ label, value, onChange, placeholder }: {
  label: string; value: string; onChange: (v: string) => void; placeholder?: string;
}) {
  return (
    <label className="block">
      <div className="text-sm font-medium mb-1">{label}</div>
      <input
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className="w-full px-3 py-2 rounded-xl bg-surface border border-border placeholder:text-muted outline-none focus:border-accent"
      />
    </label>
  );
}

function Area({ label, value, onChange, rows, hint }: {
  label: string; value: string; onChange: (v: string) => void; rows: number; hint?: string;
}) {
  return (
    <label className="block">
      <div className="text-sm font-medium mb-1">{label}</div>
      <textarea
        value={value}
        onChange={(e) => onChange(e.target.value)}
        rows={rows}
        className="w-full px-3 py-2 rounded-xl bg-surface border border-border placeholder:text-muted outline-none focus:border-accent resize-y font-chat"
      />
      {hint && <div className="text-xs text-muted mt-1 whitespace-pre-line">{hint}</div>}
    </label>
  );
}
