import { useRef, useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import { useLiveQuery } from 'dexie-react-hooks';
import { db } from '@/db/db';
import { createChat } from '@/db/repo';
import { importCharacterFile, characterToCardJson } from '@/lib/cardImport';
import Avatar from '@/components/Avatar';
import type { Character } from '@/types';

/*
 * Character gallery (docs/02 "Characters → gallery/management").
 * Phase 1: grid + search, ST card import (.png/.json), basic creator link,
 * tap → start/resume chat. The full Character Studio arrives later.
 */
export default function Characters() {
  const navigate = useNavigate();
  const fileRef = useRef<HTMLInputElement>(null);
  const [query, setQuery] = useState('');
  const [error, setError] = useState<string>();

  const characters = useLiveQuery(
    () => db.characters.filter((c) => !c.deletedAt).sortBy('name'),
    [],
  );
  const filtered = (characters ?? []).filter((c) =>
    c.name.toLowerCase().includes(query.toLowerCase()),
  );

  async function onImport(files: FileList | null) {
    if (!files?.length) return;
    setError(undefined);
    try {
      for (const file of Array.from(files)) await importCharacterFile(file);
    } catch (e) {
      setError(`Import failed: ${(e as Error).message}`);
    }
  }

  async function openChat(character: Character) {
    // Resume the most recent chat with this character, or start a new one.
    const existing = await db.chats
      .where('characterIds')
      .equals(character.id)
      .and((c) => !c.deletedAt)
      .reverse()
      .sortBy('updatedAt');
    const chat = existing[0] ?? (await createChat(character));
    navigate(`/chat/${chat.id}`);
  }

  function exportJson(character: Character) {
    const blob = new Blob([characterToCardJson(character)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${character.name}.json`;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  return (
    <div className="min-h-full bg-bg text-text font-ui">
      <header className="flex items-center gap-3 px-5 pt-4 pb-3 border-b border-border">
        <Link to="/" aria-label="Back" className="text-muted text-lg">
          ‹
        </Link>
        <h1 className="text-lg font-semibold flex-1">Characters</h1>
        <button
          onClick={() => fileRef.current?.click()}
          className="text-sm px-3 py-1.5 rounded-lg bg-surface border border-border active:bg-surface-2"
        >
          Import
        </button>
        <Link
          to="/characters/new"
          className="text-sm px-3 py-1.5 rounded-lg bg-accent text-bg font-medium"
        >
          ＋ New
        </Link>
        <input
          ref={fileRef}
          type="file"
          accept=".png,.json"
          multiple
          hidden
          onChange={(e) => onImport(e.target.files)}
        />
      </header>

      <main className="px-5 py-4">
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Search characters…"
          className="w-full mb-4 px-3 py-2 rounded-xl bg-surface border border-border placeholder:text-muted outline-none focus:border-accent"
        />
        {error && <p className="mb-3 text-sm text-red-400">{error}</p>}

        {characters && !characters.length && (
          <p className="text-muted py-8 text-center">
            No characters yet — import a SillyTavern card (PNG/JSON) or create one.
          </p>
        )}

        <div className="grid grid-cols-2 gap-3">
          {filtered.map((c) => (
            <div key={c.id} className="rounded-2xl bg-surface border border-border p-3">
              <button onClick={() => openChat(c)} className="w-full text-left">
                <Avatar assetId={c.avatarAssetId} name={c.name} className="w-full h-32 mb-2" />
                <div className="font-medium truncate">{c.name}</div>
                <div className="text-xs text-muted truncate">
                  {c.creator ? `by ${c.creator}` : c.tags.join(' · ') || '—'}
                </div>
              </button>
              <div className="flex gap-3 mt-2 text-xs text-muted">
                <Link to={`/characters/${c.id}`} className="active:text-text">
                  Edit
                </Link>
                <button onClick={() => exportJson(c)} className="active:text-text">
                  Export
                </button>
                <button
                  onClick={async () => {
                    if (confirm(`Delete ${c.name}?`))
                      await db.characters.update(c.id, { deletedAt: Date.now() });
                  }}
                  className="active:text-text"
                >
                  Delete
                </button>
              </div>
            </div>
          ))}
        </div>
      </main>
    </div>
  );
}
