import { useRef, useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import { useLiveQuery } from 'dexie-react-hooks';
import { db } from '@/db/db';
import { createChat } from '@/db/repo';
import { importCharacterFile, characterToCardJson } from '@/lib/cardImport';
import { uuid, now } from '@/lib/id';
import Avatar from '@/components/Avatar';
import Icon from '@/components/Icon';
import type { Character } from '@/types';

/*
 * Character gallery (docs/02). Tapping a character opens the QUICK MENU
 * (docs/04 "character quick menu"): chats with this character, new chat,
 * edit, duplicate, export, delete. No action fires on the bare tap.
 */
export default function Characters() {
  const fileRef = useRef<HTMLInputElement>(null);
  const [query, setQuery] = useState('');
  const [error, setError] = useState<string>();
  const [menuFor, setMenuFor] = useState<Character>();

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

  return (
    <div className="min-h-full bg-bg text-text font-ui">
      <header className="flex items-center gap-3 px-5 pt-4 pb-3 border-b border-border">
        <Link to="/" aria-label="Back" className="text-muted">
          <Icon name="chevronLeft" size={22} />
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
          className="text-sm px-3 py-1.5 rounded-lg bg-accent text-bg font-medium flex items-center gap-1"
        >
          <Icon name="plus" size={16} /> New
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
            <button
              key={c.id}
              onClick={() => setMenuFor(c)}
              className="rounded-2xl bg-surface border border-border p-3 text-left active:bg-surface-2"
            >
              <Avatar assetId={c.avatarAssetId} name={c.name} className="w-full h-32 mb-2" />
              <div className="font-medium truncate">{c.name}</div>
              <div className="text-xs text-muted truncate">
                {c.creator ? `by ${c.creator}` : c.tags.join(' · ') || '—'}
              </div>
            </button>
          ))}
        </div>
      </main>

      {menuFor && <CharacterQuickMenu character={menuFor} onClose={() => setMenuFor(undefined)} />}
    </div>
  );
}

/** Bottom-sheet quick menu: chat list + new chat + edit/duplicate/export/delete. */
function CharacterQuickMenu({ character, onClose }: { character: Character; onClose: () => void }) {
  const navigate = useNavigate();
  const chats = useLiveQuery(
    () =>
      db.chats
        .where('characterIds')
        .equals(character.id)
        .and((c) => !c.deletedAt)
        .reverse()
        .sortBy('updatedAt'),
    [character.id],
  );

  async function onNewChat() {
    const chat = await createChat(character);
    navigate(`/chat/${chat.id}`);
  }

  async function onDuplicate() {
    const t = now();
    await db.characters.add({
      ...character,
      id: uuid(),
      name: `${character.name} (copy)`,
      createdAt: t,
      updatedAt: t,
    });
    onClose();
  }

  function onExport() {
    const blob = new Blob([characterToCardJson(character)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${character.name}.json`;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  async function onDelete() {
    if (!confirm(`Delete ${character.name}? Their chats are kept.`)) return;
    await db.characters.update(character.id, { deletedAt: now() });
    onClose();
  }

  const row =
    'w-full flex items-center gap-3 px-4 py-3 text-left active:bg-surface-2 rounded-xl';

  return (
    <div className="fixed inset-0 z-40 bg-black/60" onClick={onClose}>
      <div
        className="absolute bottom-0 inset-x-0 max-h-[85%] overflow-y-auto rounded-t-2xl bg-surface border-t border-border p-3 pb-6"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center gap-3 px-2 py-2 mb-1">
          <Avatar assetId={character.avatarAssetId} name={character.name} className="w-12 h-12" />
          <div className="min-w-0 flex-1">
            <div className="font-semibold truncate">{character.name}</div>
            <div className="text-xs text-muted truncate">
              {chats?.length ? `${chats.length} chat${chats.length > 1 ? 's' : ''}` : 'No chats yet'}
            </div>
          </div>
          <button onClick={onClose} aria-label="Close" className="text-muted p-2">
            <Icon name="x" size={20} />
          </button>
        </div>

        {(chats ?? []).length > 0 && (
          <div className="mb-2 rounded-xl bg-surface-2/60 border border-border overflow-hidden">
            {(chats ?? []).slice(0, 5).map((chat) => (
              <button
                key={chat.id}
                onClick={() => navigate(`/chat/${chat.id}`)}
                className="w-full flex items-center gap-3 px-4 py-2.5 text-left active:bg-surface-2 border-b border-border last:border-b-0"
              >
                <Icon name="chat" size={16} className="text-muted" />
                <span className="flex-1 truncate text-sm">{chat.title}</span>
                <span className="text-xs text-muted shrink-0">
                  {new Date(chat.updatedAt).toLocaleDateString([], { month: 'short', day: 'numeric' })}
                </span>
              </button>
            ))}
          </div>
        )}

        <button className={row} onClick={onNewChat}>
          <Icon name="plus" className="text-accent" /> New chat
        </button>
        <button className={row} onClick={() => navigate(`/characters/${character.id}`)}>
          <Icon name="pencil" className="text-muted" /> Edit character
        </button>
        <button className={row} onClick={onDuplicate}>
          <Icon name="duplicate" className="text-muted" /> Duplicate
        </button>
        <button className={row} onClick={onExport}>
          <Icon name="download" className="text-muted" /> Export JSON
        </button>
        <button className={`${row} text-red-400`} onClick={onDelete}>
          <Icon name="trash" /> Delete
        </button>
      </div>
    </div>
  );
}
