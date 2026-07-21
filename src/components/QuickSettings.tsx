import { Link } from 'react-router-dom';
import { useLiveQuery } from 'dexie-react-hooks';
import { db } from '@/db/db';
import { getSettings, patchSettings } from '@/db/repo';
import Avatar from './Avatar';
import Icon from './Icon';
import type { Character, Chat, ChatStyle, Settings } from '@/types';

/*
 * The chat menu (docs/01, revised): opened by the floating hamburger — the
 * chat screen's ONLY chrome. Carries the character header, navigation, and
 * quick settings (display style, avatars, per-chat preset/provider).
 * Scenario override + author's note editors arrive in Phase 2.
 */
export default function QuickSettings({
  chat,
  character,
  settings,
  onSettingsChange,
  onClose,
}: {
  chat: Chat;
  character: Character;
  settings: Settings;
  onSettingsChange: (s: Settings) => void;
  onClose: () => void;
}) {
  const presets = useLiveQuery(() => db.presets.filter((p) => !p.deletedAt).sortBy('name'), []);
  const connections = useLiveQuery(
    () => db.connections.filter((c) => !c.deletedAt).sortBy('name'),
    [],
  );

  const style: ChatStyle = chat.chatStyle ?? settings.chatStyle;
  const showAvatars = settings.showAvatars[style] ?? style !== 'novel';

  async function setChatField(patch: Partial<Chat>) {
    await db.chats.update(chat.id, { ...patch, updatedAt: Date.now() });
  }

  async function setGlobal(patch: Partial<Settings>) {
    await patchSettings(patch);
    onSettingsChange(await getSettings());
  }

  const select =
    'w-full px-3 py-2 rounded-xl bg-surface-2 border border-border outline-none focus:border-accent text-sm';
  const navRow = 'flex items-center gap-3 px-1 py-2 text-sm active:text-accent';

  return (
    <div className="fixed inset-0 z-40 bg-black/60" onClick={onClose}>
      <aside
        className="absolute left-0 inset-y-0 w-72 bg-surface border-r border-border p-4 overflow-y-auto"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center gap-3 mb-4">
          <Avatar assetId={character.avatarAssetId} name={character.name} className="w-10 h-10" />
          <div className="min-w-0 flex-1">
            <div className="font-semibold truncate">{character.name}</div>
            <div className="text-xs text-muted truncate">{chat.title}</div>
          </div>
          <button onClick={onClose} aria-label="Close" className="text-muted p-1">
            <Icon name="x" size={20} />
          </button>
        </div>

        <nav className="mb-4 border-b border-border pb-3">
          <Link to="/" className={navRow}>
            <Icon name="home" size={16} className="text-muted" /> Home
          </Link>
          <Link to="/characters" className={navRow}>
            <Icon name="users" size={16} className="text-muted" /> Characters
          </Link>
          <Link to={`/characters/${character.id}`} className={navRow}>
            <Icon name="pencil" size={16} className="text-muted" /> Edit character
          </Link>
          <Link to="/settings/logs" className={navRow}>
            <Icon name="list" size={16} className="text-muted" /> Request logs
          </Link>
        </nav>

        <label className="block mb-4">
          <div className="text-sm font-medium mb-1">Display style</div>
          <select
            className={select}
            value={style}
            onChange={(e) => setChatField({ chatStyle: e.target.value as ChatStyle })}
          >
            <option value="bubble">Bubble</option>
            <option value="document">Document</option>
            <option value="discord">Discord</option>
            <option value="novel">Novel</option>
          </select>
          <div className="text-xs text-muted mt-1">Stored per chat.</div>
        </label>

        <label className="flex items-center gap-2 mb-4 text-sm">
          <input
            type="checkbox"
            checked={showAvatars}
            onChange={(e) =>
              setGlobal({ showAvatars: { ...settings.showAvatars, [style]: e.target.checked } })
            }
          />
          Show avatars ({style})
        </label>

        <label className="block mb-4">
          <div className="text-sm font-medium mb-1">Preset</div>
          <select
            className={select}
            value={chat.presetId ?? ''}
            onChange={(e) => setChatField({ presetId: e.target.value || undefined })}
          >
            <option value="">Global default</option>
            {(presets ?? []).map((p) => (
              <option key={p.id} value={p.id}>{p.name}</option>
            ))}
          </select>
        </label>

        <label className="block mb-4">
          <div className="text-sm font-medium mb-1">Provider</div>
          <select
            className={select}
            value={chat.connectionId ?? ''}
            onChange={(e) => setChatField({ connectionId: e.target.value || undefined })}
          >
            <option value="">Global default</option>
            {(connections ?? []).map((c) => (
              <option key={c.id} value={c.id}>{c.name}</option>
            ))}
          </select>
          <div className="text-xs text-muted mt-1">
            Manage providers &amp; presets in Settings.
          </div>
        </label>
      </aside>
    </div>
  );
}
