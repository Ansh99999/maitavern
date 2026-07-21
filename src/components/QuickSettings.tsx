import { useLiveQuery } from 'dexie-react-hooks';
import { db } from '@/db/db';
import { getSettings, patchSettings } from '@/db/repo';
import type { Chat, ChatStyle, Settings } from '@/types';

/*
 * Quick Settings sidebar (docs/01 top-bar ☰): swap preset/provider for THIS
 * chat (per-chat override) or globally, pick the display style, toggle
 * avatars. Scenario override + author's note editors arrive in Phase 2.
 */
export default function QuickSettings({
  chat,
  settings,
  onSettingsChange,
  onClose,
}: {
  chat: Chat;
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

  return (
    <div className="fixed inset-0 z-40 bg-black/60" onClick={onClose}>
      <aside
        className="absolute right-0 inset-y-0 w-72 bg-surface border-l border-border p-4 overflow-y-auto"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center mb-4">
          <h2 className="font-semibold flex-1">Quick settings</h2>
          <button onClick={onClose} className="text-muted px-2">✕</button>
        </div>

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
