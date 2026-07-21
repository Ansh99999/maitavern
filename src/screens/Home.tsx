import { Link, useNavigate } from 'react-router-dom';
import { useLiveQuery } from 'dexie-react-hooks';
import { timeGreeting } from '../lib/greeting';
import { db } from '@/db/db';
import { activeContent } from '@/types';
import Avatar from '@/components/Avatar';

/*
 * The "lounge" homepage (docs/02): logo + clock, Characters shelf, live Chats
 * previews, Provider/Preset status handled in Settings. The auto-rotating
 * carousel joins in a later polish pass.
 */
export default function Home() {
  const g = timeGreeting();
  const navigate = useNavigate();

  const characters = useLiveQuery(
    () => db.characters.filter((c) => !c.deletedAt).reverse().sortBy('updatedAt'),
    [],
  );
  const chats = useLiveQuery(
    async () => {
      const rows = await db.chats.filter((c) => !c.deletedAt).reverse().sortBy('updatedAt');
      return Promise.all(
        rows.slice(0, 5).map(async (chat) => {
          const msgs = await db.messages
            .where('chatId')
            .equals(chat.id)
            .and((m) => m.branchId === chat.activeBranchId && !m.deletedAt)
            .sortBy('createdAt');
          const last = msgs[msgs.length - 1];
          const character = await db.characters.get(chat.characterIds[0]);
          return { chat, last, character };
        }),
      );
    },
    [],
  );

  const mostRecentChat = chats?.[0]?.chat;

  return (
    <div className="min-h-full bg-bg text-text font-ui">
      <header className="flex items-center justify-between px-5 pt-4 pb-2">
        <div className="text-xl font-semibold tracking-tight">
          Mai<span className="text-accent">Tavern</span>
        </div>
        {/* clock → resume the most recent chat (docs/02) */}
        <button
          aria-label="Most recent chat"
          className="text-muted text-xl disabled:opacity-40"
          disabled={!mostRecentChat}
          onClick={() => mostRecentChat && navigate(`/chat/${mostRecentChat.id}`)}
        >
          🕐
        </button>
      </header>

      <main className="px-5 pb-10">
        <h1 className="text-2xl font-semibold mt-2">
          {g.text} {g.emoji}
        </h1>
        <p className="text-muted mb-6">Ready to dive back in?</p>

        <div className="flex items-baseline justify-between mb-2">
          <h2 className="font-medium">Characters</h2>
          <Link to="/characters" className="text-sm text-muted">see all ›</Link>
        </div>
        <div className="flex gap-3 overflow-x-auto pb-2 mb-5">
          {(characters ?? []).slice(0, 8).map((c) => (
            <Link key={c.id} to="/characters" className="shrink-0 w-16 text-center">
              <Avatar assetId={c.avatarAssetId} name={c.name} className="w-16 h-16" />
              <div className="text-xs text-muted truncate mt-1">{c.name}</div>
            </Link>
          ))}
          <Link
            to="/characters/new"
            className="shrink-0 w-16 h-16 rounded-xl bg-surface border border-border grid place-items-center text-muted text-2xl"
            aria-label="New character"
          >
            ＋
          </Link>
        </div>

        <div className="flex items-baseline justify-between mb-2">
          <h2 className="font-medium">Chats</h2>
        </div>
        <div className="flex flex-col gap-2 mb-6">
          {chats && !chats.length && (
            <p className="text-sm text-muted">
              No chats yet — pick a character to start one.
            </p>
          )}
          {(chats ?? []).map(({ chat, last, character }) => (
            <Link
              key={chat.id}
              to={`/chat/${chat.id}`}
              className="flex items-center gap-3 rounded-2xl bg-surface border border-border px-3 py-2.5 active:bg-surface-2"
            >
              <Avatar assetId={character?.avatarAssetId} name={chat.title} className="w-10 h-10" />
              <div className="min-w-0 flex-1">
                <div className="font-medium truncate">{chat.title}</div>
                <div className="text-sm text-muted truncate">
                  {last ? activeContent(last) : 'New chat'}
                </div>
              </div>
              <div className="text-xs text-muted shrink-0">{timeAgo(chat.updatedAt)}</div>
            </Link>
          ))}
        </div>

        <div className="grid grid-cols-2 gap-3">
          <Tile to="/library" title="Library" subtitle="lorebooks · snippets · gallery" />
          <Tile to="/settings" title="Settings" subtitle="theme · provider · preset" />
        </div>
      </main>
    </div>
  );
}

function timeAgo(ts: number): string {
  const s = Math.max(1, Math.floor((Date.now() - ts) / 1000));
  if (s < 60) return `${s}s`;
  if (s < 3600) return `${Math.floor(s / 60)}m`;
  if (s < 86400) return `${Math.floor(s / 3600)}h`;
  return `${Math.floor(s / 86400)}d`;
}

function Tile({ to, title, subtitle }: { to: string; title: string; subtitle: string }) {
  return (
    <Link
      to={to}
      className="rounded-2xl bg-surface border border-border p-4 active:bg-surface-2 transition-colors"
    >
      <div className="font-medium">{title}</div>
      <div className="text-sm text-muted">{subtitle}</div>
    </Link>
  );
}
