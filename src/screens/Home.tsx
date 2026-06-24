import { Link } from 'react-router-dom';
import { timeGreeting } from '../lib/greeting';

// The "lounge" homepage (docs/02). Phase 0 shell: logo, greeting, and nav tiles.
// Carousel, Characters shelf, and live Chats previews land in Phase 1.
export default function Home() {
  const g = timeGreeting();
  return (
    <div className="min-h-full bg-bg text-text font-ui">
      <header className="flex items-center justify-between px-5 pt-4 pb-2">
        <div className="text-xl font-semibold tracking-tight">
          Mai<span className="text-accent">Tavern</span>
        </div>
        {/* clock → most recent chat (wired in Phase 1) */}
        <button aria-label="Most recent chat" className="text-muted text-xl">
          🕐
        </button>
      </header>

      <main className="px-5 pb-10">
        <div className="mt-2 mb-6 rounded-2xl bg-surface border border-border h-40 grid place-items-center text-muted">
          carousel — recent chats (Phase 1)
        </div>

        <h1 className="text-2xl font-semibold">
          {g.text} {g.emoji}
        </h1>
        <p className="text-muted mb-6">Ready to dive back in?</p>

        <div className="grid grid-cols-2 gap-3">
          <Tile to="/library" title="Library" subtitle="lorebooks · snippets · gallery" />
          <Tile to="/settings" title="Settings" subtitle="theme · providers · memory" />
          <Tile to="/chat/demo" title="Characters" subtitle="your cast (Phase 1)" />
          <Tile to="/chat/demo" title="Chats" subtitle="recent conversations" />
        </div>
      </main>
    </div>
  );
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
