import { useState } from 'react';
import { Link } from 'react-router-dom';
import { useLiveQuery } from 'dexie-react-hooks';
import { db } from '@/db/db';
import LogSheet from '@/components/LogSheet';

/*
 * Global request-log viewer (docs/01 "global logs"): most recent generations
 * across all chats; tap one to inspect its assembled blocks + payload.
 */
export default function Logs() {
  const logs = useLiveQuery(() => db.logs.orderBy('at').reverse().limit(100).toArray(), []);
  const chats = useLiveQuery(() => db.chats.toArray(), []);
  const [openId, setOpenId] = useState<string>();

  const chatTitle = (chatId: string) => chats?.find((c) => c.id === chatId)?.title ?? '(deleted chat)';

  return (
    <div className="min-h-full bg-bg text-text font-ui">
      <header className="flex items-center gap-3 px-5 pt-4 pb-3 border-b border-border">
        <Link to="/settings" aria-label="Back" className="text-muted text-lg">‹</Link>
        <h1 className="text-lg font-semibold flex-1">Request logs</h1>
        <button
          onClick={async () => {
            if (confirm('Clear all logs?')) await db.logs.clear();
          }}
          className="text-sm text-muted"
        >
          Clear
        </button>
      </header>

      <main className="px-5 py-3">
        {logs && !logs.length && <p className="text-muted py-8 text-center">No requests logged yet.</p>}
        {(logs ?? []).map((log) => (
          <button
            key={log.id}
            onClick={() => setOpenId(log.id)}
            className="w-full text-left mb-2 rounded-xl bg-surface border border-border px-3 py-2"
          >
            <div className="flex justify-between text-sm">
              <span className="font-medium truncate">{chatTitle(log.chatId)}</span>
              <span className="text-muted shrink-0 ml-2">
                {new Date(log.at).toLocaleString([], { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' })}
              </span>
            </div>
            <div className="text-xs text-muted">
              {log.assembledBlocks.length} blocks
              {log.usage ? ` · ${log.usage.in}→${log.usage.out} tok` : ''}
              {log.ms ? ` · ${log.ms}ms` : ''}
              {log.responseRaw && (log.responseRaw as { error?: string }).error ? ' · ⚠ error' : ''}
            </div>
          </button>
        ))}
      </main>

      <LogSheet logId={openId} onClose={() => setOpenId(undefined)} />
    </div>
  );
}
