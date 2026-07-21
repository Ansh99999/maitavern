import { useEffect, useState } from 'react';
import { db } from '@/db/db';
import Icon from './Icon';
import type { LogEntry } from '@/types';

/*
 * Request log sheet (docs/01 "view request log"): the assembled blocks in
 * order, the final payload, usage and latency. Reads the LogEntry the
 * generation persisted — what you see is literally what was sent.
 */
export default function LogSheet({ logId, onClose }: { logId?: string; onClose: () => void }) {
  const [log, setLog] = useState<LogEntry | null>(null);
  const [showPayload, setShowPayload] = useState(false);

  useEffect(() => {
    if (logId) db.logs.get(logId).then((l) => setLog(l ?? null));
    else setLog(null);
  }, [logId]);

  if (!logId) return null;
  return (
    <div className="fixed inset-0 z-40 bg-black/60" onClick={onClose}>
      <div
        className="absolute bottom-0 inset-x-0 max-h-[85%] overflow-y-auto rounded-t-2xl bg-surface border-t border-border p-4"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center mb-3">
          <h2 className="font-semibold flex-1">Request log</h2>
          {log?.usage && (
            <span className="text-xs text-muted mr-3">
              {log.usage.in} in · {log.usage.out} out · {log.ms}ms
            </span>
          )}
          <button onClick={onClose} aria-label="Close" className="text-muted px-2">
            <Icon name="x" size={20} />
          </button>
        </div>
        {!log && <p className="text-muted text-sm">No log recorded for this message.</p>}
        {log?.assembledBlocks.map((b, i) => (
          <details key={i} className="mb-2 rounded-xl bg-surface-2 border border-border" open={i === 0}>
            <summary className="px-3 py-2 text-sm cursor-pointer select-none">
              <span className="font-medium">{b.title}</span>
              <span className="text-muted"> · {b.role} · ~{b.tokens} tok</span>
            </summary>
            <pre className="px-3 pb-3 text-xs whitespace-pre-wrap text-muted">{b.content || '(empty)'}</pre>
          </details>
        ))}
        {log && (
          <button
            onClick={() => setShowPayload(!showPayload)}
            className="text-xs text-accent mt-1"
          >
            {showPayload ? 'Hide' : 'Show'} raw request payload
          </button>
        )}
        {showPayload && log && (
          <pre className="mt-2 p-3 rounded-xl bg-bg text-xs whitespace-pre-wrap text-muted overflow-x-auto">
            {JSON.stringify(log.requestPayload, null, 2)}
          </pre>
        )}
      </div>
    </div>
  );
}
