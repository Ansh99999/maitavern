import { useEffect, useRef, useState } from 'react';
import { useParams } from 'react-router-dom';
import { useLiveQuery } from 'dexie-react-hooks';
import { db } from '@/db/db';
import { branchMessages, getSettings, makeMessage, touchChat } from '@/db/repo';
import { generateReply, GenerateConfigError } from '@/engine/generate';
import { activeContent } from '@/types';
import type { Message, Settings } from '@/types';
import MessageRow, { type MessageAction } from '@/components/MessageRow';
import QuickSettings from '@/components/QuickSettings';
import LogSheet from '@/components/LogSheet';
import Icon from '@/components/Icon';

/*
 * Chat interface (docs/01, revised): NO top bar — the message area runs
 * full-bleed and a single floating hamburger (top-left) opens the chat menu
 * (character info, navigation, quick settings). Send morphs into Stop while
 * streaming. Swipes/branching land in Phase 2.
 */
export default function Chat() {
  const { chatId } = useParams();
  const [settings, setSettings] = useState<Settings>();
  const [input, setInput] = useState('');
  const [streamingId, setStreamingId] = useState<string>();
  const [liveText, setLiveText] = useState('');
  const [error, setError] = useState<string>();
  const [showSidebar, setShowSidebar] = useState(false);
  const [logId, setLogId] = useState<string>();
  const [editing, setEditing] = useState<{ id: string; text: string }>();
  const abortRef = useRef<AbortController>();
  const bottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    getSettings().then(setSettings);
  }, []);

  const chat = useLiveQuery(() => (chatId ? db.chats.get(chatId) : undefined), [chatId]);
  const characterId = chat?.characterIds[0];
  const character = useLiveQuery(
    () => (characterId ? db.characters.get(characterId) : undefined),
    [characterId],
  );
  const messages = useLiveQuery(
    async () => (chat ? await branchMessages(chat) : []),
    [chat?.id, chat?.activeBranchId],
  );

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages?.length, liveText]);

  if (!chat || !character || !settings) {
    return <div className="min-h-full bg-bg text-text grid place-items-center text-muted">Loading…</div>;
  }

  const style = chat.chatStyle ?? settings.chatStyle;
  const showAvatar = settings.showAvatars[style] ?? style !== 'novel';
  const busy = streamingId !== undefined;

  async function runGeneration(target: Message, history: Message[]) {
    if (!chat || !character) return;
    setError(undefined);
    setStreamingId(target.id);
    setLiveText('');
    const controller = new AbortController();
    abortRef.current = controller;
    try {
      await generateReply({
        chat,
        character,
        target,
        history,
        onDelta: setLiveText,
        signal: controller.signal,
      });
    } catch (err) {
      if ((err as Error).name !== 'AbortError') {
        setError(
          err instanceof GenerateConfigError
            ? (err as Error).message
            : `Generation failed: ${(err as Error).message}`,
        );
        // Remove the empty placeholder so the chat doesn't collect husks.
        const row = await db.messages.get(target.id);
        if (row && !activeContent(row).trim()) await db.messages.delete(target.id);
      }
    } finally {
      setStreamingId(undefined);
      setLiveText('');
      abortRef.current = undefined;
    }
  }

  async function onSend() {
    if (!chat || !character || busy) return;
    const text = input.trim();
    if (!text) return;
    setInput('');
    const userMsg = makeMessage(chat, 'user', text);
    await db.messages.add(userMsg);
    await touchChat(chat.id);
    const history = await branchMessages(chat);
    const target = makeMessage(chat, 'assistant', '', { characterId: character.id });
    await db.messages.add(target);
    await runGeneration(target, history);
  }

  function onStop() {
    abortRef.current?.abort();
  }

  async function onAction(action: MessageAction, m: Message) {
    if (busy) return;
    switch (action) {
      case 'copy':
        await navigator.clipboard.writeText(activeContent(m));
        return;
      case 'delete':
        await db.messages.delete(m.id);
        return;
      case 'edit':
        setEditing({ id: m.id, text: activeContent(m) });
        return;
      case 'log':
        setLogId(m.swipes[m.activeSwipe]?.meta?.logId);
        return;
      case 'regenerate': {
        if (!chat) return;
        const history = (await branchMessages(chat)).filter((row) => row.createdAt < m.createdAt);
        await runGeneration(m, history);
        return;
      }
    }
  }

  async function saveEdit() {
    if (!editing) return;
    const row = await db.messages.get(editing.id);
    if (row) {
      const swipes = [...row.swipes];
      const prev = swipes[row.activeSwipe]?.content ?? '';
      swipes[row.activeSwipe] = { ...swipes[row.activeSwipe], content: editing.text, at: Date.now() };
      await db.messages.update(row.id, {
        swipes,
        editHistory: [...row.editHistory, { at: Date.now(), content: prev }],
        updatedAt: Date.now(),
      });
    }
    setEditing(undefined);
  }

  return (
    <div className="h-dvh flex flex-col bg-bg text-text font-ui relative">
      <button
        aria-label="Menu"
        onClick={() => setShowSidebar(true)}
        className="absolute top-2 left-2 z-30 p-2 rounded-full bg-surface/80 backdrop-blur border border-border text-muted active:text-text"
      >
        <Icon name="menu" size={20} />
      </button>

      <main className="flex-1 overflow-y-auto pt-12 pb-2">
        {(messages ?? []).map((m) => (
          <MessageRow
            key={m.id}
            message={m}
            liveText={m.id === streamingId ? liveText : undefined}
            style={style}
            showAvatar={showAvatar}
            authorName={m.role === 'user' ? settings.userName : character.name}
            avatarAssetId={m.role === 'user' ? undefined : character.avatarAssetId}
            isUser={m.role === 'user'}
            onAction={onAction}
            streaming={m.id === streamingId}
          />
        ))}
        {error && (
          <div className="mx-4 my-2 px-3 py-2 rounded-xl border border-red-500/40 bg-red-500/10 text-sm text-red-300">
            {error}
          </div>
        )}
        <div ref={bottomRef} />
      </main>

      <footer className="shrink-0 border-t border-border px-3 py-2 flex items-end gap-2">
        <textarea
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault();
              onSend();
            }
          }}
          placeholder="Type a message…"
          rows={Math.min(4, Math.max(1, input.split('\n').length))}
          className="flex-1 px-3 py-2 rounded-2xl bg-surface border border-border placeholder:text-muted outline-none focus:border-accent resize-none font-chat"
        />
        {busy ? (
          <button
            onClick={onStop}
            aria-label="Stop"
            className="w-10 h-10 rounded-full bg-red-500/80 text-bg grid place-items-center shrink-0"
          >
            <Icon name="stop" size={18} />
          </button>
        ) : (
          <button
            onClick={onSend}
            disabled={!input.trim()}
            aria-label="Send"
            className="w-10 h-10 rounded-full bg-accent text-bg grid place-items-center shrink-0 disabled:opacity-40"
          >
            <Icon name="send" size={18} />
          </button>
        )}
      </footer>

      {showSidebar && (
        <QuickSettings
          chat={chat}
          character={character}
          settings={settings}
          onSettingsChange={setSettings}
          onClose={() => setShowSidebar(false)}
        />
      )}
      <LogSheet logId={logId} onClose={() => setLogId(undefined)} />

      {editing && (
        <div
          className="fixed inset-0 z-40 bg-black/60 grid place-items-center p-4"
          onClick={() => setEditing(undefined)}
        >
          <div
            className="w-full max-w-lg rounded-2xl bg-surface border border-border p-4"
            onClick={(e) => e.stopPropagation()}
          >
            <h2 className="font-semibold mb-2">Edit message</h2>
            <textarea
              value={editing.text}
              onChange={(e) => setEditing({ ...editing, text: e.target.value })}
              rows={8}
              className="w-full px-3 py-2 rounded-xl bg-surface-2 border border-border outline-none focus:border-accent resize-y font-chat"
            />
            <div className="flex justify-end gap-3 mt-3 text-sm">
              <button onClick={() => setEditing(undefined)} className="px-3 py-1.5 text-muted">
                Cancel
              </button>
              <button onClick={saveEdit} className="px-4 py-1.5 rounded-lg bg-accent text-bg font-medium">
                Save
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
