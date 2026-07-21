import ReactMarkdown from 'react-markdown';
import Avatar from './Avatar';
import { activeContent } from '@/types';
import type { ChatStyle, Message } from '@/types';

/*
 * One message, rendered per display style (docs/01):
 * bubble · document · discord · novel. Avatars follow the universal model —
 * optional in every style, never cropped, aspect preserved.
 */

export interface MessageRowProps {
  message: Message;
  /** overrides stored content while this row is streaming */
  liveText?: string;
  style: ChatStyle;
  showAvatar: boolean;
  authorName: string;
  avatarAssetId?: string;
  isUser: boolean;
  onAction: (action: MessageAction, message: Message) => void;
  streaming: boolean;
}

export type MessageAction = 'regenerate' | 'edit' | 'delete' | 'copy' | 'log';

export default function MessageRow(props: MessageRowProps) {
  const text = props.liveText ?? activeContent(props.message);
  switch (props.style) {
    case 'document':
      return <DocumentRow {...props} text={text} />;
    case 'discord':
      return <DiscordRow {...props} text={text} />;
    case 'novel':
      return <NovelRow {...props} text={text} />;
    default:
      return <BubbleRow {...props} text={text} />;
  }
}

type StyledProps = MessageRowProps & { text: string };

function Actions({ message, isUser, onAction, streaming }: MessageRowProps) {
  if (streaming) return null;
  const btn = 'px-1 text-muted active:text-text';
  return (
    <span className="text-xs opacity-70 shrink-0">
      {!isUser && (
        <button className={btn} title="Regenerate" onClick={() => onAction('regenerate', message)}>
          ⟳
        </button>
      )}
      <button className={btn} title="Edit" onClick={() => onAction('edit', message)}>
        ✎
      </button>
      <button className={btn} title="Copy" onClick={() => onAction('copy', message)}>
        ⧉
      </button>
      {!isUser && (
        <button className={btn} title="Request log" onClick={() => onAction('log', message)}>
          ☰
        </button>
      )}
      <button className={btn} title="Delete" onClick={() => onAction('delete', message)}>
        🗑
      </button>
    </span>
  );
}

function Markdown({ text, streaming }: { text: string; streaming: boolean }) {
  return (
    <div className="prose-chat font-chat break-words">
      <ReactMarkdown>{text}</ReactMarkdown>
      {streaming && <span className="inline-block w-2 h-4 bg-accent animate-pulse align-text-bottom" />}
    </div>
  );
}

function BubbleRow(p: StyledProps) {
  return (
    <div className={`flex gap-2 px-3 py-1.5 ${p.isUser ? 'flex-row-reverse' : ''}`}>
      {p.showAvatar && <Avatar assetId={p.avatarAssetId} name={p.authorName} className="w-9 h-9" />}
      <div className={`max-w-[80%] ${p.isUser ? 'items-end' : ''} flex flex-col`}>
        <div className={`text-xs text-muted mb-0.5 flex gap-2 ${p.isUser ? 'flex-row-reverse' : ''}`}>
          <span>{p.authorName}</span>
          <Actions {...p} />
        </div>
        <div
          className={`rounded-2xl px-3 py-2 ${
            p.isUser ? 'bg-bubble-user rounded-tr-sm' : 'bg-bubble-bot rounded-tl-sm'
          }`}
        >
          <Markdown text={p.text} streaming={p.streaming} />
        </div>
      </div>
    </div>
  );
}

function DocumentRow(p: StyledProps) {
  return (
    <div className="px-4 py-2">
      <div className="flex items-center gap-2 mb-1">
        {p.showAvatar && <Avatar assetId={p.avatarAssetId} name={p.authorName} className="w-6 h-6" />}
        <span className="font-semibold text-sm">{p.authorName}</span>
        <Actions {...p} />
      </div>
      <Markdown text={p.text} streaming={p.streaming} />
    </div>
  );
}

function DiscordRow(p: StyledProps) {
  return (
    <div className="flex gap-3 px-3 py-1.5 hover:bg-surface/50">
      {p.showAvatar ? (
        <Avatar assetId={p.avatarAssetId} name={p.authorName} className="w-10 h-10" />
      ) : (
        <div className="w-10 shrink-0" />
      )}
      <div className="min-w-0 flex-1">
        <div className="flex items-baseline gap-2">
          <span className={`font-medium text-sm ${p.isUser ? 'text-accent' : 'text-text'}`}>
            {p.authorName}
          </span>
          <span className="text-[10px] text-muted">
            {new Date(p.message.createdAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
          </span>
          <Actions {...p} />
        </div>
        <Markdown text={p.text} streaming={p.streaming} />
      </div>
    </div>
  );
}

function NovelRow(p: StyledProps) {
  return (
    <div className="px-5 py-2 leading-relaxed">
      <div className="flex items-center gap-2">
        {p.showAvatar && <Avatar assetId={p.avatarAssetId} name={p.authorName} className="w-7 h-7 float-left mr-2" />}
        <Actions {...p} />
      </div>
      <Markdown text={p.text} streaming={p.streaming} />
    </div>
  );
}
