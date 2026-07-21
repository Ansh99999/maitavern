import { describe, expect, it } from 'vitest';
import { assemblePrompt } from './assemble';
import { adaptRequest } from './adapters';
import { DEFAULT_SETTINGS, makeCharacter, makeDefaultPreset, makeConnection } from '@/db/repo';
import type { Chat, Message } from '@/types';

const character = makeCharacter({
  name: 'Aria',
  description: 'An android bard named {{char}}.',
  personality: 'warm',
  scenario: 'A rooftop tavern.',
  mesExample: '{{user}}: hi\n{{char}}: hey',
  firstMes: 'You made it.',
  systemPrompt: 'Stay poetic, {{char}}.',
  postHistoryInstructions: 'Keep replies short.',
});

const chat: Chat = {
  id: 'c1',
  title: 'Aria',
  characterIds: [character.id],
  settingOverrides: {},
  rootBranchId: 'b1',
  activeBranchId: 'b1',
  createdAt: 0,
  updatedAt: 0,
};

function msg(role: Message['role'], content: string, at = 0): Message {
  return {
    id: `m${at}`,
    chatId: chat.id,
    branchId: 'b1',
    role,
    swipes: [{ content, at }],
    activeSwipe: 0,
    editHistory: [],
    hiddenFromContext: false,
    createdAt: at,
    updatedAt: at,
  };
}

const settings = { ...DEFAULT_SETTINGS, userName: 'Dev' };

describe('assemblePrompt', () => {
  it('orders blocks: system prompt, markers, history, post-history', () => {
    const preset = makeDefaultPreset();
    const out = assemblePrompt({
      character,
      settings,
      preset,
      history: [msg('assistant', 'You made it.', 1), msg('user', 'hi there', 2)],
    });
    expect(out.assembledBlocks[0].title).toBe('Character system prompt');
    expect(out.assembledBlocks[0].content).toBe('Stay poetic, Aria.');
    expect(out.assembledBlocks.at(-1)?.title).toBe('Post-history instructions');
    const roles = out.messages.map((m) => m.role);
    expect(roles.slice(0, 2)).toEqual(['system', 'system']); // char system + main prompt
    expect(roles.at(-1)).toBe('system'); // post-history
    expect(out.messages.at(-2)?.content).toBe('hi there');
  });

  it('expands macros with the settings user name', () => {
    const preset = makeDefaultPreset();
    const out = assemblePrompt({ character, settings, preset, history: [] });
    const main = out.messages.find((m) => m.content.includes('fictional roleplay'));
    expect(main?.content).toContain('Aria');
    expect(main?.content).toContain('Dev');
  });

  it('skips hidden messages and drops oldest history over budget', () => {
    const preset = { ...makeDefaultPreset(), contextSize: 1200, maxResponseTokens: 600 };
    const long = 'x'.repeat(700); // ~175 tokens each
    const history = [
      msg('user', long, 1),
      msg('assistant', long, 2),
      { ...msg('user', 'secret', 3), hiddenFromContext: true },
      msg('user', 'latest', 4),
    ];
    const out = assemblePrompt({ character, settings, preset, history });
    expect(out.messages.some((m) => m.content === 'secret')).toBe(false);
    expect(out.droppedCount).toBeGreaterThan(0);
    expect(out.messages.at(-2)?.content).toBe('latest'); // newest survives
  });

  it('prefers the chat scenario override', () => {
    const preset = makeDefaultPreset();
    const out = assemblePrompt({
      character,
      settings,
      preset,
      history: [],
      scenarioOverride: 'A moonlit pier.',
    });
    const scen = out.assembledBlocks.find((b) => b.title === 'Scenario');
    expect(scen?.content).toBe('Scenario: A moonlit pier.');
  });
});

describe('adaptRequest', () => {
  const preset = makeDefaultPreset();
  const messages = [
    { role: 'system' as const, content: 'sys' },
    { role: 'assistant' as const, content: 'You made it.' },
    { role: 'user' as const, content: 'hi' },
  ];

  it('maps anthropic_messages: top-level system, leading user turn, allowed params only', () => {
    const conn = makeConnection({
      method: 'anthropic_messages',
      baseUrl: 'https://api.anthropic.com/v1',
      apiKey: 'k',
      model: 'claude-sonnet-5',
    });
    const req = adaptRequest(conn, preset, messages, true);
    expect(req.url).toBe('https://api.anthropic.com/v1/messages');
    const body = req.body as Record<string, unknown>;
    expect(body.system).toBe('sys');
    const convo = body.messages as { role: string }[];
    expect(convo[0].role).toBe('user'); // injected (begin)
    expect(body.temperature).toBe(0.9);
    expect(body).not.toHaveProperty('frequency_penalty');
    expect(req.headers['x-api-key']).toBe('k');
    expect(req.headers['anthropic-dangerous-direct-browser-access']).toBe('true');
  });

  it('maps openai_chat with bearer auth and coalesced messages', () => {
    const conn = makeConnection({ apiKey: 'k2', model: 'gpt-x' });
    const req = adaptRequest(conn, preset, messages, false);
    expect(req.url).toBe('https://api.openai.com/v1/chat/completions');
    expect(req.headers.authorization).toBe('Bearer k2');
    const body = req.body as Record<string, unknown>;
    expect((body.messages as unknown[]).length).toBe(3);
    expect(body.stream).toBe(false);
    expect(body).not.toHaveProperty('top_k'); // disabled param stays out
  });
});
