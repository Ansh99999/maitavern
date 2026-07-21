import { expandMacros, type MacroContext } from '@/lib/macros';
import { estimateTokens, SAFETY_MARGIN } from '@/lib/tokens';
import { activeContent } from '@/types';
import type { Character, Message, Preset, Settings } from '@/types';

/*
 * Prompt assembler (docs/10). One deterministic pipeline:
 * ordered blocks → markers resolved → macros expanded → budget/truncate.
 * Every stage's output is captured in `assembledBlocks` for the Log viewer.
 */

export interface RoleMessage {
  role: 'system' | 'user' | 'assistant';
  content: string;
}

export interface AssembledPrompt {
  /** role-tagged list ready for a provider adapter */
  messages: RoleMessage[];
  /** what the Log viewer shows, in assembly order */
  assembledBlocks: { title: string; role: string; content: string; tokens: number }[];
  /** history messages dropped to fit the window */
  droppedCount: number;
}

export interface AssembleInput {
  character: Character;
  settings: Settings;
  preset: Preset;
  history: Message[]; // active branch, oldest→newest, hiddenFromContext already respected by caller or here
  scenarioOverride?: string;
  /** the user text being sent this turn (already appended to history by caller) */
  input?: string;
}

export function assemblePrompt(inp: AssembleInput): AssembledPrompt {
  const { character, settings, preset } = inp;
  const scenario = inp.scenarioOverride ?? character.scenario;

  const visible = inp.history.filter((m) => !m.hiddenFromContext && !m.deletedAt);
  const lastUser = [...visible].reverse().find((m) => m.role === 'user');
  const lastChar = [...visible].reverse().find((m) => m.role === 'assistant');

  const macroCtx: MacroContext = {
    char: character.name,
    user: settings.userName,
    description: character.description,
    personality: character.personality,
    scenario,
    mesExamples: character.mesExample,
    lastMessage: visible.length ? activeContent(visible[visible.length - 1]) : '',
    lastUserMessage: lastUser ? activeContent(lastUser) : '',
    lastCharMessage: lastChar ? activeContent(lastChar) : '',
    input: inp.input ?? '',
  };
  const mx = (s: string) => expandMacros(s, macroCtx).trim();

  // Resolve marker blocks → concrete content (empty markers vanish).
  const resolveMarker = (marker: string): string => {
    switch (marker) {
      case 'charDescription':
        return mx(character.description);
      case 'personality':
        return character.personality.trim() ? `${character.name}'s personality: ${mx(character.personality)}` : '';
      case 'scenario':
        return scenario.trim() ? `Scenario: ${mx(scenario)}` : '';
      case 'persona':
        return ''; // persona descriptions arrive in Phase 2; {{user}} name already flows via macros
      case 'dialogueExamples':
        return mx(character.mesExample);
      default:
        return '';
    }
  };

  const preHistory: RoleMessage[] = [];
  const postHistory: RoleMessage[] = [];
  const blocksLog: AssembledPrompt['assembledBlocks'] = [];
  let seenHistory = false;
  let historySlot = -1;

  for (const block of preset.promptBlocks) {
    if (!block.enabled) continue;
    if (block.kind === 'marker' && block.marker === 'chatHistory') {
      seenHistory = true;
      historySlot = blocksLog.length;
      blocksLog.push({ title: 'Chat history', role: 'chat', content: '', tokens: 0 });
      continue;
    }
    const content = block.kind === 'custom' ? mx(block.content ?? '') : resolveMarker(block.marker ?? '');
    if (!content) continue;
    (seenHistory ? postHistory : preHistory).push({ role: block.role, content });
    blocksLog.push({ title: block.title, role: block.role, content, tokens: estimateTokens(content) });
  }
  // Character-card system prompt / jailbreak wrap the preset blocks (ST semantics).
  if (character.systemPrompt?.trim()) {
    const content = mx(character.systemPrompt);
    preHistory.unshift({ role: 'system', content });
    blocksLog.unshift({ title: 'Character system prompt', role: 'system', content, tokens: estimateTokens(content) });
    if (historySlot >= 0) historySlot++;
  }
  if (character.postHistoryInstructions?.trim()) {
    const content = mx(character.postHistoryInstructions);
    postHistory.push({ role: 'system', content });
    blocksLog.push({ title: 'Post-history instructions', role: 'system', content, tokens: estimateTokens(content) });
  }

  // Budget: contextSize − maxResponseTokens − safety margin; drop-oldest history.
  const budget = preset.contextSize - preset.maxResponseTokens - SAFETY_MARGIN;
  const fixedSpend = [...preHistory, ...postHistory].reduce(
    (sum, b) => sum + estimateTokens(b.content),
    0,
  );

  const historyMsgs: RoleMessage[] = visible.map((m) => ({
    role: m.role === 'system' ? 'user' : m.role,
    content: mx(activeContent(m)),
  }));
  let dropped = 0;
  let historySpend = historyMsgs.reduce((s, m) => s + estimateTokens(m.content), 0);
  while (dropped < historyMsgs.length - 1 && fixedSpend + historySpend > budget) {
    historySpend -= estimateTokens(historyMsgs[dropped].content);
    dropped++;
  }
  const keptHistory = historyMsgs.slice(dropped);

  if (historySlot >= 0) {
    const joined = keptHistory.map((m) => `${m.role}: ${m.content}`).join('\n');
    blocksLog[historySlot] = {
      title: dropped ? `Chat history (${dropped} oldest dropped)` : 'Chat history',
      role: 'chat',
      content: joined,
      tokens: estimateTokens(joined),
    };
  }

  return {
    messages: [...preHistory, ...keptHistory, ...postHistory],
    assembledBlocks: blocksLog,
    droppedCount: dropped,
  };
}
