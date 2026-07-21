/*
 * Macro engine (docs/10). SillyTavern-compatible {{...}} expansion, applied
 * AFTER block assembly so positional macros see final context.
 * Non-deterministic macros ({{roll}}, {{random}}, {{pick}}) evaluate per call.
 */

export interface MacroContext {
  char: string;
  user: string;
  description?: string;
  personality?: string;
  scenario?: string;
  mesExamples?: string;
  lastMessage?: string;
  lastUserMessage?: string;
  lastCharMessage?: string;
  input?: string;
  /** chat-scoped {{setvar}}/{{getvar}} store; mutated in place */
  vars?: Record<string, string>;
  /** global {{setglobalvar}}/{{getglobalvar}} store; mutated in place */
  globalVars?: Record<string, string>;
  now?: Date; // injectable for tests
  rng?: () => number; // injectable for tests
}

const MACRO_RE = /\{\{([^{}]+)\}\}/g;

export function expandMacros(text: string, ctx: MacroContext): string {
  const rng = ctx.rng ?? Math.random;
  const d = ctx.now ?? new Date();
  // Repeat until stable so macros produced by macros (one level) resolve too.
  let out = text;
  for (let pass = 0; pass < 3; pass++) {
    const before = out;
    out = out.replace(MACRO_RE, (whole, rawBody: string) => {
      const body = rawBody.trim();
      const lower = body.toLowerCase();

      // {{// comment}} → removed
      if (body.startsWith('//')) return '';

      switch (lower) {
        case 'char':
          return ctx.char;
        case 'user':
        case 'persona':
          return ctx.user;
        case 'description':
          return ctx.description ?? '';
        case 'personality':
          return ctx.personality ?? '';
        case 'scenario':
          return ctx.scenario ?? '';
        case 'mesexamples':
          return ctx.mesExamples ?? '';
        case 'lastmessage':
          return ctx.lastMessage ?? '';
        case 'lastusermessage':
          return ctx.lastUserMessage ?? '';
        case 'lastcharmessage':
          return ctx.lastCharMessage ?? '';
        case 'input':
          return ctx.input ?? '';
        case 'newline':
          return '\n';
        case 'noop':
        case 'trim': // trim is handled as a post-step marker; strip the tag
          return '';
        case 'time':
          return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        case 'date':
          return d.toLocaleDateString();
        case 'weekday':
          return d.toLocaleDateString('en-US', { weekday: 'long' });
        case 'isotime':
          return d.toISOString().slice(11, 19);
      }

      // {{random:a,b,c}} / {{pick:a,b,c}}
      const listMatch = /^(random|pick)\s*:(.*)$/i.exec(body);
      if (listMatch) {
        const items = listMatch[2].split(',').map((s) => s.trim()).filter(Boolean);
        if (!items.length) return '';
        return items[Math.floor(rng() * items.length)];
      }

      // {{roll:dN}} / {{roll:N}}
      const rollMatch = /^roll\s*:\s*d?(\d+)$/i.exec(body);
      if (rollMatch) {
        const sides = Math.max(1, parseInt(rollMatch[1], 10));
        return String(1 + Math.floor(rng() * sides));
      }

      // Variables: {{setvar::k::v}} {{getvar::k}} {{addvar::k::n}} {{incvar::k}}
      const parts = body.split('::').map((s) => s.trim());
      const op = parts[0].toLowerCase();
      const chatVars = ctx.vars ?? (ctx.vars = {});
      const globalVars = ctx.globalVars ?? (ctx.globalVars = {});
      switch (op) {
        case 'setvar':
          if (parts.length >= 3) chatVars[parts[1]] = parts.slice(2).join('::');
          return '';
        case 'getvar':
          return chatVars[parts[1]] ?? '';
        case 'addvar': {
          const cur = parseFloat(chatVars[parts[1]] ?? '0') || 0;
          chatVars[parts[1]] = String(cur + (parseFloat(parts[2] ?? '0') || 0));
          return '';
        }
        case 'incvar': {
          const cur = parseFloat(chatVars[parts[1]] ?? '0') || 0;
          chatVars[parts[1]] = String(cur + 1);
          return '';
        }
        case 'setglobalvar':
          if (parts.length >= 3) globalVars[parts[1]] = parts.slice(2).join('::');
          return '';
        case 'getglobalvar':
          return globalVars[parts[1]] ?? '';
      }

      // Unknown macro: leave untouched so authors can see the typo.
      return whole;
    });
    if (out === before) break;
  }
  return out;
}
