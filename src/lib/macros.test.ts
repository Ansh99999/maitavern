import { describe, expect, it } from 'vitest';
import { expandMacros, type MacroContext } from './macros';

const base: MacroContext = {
  char: 'Aria',
  user: 'Dev',
  description: 'a curious android',
  scenario: 'a rainy rooftop',
  now: new Date('2026-07-20T15:04:05'),
  rng: () => 0.5,
};

describe('expandMacros', () => {
  it('expands identity and content macros', () => {
    expect(expandMacros('{{char}} meets {{user}}: {{description}}', { ...base })).toBe(
      'Aria meets Dev: a curious android',
    );
  });

  it('is case-insensitive and trims whitespace inside braces', () => {
    expect(expandMacros('{{ CHAR }} in {{Scenario}}', { ...base })).toBe('Aria in a rainy rooftop');
  });

  it('handles random/pick/roll deterministically with injected rng', () => {
    expect(expandMacros('{{random: a, b, c}}', { ...base })).toBe('b');
    expect(expandMacros('{{roll:d20}}', { ...base })).toBe('11');
  });

  it('supports chat and global variables', () => {
    const ctx = { ...base, vars: {}, globalVars: {} };
    const out = expandMacros('{{setvar::hp::10}}{{addvar::hp::5}}HP={{getvar::hp}}', ctx);
    expect(out).toBe('HP=15');
    expect(expandMacros('{{setglobalvar::mood::calm}}{{getglobalvar::mood}}', ctx)).toBe('calm');
  });

  it('strips comments and newline/noop control macros', () => {
    expect(expandMacros('a{{// hidden note}}b{{newline}}c{{noop}}', { ...base })).toBe('ab\nc');
  });

  it('leaves unknown macros untouched for author visibility', () => {
    expect(expandMacros('{{definitely_not_a_macro}}', { ...base })).toBe(
      '{{definitely_not_a_macro}}',
    );
  });

  it('resolves macros produced by variables (one extra pass)', () => {
    const ctx = { ...base, vars: { greet: 'hi {{char}}' } };
    expect(expandMacros('{{getvar::greet}}', ctx)).toBe('hi Aria');
  });
});
