import { describe, expect, it } from 'vitest';
import { cardJsonToCharacter } from './cardImport';
import { pngToCardJson } from './cardImport';

const v2Card = JSON.stringify({
  spec: 'chara_card_v2',
  spec_version: '2.0',
  data: {
    name: 'Aria',
    description: 'an android bard',
    personality: 'warm, curious',
    scenario: 'a rooftop tavern',
    first_mes: 'Hey, you made it.',
    alternate_greetings: ['Oh! Hello.'],
    mes_example: '<START>\n{{user}}: hi\n{{char}}: hey',
    system_prompt: 'Stay poetic.',
    tags: ['fantasy'],
    creator: 'dev',
  },
});

/** Build a minimal PNG containing one tEXt chunk with the given keyword/payload. */
function makePng(keyword: string, payloadB64: string): Uint8Array {
  const sig = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  const text = [...keyword].map((c) => c.charCodeAt(0)).concat(
    0,
    [...payloadB64].map((c) => c.charCodeAt(0)),
  );
  const chunk = (type: string, data: number[]) => {
    const len = [data.length >>> 24, (data.length >>> 16) & 255, (data.length >>> 8) & 255, data.length & 255];
    const t = [...type].map((c) => c.charCodeAt(0));
    return len.concat(t, data, [0, 0, 0, 0]); // CRC unchecked by our reader
  };
  return new Uint8Array(sig.concat(chunk('tEXt', text.flat() as number[]), chunk('IEND', [])));
}

describe('cardJsonToCharacter', () => {
  it('maps V2 card fields', () => {
    const c = cardJsonToCharacter(v2Card);
    expect(c.spec).toBe('chara_card_v2');
    expect(c.name).toBe('Aria');
    expect(c.firstMes).toBe('Hey, you made it.');
    expect(c.alternateGreetings).toEqual(['Oh! Hello.']);
    expect(c.systemPrompt).toBe('Stay poetic.');
    expect(c.tags).toEqual(['fantasy']);
  });

  it('accepts V1 top-level cards', () => {
    const c = cardJsonToCharacter(JSON.stringify({ name: 'Old Timer', description: 'v1' }));
    expect(c.name).toBe('Old Timer');
    expect(c.description).toBe('v1');
  });

  it('rejects cards without a name', () => {
    expect(() => cardJsonToCharacter('{"data":{}}')).toThrow(/name/);
  });
});

describe('pngToCardJson', () => {
  it('extracts a base64 chara tEXt chunk', () => {
    const png = makePng('chara', btoa(v2Card));
    expect(JSON.parse(pngToCardJson(png)).data.name).toBe('Aria');
  });

  it('throws on a PNG without a card', () => {
    const png = makePng('comment', btoa('hi'));
    expect(() => pngToCardJson(png)).toThrow(/no embedded/);
  });
});
