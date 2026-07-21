import { base64ToUtf8, readPngTextChunks } from './png';
import { makeCharacter } from '@/db/repo';
import { db } from '@/db/db';
import { now, uuid } from '@/lib/id';
import type { Character } from '@/types';

/*
 * Character Card V2/V3 import (docs/04 "Card spec mapping").
 * Accepts: card JSON (.json) or a SillyTavern PNG with an embedded card
 * (tEXt `chara` = V2, `ccv3` = V3). The PNG itself becomes the avatar asset.
 */

interface CardData {
  name?: string;
  nickname?: string;
  description?: string;
  personality?: string;
  scenario?: string;
  first_mes?: string;
  alternate_greetings?: string[];
  mes_example?: string;
  system_prompt?: string;
  post_history_instructions?: string;
  creator_notes?: string;
  tags?: string[];
  creator?: string;
  character_version?: string;
}

interface CardFile {
  spec?: string;
  spec_version?: string;
  data?: CardData;
  // V1 cards put fields at the top level:
  name?: string;
}

export function cardJsonToCharacter(json: string): Character {
  const card = JSON.parse(json) as CardFile;
  const isV3 = card.spec === 'chara_card_v3';
  // V2/V3 nest fields under `data`; V1 has them at the top level.
  const d: CardData = card.data ?? (card as CardData);
  if (!d.name?.trim()) throw new Error('Card has no character name');
  return makeCharacter({
    spec: isV3 ? 'chara_card_v3' : 'chara_card_v2',
    name: d.name,
    nickname: d.nickname || undefined,
    description: d.description ?? '',
    personality: d.personality ?? '',
    scenario: d.scenario ?? '',
    firstMes: d.first_mes ?? '',
    alternateGreetings: d.alternate_greetings ?? [],
    mesExample: d.mes_example ?? '',
    systemPrompt: d.system_prompt || undefined,
    postHistoryInstructions: d.post_history_instructions || undefined,
    creatorNotes: d.creator_notes || undefined,
    tags: d.tags ?? [],
    creator: d.creator || undefined,
    characterVersion: d.character_version || undefined,
  });
}

/** Extract the embedded card JSON from a ST PNG (V3 `ccv3` preferred over V2 `chara`). */
export function pngToCardJson(bytes: Uint8Array): string {
  const chunks = readPngTextChunks(bytes);
  const b64 = chunks['ccv3'] ?? chunks['chara'];
  if (!b64) throw new Error('PNG has no embedded character card');
  return base64ToUtf8(b64);
}

/** Import a picked file (.png or .json) → persisted Character (+ avatar asset). */
export async function importCharacterFile(file: File): Promise<Character> {
  if (/\.png$/i.test(file.name) || file.type === 'image/png') {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const character = cardJsonToCharacter(pngToCardJson(bytes));
    character.avatarAssetId = await saveAvatarBlob(new Blob([bytes], { type: 'image/png' }));
    await db.characters.add(character);
    return character;
  }
  const character = cardJsonToCharacter(await file.text());
  await db.characters.add(character);
  return character;
}

export async function saveAvatarBlob(blob: Blob): Promise<string> {
  const { width, height } = await imageSize(blob);
  const id = uuid();
  await db.galleryAssets.add({
    id,
    kind: 'avatar',
    blob,
    width,
    height,
    mime: blob.type || 'image/png',
    createdAt: now(),
  });
  return id;
}

function imageSize(blob: Blob): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      resolve({ width: img.naturalWidth, height: img.naturalHeight });
      URL.revokeObjectURL(url);
    };
    img.onerror = () => {
      resolve({ width: 0, height: 0 });
      URL.revokeObjectURL(url);
    };
    img.src = url;
  });
}

/** Export a character as V2 card JSON (PNG re-embedding arrives in Phase 4 polish). */
export function characterToCardJson(c: Character): string {
  return JSON.stringify(
    {
      spec: 'chara_card_v2',
      spec_version: '2.0',
      data: {
        name: c.name,
        description: c.description,
        personality: c.personality,
        scenario: c.scenario,
        first_mes: c.firstMes,
        alternate_greetings: c.alternateGreetings,
        mes_example: c.mesExample,
        system_prompt: c.systemPrompt ?? '',
        post_history_instructions: c.postHistoryInstructions ?? '',
        creator_notes: c.creatorNotes ?? '',
        tags: c.tags,
        creator: c.creator ?? '',
        character_version: c.characterVersion ?? '',
        extensions: {},
      },
    },
    null,
    2,
  );
}
