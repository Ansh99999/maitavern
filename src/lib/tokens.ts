/*
 * Approximate tokenizer (docs/10 default): chars/4 heuristic plus a safety
 * margin at budget time. Exact tiktoken for OpenAI-family can slot in later
 * behind the same signature.
 */
export function estimateTokens(text: string): number {
  return Math.ceil(text.length / 4);
}

export const SAFETY_MARGIN = 256;
