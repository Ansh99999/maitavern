/** UUID v4 (docs/09: IDs are UUID v4 strings, stable across export/sync). */
export function uuid(): string {
  return crypto.randomUUID();
}

export function now(): number {
  return Date.now();
}
