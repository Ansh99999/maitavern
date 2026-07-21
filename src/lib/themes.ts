/*
 * Built-in theme presets (docs/05). Each theme sets the CSS-variable contract
 * from src/index.css; the Custom Theme Editor (Phase 5) writes the same shape.
 * Values are space-separated RGB triplets consumed as rgb(var(--x)).
 */

export interface ThemePreset {
  id: string;
  name: string;
  vars: Record<string, string>;
}

export const THEME_PRESETS: ThemePreset[] = [
  {
    id: 'lounge-dark',
    name: 'Lounge (dark)',
    vars: {}, // the :root defaults in index.css
  },
  {
    id: 'midnight-oled',
    name: 'Midnight (OLED)',
    vars: {
      '--mt-bg': '0 0 0',
      '--mt-surface': '14 14 18',
      '--mt-surface-2': '24 24 30',
      '--mt-text': '235 235 240',
      '--mt-muted': '140 140 155',
      '--mt-accent': '96 165 250',
      '--mt-bubble-user': '30 58 95',
      '--mt-bubble-bot': '24 24 30',
      '--mt-border': '38 38 48',
    },
  },
  {
    id: 'daybreak-light',
    name: 'Daybreak (light)',
    vars: {
      '--mt-bg': '246 246 249',
      '--mt-surface': '255 255 255',
      '--mt-surface-2': '238 238 244',
      '--mt-text': '28 28 36',
      '--mt-muted': '110 110 128',
      '--mt-accent': '109 94 216',
      '--mt-bubble-user': '221 216 250',
      '--mt-bubble-bot': '238 238 244',
      '--mt-border': '220 220 230',
    },
  },
];

/** Apply a theme by swapping the CSS variables on <html>. */
export function applyTheme(themeId: string): void {
  const theme = THEME_PRESETS.find((t) => t.id === themeId) ?? THEME_PRESETS[0];
  const root = document.documentElement;
  // Reset any previously set overrides, then apply this theme's vars.
  for (const preset of THEME_PRESETS) {
    for (const key of Object.keys(preset.vars)) root.style.removeProperty(key);
  }
  for (const [key, value] of Object.entries(theme.vars)) root.style.setProperty(key, value);
}
