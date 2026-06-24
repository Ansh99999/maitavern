/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{ts,tsx}'],
  theme: {
    extend: {
      // Theme is driven by CSS variables (see src/index.css) so it can be
      // swapped at runtime per the Global → Character → Chat override system.
      colors: {
        bg: 'rgb(var(--mt-bg) / <alpha-value>)',
        surface: 'rgb(var(--mt-surface) / <alpha-value>)',
        'surface-2': 'rgb(var(--mt-surface-2) / <alpha-value>)',
        text: 'rgb(var(--mt-text) / <alpha-value>)',
        muted: 'rgb(var(--mt-muted) / <alpha-value>)',
        accent: 'rgb(var(--mt-accent) / <alpha-value>)',
        'bubble-user': 'rgb(var(--mt-bubble-user) / <alpha-value>)',
        'bubble-bot': 'rgb(var(--mt-bubble-bot) / <alpha-value>)',
        border: 'rgb(var(--mt-border) / <alpha-value>)',
      },
      fontFamily: {
        ui: 'var(--mt-font-ui)',
        chat: 'var(--mt-font-chat)',
      },
    },
  },
  plugins: [],
};
