import type { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'app.maitavern',
  appName: 'MaiTavern',
  webDir: 'dist',
  android: {
    // Allow plaintext HTTP so users can reach local model servers
    // (KoboldCpp / Ollama / oobabooga) on their LAN. See docs/00.
    allowMixedContent: true,
  },
  server: {
    androidScheme: 'https',
    cleartext: true,
  },
};

export default config;
