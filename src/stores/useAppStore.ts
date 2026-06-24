import { create } from 'zustand';

/*
 * Minimal app store stub. Real stores (settings, chats, providers, presets)
 * arrive in Phase 1; this proves the Zustand wiring in Phase 0.
 */
interface AppState {
  themeId: string;
  setThemeId: (id: string) => void;
}

export const useAppStore = create<AppState>((set) => ({
  themeId: 'default-dark',
  setThemeId: (id) => set({ themeId: id }),
}));
