import { useEffect, useState } from 'react';
import { Link } from 'react-router-dom';
import { getSettings, patchSettings } from '@/db/repo';
import { applyTheme, THEME_PRESETS } from '@/lib/themes';
import type { ChatStyle, Settings as SettingsShape } from '@/types';

/*
 * Settings hub (docs/05). Phase 1: Appearance (theme presets, default chat
 * style, your name) inline + Provider / Preset / Logs subscreens. The full
 * searchable, override-scoped settings tree grows here in later phases.
 */
export default function Settings() {
  const [settings, setSettings] = useState<SettingsShape>();

  useEffect(() => {
    getSettings().then(setSettings);
  }, []);

  async function patch(p: Partial<SettingsShape>) {
    const next = await patchSettings(p);
    setSettings(next);
    if (p.themeId) applyTheme(p.themeId);
  }

  if (!settings) return null;

  const select =
    'w-full px-3 py-2 rounded-xl bg-surface border border-border outline-none focus:border-accent';

  return (
    <div className="min-h-full bg-bg text-text font-ui">
      <header className="flex items-center gap-3 px-5 pt-4 pb-3 border-b border-border">
        <Link to="/" aria-label="Back" className="text-muted text-lg">‹</Link>
        <h1 className="text-lg font-semibold">Settings</h1>
      </header>

      <main className="px-5 py-4 flex flex-col gap-6 max-w-xl mx-auto pb-16">
        <section>
          <h2 className="text-sm font-semibold text-muted uppercase tracking-wide mb-3">Connections</h2>
          <div className="grid grid-cols-2 gap-3">
            <Card to="/settings/provider" title="Provider" subtitle="API endpoint · key · model" />
            <Card to="/settings/preset" title="Preset" subtitle="sampler · context · blocks" />
          </div>
        </section>

        <section>
          <h2 className="text-sm font-semibold text-muted uppercase tracking-wide mb-3">Appearance</h2>
          <label className="block mb-4">
            <div className="text-sm font-medium mb-1">Theme</div>
            <select
              className={select}
              value={settings.themeId}
              onChange={(e) => patch({ themeId: e.target.value })}
            >
              {THEME_PRESETS.map((t) => (
                <option key={t.id} value={t.id}>{t.name}</option>
              ))}
            </select>
          </label>
          <label className="block mb-4">
            <div className="text-sm font-medium mb-1">Default chat style</div>
            <select
              className={select}
              value={settings.chatStyle}
              onChange={(e) => patch({ chatStyle: e.target.value as ChatStyle })}
            >
              <option value="bubble">Bubble</option>
              <option value="document">Document</option>
              <option value="discord">Discord</option>
              <option value="novel">Novel</option>
            </select>
            <div className="text-xs text-muted mt-1">Individual chats can override this.</div>
          </label>
          <label className="block">
            <div className="text-sm font-medium mb-1">Your name</div>
            <input
              className={select}
              value={settings.userName}
              onChange={(e) => patch({ userName: e.target.value })}
              placeholder="You"
            />
            <div className="text-xs text-muted mt-1">
              Used for {'{{user}}'} in prompts. Full personas arrive in Phase 2.
            </div>
          </label>
        </section>

        <section>
          <h2 className="text-sm font-semibold text-muted uppercase tracking-wide mb-3">Data</h2>
          <div className="grid grid-cols-2 gap-3">
            <Card to="/settings/logs" title="Request logs" subtitle="assembled prompts · usage" />
          </div>
        </section>
      </main>
    </div>
  );
}

function Card({ to, title, subtitle }: { to: string; title: string; subtitle: string }) {
  return (
    <Link
      to={to}
      className="rounded-2xl bg-surface border border-border p-4 active:bg-surface-2 transition-colors"
    >
      <div className="font-medium">{title}</div>
      <div className="text-sm text-muted">{subtitle}</div>
    </Link>
  );
}
