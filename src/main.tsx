import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import './index.css';
import { getSettings, seed } from './db/repo';
import { applyTheme } from './lib/themes';

// First-run seeding (default preset + settings) and theme application happen
// before render so the initial paint uses the right palette.
async function boot() {
  await seed();
  const settings = await getSettings();
  applyTheme(settings.themeId);
  ReactDOM.createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>,
  );
}

void boot();
