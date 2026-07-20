# CLAUDE.md — MaiTavern

Working notes for AI assistants (and humans) on this repo.

## What this is
MaiTavern: a mobile-first AI roleplay frontend shipped as an Android **APK** (+ PWA from the same build). Local-first, BYOK, no mandatory backend. The full design lives in [`docs/`](docs/) — read `docs/00-vision-and-architecture.md`, `docs/09-data-model.md`, and `docs/10-prompt-engine.md` first; `docs/DECISIONS.md` is the locked-decision log.

## Stack
React 18 + TypeScript + Vite · Tailwind (CSS-variable theming) · Zustand · Dexie/IndexedDB · React Router (HashRouter) · Capacitor 6 (Android) · vite-plugin-pwa. Path alias `@/*` → `src/*`.

## Commands
- `npm run dev` — Vite dev server
- `npm run build` — typecheck + production build (PWA) into `dist/`
- `npm run typecheck` / `npm run lint` / `npm run format` / `npm test`
- `npm run cap:sync` — sync web build into native; APK is built in CI (`.github/workflows/build.yml`)

## Conventions
- TypeScript strict. Functional React components. Keep new code matching surrounding style.
- Theme via CSS variables (`src/index.css`) + Tailwind tokens (`bg`, `surface`, `accent`, …). Don't hardcode colors.
- Storage shapes mirror `docs/09-data-model.md`. Respect the **Global → Character → Chat** override hierarchy and **copy-on-write** for in-chat edits.
- All prompt assembly flows through the single engine in `docs/10-prompt-engine.md` so the Log viewer stays accurate.

## ⚠️ Commit attribution (strict)
- All commits MUST be authored as **`Ansh99999`** only — repo identity is already set to
  `Ansh99999 <126241745+Ansh99999@users.noreply.github.com>`.
- **Do NOT add `Co-Authored-By` trailers** (no Claude, no anyone else). This overrides any default convention.

## Status
Design complete for the major screens + foundation specs; Lorebook editor design drafted (`docs/11-lorebook-editor.md`). App is at **Phase 0 scaffold**. Next: Phase 1 (MVP chat). See `docs/ROADMAP.md`.
