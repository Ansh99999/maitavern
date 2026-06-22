# 01 — Chat Interface

Mobile-first, portrait. Three regions: a slim top bar, a maximal message area, and a fixed composer.

```
┌────────────────────────────────────────────┐
│ ‹   Aria                          [☰]      │  Top bar (slim)
├────────────────────────────────────────────┤
│              MESSAGE AREA                   │  Max space, virtualized
├────────────────────────────────────────────┤
│ (P)  Type a message…           [⋮]  [➤]   │  Composer
└────────────────────────────────────────────┘
```

- **Top bar:** back · character name (+ tiny avatar) · `☰` opens the **Quick Settings sidebar** (swap preset/provider, author's note, lorebooks, theme, global logs).

## Display styles (user-selectable; remembered per style)

Four styles: **bubble · document · discord · novel**.

| Style | Avatars | Name | Alignment | Best for |
|---|---|---|---|---|
| Bubble | beside bubble | above bubble | bot-left / user-right | messenger feel |
| Document | small/hidden | bold header | all left, full-width | long-form reading |
| Discord | fixed left, grouped | name + timestamp inline | all left | channel-style RP |
| Novel | optional | optional | continuous prose | immersive reading |

## Avatars (universal model, all four modes)

- **Show/Hide** toggle in every mode, remembered **per style**. Defaults: on for bubble/discord/document, off for novel — all flippable. **Novel can show avatars** (avatar at the start of the speaker's prose passage).
- **Never cropped or distorted** — always native aspect ratio (`object-fit: contain`), rounded corners ok, no circle-crop.
- **Two sizing modes:** **Default/Auto** (the layout accommodates the avatar's native aspect ratio — 16:9 → wide, 4:3 → taller, 1:1 → square, 2:3 → tall) or **Custom** (slider-set size). Both preserve aspect ratio.
- **Text relation, user choice in every mode:** wrap-around the avatar, or begin below it.

### Avatar Viewer

Tapping a message avatar opens a **floating, draggable** window of that avatar (user or char): pinch / `＋－` **zoom**, pan, aspect preserved, **✕** top-right to close (tap-outside also dismisses), and a shortcut to edit the avatar in Settings.

## Message anatomy

```
┌────┐  Aria          ⟳  ✎  🗑   ⋮
│img │  ┌──────────────────────────┐
└────┘  │ message text (markdown)   │
        └──────────────────────────┘
                              ‹ 2/3 ›   swipe controls (last bot message)
```

- **Name** — toggleable.
- **Inline actions** (reveal on tap/long-press by default; "always show" setting): bot → `⟳ regenerate · ✎ edit · 🗑 delete`; user → `✎ edit · 🗑 delete`.
- **Swipes** — `‹ n/m ›` on the last bot message (swipe gesture too).
- **⋮ overflow** (also via long-press): fork/branch here · view request log (assembled prompt + raw req/resp + tokens) · copy · checkpoint/bookmark · hide-from-context (mute) · truncate here · save-to-memory/lorebook · [Phase 5: translate, TTS] · delete.

## Composer

```
(P)  Type a message…                    [⋮]  [➤]
 │                                        │    │
 └ persona avatar — tap to switch         │    └ Send / Stop
   (long-press = avatar viewer)           └ Compose actions (combined sheet)
```

- **Persona avatar (left)** — tap to switch persona mid-chat.
- **Send button** — disabled when input is empty; **during generation it morphs into a Stop ◼** button (abort streaming) while still blocking a second send.
- **⋮ Compose actions (combined sheet):** generation actions on top — impersonate · continue · regenerate-last · guided reply · insert quick-reply/macro · author's-note quick-edit · new chat · toggle streaming — with a button at the bottom that **expands into the full Quick Settings sidebar**.

**Split rule:** per-message `⋮` acts on one message; composer `⋮` acts on the next turn/session.

## Component map

`ChatScreen` → `ChatTopBar` · `MessageList` (virtualized) → `MessageRow` (style variant) = `MessageAvatar` + `MessageHeader` (name + inline actions + ⋮) + `MessageBody` (markdown/edit) + `SwipeControls` · `Composer` = `PersonaAvatarButton` + `MessageInput` + `ComposeActionsButton` + `SendButton` · overlays: `AvatarViewer`, `MessageActionSheet`, `PersonaSwitcher`, `ComposeActionsSheet`, `QuickSettingsSidebar`.
