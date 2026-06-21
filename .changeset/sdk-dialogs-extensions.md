---
'@casualoffice/sheets': minor
---

The SDK chrome now has **built-in dialogs** and a **host extension API** — the two things that made it a viewer rather than a real editor.

- **Dialog host**: chrome controls open a built-in dialog by default. Ships **Format Cells** (Number/Alignment/Font/Border/Fill) and wires **Find & Replace** into the dialog host. A `Dialog` primitive (portal + focus-trap + Escape/backdrop) backs them.
- **Extension API** (`extensions` prop on `<CasualSheets>`, also forwarded through the iframe chrome): hosts add their **own** toolbar buttons, menu items, side panels, and dialogs on top of the built-ins — or **override** a built-in dialog by kind. Resolution order: host React override (`extensions.dialogs[kind]`) → host-owned via `onDialogRequest`/`hostOwnedDialogs` → SDK built-in. `onDialogRequest` stays backward-compatible. New exports: `ChromeExtensions`, `ToolbarExtension`, `MenuExtension`, `PanelExtension`, `DialogExtension`, `DialogComponentProps`, `PanelComponentProps`, `DialogKind`.
- **Fix**: the sheet-tab right-click menu rendered behind the bars below the tab strip (stacking-context trap) — it's now portaled to `<body>` and opens upward, fully visible.

Remaining built-in dialogs (Goal Seek, Name Manager, Page Setup, Paste Special, Insert Cells, Properties, Keyboard Shortcuts, About, Watermark, Insert Chart/PivotTable) and the side-panel rail land next.
