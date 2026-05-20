# Excel UX Gap Analysis

This document tracks how close `casual-sheets` feels to Excel for a normal user.

It is intentionally practical:

- What a normal Excel user expects
- What we already have
- What is missing
- What feels different from Excel
- What should be fixed first

Evidence is based on the current codebase and E2E coverage, especially:

- `tests/e2e/excel-shortcut-coverage.spec.ts`
- `tests/e2e/excel-shortcuts.spec.ts`
- `tests/e2e/excel-shortcuts-polish.spec.ts`
- `tests/e2e/paste-from-excel.spec.ts`
- related E2E suites for charts, pivots, tables, formula bar, print, xlsx round-trip, and co-editing

## Summary

The product is already strong on core spreadsheet interaction:

- grid editing
- formula bar
- sheet navigation
- selection
- formatting basics
- sort/filter/table/pivot/chart flows
- xlsx import/export
- print/page setup

The biggest remaining Excel-fit gaps are:

1. external Excel paste fidelity
2. `Ctrl+1` Format Cells
3. `Ctrl+G` Go To
4. function discovery UX (`Shift+F3`, argument helper)
5. keyboard context/focus affordances
6. paste special / formatting-only paste
7. flash fill
8. multi-range selection mode

## Current Validation Status

Latest local full E2E run:

- `297 passed`
- `15 skipped`
- `0 failed`

The skipped cases mostly represent known Excel-parity gaps tracked below.

## Rating

High-level fit for a normal Excel user:

- Core editing/navigation: strong
- Formatting workflows: moderate
- Data-entry shortcuts: strong
- Advanced desktop-Excel affordances: partial
- External paste/import muscle-memory: partial
- Command discoverability: partial

Overall:

- Good spreadsheet UX
- Not yet full Excel-familiar UX

## What Already Matches Normal Excel Expectations

| Area | Status | Evidence | Notes |
| --- | --- | --- | --- |
| New workbook | Implemented | `tests/e2e/excel-shortcuts.spec.ts` | `Ctrl+N` |
| Open file | Implemented | `tests/e2e/excel-shortcuts.spec.ts`, `tests/e2e/xlsx-open.spec.ts` | `Ctrl+O` flow exists |
| Save As xlsx | Implemented | `tests/e2e/excel-shortcut-coverage.spec.ts` | `Alt+F2` |
| Print / Page Setup | Implemented | `tests/e2e/page-setup.spec.ts`, `tests/e2e/print.spec.ts` | `Ctrl+P` supported |
| Undo / Redo | Implemented | `tests/e2e/excel-shortcut-coverage.spec.ts`, `tests/e2e/wave-a.spec.ts` | keyboard + toolbar |
| Formula bar | Implemented | `tests/e2e/formula-bar.spec.ts` | strong coverage |
| Name Box | Implemented | `tests/e2e/formula-bar.spec.ts` | navigate/select ranges |
| Move across sheets | Implemented | `tests/e2e/excel-shortcuts.spec.ts` | `Ctrl+PageUp/PageDown` |
| Insert new sheet | Implemented | `tests/e2e/excel-shortcuts.spec.ts`, `tests/e2e/excel-shortcut-coverage.spec.ts` | `Shift+F11` |
| Jump to A1 / used range end | Implemented | `tests/e2e/excel-shortcuts.spec.ts`, `tests/e2e/excel-shortcut-coverage.spec.ts` | `Ctrl+Home`, `Ctrl+End` |
| Edit active cell | Implemented | `tests/e2e/excel-shortcuts-polish.spec.ts`, `tests/e2e/text-edit.spec.ts` | `F2` |
| Entire row / column select | Implemented | `tests/e2e/excel-shortcuts-polish.spec.ts`, `tests/e2e/excel-shortcut-coverage.spec.ts` | `Ctrl+Space`, `Shift+Space` |
| Insert / delete row/column dialogs | Implemented | `tests/e2e/excel-shortcuts-polish.spec.ts`, `tests/e2e/excel-shortcut-coverage.spec.ts` | `Ctrl++`, `Ctrl+-` |
| Hide / unhide rows and columns | Implemented | `tests/e2e/excel-shortcut-coverage.spec.ts` | `Ctrl+9`, `Ctrl+Shift+9`, `Ctrl+0`, `Ctrl+Shift+0` |
| Bold / formatting basics | Implemented | `tests/e2e/home-ribbon.spec.ts`, `tests/e2e/ribbon-tabs.spec.ts`, `tests/e2e/excel-shortcut-coverage.spec.ts` | bold, wrap, alignment, borders, number formats |
| Date / time insertion | Implemented | `tests/e2e/excel-shortcut-coverage.spec.ts`, `tests/e2e/excel-shortcuts.spec.ts` | `Ctrl+;`, `Ctrl+Shift+;` |
| AutoSum / recalc | Implemented | `tests/e2e/excel-shortcut-coverage.spec.ts` | `Alt+=`, `F9` |
| Sort / filter basics | Implemented | `tests/e2e/ribbon-tabs.spec.ts`, `tests/e2e/excel-shortcut-coverage.spec.ts`, `tests/e2e/show-all-rows.spec.ts` | filter UX still less robust than Excel |
| Format as Table | Implemented | `tests/e2e/format-as-table.spec.ts`, `tests/e2e/tables-panel.spec.ts` | good coverage |
| Charts | Implemented | `tests/e2e/charts-p*.spec.ts` | beyond typical web-sheet baseline |
| Pivots | Implemented | `tests/e2e/pivots-p0.spec.ts` | uncommon for early-stage sheet products |
| Freeze panes | Implemented | `tests/e2e/freeze-additive.spec.ts`, `tests/e2e/wave-b.spec.ts` | |
| Hidden sheets | Implemented | `tests/e2e/hidden-sheets.spec.ts` | |
| Outline / grouping | Implemented | `tests/e2e/outline.spec.ts` | |
| Cross-sheet formulas | Implemented | `tests/e2e/cross-sheet-refs.spec.ts`, `tests/e2e/cross-sheet-picker.spec.ts` | |
| xlsx round-trip | Implemented | `tests/e2e/xlsx-round-trip.spec.ts`, `tests/e2e/xlsx-defined-names.spec.ts` | strong coverage for supported features |

## UX / Feature Gaps That Will Stand Out to Excel Users

| Feature / UX | Current status | Evidence | User impact | Difficulty | Priority | Recommended fix |
| --- | --- | --- | --- | --- | --- | --- |
| Paste from desktop Excel with high fidelity | Missing / partial | `tests/e2e/paste-from-excel.spec.ts` has `fixme` cases | Very high | Medium | P0 | Preserve values, bold headers, merges, and `x:fmla` formulas from Excel HTML clipboard |
| `Ctrl+1` Format Cells dialog | Missing | `excel-shortcut-coverage.spec.ts` `test.fixme('Ctrl+1 ...')` | Very high | Medium | P0 | Add Number / Alignment / Font / Border / Fill tabs; start compact |
| `Ctrl+G` Go To | Missing | `excel-shortcut-coverage.spec.ts` `test.fixme('Ctrl+G ...')` | High | Low | P0 | Focus Name Box first; later add true Go To dialog |
| Insert Function dialog (`Shift+F3`) | Missing | `excel-shortcut-coverage.spec.ts` | High | Medium | P1 | Searchable function picker with category + description |
| Function argument helper (`Ctrl+Shift+A`) | Missing | `excel-shortcut-coverage.spec.ts` | High | Medium | P1 | Show argument tooltip / insertion helper while editing formulas |
| Paste formatting only (`Ctrl+Shift+V`) | Missing | `excel-shortcut-coverage.spec.ts` | High | Medium | P1 | Add Paste Special path and menu exposure |
| Multi-range selection mode (`Shift+F8`) | Missing | `excel-shortcut-coverage.spec.ts` | Medium | Medium | P1 | Support additive non-adjacent range selection |
| Re-apply filter (`Ctrl+Alt+L`) | Missing | `excel-shortcut-coverage.spec.ts` | Medium | Low | P2 | Re-run current filter model over range |
| Flash Fill (`Ctrl+E`) | Missing | `excel-shortcut-coverage.spec.ts` | Medium | High | P2 | Start with simple pattern-based fill for common cases |
| Keyboard context menu (`Shift+F10`) | Missing | `excel-shortcut-coverage.spec.ts` | Medium | Low | P2 | Open current cell/range context menu from keyboard |
| Ribbon/grid focus transfer (`Ctrl+F6`) | Missing | `excel-shortcut-coverage.spec.ts` | Medium | Medium | P2 | Formalize focus zones and cycling order |
| Search / Tell Me (`Alt+Q`) | Missing | `excel-shortcut-coverage.spec.ts` | Medium | Medium | P3 | Add command palette / action search |
| AutoFilter state is not exposed cleanly | Partial | `Ctrl+Shift+L` currently smoke-tested, not deep-asserted | Medium | Medium | P2 | Provide stable filter read model / UI signal for tests and UX |

## Features a Normal Excel User Commonly Uses That Are Not There Yet

This is the practical missing list, not the entire Excel feature surface.

| Feature | Present? | Notes |
| --- | --- | --- |
| Go To (`Ctrl+G`) | No | Name Box exists, but Excel-style Go To flow does not |
| Format Cells dialog | No | major parity gap |
| Insert Function dialog | No | function search/discovery weaker than Excel |
| Function arguments helper | No | argument editing assistance missing |
| Paste Special / formatting-only paste | No | common workflow gap |
| Robust external Excel paste fidelity | Partial | major migration friction |
| Keyboard context menu | No | mouse path exists, keyboard parity missing |
| Multi-range selection mode | No | advanced but familiar Excel behavior |
| Flash Fill | No | common on modern Excel workflows |
| Tell Me / command search | No | useful for discoverability and keyboard-driven users |

## Places Where UX Differs From Excel Even When Feature Exists

| Area | Difference today | Why it matters | Fix direction |
| --- | --- | --- | --- |
| AutoFilter | Present, but readback / state signal is not robust | feels less trustworthy and is harder to test | stable filter model exposure, clearer UI marker |
| Table creation | Works, but current test coverage treats it partly as smoke behavior | suggests the API/observable model is less stable than ideal | expose table state more directly |
| Print | Web-native print approximation, not true Excel page-layout fidelity | Excel users expect page setup to map closer to output | expand print rendering fidelity over time |
| Clipboard flows | Internal copy/paste is fine; desktop Excel HTML paste is incomplete | migration workflow pain | prioritize external clipboard compatibility |
| Command discovery | many commands exist only through menus or custom placements | Excel users rely on known dialogs and accelerators | add canonical dialogs and shortcut parity |

## Features That Set This Apart From Excel

These are positive product differences, not parity problems.

| Feature | Why it stands out | Evidence |
| --- | --- | --- |
| Real-time co-editing | Native multi-user collaboration is a major product differentiator | `tests/e2e/coedit.spec.ts`, `tests/e2e/coedit-share.spec.ts` |
| Share / password / view-only flows | More lightweight than desktop Excel sharing setups | `tests/e2e/coedit-share.spec.ts` |
| Autosave restore UX | Friendly browser-native recovery flow | `tests/e2e/autosave.spec.ts` |
| Responsive web shell | Better phone / tablet story than desktop Excel | `tests/e2e/mobile.spec.ts`, `tests/e2e/responsive-layout.spec.ts` |
| Side panels for charts / tables / outline | More app-like, discoverable web affordances | multiple panel E2Es |
| Rich xlsx sidecar strategy | Preserves charts/pivots/defined-names beyond plain ExcelJS baseline | `tests/e2e/xlsx-defined-names.spec.ts`, `tests/e2e/charts-p*.spec.ts`, `tests/e2e/pivots-p0.spec.ts` |

## Shortcut Parity Snapshot

### Good parity now

- `Ctrl+N`
- `Ctrl+O`
- `Ctrl+P`
- `Ctrl+F`
- `Ctrl+B`
- `Ctrl+Home`
- `Ctrl+End`
- `Ctrl+PageUp/PageDown`
- `Shift+F11`
- `Ctrl+Space`
- `Shift+Space`
- `Ctrl++`
- `Ctrl+-`
- `Ctrl+9` / `Ctrl+Shift+9`
- `Ctrl+0` / `Ctrl+Shift+0`
- `Ctrl+;`
- `Ctrl+Shift+;`
- `Alt+=`
- `F2`
- `F9`

### Missing or partial parity

- `Ctrl+G`
- `Shift+F3`
- `Shift+F10`
- `Alt+Q`
- `Ctrl+Alt+L`
- `Ctrl+Shift+A`
- `Ctrl+Shift+V`
- `Ctrl+1`
- `Ctrl+E`
- `Shift+F8`

## Recommended Fix Order

### P0

1. External Excel paste fidelity
2. `Ctrl+1` Format Cells
3. `Ctrl+G` Go To

### P1

4. Insert Function dialog
5. Function argument helper
6. Paste Special / formatting-only paste
7. Multi-range selection

### P2

8. Re-apply filter
9. Keyboard context menu
10. Focus-zone navigation
11. Stabilize filter UX/read model
12. Flash Fill

### P3

13. Command palette / Tell Me

## Product Recommendation

If the explicit goal is “a normal Excel user should feel at home in the first 5 minutes,” focus on these:

- desktop Excel paste
- Format Cells
- Go To
- function discovery

Those are the gaps users feel immediately.

If the goal is “better than Excel for collaborative web work,” the product is already differentiated:

- co-editing
- sharing
- autosave recovery
- mobile/responsive behavior

The best strategy is not to copy all of Excel first. It is:

1. close the top muscle-memory gaps
2. keep the existing web-native strengths
3. avoid introducing heavy desktop-style complexity where a simpler web UX is already better

