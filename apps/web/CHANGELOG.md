# @sheet/web

## 0.4.6

### Patch Changes

- Updated dependencies [8bb728f]
- Updated dependencies [a3493fb]
- Updated dependencies [a921179]
  - @casualoffice/sheets@0.15.0

## 0.4.5

### Patch Changes

- Updated dependencies [224fc2c]
- Updated dependencies [33ded85]
- Updated dependencies [dfc8e6b]
  - @casualoffice/sheets@0.14.0

## 0.4.4

### Patch Changes

- Updated dependencies [2846392]
  - @casualoffice/sheets@0.13.0

## 0.4.3

### Patch Changes

- Updated dependencies [8b35360]
- Updated dependencies [58ce6a0]
- Updated dependencies [1adc983]
  - @casualoffice/sheets@0.12.0

## 0.4.2

### Patch Changes

- Updated dependencies [971ad7d]
  - @casualoffice/sheets@0.11.1

## 0.4.1

### Patch Changes

- Updated dependencies [6c8a94e]
- Updated dependencies [3c93042]
  - @casualoffice/sheets@0.11.0

## 0.4.0

### Minor Changes

- 4fd30c5: apps/web shares the SDK editor core (Phase 3 step 1)

  `apps/web` no longer hand-rolls its Univer bootstrap — `UniverSheet.tsx` now
  renders `<CasualSheets chrome="none">` from `@casualoffice/sheets`, sharing the
  SDK's Univer boot, plugin set, formula engine, and snapshot/API. The app keeps
  its rich shell (ribbon, charts, pivots, panels, dialogs) and layers its extras
  on top: crosshair-highlight + zen-editor + Merge/Unmerge context menu via
  `onBeforeCreateUnit`, off-main compute via `formula={{ worker }}`, and the
  paste-merge hook / dev helpers / zoom-shortcut override via `onReady`. One Univer
  bootstrap now serves both the app and third-party SDK hosts.

### Patch Changes

- Updated dependencies [49a3215]
- Updated dependencies [5256f3d]
- Updated dependencies [7f42243]
- Updated dependencies [29744e8]
- Updated dependencies [ce87187]
- Updated dependencies [99b617f]
- Updated dependencies [f6b1b24]
- Updated dependencies [67e0d55]
- Updated dependencies [7816a5d]
- Updated dependencies [1495444]
- Updated dependencies [838ce1b]
- Updated dependencies [3d9d0b5]
- Updated dependencies [35abbab]
- Updated dependencies [91ff777]
- Updated dependencies [f8b05b4]
- Updated dependencies [a090e65]
- Updated dependencies [65124b4]
- Updated dependencies [53b87fe]
- Updated dependencies [ea014be]
- Updated dependencies [f0d5779]
- Updated dependencies [c007f64]
- Updated dependencies [161aa91]
- Updated dependencies [3c5a990]
  - @casualoffice/sheets@0.10.0

## 0.3.13

### Patch Changes

- Updated dependencies [652068f]
- Updated dependencies [f93fa6c]
- Updated dependencies [d3f9be6]
- Updated dependencies [1da029e]
- Updated dependencies [2381fb4]
  - @casualoffice/sheets@0.9.0

## 0.3.12

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.8.0

## 0.3.11

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.7.0

## 0.3.10

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.6.0

## 0.3.9

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.5.7

## 0.3.8

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.5.6

## 0.3.7

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.5.5

## 0.3.6

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.5.4

## 0.3.5

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.5.3

## 0.3.4

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.5.2

## 0.3.3

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.5.1

## 0.3.2

### Patch Changes

- Updated dependencies [e044efd]
  - @casualoffice/sheets@0.5.0

## 0.3.1

### Patch Changes

- Updated dependencies
  - @casualoffice/sheets@0.4.0
