# ODS fidelity tracker

Working tracker for the ODS round-trip fidelity pass. This file is the execution checklist, not a brainstorm.

## Ground rules

- Follow [`CLAUDE.md`](/Users/sachin/Desktop/melp/services/sheet/CLAUDE.md:1) and keep docs terse and decision-oriented.
- Prefer native ODS metadata where SheetJS exposes it.
- If the upstream SheetJS ODS parser does not expose a feature, record that constraint here before attempting a workaround.
- Do not mark a task complete until:
  - the implementation lands
  - a Playwright test covers the supported behavior
  - this tracker is updated with the result

## Search map

Use these first before broad repo search:

- ODS codec: [`apps/web/src/ods/index.ts`](/Users/sachin/Desktop/melp/services/sheet/apps/web/src/ods/index.ts:1)
- Existing ODS smoke spec: [`tests/e2e/ods-csv-tsv.spec.ts`](/Users/sachin/Desktop/melp/services/sheet/tests/e2e/ods-csv-tsv.spec.ts:1)
- xlsx audit reference: [`tests/e2e/xlsx-lossiness-audit.spec.ts`](/Users/sachin/Desktop/melp/services/sheet/tests/e2e/xlsx-lossiness-audit.spec.ts:1)
- xlsx style mapping reference: [`apps/web/src/xlsx/style-mapping.ts`](/Users/sachin/Desktop/melp/services/sheet/apps/web/src/xlsx/style-mapping.ts:1)
- xlsx parse/export metadata reference:
  - [`apps/web/src/xlsx/parse-impl.ts`](/Users/sachin/Desktop/melp/services/sheet/apps/web/src/xlsx/parse-impl.ts:1)
  - [`apps/web/src/xlsx/export-impl.ts`](/Users/sachin/Desktop/melp/services/sheet/apps/web/src/xlsx/export-impl.ts:1)

## Status legend

- `todo`: not started
- `doing`: actively in progress
- `blocked`: upstream/library constraint identified
- `done`: implemented, tested, tracker updated

## Work items

| ID | Status | Task | Deliverable | Required Playwright coverage |
| --- | --- | --- | --- | --- |
| ODS-01 | done | Add ODS round-trip lossiness audit | `tests/e2e/ods-lossiness-audit.spec.ts` + `docs/ods-lossiness.md` | Audit spec writes report and locks features already expected to survive |
| ODS-02 | done | Lock formula round-trip behavior | ODS import/export keeps formula text and cached values for supported cases | Cross-sheet + cached-result round-trip coverage |
| ODS-03 | done | Preserve ODS number formats | Map SheetJS `cell.z` into snapshot style data and back out | Probe at least currency, percent, and date formats |
| ODS-04 | done | Preserve row heights and column widths | Parse ODS `content.xml` dimension styles on import and emit `!rows` / `!cols` metadata on export | Dedicated Playwright round-trip coverage plus audit lock |
| ODS-05 | done | Preserve frozen panes via ODS settings.xml | Read/write freeze metadata below the SheetJS worksheet surface | Dedicated Playwright round-trip coverage |
| ODS-06 | blocked | Evaluate minimal style subset beyond number format | Check whether ODS parser exposes font/fill/alignment metadata we can map | Coverage only for fields that survive parser + writer |
| ODS-07 | done | Preserve ODS defined names | Map SheetJS ODS workbook names into `SHEET_DEFINED_NAME_PLUGIN` and export them back out | Import + export Playwright coverage plus audit lock |
| ODS-08 | done | Preserve ODS hyperlinks | Map ODS `cell.l.Target` into inline snapshot hyperlink bodies and export them back out | Import + export Playwright coverage plus audit lock |
| ODS-09 | done | Investigate ODS comments | Determine whether Univer snapshot has a stable comment path we can map from ODS `cell.c` and back | Dedicated Playwright coverage plus audit lock if viable |
| ODS-10 | todo | Triage remaining advanced ODS structures | Decide whether charts and page setup are implementable or out of scope after core fidelity work | Audit entries and scoped follow-up tasks |

## Notes from initial code read

- Current ODS code already preserves formula text in addition to cached values; the stale file header should not be treated as truth.
- Quick local probing against `@e965/xlsx` suggests:
  - ODS formula text is exposed through `cell.f`
  - ODS number formats are exposed through `cell.z`
  - ODS parser may not expose row heights, column widths, or general cell style objects
  - ODS writer emits row and column styles, but parser support appears incomplete

## Execution log

### ODS-01

- Status: `done`
- Completed:
  - added [`tests/e2e/ods-lossiness-audit.spec.ts`](/Users/sachin/Desktop/melp/services/sheet/tests/e2e/ods-lossiness-audit.spec.ts:1)
  - generated [`docs/ods-lossiness.md`](/Users/sachin/Desktop/melp/services/sheet/docs/ods-lossiness.md:1)
  - locked current supported behavior: values, formulas, sheet order/names, merges
- Result snapshot:
  - `8` probes survived
  - `5` probes dropped
  - dropped now: number formats, hyperlinks, comments, defined names

### ODS-03

- Status: `done`
- Completed:
  - mapped ODS `cell.z` into snapshot `styles` on import
  - emitted `IStyleData.n.pattern` back to ODS `cell.z` on export
  - added direct Playwright coverage in [`tests/e2e/ods-csv-tsv.spec.ts`](/Users/sachin/Desktop/melp/services/sheet/tests/e2e/ods-csv-tsv.spec.ts:55)
  - promoted number-format probes in the audit from informational to blocking

### ODS-02

- Status: `done`
- Completed:
  - the audit now locks formula text and cached-result survival for both same-sheet and cross-sheet formulas
  - no codec change was required because the current ODS path already preserved those cases

### ODS-07

- Status: `done`
- Completed:
  - mapped ODS workbook names into the `SHEET_DEFINED_NAME_PLUGIN` snapshot resource on import
  - emitted the same resource back into ODS workbook names on export
  - added dedicated Playwright coverage in [`tests/e2e/ods-defined-names.spec.ts`](/Users/sachin/Desktop/melp/services/sheet/tests/e2e/ods-defined-names.spec.ts:1)
  - promoted the audit defined-name probe from informational to blocking

### ODS-08

- Status: `done`
- Completed:
  - mapped ODS `cell.l.Target` into inline snapshot hyperlink bodies on import
  - emitted inline snapshot hyperlinks back to ODS cell hyperlinks on export
  - added dedicated Playwright coverage in [`tests/e2e/ods-hyperlinks.spec.ts`](/Users/sachin/Desktop/melp/services/sheet/tests/e2e/ods-hyperlinks.spec.ts:1)
  - promoted the audit hyperlink probe from informational to blocking

### Remaining audit gap

- None. The current generated audit at [`docs/ods-lossiness.md`](/Users/sachin/Desktop/melp/services/sheet/docs/ods-lossiness.md:1) is fully green for the features the parser currently exposes.

### ODS-09

- Status: `done`
- Completed:
  - mapped ODS `cell.c` comments into `SHEET_NOTE_PLUGIN` snapshot resources on import
  - emitted `SHEET_NOTE_PLUGIN` note resources back to ODS cell comments on export
  - added dedicated Playwright coverage in [`tests/e2e/ods-comments.spec.ts`](/Users/sachin/Desktop/melp/services/sheet/tests/e2e/ods-comments.spec.ts:1)
  - the audit comment probe is now green

### ODS-04

- Status: `done`
- Completed:
  - implemented ODS `content.xml` parsing for row-height and column-width styles in [`apps/web/src/ods/index.ts`](/Users/sachin/Desktop/melp/services/sheet/apps/web/src/ods/index.ts:1)
  - emitted snapshot `rowData` / `columnData` back through SheetJS `!rows` / `!cols` metadata with the pixel fields the ODS writer expects
  - added dedicated round-trip coverage in [`tests/e2e/ods-dimensions-roundtrip.spec.ts`](/Users/sachin/Desktop/melp/services/sheet/tests/e2e/ods-dimensions-roundtrip.spec.ts:1)
  - promoted dimensions into the main ODS audit and regenerated [`docs/ods-lossiness.md`](/Users/sachin/Desktop/melp/services/sheet/docs/ods-lossiness.md:1)

### ODS-06

- Status: `blocked`
- Evidence:
  - [`tests/e2e/ods-parser-viability.spec.ts`](/Users/sachin/Desktop/melp/services/sheet/tests/e2e/ods-parser-viability.spec.ts:27) shows the current SheetJS ODS read path does not surface general cell style objects for font, fill, or alignment.
- Implication:
  - a broader style subset is blocked on upstream parser support or a lower-level ODS XML pass

### ODS-05

- Status: `done`
- Completed:
  - implemented package-level ODS `settings.xml` parsing and writing for freeze panes in [`apps/web/src/ods/index.ts`](/Users/sachin/Desktop/melp/services/sheet/apps/web/src/ods/index.ts:1)
  - added dedicated round-trip coverage in [`tests/e2e/ods-freeze-roundtrip.spec.ts`](/Users/sachin/Desktop/melp/services/sheet/tests/e2e/ods-freeze-roundtrip.spec.ts:1)
  - kept [`tests/e2e/ods-freeze-viability.spec.ts`](/Users/sachin/Desktop/melp/services/sheet/tests/e2e/ods-freeze-viability.spec.ts:1) as the rationale for why this has to bypass the SheetJS worksheet model
