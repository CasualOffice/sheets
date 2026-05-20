# ODS fidelity milestone plan

This breaks ODS fidelity follow-up work into an umbrella milestone under issue `#35`.
It is intentionally ordered to establish observability first, then lock in the highest-value
round-trip metadata before taking on broader formatting work.

## Current state from code

Verified in [`apps/web/src/ods/index.ts`](/Users/sachin/Desktop/melp/services/sheet/apps/web/src/ods/index.ts:1) and [`tests/e2e/ods-csv-tsv.spec.ts`](/Users/sachin/Desktop/melp/services/sheet/tests/e2e/ods-csv-tsv.spec.ts:1):

- Preserved now:
  - cell values
  - basic formula text plus cached results
  - sheet order and names
  - merges
- Still missing or unverified:
  - a generated ODS lossiness report equivalent to [`docs/xlsx-lossiness.md`](/Users/sachin/Desktop/melp/services/sheet/docs/xlsx-lossiness.md:1)
  - explicit formula audit coverage for cross-sheet and edge-case formulas
  - column widths and row heights
  - frozen panes and related sheet view metadata
  - styles
  - charts and richer document structures

## Suggested implementation order

1. Build the ODS lossiness audit harness and generated report.
2. Use the audit to lock current formula behavior and close any formula gaps it exposes.
3. Round-trip row heights and column widths through SheetJS metadata.
4. Add frozen-pane and other sheet-property mapping where the ODS model exposes them.
5. Add a narrow style subset: font, fill, alignment, number format.
6. Re-evaluate charts, tables, defined names, comments, hyperlinks, and other advanced ODS structures based on audit findings and product value.

## Suggested issues

### 1. ODS: add round-trip lossiness audit and generated report

- Type: foundational
- Effort: 2 to 3 days
- Depends on: none
- Why first: it turns fidelity work from anecdotal behavior into a reproducible contract.

Scope:

- Add a Playwright or Node-driven audit similar to [`tests/e2e/xlsx-lossiness-audit.spec.ts`](/Users/sachin/Desktop/melp/services/sheet/tests/e2e/xlsx-lossiness-audit.spec.ts:1).
- Generate `docs/ods-lossiness.md` on every run.
- Probe at least:
  - values
  - formulas and cached results
  - sheet order and names
  - merges
  - column widths and row heights
  - frozen panes
  - style subset candidates
  - charts or chart loss detection if practical

Acceptance criteria:

- A single audit command produces a checked-in markdown report.
- The test hard-fails on features already expected to survive today.
- The report clearly separates preserved, partial, and dropped ODS features.

Suggested GitHub issue title:

`ODS: add round-trip lossiness audit and generated report`

Suggested body:

```md
## Summary

Add an ODS-specific round-trip lossiness audit, equivalent to the existing xlsx audit, so we can measure and prioritize fidelity work from a generated report instead of ad hoc manual checks.

## Scope

- add an ODS audit harness under `tests/e2e/`
- generate `docs/ods-lossiness.md`
- probe values, formulas, cached results, sheet names/order, merges, dimensions, frozen panes, styles, and chart loss where practical
- lock currently supported ODS behavior so regressions fail loudly

## Acceptance criteria

- running the audit writes `docs/ods-lossiness.md`
- the report marks each probe as preserved, partial, or dropped
- currently supported behavior is asserted in the test, not just documented

## Notes

`apps/web/src/ods/index.ts` already preserves basic formula text, so the audit should treat formulas as a feature to verify and expand, not as greenfield support.
```

### 2. ODS: lock formula round-trip behavior and close audit gaps

- Type: correctness
- Effort: 1 to 2 days
- Depends on: issue 1
- Why second: formulas are high-value and already partially implemented.

Scope:

- Confirm current formula import/export coverage in [`apps/web/src/ods/index.ts`](/Users/sachin/Desktop/melp/services/sheet/apps/web/src/ods/index.ts:37).
- Add targeted tests for:
  - basic formulas
  - cross-sheet formulas
  - cached results surviving export
  - formula cells with no cached value
- Close whatever gaps the audit reveals instead of broad speculative work.

Acceptance criteria:

- Formula probes in the ODS audit pass for the supported cases.
- Formula text round-trips without being downgraded to value-only cells.
- Cached results still survive for consumers that depend on them.

Suggested GitHub issue title:

`ODS: lock formula round-trip behavior and close audit gaps`

Suggested body:

```md
## Summary

ODS formula text is already partially preserved in `apps/web/src/ods/index.ts`, but we do not yet have deliberate audit coverage around its boundaries. Lock that behavior in and fix any gaps the audit exposes.

## Scope

- add focused ODS formula probes
- verify basic formulas, cross-sheet formulas, and cached results
- ensure formula cells are exported as formulas, not value-only cells
- fix any gaps revealed by the new audit

## Acceptance criteria

- supported formula cases are covered by automated tests
- formula text survives a full ODS round-trip
- cached values are preserved where available
```

### 3. ODS: round-trip row heights and column widths

- Type: fidelity
- Effort: 1 to 2 days
- Depends on: issue 1
- Why third: dimensions are visible to users and map cleanly to SheetJS sheet metadata.

Scope:

- Read and write SheetJS `!rows` and `!cols` metadata in the ODS path.
- Map those dimensions into Univer snapshot shape and back.
- Add regression coverage in the new audit plus a focused spec if needed.

Acceptance criteria:

- Explicit row heights survive ODS import and re-export.
- Explicit column widths survive ODS import and re-export.
- Default sheet sizing behavior is unchanged for sheets with no custom metadata.

Suggested GitHub issue title:

`ODS: round-trip row heights and column widths`

Suggested body:

```md
## Summary

Preserve explicit row heights and column widths in the ODS pipeline by mapping SheetJS dimension metadata into our workbook snapshot model and back out on export.

## Scope

- read ODS row and column metadata from SheetJS
- map dimensions into the workbook snapshot
- write the same metadata back during ODS export
- cover the behavior in the ODS audit

## Acceptance criteria

- custom row heights survive round-trip
- custom column widths survive round-trip
- sheets without custom dimensions behave exactly as before
```

### 4. ODS: preserve frozen panes and exposed sheet-view metadata

- Type: fidelity
- Effort: 1 to 2 days
- Depends on: issue 1
- Why fourth: likely lower surface area than styles and strongly user-visible.

Scope:

- Investigate what SheetJS exposes for ODS sheet views.
- Preserve frozen panes if available.
- Include any low-risk adjacent sheet properties that map cleanly through the same path.

Acceptance criteria:

- Freeze rows and freeze columns survive ODS round-trip when represented in the source file.
- Unsupported view metadata remains explicitly documented in the audit report.

Suggested GitHub issue title:

`ODS: preserve frozen panes and exposed sheet-view metadata`

Suggested body:

```md
## Summary

Add ODS sheet-view fidelity for frozen panes and any adjacent sheet properties that SheetJS exposes in a stable way.

## Scope

- inspect SheetJS ODS sheet-view metadata
- map frozen rows and columns into the snapshot model
- export the same metadata back out to ODS
- document anything still unsupported in the audit report

## Acceptance criteria

- frozen panes survive a full ODS round-trip
- unsupported view metadata is explicitly called out, not silently assumed
```

### 5. ODS: preserve a minimal style subset

- Type: fidelity
- Effort: 3 to 5 days
- Depends on: issues 1, 2, 3, and 4
- Why fifth: styles have the broadest mapping surface and should start narrow.

Scope:

- Preserve only:
  - font
  - fill
  - alignment
  - number format
- Reuse the xlsx audit style categories as the reporting model where possible.
- Avoid broad style ambitions like borders, conditional formatting, or full theme fidelity in this pass unless they fall out almost for free.

Acceptance criteria:

- The selected style subset survives ODS round-trip in the audit.
- Non-goal style classes remain documented as unsupported.
- The implementation does not regress plain value/formula fidelity.

Suggested GitHub issue title:

`ODS: preserve a minimal style subset (font, fill, alignment, number format)`

Suggested body:

```md
## Summary

Add a deliberate first-pass style mapping for ODS. Keep the scope narrow to the style fields users will notice most often and that are most likely to map cleanly through SheetJS and Univer.

## Scope

- preserve font
- preserve fill
- preserve alignment
- preserve number format
- add audit probes for each style class

## Non-goals

- full border fidelity
- conditional formatting
- theme fidelity
- charts and drawing-layer formatting

## Acceptance criteria

- the selected style subset survives ODS round-trip
- unsupported style classes remain explicit in the audit report
- no regressions to existing ODS value, formula, sheet, or merge behavior
```

### 6. ODS: evaluate advanced structures after the core fidelity pass

- Type: follow-up decision
- Effort: 1 day investigation, implementation TBD
- Depends on: issues 1 through 5
- Why last: advanced structures should be prioritized with data, not intuition.

Scope:

- Review audit findings and product demand for:
  - charts
  - tables
  - defined names
  - comments
  - hyperlinks
  - page setup
- Decide which of these deserve implementation issues versus explicit non-goals.

Acceptance criteria:

- A short follow-up design note or issue comment records the keep/defer decisions.
- Any chosen feature gets its own implementation issue with a bounded scope.

Suggested GitHub issue title:

`ODS: evaluate advanced structures after the core fidelity pass`

Suggested body:

```md
## Summary

After the core ODS fidelity work lands, use the audit results to decide which advanced structures are worth implementing next instead of guessing up front.

## Review candidates

- charts
- tables
- defined names
- comments
- hyperlinks
- page setup and print metadata

## Acceptance criteria

- we have a written keep/defer decision for each candidate
- any approved follow-up work is split into bounded implementation issues
```

## Milestone recommendation

Suggested milestone name:

`ODS fidelity`

Suggested milestone description:

```md
Improve ODS round-trip fidelity in measured steps: add an audit first, then close the highest-value metadata and formatting gaps with regression coverage.
```
