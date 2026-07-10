/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * ConditionalFormattingDialog — the SDK chrome's built-in Conditional Formatting
 * modal. Sibling of DataValidationDialog / FormatCellsDialog: it reads the active
 * A1 selection off the FUniver facade, gathers a rule via a small form, and
 * applies it through the `@univerjs/sheets-conditional-formatting` facade.
 *
 * Pattern (grounded in
 * `sheets-conditional-formatting/lib/types/facade/*.d.ts`):
 *   worksheet.newConditionalFormattingRule()   // FConditionalFormattingBuilder
 *     .<condition>()                            // → ConditionalFormatHighlightRuleBuilder
 *     .setBackground(fill) / .setFontColor(text)
 *     .setRanges([range.getRange()])
 *     .build()                                  // → IConditionFormattingRule
 *   worksheet.addConditionalFormattingRule(rule)
 *
 * Condition families (each → a real builder method, verified in
 * f-conditional-formatting-builder.d.ts):
 *   - Cell value: greater / greaterEqual / less / lessEqual / equal / notEqual /
 *     between / notBetween → `.whenNumber*` (L461-609)
 *   - Text contains          → `.whenTextContains(text)` (L630)
 *   - Top N / Bottom N       → `.setRank({ isBottom, isPercent, value })` (L227)
 *   - Color scale            → `.setColorScale(config)` (L853) — its own visual,
 *     no fill/text format
 *
 * The "Clear rules" action calls `worksheet.clearConditionalFormatRules()`
 * (f-worksheet.d.ts L172) — the whole-sheet clear the facade exposes (there's no
 * range-scoped clear on FWorksheet; FRange.clearConditionalFormatRules exists but
 * is deprecated — see limitations).
 *
 * Mounted by `<DialogHost>` when `openDialog('conditional-formatting')` is called
 * and no host override is registered.
 */

import { useMemo, useState, type CSSProperties } from 'react';
// Side-effect import: installs `FWorksheet.newConditionalFormattingRule /
// addConditionalFormattingRule / clearConditionalFormatRules` (and the FRange
// mixin) on the facade prototypes AND augments the `@univerjs/sheets/facade`
// FWorksheet/FRange types. The `cf` plugin group registers the plugin but does
// NOT import this facade module, so without it the builder calls below are
// undefined at runtime. Mirrors DataValidationDialog's `/facade` side-effect.
import '@univerjs/sheets-conditional-formatting/facade';
import type { DialogComponentProps } from './extensions';
import type { CasualSheetsAPI } from '../sheets/api';
import { Dialog } from './Dialog';
import {
  DIALOG_BTN_PRIMARY_STYLE,
  DIALOG_BTN_SECONDARY_STYLE,
  DIALOG_FIELD_STYLE,
  DIALOG_INPUT_STYLE,
  DIALOG_LABEL_STYLE,
} from './dialog-styles';

/** Rule families the dialog can build. */
type RuleType = 'cellValue' | 'textContains' | 'topN' | 'bottomN' | 'colorScale';

/** Operators for the cell-value family. */
type Operator =
  | 'greater'
  | 'greaterEqual'
  | 'less'
  | 'lessEqual'
  | 'equal'
  | 'notEqual'
  | 'between'
  | 'notBetween';

const RULE_TYPE_OPTIONS: Array<{ value: RuleType; label: string }> = [
  { value: 'cellValue', label: 'Cell value' },
  { value: 'textContains', label: 'Text contains' },
  { value: 'topN', label: 'Top N' },
  { value: 'bottomN', label: 'Bottom N' },
  { value: 'colorScale', label: 'Color scale' },
];

const OPERATOR_OPTIONS: Array<{ value: Operator; label: string }> = [
  { value: 'greater', label: 'Greater than' },
  { value: 'greaterEqual', label: 'Greater than or equal to' },
  { value: 'less', label: 'Less than' },
  { value: 'lessEqual', label: 'Less than or equal to' },
  { value: 'equal', label: 'Equal to' },
  { value: 'notEqual', label: 'Not equal to' },
  { value: 'between', label: 'Between' },
  { value: 'notBetween', label: 'Not between' },
];

/** True when the operator needs a second operand. */
function isRangeOperator(op: Operator): boolean {
  return op === 'between' || op === 'notBetween';
}

interface DialogState {
  ruleType: RuleType;
  operator: Operator;
  /** Cell-value first operand (raw string; parsed on apply). */
  operand1: string;
  /** Second operand, used only for between / notBetween. */
  operand2: string;
  /** Substring for the text-contains family. */
  text: string;
  /** N for top/bottom N. */
  rankN: string;
  /** Interpret rankN as a percent instead of a count. */
  rankPercent: boolean;
  /** Highlight fill color. */
  fillColor: string;
  /** Highlight text color. */
  textColor: string;
  /** Color-scale endpoints (2-color: min → max). */
  scaleMinColor: string;
  scaleMaxColor: string;
}

const INITIAL_STATE: DialogState = {
  ruleType: 'cellValue',
  operator: 'greater',
  operand1: '',
  operand2: '',
  text: '',
  rankN: '10',
  rankPercent: false,
  fillColor: '#fce8b2',
  textColor: '#7f6000',
  scaleMinColor: '#ffffff',
  scaleMaxColor: '#57bb8a',
};

/** The active FRange, or null when there is no selection. */
function activeRange(api: CasualSheetsAPI) {
  return api.univer.getActiveWorkbook()?.getActiveSheet()?.getActiveRange() ?? null;
}

/** The active FWorksheet, or null. */
function activeSheet(api: CasualSheetsAPI) {
  return api.univer.getActiveWorkbook()?.getActiveSheet() ?? null;
}

/** True when the rule family carries a fill/text highlight format. */
function usesHighlight(type: RuleType): boolean {
  return type !== 'colorScale';
}

/**
 * Build a conditional-formatting rule from the form via the facade builder and
 * add it to the active sheet, scoped to the active range. Returns false when
 * there's no selection or the inputs are unusable.
 */
function applyRule(api: CasualSheetsAPI, s: DialogState): boolean {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return false;

  // The conditional-formatting facade contributes these to FWorksheet at
  // runtime; typed loosely here so the SDK doesn't hard-depend on the ambient
  // FWorksheet augmentation at this call site.
  const cfSheet = sheet as unknown as {
    newConditionalFormattingRule: () => CfBuilder;
    addConditionalFormattingRule: (rule: unknown) => unknown;
  };
  const iRange = (range as unknown as { getRange: () => unknown }).getRange();

  const n1 = Number(s.operand1);
  const n2 = Number(s.operand2);
  const rankN = Number(s.rankN);

  let builder = cfSheet.newConditionalFormattingRule();

  switch (s.ruleType) {
    case 'cellValue': {
      if (!Number.isFinite(n1)) return false;
      switch (s.operator) {
        case 'greater':
          builder = builder.whenNumberGreaterThan(n1);
          break;
        case 'greaterEqual':
          builder = builder.whenNumberGreaterThanOrEqualTo(n1);
          break;
        case 'less':
          builder = builder.whenNumberLessThan(n1);
          break;
        case 'lessEqual':
          builder = builder.whenNumberLessThanOrEqualTo(n1);
          break;
        case 'equal':
          builder = builder.whenNumberEqualTo(n1);
          break;
        case 'notEqual':
          builder = builder.whenNumberNotEqualTo(n1);
          break;
        case 'between':
          if (!Number.isFinite(n2)) return false;
          builder = builder.whenNumberBetween(n1, n2);
          break;
        case 'notBetween':
          if (!Number.isFinite(n2)) return false;
          builder = builder.whenNumberNotBetween(n1, n2);
          break;
      }
      break;
    }
    case 'textContains': {
      const text = s.text.trim();
      if (text.length === 0) return false;
      builder = builder.whenTextContains(text);
      break;
    }
    case 'topN':
    case 'bottomN': {
      if (!Number.isFinite(rankN) || rankN <= 0) return false;
      builder = builder.setRank({
        isBottom: s.ruleType === 'bottomN',
        isPercent: s.rankPercent,
        value: rankN,
      });
      break;
    }
    case 'colorScale': {
      // 2-color scale: min endpoint → max endpoint. Value types 'min'/'max'
      // are the CFValueType string literals (base/const.d.ts CFValueType).
      builder = builder.setColorScale([
        { index: 0, color: s.scaleMinColor, value: { type: 'min' } },
        { index: 1, color: s.scaleMaxColor, value: { type: 'max' } },
      ]);
      break;
    }
  }

  if (usesHighlight(s.ruleType)) {
    builder = builder.setBackground(s.fillColor);
    builder = builder.setFontColor(s.textColor);
  }

  const rule = builder.setRanges([iRange]).build();
  cfSheet.addConditionalFormattingRule(rule);
  return true;
}

/**
 * The subset of the conditional-formatting builder chain we use, typed here so
 * the call site is checked without importing the facade's private builder
 * classes. Every method is verified in
 * `sheets-conditional-formatting/lib/types/facade/f-conditional-formatting-builder.d.ts`.
 */
interface CfBuilder {
  whenNumberGreaterThan: (v: number) => CfBuilder;
  whenNumberGreaterThanOrEqualTo: (v: number) => CfBuilder;
  whenNumberLessThan: (v: number) => CfBuilder;
  whenNumberLessThanOrEqualTo: (v: number) => CfBuilder;
  whenNumberEqualTo: (v: number) => CfBuilder;
  whenNumberNotEqualTo: (v: number) => CfBuilder;
  whenNumberBetween: (a: number, b: number) => CfBuilder;
  whenNumberNotBetween: (a: number, b: number) => CfBuilder;
  whenTextContains: (t: string) => CfBuilder;
  setRank: (config: { isBottom: boolean; isPercent: boolean; value: number }) => CfBuilder;
  setColorScale: (
    config: Array<{ index: number; color: string; value: { type: string; value?: number } }>,
  ) => CfBuilder;
  setBackground: (color?: string) => CfBuilder;
  setFontColor: (color?: string) => CfBuilder;
  setRanges: (ranges: unknown[]) => CfBuilder;
  build: () => unknown;
}

const RANGE_NOTE_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-muted, #605e5c)',
  marginBottom: 12,
};

const CHECK_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 6,
  marginBottom: 8,
  cursor: 'pointer',
};

const COLOR_INPUT_STYLE: CSSProperties = {
  width: 48,
  height: 30,
  padding: 2,
  border: '1px solid var(--cs-chrome-border, #cdd3db)',
  borderRadius: 6,
  background: 'var(--cs-chrome-input-bg, #fff)',
  cursor: 'pointer',
};

const TWO_COL_STYLE: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '1fr 1fr',
  gap: 8,
};

export function ConditionalFormattingDialog({ api, onClose }: DialogComponentProps) {
  const [state, setState] = useState<DialogState>(INITIAL_STATE);

  // Read the selection once for the header hint. `getA1Notation` off the live
  // FRange (verified in @univerjs/sheets/facade f-range.d.ts) gives the
  // user-facing A1 label, e.g. "A1:B2".
  const rangeLabel = useMemo(() => {
    const fRange = activeRange(api) as unknown as { getA1Notation?: () => string } | null;
    return fRange?.getA1Notation?.() ?? null;
  }, [api]);

  const hasSelection = activeRange(api) !== null;

  const update = <K extends keyof DialogState>(key: K, value: DialogState[K]) =>
    setState((prev) => ({ ...prev, [key]: value }));

  const apply = () => {
    if (applyRule(api, state)) onClose();
  };

  const clearRules = () => {
    const sheet = activeSheet(api) as unknown as {
      clearConditionalFormatRules?: () => unknown;
    } | null;
    sheet?.clearConditionalFormatRules?.();
    onClose();
  };

  const showOperator = state.ruleType === 'cellValue';
  const showSecondOperand = showOperator && isRangeOperator(state.operator);
  const isRank = state.ruleType === 'topN' || state.ruleType === 'bottomN';

  return (
    <Dialog
      title="Conditional formatting"
      onClose={onClose}
      width={440}
      data-testid="cs-conditional-formatting-dialog"
      footer={
        <>
          <button
            type="button"
            style={DIALOG_BTN_SECONDARY_STYLE}
            data-testid="cs-conditional-formatting-clear"
            onClick={clearRules}
          >
            Clear rules
          </button>
          <span style={{ flex: 1 }} />
          <button type="button" style={DIALOG_BTN_SECONDARY_STYLE} onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            style={DIALOG_BTN_PRIMARY_STYLE}
            data-testid="cs-conditional-formatting-apply"
            disabled={!hasSelection}
            onClick={apply}
          >
            Apply
          </button>
        </>
      }
    >
      {hasSelection ? (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-conditional-formatting-range">
          Applies to <strong>{rangeLabel ?? 'the current selection'}</strong>
        </div>
      ) : (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-conditional-formatting-no-selection">
          Select one or more cells first, then reopen this dialog.
        </div>
      )}

      <label style={DIALOG_FIELD_STYLE}>
        <span style={DIALOG_LABEL_STYLE}>Condition type</span>
        <select
          style={DIALOG_INPUT_STYLE}
          data-testid="cs-conditional-formatting-type"
          value={state.ruleType}
          onChange={(e) => update('ruleType', e.target.value as RuleType)}
        >
          {RULE_TYPE_OPTIONS.map((opt) => (
            <option key={opt.value} value={opt.value}>
              {opt.label}
            </option>
          ))}
        </select>
      </label>

      {showOperator && (
        <>
          <label style={DIALOG_FIELD_STYLE}>
            <span style={DIALOG_LABEL_STYLE}>Condition</span>
            <select
              style={DIALOG_INPUT_STYLE}
              data-testid="cs-conditional-formatting-operator"
              value={state.operator}
              onChange={(e) => update('operator', e.target.value as Operator)}
            >
              {OPERATOR_OPTIONS.map((opt) => (
                <option key={opt.value} value={opt.value}>
                  {opt.label}
                </option>
              ))}
            </select>
          </label>

          <div style={showSecondOperand ? TWO_COL_STYLE : undefined}>
            <label style={DIALOG_FIELD_STYLE}>
              <span style={DIALOG_LABEL_STYLE}>{showSecondOperand ? 'Minimum' : 'Value'}</span>
              <input
                style={DIALOG_INPUT_STYLE}
                data-testid="cs-conditional-formatting-operand1"
                type="number"
                value={state.operand1}
                onChange={(e) => update('operand1', e.target.value)}
              />
            </label>
            {showSecondOperand && (
              <label style={DIALOG_FIELD_STYLE}>
                <span style={DIALOG_LABEL_STYLE}>Maximum</span>
                <input
                  style={DIALOG_INPUT_STYLE}
                  data-testid="cs-conditional-formatting-operand2"
                  type="number"
                  value={state.operand2}
                  onChange={(e) => update('operand2', e.target.value)}
                />
              </label>
            )}
          </div>
        </>
      )}

      {state.ruleType === 'textContains' && (
        <label style={DIALOG_FIELD_STYLE}>
          <span style={DIALOG_LABEL_STYLE}>Text contains</span>
          <input
            style={DIALOG_INPUT_STYLE}
            data-testid="cs-conditional-formatting-text"
            value={state.text}
            placeholder="e.g. urgent"
            onChange={(e) => update('text', e.target.value)}
          />
        </label>
      )}

      {isRank && (
        <>
          <label style={DIALOG_FIELD_STYLE}>
            <span style={DIALOG_LABEL_STYLE}>
              {state.ruleType === 'topN' ? 'Top' : 'Bottom'}{' '}
              {state.rankPercent ? 'percent' : 'count'}
            </span>
            <input
              style={DIALOG_INPUT_STYLE}
              data-testid="cs-conditional-formatting-rank-n"
              type="number"
              min={1}
              value={state.rankN}
              onChange={(e) => update('rankN', e.target.value)}
            />
          </label>
          <label style={CHECK_STYLE}>
            <input
              type="checkbox"
              data-testid="cs-conditional-formatting-rank-percent"
              checked={state.rankPercent}
              onChange={(e) => update('rankPercent', e.target.checked)}
            />
            <span>Interpret as percent</span>
          </label>
        </>
      )}

      {state.ruleType === 'colorScale' ? (
        <div style={TWO_COL_STYLE}>
          <label style={DIALOG_FIELD_STYLE}>
            <span style={DIALOG_LABEL_STYLE}>Min color</span>
            <input
              style={COLOR_INPUT_STYLE}
              data-testid="cs-conditional-formatting-scale-min"
              type="color"
              value={state.scaleMinColor}
              onChange={(e) => update('scaleMinColor', e.target.value)}
            />
          </label>
          <label style={DIALOG_FIELD_STYLE}>
            <span style={DIALOG_LABEL_STYLE}>Max color</span>
            <input
              style={COLOR_INPUT_STYLE}
              data-testid="cs-conditional-formatting-scale-max"
              type="color"
              value={state.scaleMaxColor}
              onChange={(e) => update('scaleMaxColor', e.target.value)}
            />
          </label>
        </div>
      ) : (
        <div style={TWO_COL_STYLE}>
          <label style={DIALOG_FIELD_STYLE}>
            <span style={DIALOG_LABEL_STYLE}>Fill color</span>
            <input
              style={COLOR_INPUT_STYLE}
              data-testid="cs-conditional-formatting-fill-color"
              type="color"
              value={state.fillColor}
              onChange={(e) => update('fillColor', e.target.value)}
            />
          </label>
          <label style={DIALOG_FIELD_STYLE}>
            <span style={DIALOG_LABEL_STYLE}>Text color</span>
            <input
              style={COLOR_INPUT_STYLE}
              data-testid="cs-conditional-formatting-text-color"
              type="color"
              value={state.textColor}
              onChange={(e) => update('textColor', e.target.value)}
            />
          </label>
        </div>
      )}
    </Dialog>
  );
}
