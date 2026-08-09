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
 * Pure command-policy primitives. This module deliberately has no Univer value
 * imports: the policy can be unit-tested under the SDK's bare node:test runner,
 * while `<CasualSheets>` supplies the actual `beforeCommandExecuted` wiring.
 */

import { clonePlainCommandData } from './command-context';

/** A local command about to enter Univer's command bus. */
export interface BeforeCommandContext {
  /** Univer command / operation / mutation id. */
  readonly id: string;
  /** Parameters passed to the command. Treat as read-only. */
  readonly params?: object;
}

/**
 * Synchronous host policy. Return `false` to veto the command; `true` or
 * `undefined` lets it run. The callback sees every local engine command,
 * including commands reached through shortcuts, context menus, facades, the
 * raw `api.univer` escape hatch, and nested command/mutation execution.
 */
export type BeforeCommandPolicy = (command: BeforeCommandContext) => boolean | void;

/** The execution metadata we need from Univer's `IExecutionOptions`. */
export interface CommandPolicyExecutionOptions {
  fromCollab?: boolean;
  fromChangeset?: boolean;
  onlyLocal?: boolean;
  fromFormula?: boolean;
}

/** Inputs are getters so reactive props take effect without remounting Univer. */
export interface CommandPolicySource {
  features: () => Readonly<Record<string, boolean>> | undefined;
  beforeCommand: () => BeforeCommandPolicy | undefined;
}

type FeatureMatcher = (id: string, params: object | undefined) => boolean;

const exact = (...ids: string[]): FeatureMatcher => {
  const allowed = new Set(ids);
  return (id) => allowed.has(id);
};

const contains = (...parts: string[]): FeatureMatcher => {
  const normalized = parts.map((part) => part.toLowerCase());
  return (id) => {
    const candidate = id.toLowerCase();
    return normalized.some((part) => candidate.includes(part));
  };
};

const STYLE_FEATURES: Readonly<Record<string, string>> = {
  ff: 'font',
  fs: 'font',
  bl: 'font-style',
  it: 'font-style',
  ul: 'font-style',
  bbl: 'font-style',
  st: 'font-style',
  ol: 'font-style',
  va: 'font-style',
  cl: 'color',
  bg: 'color',
  bd: 'borders',
  ht: 'alignment',
  vt: 'alignment',
  tb: 'alignment',
  tr: 'alignment',
  td: 'alignment',
  pd: 'alignment',
  n: 'number',
};

const FORMATTING_FEATURES = [
  'font',
  'font-style',
  'color',
  'borders',
  'alignment',
  'number',
] as const;

/**
 * Command-backed feature flags already exposed by the SDK chrome. UI-only
 * flags (`file`, `help`, `branding`) and SDK-resource dialogs (charts, pivots,
 * sparklines) intentionally have no entry because they do not dispatch an
 * engine command. Unknown feature keys remain forward-compatible no-ops.
 */
const FEATURE_MATCHERS: Readonly<Record<string, readonly FeatureMatcher[]>> = {
  history: [exact('univer.command.undo', 'univer.command.redo')],
  clipboard: [
    exact('univer.command.copy', 'univer.command.cut'),
    // Paste has several plugin-defined variants (formula, value, formatting,
    // image, optional-paste). They are one clipboard capability and share the
    // stable "paste" token even when a future plugin adds another variant.
    contains('paste'),
  ],
  'format-painter': [contains('format-painter')],
  font: [contains('set-range-font', 'set-font-family', 'set-font-size')],
  'font-style': [
    contains(
      'set-range-bold',
      'set-range-italic',
      'set-range-underline',
      'set-range-stroke',
      'set-range-subscript',
      'set-range-superscript',
      'set-bold',
      'set-italic',
      'set-underline',
      'set-stroke',
      'set-overline',
    ),
  ],
  color: [contains('text-color', 'background-color')],
  borders: [contains('set-border')],
  alignment: [contains('text-align', 'text-wrap', 'text-rotation')],
  merge: [contains('worksheet-merge')],
  number: [contains('.numfmt.', 'mutation.set.numfmt', 'mutation.remove.numfmt')],
  'clear-format': [
    exact('sheet.command.clear-selection-format', 'sheet.command.clear-selection-all'),
  ],
  tables: [
    contains('.add-table', '.delete-table', 'set-table', 'sheet-table', '.table-', 'table-theme'),
  ],
  conditionalFormatting: [
    contains('conditional-rule', 'conditional-formatting', 'conditional.formatting'),
  ],
  filter: [contains('filter')],
  dataValidation: [contains('data-validation', 'dataValidation')],
};

function asRecord(value: unknown): Record<string, unknown> | undefined {
  return value !== null && typeof value === 'object'
    ? (value as Record<string, unknown>)
    : undefined;
}

interface StyleFeatureInspection {
  readonly features: Set<string>;
  /** A style-registry id cannot be decomposed without workbook state. */
  readonly opaqueStyle: boolean;
  /** `s: null` or a null cell clears every style property. */
  readonly clearsStyle: boolean;
}

// Style payloads may be sparse matrices covering an entire worksheet. The
// policy runs synchronously on the browser main thread, so inspect only a
// bounded prefix and fail closed if a restricted formatter cannot be proven
// absent. This is a work budget (rows + cells + style keys), not a file limit.
const MAX_STYLE_INSPECTION_ENTRIES = 10_000;

function styleFeaturesFromParams(
  params: object | undefined,
  disabledFormatting: ReadonlySet<string>,
): StyleFeatureInspection {
  const features = new Set<string>();
  const record = asRecord(params);
  if (!record) return { features, opaqueStyle: false, clearsStyle: false };
  let opaqueStyle = false;
  let clearsStyle = false;
  let remainingEntries = MAX_STYLE_INSPECTION_ENTRIES;
  const consumeEntry = () => {
    remainingEntries -= 1;
    if (remainingEntries >= 0) return true;
    opaqueStyle = true;
    return false;
  };
  const inspection = (): StyleFeatureInspection => ({
    features,
    opaqueStyle,
    clearsStyle,
  });

  // Direct FRange style calls enter as `sheet.command.set-style` with
  // `{ style: { type } }`.
  const directStyle = record.style;
  const styleType = asRecord(directStyle)?.type;
  if (typeof styleType === 'string') {
    const feature = STYLE_FEATURES[styleType];
    if (!feature) {
      opaqueStyle = true;
      return inspection();
    }
    features.add(feature);
    if (disabledFormatting.has(feature)) return inspection();
  } else if (asRecord(directStyle)) {
    // An object-shaped style without a recognized `type` is an unknown facade
    // or plugin payload. Do not guess which formatter it can mutate.
    opaqueStyle = true;
    return inspection();
  } else if (directStyle !== undefined && !asRecord(directStyle)) {
    opaqueStyle = directStyle !== null;
    clearsStyle = directStyle === null;
    return inspection();
  }

  // Recorded/replayed style mutations enter as
  // `sheet.mutation.set-range-values` with a sparse `cellValue` matrix. Inspect
  // only the style keys; values/formulas remain unrelated to formatting flags.
  const matrix = asRecord(record.cellValue) ?? asRecord(record.value);
  if (!matrix) return inspection();
  for (const rowKey in matrix) {
    if (!Object.prototype.hasOwnProperty.call(matrix, rowKey)) continue;
    if (!consumeEntry()) return inspection();
    const row = matrix[rowKey];
    const cells = asRecord(row);
    if (!cells) continue;
    for (const cellKey in cells) {
      if (!Object.prototype.hasOwnProperty.call(cells, cellKey)) continue;
      if (!consumeEntry()) return inspection();
      const cell = cells[cellKey];
      if (cell === null) {
        clearsStyle = true;
        return inspection();
      }
      const styleValue = asRecord(cell)?.s;
      if (styleValue === null) {
        clearsStyle = true;
        return inspection();
      }
      const style = asRecord(styleValue);
      if (!style) {
        if (styleValue !== undefined) {
          opaqueStyle = true;
          return inspection();
        }
        continue;
      }
      for (const key in style) {
        if (!Object.prototype.hasOwnProperty.call(style, key)) continue;
        if (!consumeEntry()) return inspection();
        const feature = STYLE_FEATURES[key];
        if (!feature) {
          opaqueStyle = true;
          return inspection();
        }
        features.add(feature);
        if (disabledFormatting.has(feature)) return inspection();
      }
    }
  }
  return inspection();
}

function disabledFormattingFeature(
  features: Readonly<Record<string, boolean>>,
  except: ReadonlySet<string> = new Set(),
): string | undefined {
  return FORMATTING_FEATURES.find((feature) => !except.has(feature) && features[feature] === false);
}

function firstDisabledFeature(
  features: Readonly<Record<string, boolean>>,
  candidates: readonly string[],
): string | undefined {
  return candidates.find((feature) => features[feature] === false);
}

/**
 * Some Univer commands calculate and execute several mutations sequentially.
 * A nested veto is too late: earlier mutations have already changed the sheet
 * and Univer does not add an undo record for the failed compound action. Block
 * these commands before their handler starts whenever they could touch a
 * disabled downstream capability.
 */
function disabledCompoundFeature(
  command: BeforeCommandContext,
  features: Readonly<Record<string, boolean>>,
): string | undefined {
  const id = command.id.toLowerCase();

  // Row/column/range commands execute their core structural mutation before
  // ref-range interceptors append merge/table/CF/filter/DV maintenance. If a
  // child policy vetoes one of those maintenance mutations, the structure has
  // already changed and Univer deliberately skips the undo record. Without
  // workbook state here, the only atomic choice is to reject these roots when
  // any affected capability is disabled.
  if (id === 'sheet.command.reorder-range') {
    // Reorder/sort has no merge/table/filter interceptor in Univer 0.25.
    // Conditional-formatting formula ranges and data validation do append
    // durable ref-range mutations after the core reorder.
    return firstDisabledFeature(features, ['conditionalFormatting', 'dataValidation']);
  }
  const rowOrColumnShapeRoot =
    /^sheet\.command\.(?:insert-(?:row|col)(?:-|$)|remove-(?:row|col)(?:-|$))/.test(id);
  if (rowOrColumnShapeRoot) {
    return firstDisabledFeature(features, [
      'merge',
      'tables',
      'conditionalFormatting',
      'filter',
      'dataValidation',
    ]);
  }
  if (/^sheet\.command\.move-(?:rows|cols|range)$/.test(id)) {
    return firstDisabledFeature(features, [
      'merge',
      'conditionalFormatting',
      'filter',
      'dataValidation',
    ]);
  }
  if (
    /^sheet\.command\.(?:insert-range-move-(?:down|right)(?:-|$)|delete-range-move-(?:up|left)(?:-|$))/.test(
      id,
    )
  ) {
    return firstDisabledFeature(features, ['merge', 'conditionalFormatting', 'dataValidation']);
  }

  // Sheet removal runs the base mutation before plugin cleanup mutations.
  if (id === 'sheet.command.remove-sheet') {
    return firstDisabledFeature(features, [
      'tables',
      'conditionalFormatting',
      'filter',
      'dataValidation',
    ]);
  }

  if (id === 'sheet.command.copy-sheet') {
    // CopySheet clones cell styles and merges directly into the inserted sheet;
    // large sheets then stream ignored-result SetRangeValues chunks. Preflight
    // every copied durable capability before the new sheet can be partial.
    return (
      disabledFormattingFeature(features) ??
      firstDisabledFeature(features, [
        'merge',
        'tables',
        'conditionalFormatting',
        'filter',
        'dataValidation',
      ])
    );
  }

  if (id === 'sheet.command.remove-worksheet-merge') {
    // Unmerge removes the merge first, then copies the top-left cell's full
    // style across the newly independent cells. The style payload depends on
    // workbook state, so every restricted formatter must be preflighted.
    return features.merge === false ? 'merge' : disabledFormattingFeature(features);
  }

  if (id === 'univer.command.undo' || id === 'univer.command.redo') {
    // History entries can contain several mutations and sequenceExecute stops
    // on the first veto. Once restrictions change reactively, we cannot inspect
    // the pending history item before earlier mutations run, so disable history
    // conservatively whenever a capability represented in history is disabled.
    return (
      disabledFormattingFeature(features) ??
      firstDisabledFeature(features, [
        'clipboard',
        'format-painter',
        'clear-format',
        'merge',
        'tables',
        'conditionalFormatting',
        'filter',
        'dataValidation',
      ])
    );
  }

  if (id === 'sheet.command.clear-selection-format') {
    return (
      firstDisabledFeature(features, ['clear-format', 'merge', 'conditionalFormatting']) ??
      disabledFormattingFeature(features)
    );
  }

  if (id === 'sheet.command.clear-selection-all') {
    return (
      firstDisabledFeature(features, [
        'clear-format',
        'merge',
        'conditionalFormatting',
        'dataValidation',
      ]) ?? disabledFormattingFeature(features)
    );
  }

  if (id.includes('format-painter')) {
    return (
      firstDisabledFeature(features, ['format-painter', 'merge', 'conditionalFormatting']) ??
      disabledFormattingFeature(features)
    );
  }

  if (id === 'sheet.command.auto-fill') {
    const params = asRecord(command.params);
    const applyType = String(params?.applyType ?? '').toUpperCase();

    // Every mode emits set-range-values payloads containing a style snapshot.
    // NO_FORMAT restores the target's existing style rather than changing it,
    // but the synchronous child policy cannot prove that without workbook
    // state. Conservatively preflight restricted formatting before the handler
    // clears any target cells.
    const formattingFeature = disabledFormattingFeature(features);
    if (formattingFeature) return formattingFeature;

    // NO_FORMAT does not run merge/CF/DV fill hooks. COPY, SERIES,
    // ONLY_FORMAT, and the data-dependent default can emit all three after the
    // leading clear/set-range mutations, so they must be blocked at the root.
    if (applyType.includes('NO_FORMAT')) return undefined;
    return firstDisabledFeature(features, ['merge', 'conditionalFormatting', 'dataValidation']);
  }

  if (id === 'sheet.command.refill') {
    // Refill first replays the previous fill's undo mutations, then applies a
    // newly-selected fill mode. The prior/current payloads are held in engine
    // state rather than command params, so conservatively preflight every
    // downstream capability before either phase begins.
    return (
      disabledFormattingFeature(features) ??
      firstDisabledFeature(features, ['merge', 'conditionalFormatting', 'dataValidation'])
    );
  }

  if (!id.includes('paste')) return undefined;

  if (features.clipboard === false) return 'clipboard';

  if (id === 'sheet.command.optional-paste') {
    // Choosing another mode first replays the previous paste's undo list. That
    // historical payload is not present in this command, so even a new
    // value-only mode must preflight every capability the previous paste may
    // have touched before any undo mutation runs.
    return (
      disabledFormattingFeature(features) ??
      firstDisabledFeature(features, ['merge', 'conditionalFormatting', 'dataValidation'])
    );
  }

  const params = asRecord(command.params);
  // Univer exposes the same modes through enum keys (`PASTE_VALUE`) and enum
  // values (`sheet.command.paste-value`) depending on the caller. Normalize
  // both shapes before deciding whether this is a mutation-safe paste mode.
  const pasteMode = String(params?.value ?? params?.type ?? '')
    .toLowerCase()
    .replace(/_/g, '-');
  // Value/formula-only paste never applies styles or merge mutations. Column
  // width paste is likewise outside the currently enforced feature set.
  if (
    pasteMode.includes('paste-value') ||
    pasteMode.includes('paste-formula') ||
    pasteMode.includes('paste-col-width') ||
    id.includes('paste-value') ||
    id.includes('paste-formula') ||
    id.includes('paste-col-width')
  ) {
    return undefined;
  }

  // "Besides border" still serializes `s.bd = null` in its leading
  // set-range-values mutation. The independent child policy therefore cannot
  // distinguish that mode from a border write. Preflight borders here too so a
  // late child veto cannot leave earlier paste hooks applied.
  const formattingFeature = disabledFormattingFeature(features);
  if (formattingFeature) return formattingFeature;

  // Default, format-only, and "besides border" paste all calculate merge
  // mutations. Their feature hooks differ for CF/DV, so preflight those modes
  // explicitly before their nested mutations begin.
  if (features.merge === false) return 'merge';

  const formatOnly = pasteMode.includes('paste-format') || id.includes('paste-format');
  if (features.conditionalFormatting === false) return 'conditionalFormatting';
  if (!formatOnly && features.dataValidation === false) return 'dataValidation';
  return undefined;
}

/** Return the first explicitly-disabled feature governing this command. */
export function disabledFeatureForCommand(
  command: BeforeCommandContext,
  features: Readonly<Record<string, boolean>> | undefined,
): string | undefined {
  if (!features) return undefined;

  // `params` is serializable caller input. In particular, Univer's `trigger`
  // field is useful attribution metadata but is not trusted provenance: raw
  // command dispatch and replay APIs can forge it. Always judge the command id
  // that actually reached the bus, and inspect the actual mutation payload.
  // Univer's execution stack is global rather than async-context-local, so it
  // is not safe provenance either: a concurrent command can appear nested
  // under an awaiting root. Compound handlers must preflight their own root.
  const compoundFeature = disabledCompoundFeature(command, features);
  if (compoundFeature) return compoundFeature;

  const disabledFormatting = new Set(
    FORMATTING_FEATURES.filter((feature) => features[feature] === false),
  );
  const styleInspection =
    disabledFormatting.size > 0
      ? styleFeaturesFromParams(command.params, disabledFormatting)
      : { features: new Set<string>(), opaqueStyle: false, clearsStyle: false };

  for (const [feature, enabled] of Object.entries(features)) {
    if (enabled !== false) continue;
    if (styleInspection.features.has(feature)) return feature;
    if (
      (styleInspection.opaqueStyle || styleInspection.clearsStyle) &&
      FORMATTING_FEATURES.includes(feature as (typeof FORMATTING_FEATURES)[number])
    ) {
      return feature;
    }
    const matchers = FEATURE_MATCHERS[feature];
    if (matchers?.some((matcher) => matcher(command.id, command.params))) {
      return feature;
    }
  }
  return undefined;
}

/**
 * Evaluate the live feature map + host policy. Authoritative collaboration,
 * snapshot replay, and local formula-result writes bypass local capability
 * policy so a restricted client still converges and recalculates. Formula
 * engine mutations carry a `formula.mutation.*` id plus `onlyLocal`; their
 * cell-write children carry `onlyLocal && fromFormula`. Public dispatch strips
 * all of those provenance options.
 */
export function createCommandPolicy(source: CommandPolicySource) {
  return (
    command: BeforeCommandContext,
    options?: CommandPolicyExecutionOptions,
  ): { allowed: true } | { allowed: false; feature?: string } => {
    if (
      options?.fromCollab ||
      options?.fromChangeset ||
      (options?.onlyLocal && (options.fromFormula || command.id.startsWith('formula.mutation.')))
    ) {
      return { allowed: true };
    }

    // Call the public hook for every local command even when a built-in feature
    // also blocks it. `true` is permission to continue evaluating policy, not
    // an override for an explicitly-disabled capability.
    const hostAllowed = source.beforeCommand()?.(clonePlainCommandData(command)) !== false;
    const feature = disabledFeatureForCommand(command, source.features());
    if (feature) return { allowed: false, feature };
    if (!hostAllowed) return { allowed: false };
    return { allowed: true };
  };
}
