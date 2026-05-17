/**
 * A single-axis outline group. Single-level for v1 — no nesting, no level
 * field. Stored per axis (rows / cols) per sheet.
 */
export type OutlineGroup = {
  /** Random id — UI uses it to address groups without depending on order. */
  id: string;
  /** Inclusive zero-based start index (row or column). */
  start: number;
  /** Inclusive zero-based end index. */
  end: number;
  /** True if the group is currently collapsed (rows/cols hidden). */
  collapsed: boolean;
};

export type SheetOutline = {
  rows: OutlineGroup[];
  cols: OutlineGroup[];
};

export type OutlineState = Record<string, SheetOutline>;

/**
 * Plugin-resource name we use when stashing outline state in
 * `IWorkbookData.resources`. Picked to be distinct from any Univer-native
 * resource key so we don't ever collide.
 */
export const OUTLINE_RESOURCE_NAME = '__casual_sheets_outline__';

/**
 * Serialized shape stored in the resource. Matches `OutlineState` 1:1; kept
 * as a stand-alone type so future schema migrations can branch on it.
 */
export type OutlineResourceV1 = {
  v: 1;
  sheets: OutlineState;
};
