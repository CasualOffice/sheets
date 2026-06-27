/**
 * Text to Columns — pure delimiter helpers (preview + mask building).
 *
 * The actual split is performed by the Univer fork's
 * `sheet.command.split-text-to-columns` (which inserts columns + is undoable).
 * This module only (a) builds the delimiter bitmask the command expects and
 * (b) reproduces its split for the wizard's live preview — using the SAME
 * character-class regex the command builds, so what the user sees is what they
 * get. Mirrors `vendor/.../basics/split-range-text.ts`.
 */

// Matches SplitDelimiterEnum in the fork (a bitmask).
export const DELIMITER = {
  tab: 1,
  comma: 2,
  semicolon: 4,
  space: 8,
  custom: 16,
} as const;

export interface DelimiterOptions {
  tab: boolean;
  comma: boolean;
  semicolon: boolean;
  space: boolean;
  /** Single-character custom delimiter; empty/undefined disables it. */
  custom?: string;
}

export function buildDelimiterMask(opts: DelimiterOptions): number {
  let mask = 0;
  if (opts.tab) mask |= DELIMITER.tab;
  if (opts.comma) mask |= DELIMITER.comma;
  if (opts.semicolon) mask |= DELIMITER.semicolon;
  if (opts.space) mask |= DELIMITER.space;
  if (opts.custom && opts.custom.length > 0) mask |= DELIMITER.custom;
  return mask;
}

/** True when at least one delimiter is active (otherwise there's nothing to split on). */
export function hasActiveDelimiter(opts: DelimiterOptions): boolean {
  return opts.tab || opts.comma || opts.semicolon || opts.space || !!opts.custom;
}

const escapeRegExp = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** Build the character-class regex the fork uses (custom char first). */
function delimiterRegex(opts: DelimiterOptions, treatMultipleAsOne: boolean): RegExp {
  let str = '';
  if (opts.custom && opts.custom.length > 0) str += escapeRegExp(opts.custom[0]);
  if (opts.tab) str += '\\t';
  if (opts.comma) str += ',';
  if (opts.semicolon) str += ';';
  if (opts.space) str += ' ';
  return new RegExp(`[${str}]${treatMultipleAsOne ? '+' : ''}`);
}

/**
 * Split sample rows for the preview. Returns one string[] per input row; with
 * no active delimiter each row stays a single column (matches the regex's
 * empty character class, which never matches).
 */
export function splitPreview(
  samples: string[],
  opts: DelimiterOptions,
  treatMultipleAsOne: boolean,
): string[][] {
  if (!hasActiveDelimiter(opts)) return samples.map((s) => [s]);
  const re = delimiterRegex(opts, treatMultipleAsOne);
  return samples.map((s) => s.split(re));
}
