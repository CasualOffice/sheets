/*
 * Copyright (c) 2026 Casual Office. All rights reserved.
 */

/** Anthropic-compatible tool definition (matches SheetsTool / DocOpsTool). */
export interface DocOpsTool {
  name: string;
  description: string;
  input_schema: { type: 'object'; properties: Record<string, unknown>; required?: string[] };
}

export type DocOpsSuccess<T = unknown> = {
  ok: true;
  data?: T;
  changedBlockIds?: string[];
  diffSummary?: string;
  suggestionId?: string;
};
export type DocOpsError = { ok: false; code: string; message: string; retryable: boolean };
export type DocOpsResult<T = unknown> = DocOpsSuccess<T> | DocOpsError;
