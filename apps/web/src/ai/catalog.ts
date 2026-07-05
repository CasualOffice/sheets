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
 * Sheets DocOps tool catalog — Phase 0 parity with the document editor.
 * 5 read tools + 2 write tools.
 * Sent verbatim to the Anthropic messages API as the `tools` array.
 */

export interface SheetsTool {
  name: string;
  description: string;
  input_schema: {
    type: 'object';
    properties: Record<string, unknown>;
    required?: string[];
  };
}

export const SHEETS_CATALOG: SheetsTool[] = [
  {
    name: 'get_workbook_info',
    description:
      'List all sheets: name, id, how many rows and columns of DATA each contains (dataRows/dataColumns — not the empty grid size), its used range in A1 (dataRange), and whether it is empty (isEmpty) or active. Call this first to orient yourself, then read the used range with get_cell_range to see the actual values.',
    input_schema: {
      type: 'object',
      properties: {},
    },
  },
  {
    name: 'get_selection',
    description:
      'Return the current selected range: A1 notation and cell values. Call this to know what the user has selected before making edits.',
    input_schema: {
      type: 'object',
      properties: {},
    },
  },
  {
    name: 'get_cell_range',
    description:
      'Get values from a specific range of cells (A1 notation, e.g. "A1:D10" or "Sheet2!B2:E5"). Returns a 2D array of cell values.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'A1 notation of the range to read, e.g. "A1:D10" or "Sheet2!B2:E5".',
        },
      },
      required: ['range'],
    },
  },
  {
    name: 'get_sheet_stats',
    description:
      'Return the active sheet data extent: rowCount and columnCount of data, non-empty cell count, the used range in A1 (dataRange), and isEmpty. Use the dataRange with get_cell_range to read the actual values before summarizing or analyzing.',
    input_schema: {
      type: 'object',
      properties: {},
    },
  },
  {
    name: 'find_in_sheet',
    description:
      'Search for text or a value in the active sheet. Returns matching cell addresses and their values.',
    input_schema: {
      type: 'object',
      properties: {
        query: {
          type: 'string',
          description: 'The text to search for (case-insensitive).',
        },
        limit: {
          type: 'number',
          description: 'Maximum number of matches to return. Defaults to 20.',
        },
      },
      required: ['query'],
    },
  },

  // ── Write tools ────────────────────────────────────────────────────────────
  {
    name: 'set_cell_values',
    description:
      'Set values in a range of cells. Pass values as a 2D array matching the range shape. ' +
      'Use for plain text or numbers. For formulas, use set_formula instead.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'A1 notation of the target range, e.g. "A1:C3" or "B2".',
        },
        values: {
          type: 'array',
          items: { type: 'array', items: { type: ['string', 'number'] } },
          description:
            'Values to write — outer array is rows, inner is columns. Must match the range dimensions. Use numbers for numeric values (e.g. 100, not "100") so they compute correctly.',
        },
      },
      required: ['range', 'values'],
    },
  },
  {
    name: 'set_formula',
    description:
      'Set a formula in a single cell. The formula can start with "=" or without (both are accepted). ' +
      'Example: "=SUM(A1:A10)" or "SUM(A1:A10)".',
    input_schema: {
      type: 'object',
      properties: {
        cell: {
          type: 'string',
          description: 'A1 address of the target cell, e.g. "B5" or "Sheet2!C3".',
        },
        formula: {
          type: 'string',
          description: 'The formula to set. Leading "=" is optional.',
        },
      },
      required: ['cell', 'formula'],
    },
  },
];
