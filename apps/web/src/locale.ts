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

import { LocaleType, Tools } from '@univerjs/core';

import UniverSheetsEnUS from '@univerjs/sheets/locale/en-US';
import UniverSheetsUIEnUS from '@univerjs/sheets-ui/locale/en-US';
import UniverSheetsFormulaEnUS from '@univerjs/sheets-formula/locale/en-US';
import UniverSheetsFormulaUIEnUS from '@univerjs/sheets-formula-ui/locale/en-US';
import UniverSheetsSortUIEnUS from '@univerjs/sheets-sort-ui/locale/en-US';
import UniverSheetsFilterUIEnUS from '@univerjs/sheets-filter-ui/locale/en-US';
import UniverSheetsNumfmtUIEnUS from '@univerjs/sheets-numfmt-ui/locale/en-US';
import UniverFindReplaceEnUS from '@univerjs/find-replace/locale/en-US';
// Univer 0.25 consolidated @univerjs/sheets-find-replace's locale strings into the
// base @univerjs/find-replace package; the sheets-find-replace package no longer ships
// its own `locale/` (the import broke the prod Rollup build). UniverFindReplaceEnUS now
// covers both.
import UniverSheetsConditionalFormattingUIEnUS from '@univerjs/sheets-conditional-formatting-ui/locale/en-US';
// The data-validation *UI* package only ships the panel chrome strings
// (`sheets-data-validation-ui.*`). The validator titles, operator names,
// rule names and error messages live in the base `@univerjs/data-validation`
// (`data-validation.*`) and `@univerjs/sheets-data-validation`
// (`sheets-data-validation.*`) packages — without these the DV panel's Type /
// Operator selectors and the cell error messages render raw i18n keys.
import UniverDataValidationEnUS from '@univerjs/data-validation/locale/en-US';
import UniverSheetsDataValidationEnUS from '@univerjs/sheets-data-validation/locale/en-US';
import UniverSheetsDataValidationUIEnUS from '@univerjs/sheets-data-validation-ui/locale/en-US';
// Same base-vs-UI split as data-validation: these `-ui` packages ship only the
// panel chrome, while error toasts + generated labels (filter range errors,
// table name/column prefixes + validation, hyperlink ref errors) live in the
// base packages' `*.` namespaces. Without the base locales they render raw keys.
import UniverSheetsFilterEnUS from '@univerjs/sheets-filter/locale/en-US';
import UniverSheetsHyperLinkEnUS from '@univerjs/sheets-hyper-link/locale/en-US';
import UniverSheetsTableEnUS from '@univerjs/sheets-table/locale/en-US';
import UniverSheetsHyperLinkUIEnUS from '@univerjs/sheets-hyper-link-ui/locale/en-US';
import UniverSheetsNoteUIEnUS from '@univerjs/sheets-note-ui/locale/en-US';
import UniverSheetsTableUIEnUS from '@univerjs/sheets-table-ui/locale/en-US';
import UniverThreadCommentUIEnUS from '@univerjs/thread-comment-ui/locale/en-US';
import UniverSheetsThreadCommentUIEnUS from '@univerjs/sheets-thread-comment-ui/locale/en-US';
import UniverDrawingUIEnUS from '@univerjs/drawing-ui/locale/en-US';
import UniverSheetsDrawingUIEnUS from '@univerjs/sheets-drawing-ui/locale/en-US';
import UniverSheetsCrosshairHighlightEnUS from '@univerjs/sheets-crosshair-highlight/locale/en-US';
import UniverSheetsZenEditorEnUS from '@univerjs/sheets-zen-editor/locale/en-US';
import UniverDocsUIEnUS from '@univerjs/docs-ui/locale/en-US';
import UniverUIEnUS from '@univerjs/ui/locale/en-US';

const enUS = Tools.deepMerge(
  {},
  UniverSheetsEnUS,
  UniverSheetsUIEnUS,
  UniverSheetsFormulaEnUS,
  UniverSheetsFormulaUIEnUS,
  UniverSheetsSortUIEnUS,
  UniverSheetsFilterUIEnUS,
  UniverSheetsNumfmtUIEnUS,
  UniverFindReplaceEnUS,
  UniverSheetsConditionalFormattingUIEnUS,
  UniverDataValidationEnUS,
  UniverSheetsDataValidationEnUS,
  UniverSheetsDataValidationUIEnUS,
  UniverSheetsFilterEnUS,
  UniverSheetsHyperLinkEnUS,
  UniverSheetsTableEnUS,
  UniverSheetsHyperLinkUIEnUS,
  UniverSheetsNoteUIEnUS,
  UniverSheetsTableUIEnUS,
  UniverThreadCommentUIEnUS,
  UniverSheetsThreadCommentUIEnUS,
  UniverDrawingUIEnUS,
  UniverSheetsDrawingUIEnUS,
  UniverSheetsCrosshairHighlightEnUS,
  UniverSheetsZenEditorEnUS,
  UniverDocsUIEnUS,
  UniverUIEnUS,
  // Explicit fallback for the sheets-numfmt-ui "number stored as text" toast
  // (title + message). These were observed rendering as raw keys
  // (`sheets-numfmt-ui.info.error` / `.forceStringInfo`) in the production
  // build despite the bundle being merged above; pinning them here guarantees
  // they resolve. deepMerge applies later args last, so this wins.
  {
    'sheets-numfmt-ui': {
      info: {
        error: 'Error',
        forceStringInfo: 'Number stored as text',
      },
    },
  },
);

export const LOCALES = {
  [LocaleType.EN_US]: enUS,
};
