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
 * en-US locale bundle for the in-iframe runtime.
 *
 * The embed runs the FULL feature set (`lazyPlugins={true}`): the base editor
 * (render + formula engines, UI, docs/docs-ui, sheets, sheets-ui,
 * sheets-formula-ui, numfmt) PLUS the lazily-registered feature plugins
 * (tables, sort, filter, conditional formatting, data validation, drawing,
 * hyperlinks, notes, thread comments, find/replace). Univer's `LocaleService`
 * resolves every UI string against this map; without a plugin's strings its UI
 * renders raw i18n keys (e.g. the comment panel showed
 * `thread-comment-ui.editor.reply` instead of "Reply"). So we MUST seed strings
 * for every plugin the runtime can register — matching the reference app's
 * `apps/web/src/locale.ts` set.
 *
 * (The React `<CasualSheets>` host path supplies its own `locales` prop; the
 * iframe has no host to pass one, so it bundles this full set.)
 */

import { LocaleType, Tools } from '@univerjs/core';

import UniverUIEnUS from '@univerjs/ui/locale/en-US';
import UniverDocsUIEnUS from '@univerjs/docs-ui/locale/en-US';
import UniverSheetsEnUS from '@univerjs/sheets/locale/en-US';
import UniverSheetsUIEnUS from '@univerjs/sheets-ui/locale/en-US';
import UniverSheetsFormulaUIEnUS from '@univerjs/sheets-formula-ui/locale/en-US';
import UniverSheetsNumfmtUIEnUS from '@univerjs/sheets-numfmt-ui/locale/en-US';
// Feature-plugin strings — these load lazily, so without their locales the
// feature UIs (comment panel, table dialog, filter, sort, CF, data validation,
// images, hyperlinks, notes, find/replace) render raw i18n keys.
import UniverSheetsTableUIEnUS from '@univerjs/sheets-table-ui/locale/en-US';
import UniverSheetsSortUIEnUS from '@univerjs/sheets-sort-ui/locale/en-US';
import UniverSheetsFilterUIEnUS from '@univerjs/sheets-filter-ui/locale/en-US';
import UniverSheetsCfUIEnUS from '@univerjs/sheets-conditional-formatting-ui/locale/en-US';
// The DV *UI* package ships only the panel chrome (`sheets-data-validation-ui.*`).
// Validator titles, operator names, rule names and error messages live in the
// base `@univerjs/data-validation` + `@univerjs/sheets-data-validation` packages;
// without them the DV Type/Operator selectors and cell errors render raw i18n keys.
import UniverDvEnUS from '@univerjs/data-validation/locale/en-US';
import UniverSheetsDvEnUS from '@univerjs/sheets-data-validation/locale/en-US';
import UniverSheetsDvUIEnUS from '@univerjs/sheets-data-validation-ui/locale/en-US';
// Same base-vs-UI split: the `-ui` packages above ship only panel chrome, while
// error toasts + generated labels (filter range errors, table name/column
// prefixes + validation, hyperlink ref errors) live in the base packages'
// namespaces. Without the base locales they render raw i18n keys.
import UniverSheetsFilterEnUS from '@univerjs/sheets-filter/locale/en-US';
import UniverSheetsHyperLinkEnUS from '@univerjs/sheets-hyper-link/locale/en-US';
import UniverSheetsTableEnUS from '@univerjs/sheets-table/locale/en-US';
import UniverSheetsDrawingUIEnUS from '@univerjs/sheets-drawing-ui/locale/en-US';
import UniverDrawingUIEnUS from '@univerjs/drawing-ui/locale/en-US';
import UniverSheetsHyperLinkUIEnUS from '@univerjs/sheets-hyper-link-ui/locale/en-US';
import UniverSheetsNoteUIEnUS from '@univerjs/sheets-note-ui/locale/en-US';
import UniverThreadCommentUIEnUS from '@univerjs/thread-comment-ui/locale/en-US';
import UniverSheetsThreadCommentUIEnUS from '@univerjs/sheets-thread-comment-ui/locale/en-US';
import UniverFindReplaceEnUS from '@univerjs/find-replace/locale/en-US';

const enUS = Tools.deepMerge(
  {},
  UniverUIEnUS,
  UniverDocsUIEnUS,
  UniverSheetsEnUS,
  UniverSheetsUIEnUS,
  UniverSheetsFormulaUIEnUS,
  UniverSheetsNumfmtUIEnUS,
  UniverSheetsTableUIEnUS,
  UniverSheetsSortUIEnUS,
  UniverSheetsFilterUIEnUS,
  UniverSheetsCfUIEnUS,
  UniverDvEnUS,
  UniverSheetsDvEnUS,
  UniverSheetsDvUIEnUS,
  UniverSheetsFilterEnUS,
  UniverSheetsHyperLinkEnUS,
  UniverSheetsTableEnUS,
  UniverDrawingUIEnUS,
  UniverSheetsDrawingUIEnUS,
  UniverSheetsHyperLinkUIEnUS,
  UniverSheetsNoteUIEnUS,
  UniverThreadCommentUIEnUS,
  UniverSheetsThreadCommentUIEnUS,
  UniverFindReplaceEnUS,
);

export const EMBED_LOCALES = {
  [LocaleType.EN_US]: enUS,
};
