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
 * Facade extension side-effect imports — each module mutates FUniver /
 * FWorkbook / FWorksheet prototypes to add domain methods (e.g. addTable,
 * createFilter, sort). Order doesn't matter; importing twice is a no-op.
 *
 * If you add a new sheets-* plugin that ships a /facade entrypoint, register
 * it here so its facade methods are available before any caller hits them.
 */
import '@univerjs/sheets/facade';
import '@univerjs/sheets-ui/facade';
import '@univerjs/sheets-formula/facade';
import '@univerjs/sheets-numfmt/facade';
import '@univerjs/sheets-sort/facade';
import '@univerjs/sheets-filter/facade';
import '@univerjs/sheets-table/facade';
import '@univerjs/docs-ui/facade';
import '@univerjs/ui/facade';
import '@univerjs/engine-formula/facade';
