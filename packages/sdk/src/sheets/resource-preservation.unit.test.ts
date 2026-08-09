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

import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import { createWorkbookResourcePreserver } from './resource-preservation';

test('preserves every imported resource missing from a live Univer save', () => {
  const preserver = createWorkbookResourcePreserver([
    { name: '__casual_sheets_page_setup__', data: '{"orientation":"landscape"}' },
    { name: '__casual_sheets_tables__', data: '{"tables":["Sales"]}' },
    { name: '__casual_sheets_xlsx_passthrough__', data: '{"drawings":{}}' },
    { name: 'HOST_OPAQUE_RESOURCE', data: 'opaque-payload' },
    { name: 'SHEET_DATA_VALIDATION_PLUGIN', data: 'imported-validation' },
  ]);

  assert.deepEqual(
    preserver.merge([
      { name: 'SHEET_DATA_VALIDATION_PLUGIN', data: 'live-validation' },
      { name: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN', data: 'live-formatting' },
    ]),
    [
      { name: 'SHEET_DATA_VALIDATION_PLUGIN', data: 'live-validation' },
      { name: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN', data: 'live-formatting' },
      { name: '__casual_sheets_page_setup__', data: '{"orientation":"landscape"}' },
      { name: '__casual_sheets_tables__', data: '{"tables":["Sales"]}' },
      { name: '__casual_sheets_xlsx_passthrough__', data: '{"drawings":{}}' },
      { name: 'HOST_OPAQUE_RESOURCE', data: 'opaque-payload' },
    ],
  );
});

test('promotes live resources and resets when a new workbook is loaded', () => {
  const preserver = createWorkbookResourcePreserver([
    { name: 'PLUGIN', data: 'imported-plugin' },
    { name: 'OPAQUE', data: 'imported-opaque' },
  ]);

  assert.deepEqual(preserver.merge([{ name: 'PLUGIN', data: 'edited' }]), [
    { name: 'PLUGIN', data: 'edited' },
    { name: 'OPAQUE', data: 'imported-opaque' },
  ]);
  // Once Univer has emitted a resource, later absence is a real runtime
  // deletion. Opaque resources Univer never emitted stay preserved.
  assert.deepEqual(preserver.merge(undefined), [{ name: 'OPAQUE', data: 'imported-opaque' }]);

  preserver.reset([{ name: 'REPLACEMENT', data: '' }]);
  assert.deepEqual(preserver.merge(undefined), [{ name: 'REPLACEMENT', data: '' }]);
});
