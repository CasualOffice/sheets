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

import { cleanCommandExecutionStackForVeto } from './command-veto';

test('a veto removes only the exact current command from Univer 0.25 stack shape', () => {
  const parent = { id: 'parent' };
  const vetoed = { id: 'vetoed' };
  const service = { _commandExecutionStack: [parent, vetoed] };

  assert.equal(cleanCommandExecutionStackForVeto(service, vetoed), true);
  assert.deepEqual(service._commandExecutionStack, [parent]);
});

test('stack cleanup is harmless after Univer already cleaned or changed internals', () => {
  const vetoed = { id: 'vetoed' };
  assert.equal(cleanCommandExecutionStackForVeto({ _commandExecutionStack: [] }, vetoed), false);
  assert.equal(cleanCommandExecutionStackForVeto({}, vetoed), false);
  assert.equal(cleanCommandExecutionStackForVeto(undefined, vetoed), false);
});
