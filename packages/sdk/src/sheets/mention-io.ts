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
 * `CasualMentionIOService` — the `IMentionIOService` override registered by
 * CasualSheets. Thin by design: it owns the only @univerjs *value* import
 * (`MentionType`) and delegates all logic to the pure, unit-tested
 * `resolveMentionList` in `mention-source.ts`.
 */
import { MentionType } from '@univerjs/core';
import type { IListMentionParam, IListMentionResponse, IMentionIOService } from '@univerjs/core';
import { resolveMentionList } from './mention-source';

export class CasualMentionIOService implements IMentionIOService {
  list(params: IListMentionParam): Promise<IListMentionResponse> {
    return resolveMentionList(params, MentionType.PERSON);
  }
}
