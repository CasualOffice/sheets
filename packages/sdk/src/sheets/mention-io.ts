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
