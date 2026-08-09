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

import type { IWorkbookData } from '@univerjs/core';

type WorkbookResources = IWorkbookData['resources'];
type WorkbookResource = NonNullable<WorkbookResources>[number];

const copyResources = (resources: WorkbookResources): WorkbookResource[] =>
  (resources ?? []).map((resource) => ({ ...resource }));

/**
 * Keeps snapshot resources that Univer cannot serialize itself.
 *
 * Univer's live `save()` result only contains resources owned by registered
 * plugins. Imported xlsx sidecars (page setup, tables, opaque OOXML, and host
 * resources) therefore need an SDK-level shadow until a live save supplies a
 * replacement with the same name. Live entries always win. Once Univer has
 * emitted a name, later absence is treated as a real deletion; only resources
 * the runtime has never emitted remain shadowed.
 */
export function createWorkbookResourcePreserver(initial: WorkbookResources) {
  let preserved = copyResources(initial);
  let runtimeManagedNames = new Set<string>();

  return {
    reset(resources: WorkbookResources): void {
      preserved = copyResources(resources);
      runtimeManagedNames = new Set();
    },

    merge(liveResources: WorkbookResources): WorkbookResources {
      const live = copyResources(liveResources);
      const liveNames = new Set(live.map((resource) => resource.name));
      for (const name of liveNames) runtimeManagedNames.add(name);
      const merged = [
        ...live,
        ...preserved.filter(
          (resource) => !liveNames.has(resource.name) && !runtimeManagedNames.has(resource.name),
        ),
      ];

      preserved = copyResources(merged);
      return merged.length > 0 ? merged : undefined;
    },
  };
}
