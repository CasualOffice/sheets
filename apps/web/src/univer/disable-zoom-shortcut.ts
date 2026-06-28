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
 * Disable Univer's Ctrl+- / Ctrl+= zoom shortcuts in favour of our
 * Excel-style Insert/Delete-cells dialogs. The dialogs are owned by
 * MenuBar.tsx and registered via a window-level keydown listener;
 * Univer's IShortcutService also fires for the same binding, which
 * caused BOTH the dialog to open AND the canvas to zoom — confusing,
 * and the unsolicited zoom drifted user state.
 *
 * Univer's shortcut service picks ONE candidate per keydown (sorted
 * by priority desc; first that passes preconditions). Registering a
 * NilCommand binding with a very high priority displaces Univer's
 * built-in zoom-in / zoom-out shortcuts at those bindings without
 * having to reach in and find the original disposables (which
 * Univer's service doesn't expose).
 *
 * Why not change OUR binding instead? Ctrl+- is the Excel standard
 * for Delete Cells; users with muscle memory expect it. Keeping it
 * and disabling Univer's zoom binding is the more compatible choice.
 */
import type { FUniver } from '@univerjs/core/facade';
import { NilCommand } from '@univerjs/core';
import { IShortcutService, KeyCode, MetaKeys } from '@univerjs/ui';

const HIGH_PRIORITY = 1000;

export function disableUniverZoomShortcut(api: FUniver): () => void {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const injector = (api as any)._injector as
    | { get: (token: unknown) => unknown }
    | undefined;
  if (!injector) return () => {};
  let svc:
    | {
        registerShortcut: (item: unknown) => { dispose: () => void };
      }
    | undefined;
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    svc = injector.get(IShortcutService) as any;
  } catch (err) {
    console.warn('[shortcuts] IShortcutService unavailable', err);
    return () => {};
  }
  if (!svc) return () => {};

  const disposers: Array<() => void> = [];
  const grab = (binding: number, descr: string) => {
    try {
      const d = svc!.registerShortcut({
        id: NilCommand.id,
        binding,
        description: `[overridden] ${descr}`,
        priority: HIGH_PRIORITY,
      });
      disposers.push(() => d?.dispose?.());
    } catch (err) {
      console.warn('[shortcuts] failed to override', descr, err);
    }
  };

  // Zoom-out: Ctrl+-
  grab(KeyCode.MINUS | MetaKeys.CTRL_COMMAND, 'Ctrl+- (zoom-out)');
  // Zoom-in: Ctrl+=  (= because + needs shift on most layouts; Univer
  // also binds the EQUAL key without shift). Our Insert Cells dialog
  // accepts NumpadAdd / '=' with shift / '+' — handled in MenuBar.
  grab(KeyCode.EQUAL | MetaKeys.CTRL_COMMAND, 'Ctrl+= (zoom-in)');
  // Reset zoom: Ctrl+0 (Excel uses this too but for different
  // behaviour). Leave alone for now — no conflict with our shortcuts.

  return () => {
    for (const d of disposers) {
      try {
        d();
      } catch {
        /* dispose double-call is fine */
      }
    }
  };
}
