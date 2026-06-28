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

import type { FUniver } from '@univerjs/core/facade';
import { WatermarkService } from '@univerjs/watermark';
import { IWatermarkTypeEnum } from '@univerjs/engine-render';

/**
 * Watermark plumbing for the View menu.
 *
 * The app hides Univer's ribbon, so the stock watermark panel never mounts.
 * Instead we drive the `WatermarkService` (registered by `UniverWatermarkPlugin`
 * in `univer/extra-plugins.ts`) directly from a small config dialog
 * (`WatermarkDialog.tsx`): apply writes a repeating diagonal text layer, off
 * clears it. The service persists the *active* config to `ILocalStorageService`
 * under `UNIVER_WATERMARK_STORAGE_KEY`, and the plugin re-reads it on boot, so
 * an applied watermark survives a reload for free.
 *
 * Separately we persist the user's *last chosen* text + opacity to `localStorage`
 * under `cs-watermark` so the dialog re-opens with their preference even after
 * the watermark has been turned off (the service drops its key entirely on
 * delete, so there's nothing to read back once it's off).
 */

export const DEFAULT_WATERMARK_TEXT = 'CONFIDENTIAL';
export const DEFAULT_WATERMARK_OPACITY = 0.12;

const PREF_KEY = 'cs-watermark';

export type WatermarkConfig = {
  /** Custom text. Falls back to DEFAULT_WATERMARK_TEXT when blank. */
  text: string;
  /** 0–1. Falls back to DEFAULT_WATERMARK_OPACITY when out of range. */
  opacity: number;
};

/** The user's last chosen text + opacity (independent of on/off state). */
export function loadWatermarkPref(): WatermarkConfig {
  const fallback: WatermarkConfig = {
    text: DEFAULT_WATERMARK_TEXT,
    opacity: DEFAULT_WATERMARK_OPACITY,
  };
  try {
    const raw = localStorage.getItem(PREF_KEY);
    if (!raw) return fallback;
    const parsed = JSON.parse(raw) as Partial<WatermarkConfig>;
    return {
      text: typeof parsed.text === 'string' && parsed.text.trim() ? parsed.text : fallback.text,
      opacity:
        typeof parsed.opacity === 'number' && parsed.opacity > 0 && parsed.opacity <= 1
          ? parsed.opacity
          : fallback.opacity,
    };
  } catch {
    return fallback;
  }
}

function saveWatermarkPref(config: WatermarkConfig): void {
  try {
    localStorage.setItem(PREF_KEY, JSON.stringify(config));
  } catch {
    /* storage disabled / quota — preference is best-effort */
  }
}

// Reach the service the same way the rest of the shell reaches Univer
// internals — via the facade's private `_injector` (see App.tsx /
// useAutosave.ts). Returns undefined before the workbook unit exists or if a
// Univer bump renames the back door.
function watermarkService(api: FUniver): WatermarkService | undefined {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const injector = (api as any)._injector as { get: (t: unknown) => unknown } | undefined;
  return injector?.get(WatermarkService) as WatermarkService | undefined;
}

/**
 * Apply a text watermark with the given config, or clear it when `on` is false.
 * The chosen text/opacity is persisted to `localStorage` either way, so the
 * dialog remembers the user's last choice across an off/on cycle and reloads.
 */
export function applyWatermark(api: FUniver, on: boolean, config: WatermarkConfig): void {
  const text = config.text.trim() || DEFAULT_WATERMARK_TEXT;
  const opacity =
    config.opacity > 0 && config.opacity <= 1 ? config.opacity : DEFAULT_WATERMARK_OPACITY;
  saveWatermarkPref({ text, opacity });

  const svc = watermarkService(api);
  if (!svc) return;
  if (on) {
    svc.updateWatermarkConfig({
      type: IWatermarkTypeEnum.Text,
      config: {
        text: {
          content: text,
          fontSize: 24,
          color: 'rgb(120,120,120)',
          bold: true,
          italic: false,
          direction: 'ltr',
          x: 60,
          y: 36,
          repeat: true,
          spacingX: 240,
          spacingY: 160,
          rotate: -30,
          opacity,
        },
      },
    });
  } else {
    svc.deleteWatermarkConfig();
  }
}

/** Whether a watermark is currently applied (used to seed the menu's checked state). */
export async function isWatermarkOn(api: FUniver): Promise<boolean> {
  const svc = watermarkService(api);
  if (!svc) return false;
  const config = await svc.getWatermarkConfig();
  return !!config?.config?.text?.content;
}

/** The text + opacity of the currently applied watermark, if any. */
export async function getAppliedWatermark(api: FUniver): Promise<WatermarkConfig | null> {
  const svc = watermarkService(api);
  if (!svc) return null;
  const config = await svc.getWatermarkConfig();
  const text = config?.config?.text;
  if (!text?.content) return null;
  return {
    text: text.content,
    opacity:
      typeof text.opacity === 'number' && text.opacity > 0 && text.opacity <= 1
        ? text.opacity
        : DEFAULT_WATERMARK_OPACITY,
  };
}
