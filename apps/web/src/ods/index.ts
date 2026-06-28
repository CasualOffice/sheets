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
 * OpenDocument Spreadsheet (.ods) I/O.
 *
 * Univer's own xlsx codec (apps/web/src/xlsx/) stays the primary format —
 * this module exists so users can open files saved from LibreOffice Calc /
 * OpenOffice Calc / Google Sheets's .ods export and save back to the same
 * format. We use the SheetJS Community fork (`@e965/xlsx`, Apache-2.0)
 * because it's the only well-maintained library that handles .ods on npm.
 *
 * Scope:
 *   - Values + formulas + cached formula results
 *   - Sheet order + names
 *   - Merges
 *   - Column widths / row heights
 *   - Hidden rows / columns
 *   - Frozen panes
 *   - Basic styles (font, fill, alignment, number formats)
 *   - Hyperlinks, comments, defined names
 *
 * Sheet visibility is not reliably represented by the ODS library model we
 * use here, so hidden sheet state remains out of scope for now.
 */
import * as XLSX from '@e965/xlsx';
import {
  CustomRangeType,
  LocaleType,
  type ICellData,
  type IRange,
  type IStyleData,
  type IWorkbookData,
} from '@univerjs/core';
import { INITIAL_COLUMNS, INITIAL_ROWS, UNIVER_VERSION } from '../snapshot';

type TabularFormat = 'ods' | 'csv' | 'tsv' | 'psv';
const DEFINED_NAMES_RESOURCE = 'SHEET_DEFINED_NAME_PLUGIN';
const NOTE_RESOURCE = 'SHEET_NOTE_PLUGIN';
let hyperlinkIdCounter = 0;
const nextHyperlinkId = () =>
  `hl-${Date.now().toString(36)}-${(hyperlinkIdCounter++).toString(36)}`;

function buildHyperlinkBody(display: string, url: string, id: string): ICellData['p'] {
  const dataStream = `${display}\r\n`;
  return {
    id: '__INTERNAL_EDITOR__DOCS_NORMAL',
    documentStyle: {},
    body: {
      dataStream,
      customRanges: [
        {
          startIndex: 0,
          endIndex: display.length - 1,
          rangeType: CustomRangeType.HYPERLINK,
          rangeId: id,
          properties: { url },
        },
      ],
      paragraphs: [{ startIndex: display.length }],
      sectionBreaks: [{ startIndex: display.length + 1 }],
      textRuns: [],
    },
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any;
}

function getHyperlinkFromCell(cell: {
  p?: ICellData['p'];
}): { url: string; display?: string } | undefined {
  const body = cell.p?.body;
  const ranges = body?.customRanges ?? [];
  for (const cr of ranges) {
    if (cr.rangeType !== CustomRangeType.HYPERLINK) continue;
    const url = cr.properties?.url;
    if (typeof url !== 'string' || !url) continue;
    const dataStream = body?.dataStream ?? '';
    const display = dataStream.slice(cr.startIndex, cr.endIndex + 1);
    return { url, ...(display ? { display } : {}) };
  }
  return undefined;
}

type FreezeMeta = { xSplit: number; ySplit: number; startRow: number; startColumn: number };
type DimensionMeta = {
  columnsBySheetName: Record<string, Record<number, { w?: number; hd?: number }>>;
  rowsBySheetName: Record<string, Record<number, { h?: number; hd?: number }>>;
};
type CellStyleMeta = {
  stylesByName: Record<string, IStyleData>;
  refsBySheetName: Record<string, Record<number, Record<number, string>>>;
};
type ZipContainer = {
  FullPaths: string[];
  FileIndex: Array<{ content?: unknown }>;
};
type SnapshotCell = {
  v?: string | number | boolean;
  f?: string;
  s?: string | IStyleData;
  p?: ICellData['p'];
};

function getZipEntry(cfb: ZipContainer, path: string): { content?: unknown } | undefined {
  const fullPath = cfb.FullPaths.find((p: string) => p === `Root Entry/${path}`);
  if (!fullPath) return undefined;
  const idx = cfb.FullPaths.indexOf(fullPath);
  return idx >= 0 ? cfb.FileIndex[idx] : undefined;
}

function decodeZipText(content: unknown): string {
  if (typeof content === 'string') return content;
  if (content instanceof Uint8Array) return new TextDecoder().decode(content);
  if (content instanceof ArrayBuffer) return new TextDecoder().decode(new Uint8Array(content));
  return new TextDecoder().decode(new Uint8Array(content as ArrayLike<number>));
}

function encodeZipText(text: string): Uint8Array {
  return new TextEncoder().encode(text);
}

function parseFreezeSettingsXml(xml: string): Record<string, FreezeMeta> {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const out: Record<string, FreezeMeta> = {};

  const entries = Array.from(doc.getElementsByTagName('*')).filter(
    (el) =>
      el.localName === 'config-item-map-entry' &&
      el.parentElement?.localName === 'config-item-map-named' &&
      el.parentElement?.getAttribute('config:name') === 'Tables',
  );

  const getConfig = (entry: Element, name: string): string | undefined => {
    const node = Array.from(entry.children).find(
      (child) => child.localName === 'config-item' && child.getAttribute('config:name') === name,
    );
    return node?.textContent?.trim() || undefined;
  };

  for (const entry of entries) {
    const sheetName = entry.getAttribute('config:name');
    if (!sheetName) continue;

    const xMode = Number(getConfig(entry, 'HorizontalSplitMode') ?? 0);
    const yMode = Number(getConfig(entry, 'VerticalSplitMode') ?? 0);
    const xSplit = Number(getConfig(entry, 'HorizontalSplitPosition') ?? 0);
    const ySplit = Number(getConfig(entry, 'VerticalSplitPosition') ?? 0);

    if ((xMode !== 2 && yMode !== 2) || (xSplit <= 0 && ySplit <= 0)) continue;

    out[sheetName] = {
      xSplit,
      ySplit,
      startRow: ySplit > 0 ? ySplit : -1,
      startColumn: xSplit > 0 ? xSplit : -1,
    };
  }

  return out;
}

function parsePixelSize(raw: string | null): number | undefined {
  if (!raw) return undefined;
  const m = /^\s*([0-9]+(?:\.[0-9]+)?)px\s*$/.exec(raw);
  if (!m) return undefined;
  return Math.round(Number(m[1]));
}

function parseDimensionStylesFromContentXml(xml: string): DimensionMeta {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const columnStyleSizes = new Map<string, number>();
  const rowStyleSizes = new Map<string, number>();
  const columnsBySheetName: DimensionMeta['columnsBySheetName'] = {};
  const rowsBySheetName: DimensionMeta['rowsBySheetName'] = {};

  for (const style of Array.from(doc.getElementsByTagName('*'))) {
    if (style.localName !== 'style') continue;
    const name = style.getAttribute('style:name');
    const family = style.getAttribute('style:family');
    if (!name || !family) continue;
    if (family === 'table-column') {
      const props = Array.from(style.children).find(
        (child) => child.localName === 'table-column-properties',
      );
      const width = parsePixelSize(props?.getAttribute('style:column-width') ?? null);
      if (width !== undefined) columnStyleSizes.set(name, width);
    }
    if (family === 'table-row') {
      const props = Array.from(style.children).find(
        (child) => child.localName === 'table-row-properties',
      );
      const height = parsePixelSize(props?.getAttribute('style:row-height') ?? null);
      if (height !== undefined) rowStyleSizes.set(name, height);
    }
  }

  for (const table of Array.from(doc.getElementsByTagName('*')).filter(
    (el) => el.localName === 'table',
  )) {
    const sheetName = table.getAttribute('table:name');
    if (!sheetName) continue;

    let colIndex = 0;
    let rowIndex = 0;
    for (const child of Array.from(table.children)) {
      if (child.localName === 'table-column') {
        const repeated = Number(child.getAttribute('table:number-columns-repeated') ?? '1') || 1;
        const styleName = child.getAttribute('table:style-name');
        const width = styleName ? columnStyleSizes.get(styleName) : undefined;
        const hidden = child.getAttribute('table:visibility');
        for (let i = 0; i < repeated; i++) {
          if (width !== undefined || (hidden && hidden !== 'visible')) {
            columnsBySheetName[sheetName] ??= {};
            columnsBySheetName[sheetName][colIndex] = {
              ...(width !== undefined ? { w: width } : {}),
              ...(hidden && hidden !== 'visible' ? { hd: 1 } : {}),
            };
          }
          colIndex++;
        }
      }

      if (child.localName === 'table-row') {
        const repeated = Number(child.getAttribute('table:number-rows-repeated') ?? '1') || 1;
        const styleName = child.getAttribute('table:style-name');
        const height = styleName ? rowStyleSizes.get(styleName) : undefined;
        const hidden = child.getAttribute('table:visibility');
        for (let i = 0; i < repeated; i++) {
          if (height !== undefined || (hidden && hidden !== 'visible')) {
            rowsBySheetName[sheetName] ??= {};
            rowsBySheetName[sheetName][rowIndex] = {
              ...(height !== undefined ? { h: height } : {}),
              ...(hidden && hidden !== 'visible' ? { hd: 1 } : {}),
            };
          }
          rowIndex++;
        }
      }
    }
  }

  return { columnsBySheetName, rowsBySheetName };
}

function normalizeColor(raw: string | null | undefined): string | undefined {
  if (!raw) return undefined;
  const m = /^#?([0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/.exec(raw.trim());
  if (!m) return undefined;
  const hex = m[1].length === 8 ? m[1].slice(2) : m[1];
  return `#${hex.toLowerCase()}`;
}

function parseCellStylesFromContentXml(xml: string): CellStyleMeta {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const stylesByName: CellStyleMeta['stylesByName'] = {};
  const refsBySheetName: CellStyleMeta['refsBySheetName'] = {};

  for (const style of Array.from(doc.getElementsByTagName('*'))) {
    if (style.localName !== 'style') continue;
    if (style.getAttribute('style:family') !== 'table-cell') continue;
    const name = style.getAttribute('style:name');
    if (!name) continue;

    const parsed: IStyleData = {};
    const textProps = Array.from(style.children).find(
      (child) => child.localName === 'text-properties',
    );
    const paraProps = Array.from(style.children).find(
      (child) => child.localName === 'paragraph-properties',
    );
    const cellProps = Array.from(style.children).find(
      (child) => child.localName === 'table-cell-properties',
    );

    if (textProps?.getAttribute('fo:font-weight') === 'bold') parsed.bl = 1;
    const fontColor = normalizeColor(textProps?.getAttribute('fo:color'));
    if (fontColor) parsed.cl = { rgb: fontColor };

    const fillColor = normalizeColor(cellProps?.getAttribute('fo:background-color'));
    if (fillColor) parsed.bg = { rgb: fillColor };

    const hAlign = paraProps?.getAttribute('fo:text-align');
    if (hAlign === 'left') parsed.ht = 1;
    else if (hAlign === 'center') parsed.ht = 2;
    else if (hAlign === 'right') parsed.ht = 3;

    const vAlign = cellProps?.getAttribute('style:vertical-align');
    if (vAlign === 'top') parsed.vt = 1;
    else if (vAlign === 'middle') parsed.vt = 2;
    else if (vAlign === 'bottom') parsed.vt = 3;

    if (Object.keys(parsed).length > 0) stylesByName[name] = parsed;
  }

  for (const table of Array.from(doc.getElementsByTagName('*')).filter(
    (el) => el.localName === 'table',
  )) {
    const sheetName = table.getAttribute('table:name');
    if (!sheetName) continue;

    let rowIndex = 0;
    for (const rowEl of Array.from(table.children).filter((el) => el.localName === 'table-row')) {
      const rowRepeat = Number(rowEl.getAttribute('table:number-rows-repeated') ?? '1') || 1;
      for (let rowOffset = 0; rowOffset < rowRepeat; rowOffset++) {
        let colIndex = 0;
        for (const cellEl of Array.from(rowEl.children)) {
          const repeat = Number(cellEl.getAttribute('table:number-columns-repeated') ?? '1') || 1;
          const styleName = cellEl.getAttribute('table:style-name');
          const isCovered = cellEl.localName === 'covered-table-cell';
          for (let colOffset = 0; colOffset < repeat; colOffset++) {
            if (!isCovered && styleName && stylesByName[styleName]) {
              refsBySheetName[sheetName] ??= {};
              refsBySheetName[sheetName][rowIndex] ??= {};
              refsBySheetName[sheetName][rowIndex][colIndex] = styleName;
            }
            colIndex++;
          }
        }
        rowIndex++;
      }
    }
  }

  return { stylesByName, refsBySheetName };
}

function readFreezeSettings(
  buffer: ArrayBuffer,
  format: TabularFormat,
): Record<string, FreezeMeta> {
  if (format !== 'ods') return {};
  try {
    const cfb = XLSX.CFB.read(new Uint8Array(buffer), { type: 'array' });
    const settings = getZipEntry(cfb, 'settings.xml');
    if (!settings?.content) return {};
    return parseFreezeSettingsXml(decodeZipText(settings.content));
  } catch {
    return {};
  }
}

function readDimensionSettings(buffer: ArrayBuffer, format: TabularFormat): DimensionMeta {
  if (format !== 'ods') return { columnsBySheetName: {}, rowsBySheetName: {} };
  try {
    const cfb = XLSX.CFB.read(new Uint8Array(buffer), { type: 'array' });
    const content = getZipEntry(cfb, 'content.xml');
    if (!content?.content) return { columnsBySheetName: {}, rowsBySheetName: {} };
    return parseDimensionStylesFromContentXml(decodeZipText(content.content));
  } catch {
    return { columnsBySheetName: {}, rowsBySheetName: {} };
  }
}

function readCellStyles(buffer: ArrayBuffer, format: TabularFormat): CellStyleMeta {
  if (format !== 'ods') return { stylesByName: {}, refsBySheetName: {} };
  try {
    const cfb = XLSX.CFB.read(new Uint8Array(buffer), { type: 'array' });
    const content = getZipEntry(cfb, 'content.xml');
    if (!content?.content) return { stylesByName: {}, refsBySheetName: {} };
    return parseCellStylesFromContentXml(decodeZipText(content.content));
  } catch {
    return { stylesByName: {}, refsBySheetName: {} };
  }
}

function upsertZipText(cfb: ZipContainer, path: string, text: string): void {
  const existing = getZipEntry(cfb, path);
  if (existing) XLSX.CFB.utils.cfb_del(cfb, path);
  XLSX.CFB.utils.cfb_add(cfb, path, encodeZipText(text));
}

function ensureManifestEntry(manifestXml: string, fullPath: string, mediaType: string): string {
  if (manifestXml.includes(`manifest:full-path="${fullPath}"`)) return manifestXml;
  const entry = `  <manifest:file-entry manifest:full-path="${fullPath}" manifest:media-type="${mediaType}"/>\n`;
  return manifestXml.replace('</manifest:manifest>', `${entry}</manifest:manifest>`);
}

function buildFreezeSettingsXml(data: IWorkbookData): string | null {
  const sheetEntries = data.sheetOrder
    .map((sheetId) => {
      const wsd = data.sheets[sheetId];
      const freeze = wsd?.freeze;
      const sheetName = wsd?.name;
      if (!sheetName || !freeze || (freeze.xSplit <= 0 && freeze.ySplit <= 0)) return null;
      return [
        `          <config:config-item-map-entry config:name="${escapeXmlAttr(sheetName)}">`,
        `            <config:config-item config:name="HorizontalSplitMode" config:type="short">${freeze.xSplit > 0 ? 2 : 0}</config:config-item>`,
        `            <config:config-item config:name="VerticalSplitMode" config:type="short">${freeze.ySplit > 0 ? 2 : 0}</config:config-item>`,
        `            <config:config-item config:name="HorizontalSplitPosition" config:type="int">${Math.max(0, freeze.xSplit || 0)}</config:config-item>`,
        `            <config:config-item config:name="VerticalSplitPosition" config:type="int">${Math.max(0, freeze.ySplit || 0)}</config:config-item>`,
        '            <config:config-item config:name="ActiveSplitRange" config:type="short">2</config:config-item>',
        '            <config:config-item config:name="PositionLeft" config:type="int">0</config:config-item>',
        `            <config:config-item config:name="PositionRight" config:type="int">${Math.max(0, freeze.xSplit || 0)}</config:config-item>`,
        '            <config:config-item config:name="PositionTop" config:type="int">0</config:config-item>',
        `            <config:config-item config:name="PositionBottom" config:type="int">${Math.max(0, freeze.ySplit || 0)}</config:config-item>`,
        '            <config:config-item config:name="ZoomType" config:type="short">0</config:config-item>',
        '            <config:config-item config:name="ZoomValue" config:type="int">100</config:config-item>',
        '            <config:config-item config:name="PageViewZoomValue" config:type="int">60</config:config-item>',
        '            <config:config-item config:name="ShowGrid" config:type="boolean">true</config:config-item>',
        '          </config:config-item-map-entry>',
      ].join('\n');
    })
    .filter(Boolean)
    .join('\n');

  if (!sheetEntries) return null;

  const activeSheetName = data.sheets[data.sheetOrder[0]]?.name ?? 'Sheet1';
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" office:version="1.2">',
    '  <office:settings>',
    '    <config:config-item-set config:name="ooo:view-settings">',
    '      <config:config-item-map-indexed config:name="Views">',
    '        <config:config-item-map-entry>',
    '          <config:config-item config:name="ViewId" config:type="string">view1</config:config-item>',
    '          <config:config-item-map-named config:name="Tables">',
    sheetEntries,
    '          </config:config-item-map-named>',
    `          <config:config-item config:name="ActiveTable" config:type="string">${escapeXmlText(activeSheetName)}</config:config-item>`,
    '        </config:config-item-map-entry>',
    '      </config:config-item-map-indexed>',
    '    </config:config-item-set>',
    '  </office:settings>',
    '</office:document-settings>',
  ].join('\n');
}

function patchOdsSettings(buffer: ArrayBuffer, data: IWorkbookData): ArrayBuffer {
  const settingsXml = buildFreezeSettingsXml(data);
  if (!settingsXml) return buffer;

  const cfb = XLSX.CFB.read(new Uint8Array(buffer), { type: 'array' });
  upsertZipText(cfb, 'settings.xml', settingsXml);

  const manifestEntry = getZipEntry(cfb, 'META-INF/manifest.xml');
  if (manifestEntry?.content) {
    const manifestXml = decodeZipText(manifestEntry.content);
    const nextManifest = ensureManifestEntry(manifestXml, 'settings.xml', 'text/xml');
    upsertZipText(cfb, 'META-INF/manifest.xml', nextManifest);
  }

  return XLSX.CFB.write(cfb, { fileType: 'zip', type: 'array' }) as ArrayBuffer;
}

function buildOdsCellStyles(
  data: IWorkbookData,
  contentXml: string,
): { nextXml: string; changed: boolean } {
  const doc = new DOMParser().parseFromString(contentXml, 'application/xml');
  const automaticStyles = Array.from(doc.getElementsByTagName('*')).find(
    (el) => el.localName === 'automatic-styles',
  );
  if (!automaticStyles) return { nextXml: contentXml, changed: false };

  const existingDataStyleByCellStyle = new Map<string, string | undefined>();
  for (const style of Array.from(automaticStyles.children)) {
    if (style.localName !== 'style') continue;
    if (style.getAttribute('style:family') !== 'table-cell') continue;
    const styleName = style.getAttribute('style:name');
    if (!styleName) continue;
    existingDataStyleByCellStyle.set(
      styleName,
      style.getAttribute('style:data-style-name') ?? undefined,
    );
  }

  const styleNameByKey = new Map<string, string>();
  let styleCounter = 0;
  let changed = false;

  const createStyleNode = (styleName: string, style: IStyleData, dataStyleName?: string) => {
    const node = doc.createElementNS(
      'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
      'style:style',
    );
    node.setAttribute('style:name', styleName);
    node.setAttribute('style:family', 'table-cell');
    node.setAttribute('style:parent-style-name', 'Default');
    if (dataStyleName) node.setAttribute('style:data-style-name', dataStyleName);

    if (style.bg || style.vt) {
      const props = doc.createElementNS(
        'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
        'style:table-cell-properties',
      );
      if (style.bg && typeof style.bg === 'object' && 'rgb' in style.bg && style.bg.rgb) {
        props.setAttribute('fo:background-color', style.bg.rgb);
      }
      if (style.vt === 1) props.setAttribute('style:vertical-align', 'top');
      if (style.vt === 2) props.setAttribute('style:vertical-align', 'middle');
      if (style.vt === 3) props.setAttribute('style:vertical-align', 'bottom');
      if (props.attributes.length > 0) node.appendChild(props);
    }

    if (style.ht) {
      const props = doc.createElementNS(
        'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
        'style:paragraph-properties',
      );
      if (style.ht === 1) props.setAttribute('fo:text-align', 'left');
      if (style.ht === 2) props.setAttribute('fo:text-align', 'center');
      if (style.ht === 3) props.setAttribute('fo:text-align', 'right');
      node.appendChild(props);
    }

    if (style.bl || style.cl) {
      const props = doc.createElementNS(
        'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
        'style:text-properties',
      );
      if (style.bl === 1) props.setAttribute('fo:font-weight', 'bold');
      if (style.cl && typeof style.cl === 'object' && 'rgb' in style.cl && style.cl.rgb) {
        props.setAttribute('fo:color', style.cl.rgb);
      }
      if (props.attributes.length > 0) node.appendChild(props);
    }

    automaticStyles.appendChild(node);
  };

  const tables = Array.from(doc.getElementsByTagName('*')).filter((el) => el.localName === 'table');
  for (const table of tables) {
    const sheetName = table.getAttribute('table:name');
    if (!sheetName) continue;
    const sheetId = (data.sheetOrder ?? []).find((id) => data.sheets[id]?.name === sheetName);
    const sheet = sheetId ? data.sheets[sheetId] : undefined;
    if (!sheet?.cellData) continue;

    let rowIndex = 0;
    for (const rowEl of Array.from(table.children).filter((el) => el.localName === 'table-row')) {
      const rowRepeat = Number(rowEl.getAttribute('table:number-rows-repeated') ?? '1') || 1;
      if (rowRepeat !== 1) {
        rowIndex += rowRepeat;
        continue;
      }

      let colIndex = 0;
      for (const cellEl of Array.from(rowEl.children)) {
        const repeat = Number(cellEl.getAttribute('table:number-columns-repeated') ?? '1') || 1;
        if (repeat !== 1 || cellEl.localName === 'covered-table-cell') {
          colIndex += repeat;
          continue;
        }

        const cell = (
          sheet.cellData as Record<number, Record<number, { s?: string | IStyleData }>>
        )?.[rowIndex]?.[colIndex];
        const styleRef = cell?.s;
        const style =
          typeof styleRef === 'string'
            ? ((data.styles?.[styleRef] ?? undefined) as IStyleData | undefined)
            : (styleRef as IStyleData | undefined);
        if (!style) {
          colIndex++;
          continue;
        }

        const hasSubset = Boolean(style.bl || style.cl || style.bg || style.ht || style.vt);
        if (!hasSubset) {
          colIndex++;
          continue;
        }

        const existingStyleName = cellEl.getAttribute('table:style-name') ?? undefined;
        const dataStyleName = existingStyleName
          ? existingDataStyleByCellStyle.get(existingStyleName)
          : undefined;
        const key = JSON.stringify({
          bl: style.bl,
          cl: style.cl,
          bg: style.bg,
          ht: style.ht,
          vt: style.vt,
          dataStyleName,
        });
        let nextStyleName = styleNameByKey.get(key);
        if (!nextStyleName) {
          nextStyleName = `cs${styleCounter++}`;
          styleNameByKey.set(key, nextStyleName);
          createStyleNode(nextStyleName, style, dataStyleName);
        }
        cellEl.setAttribute('table:style-name', nextStyleName);
        changed = true;
        colIndex++;
      }
      rowIndex++;
    }
  }

  return {
    nextXml: changed ? new XMLSerializer().serializeToString(doc) : contentXml,
    changed,
  };
}

function patchOdsContentStyles(buffer: ArrayBuffer, data: IWorkbookData): ArrayBuffer {
  const cfb = XLSX.CFB.read(new Uint8Array(buffer), { type: 'array' });
  const contentEntry = getZipEntry(cfb, 'content.xml');
  if (!contentEntry?.content) return buffer;

  const contentXml = decodeZipText(contentEntry.content);
  const { nextXml, changed } = buildOdsCellStyles(data, contentXml);
  if (!changed) return buffer;

  upsertZipText(cfb, 'content.xml', nextXml);
  return XLSX.CFB.write(cfb, { fileType: 'zip', type: 'array' }) as ArrayBuffer;
}

function patchOdsHiddenDimensions(buffer: ArrayBuffer, data: IWorkbookData): ArrayBuffer {
  const cfb = XLSX.CFB.read(new Uint8Array(buffer), { type: 'array' });
  const contentEntry = getZipEntry(cfb, 'content.xml');
  if (!contentEntry?.content) return buffer;

  const contentXml = decodeZipText(contentEntry.content);
  const doc = new DOMParser().parseFromString(contentXml, 'application/xml');
  const tables = Array.from(doc.getElementsByTagName('*')).filter((el) => el.localName === 'table');
  let changed = false;

  for (const table of tables) {
    const sheetName = table.getAttribute('table:name');
    if (!sheetName) continue;
    const sheetId = data.sheetOrder.find((id) => data.sheets[id]?.name === sheetName);
    const sheet = sheetId ? data.sheets[sheetId] : undefined;
    if (!sheet) continue;

    const columnData = (sheet.columnData ?? {}) as Record<string, { hd?: number }>;
    const rowData = (sheet.rowData ?? {}) as Record<string, { hd?: number }>;

    let colIndex = 0;
    let rowIndex = 0;
    for (const child of Array.from(table.children)) {
      if (child.localName === 'table-column') {
        const repeated = Number(child.getAttribute('table:number-columns-repeated') ?? '1') || 1;
        let hidden = false;
        for (let i = 0; i < repeated; i++) {
          if (columnData[colIndex + i]?.hd === 1) {
            hidden = true;
            break;
          }
        }
        if (hidden) {
          child.setAttribute('table:visibility', 'collapse');
          changed = true;
        }
        colIndex += repeated;
      }

      if (child.localName === 'table-row') {
        const repeated = Number(child.getAttribute('table:number-rows-repeated') ?? '1') || 1;
        let hidden = false;
        for (let i = 0; i < repeated; i++) {
          if (rowData[rowIndex + i]?.hd === 1) {
            hidden = true;
            break;
          }
        }
        if (hidden) {
          child.setAttribute('table:visibility', 'collapse');
          changed = true;
        }
        rowIndex += repeated;
      }
    }
  }

  if (!changed) return buffer;
  upsertZipText(cfb, 'content.xml', new XMLSerializer().serializeToString(doc));
  return XLSX.CFB.write(cfb, { fileType: 'zip', type: 'array' }) as ArrayBuffer;
}

function escapeXmlText(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeXmlAttr(value: string): string {
  return escapeXmlText(value).replace(/"/g, '&quot;');
}

function mergeDefinedNamesFromOds(
  wb: XLSX.WorkBook,
  resources: IWorkbookData['resources'],
): IWorkbookData['resources'] {
  const existing = resources?.find((r) => r.name === DEFINED_NAMES_RESOURCE);
  if (existing) return resources;

  const model = (wb.Workbook?.Names ?? []) as Array<{ Name?: string; Ref?: string }>;
  if (!model.length) return resources;

  const map: Record<string, { id: string; name: string; formulaOrRefString: string }> = {};
  let i = 0;
  for (const dn of model) {
    if (!dn?.Name || !dn.Ref) continue;
    const id = `dn-${i++}`;
    map[id] = {
      id,
      name: dn.Name,
      formulaOrRefString: dn.Ref,
    };
  }
  if (Object.keys(map).length === 0) return resources;

  const next = [...(resources ?? [])];
  next.push({ name: DEFINED_NAMES_RESOURCE, data: JSON.stringify(map) });
  return next;
}

function writeDefinedNamesToOds(wb: XLSX.WorkBook, data: IWorkbookData): void {
  const res = data.resources?.find((r) => r.name === DEFINED_NAMES_RESOURCE);
  if (!res?.data) return;

  let map: Record<string, { name?: string; formulaOrRefString?: string }>;
  try {
    map = JSON.parse(res.data);
  } catch {
    return;
  }

  const names: Array<{ Name: string; Ref: string }> = [];
  for (const entry of Object.values(map ?? {})) {
    const name = entry?.name;
    const ref = entry?.formulaOrRefString;
    if (!name || !ref) continue;
    for (const piece of ref.split(',')) {
      const cleaned = piece.trim();
      if (!cleaned) continue;
      names.push({ Name: name, Ref: cleaned });
    }
  }
  if (names.length === 0) return;

  wb.Workbook ??= {};
  wb.Workbook.Names = names;
}

function mergeNotesFromOds(
  wb: XLSX.WorkBook,
  sheetIdByName: Record<string, string>,
  resources: IWorkbookData['resources'],
): IWorkbookData['resources'] {
  const existing = resources?.find((r) => r.name === NOTE_RESOURCE);
  if (existing) return resources;

  const noteData: Record<
    string,
    Record<
      number,
      Record<
        number,
        { id: string; row: number; col: number; width: number; height: number; note: string }
      >
    >
  > = {};
  let counter = 0;

  for (const [sheetName, sheet] of Object.entries(wb.Sheets ?? {})) {
    const sheetId = sheetIdByName[sheetName];
    if (!sheetId || !sheet || !sheet['!ref']) continue;
    const range = XLSX.utils.decode_range(sheet['!ref']);
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[addr];
        const text = cell?.c?.map((entry: XLSX.Comment) => entry.t).join('\n') ?? '';
        if (!text) continue;
        noteData[sheetId] ??= {};
        noteData[sheetId][r] ??= {};
        noteData[sheetId][r][c] = {
          id: `ods-note-${counter++}`,
          row: r,
          col: c,
          width: 160,
          height: 72,
          note: text,
        };
      }
    }
  }

  if (Object.keys(noteData).length === 0) return resources;
  const next = [...(resources ?? [])];
  next.push({ name: NOTE_RESOURCE, data: JSON.stringify(noteData) });
  return next;
}

function writeNotesToOds(wb: XLSX.WorkBook, data: IWorkbookData): void {
  const res = data.resources?.find((r) => r.name === NOTE_RESOURCE);
  if (!res?.data) return;

  let parsed: Record<string, Record<number, Record<number, { note?: string }>>>;
  try {
    parsed = JSON.parse(res.data);
  } catch {
    return;
  }

  const nameBySheetId = new Map<string, string>();
  for (const sheetId of data.sheetOrder) {
    const sheetName = data.sheets[sheetId]?.name;
    if (sheetName) nameBySheetId.set(sheetId, sheetName);
  }

  for (const [sheetId, rows] of Object.entries(parsed ?? {})) {
    const sheetName = nameBySheetId.get(sheetId);
    if (!sheetName) continue;
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    for (const [rKey, cols] of Object.entries(rows ?? {})) {
      const r = Number(rKey);
      for (const [cKey, note] of Object.entries(cols ?? {})) {
        const c = Number(cKey);
        const text = note?.note;
        if (!text) continue;
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = (ws[addr] ??= { t: 's', v: '' });
        cell.c = [{ a: '', t: text }];
      }
    }
  }
}

function readWorkbook(buffer: ArrayBuffer, format: TabularFormat): XLSX.WorkBook {
  if (format === 'tsv') {
    const text = new TextDecoder().decode(new Uint8Array(buffer));
    return XLSX.read(text, { type: 'string', FS: '\t' });
  }
  if (format === 'psv') {
    const text = new TextDecoder().decode(new Uint8Array(buffer));
    return XLSX.read(text, { type: 'string', FS: '|' });
  }
  if (format === 'csv') {
    const text = new TextDecoder().decode(new Uint8Array(buffer));
    return XLSX.read(text, { type: 'string' });
  }
  return XLSX.read(buffer, { type: 'array' });
}

export async function odsToWorkbookData(buffer: ArrayBuffer): Promise<IWorkbookData> {
  return tabularToWorkbookData(buffer, 'ods');
}

export async function csvToWorkbookData(buffer: ArrayBuffer): Promise<IWorkbookData> {
  return tabularToWorkbookData(buffer, 'csv');
}

export async function tsvToWorkbookData(buffer: ArrayBuffer): Promise<IWorkbookData> {
  return tabularToWorkbookData(buffer, 'tsv');
}

export async function psvToWorkbookData(buffer: ArrayBuffer): Promise<IWorkbookData> {
  return tabularToWorkbookData(buffer, 'psv');
}

async function tabularToWorkbookData(
  buffer: ArrayBuffer,
  format: TabularFormat,
): Promise<IWorkbookData> {
  const wb = readWorkbook(buffer, format);
  const freezeBySheetName = readFreezeSettings(buffer, format);
  const dimensionsBySheetName = readDimensionSettings(buffer, format);
  const cellStylesBySheetName = readCellStyles(buffer, format);
  const id = `wb-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  const sheetOrder: string[] = [];
  const sheets: IWorkbookData['sheets'] = {};
  const sheetIdByName: Record<string, string> = {};
  const styles: Record<string, IStyleData | null> = {};
  let styleCounter = 0;
  let resources: IWorkbookData['resources'] | undefined;

  const styleByKey = new Map<string, string>();
  const internStyle = (style: IStyleData | undefined): string | undefined => {
    if (!style) return undefined;
    const key = JSON.stringify(style);
    const existing = styleByKey.get(key);
    if (existing) return existing;
    const styleId = `s${styleCounter++}`;
    styleByKey.set(key, styleId);
    styles[styleId] = style;
    return styleId;
  };

  for (const name of wb.SheetNames) {
    const sheet = wb.Sheets[name];
    if (!sheet || !sheet['!ref']) {
      const sheetId = `sheet-${sheetOrder.length + 1}`;
      sheetOrder.push(sheetId);
      sheetIdByName[name] = sheetId;
      sheets[sheetId] = {
        id: sheetId,
        name,
        cellData: {},
        rowCount: INITIAL_ROWS,
        columnCount: INITIAL_COLUMNS,
      };
      continue;
    }

    const sheetId = `sheet-${sheetOrder.length + 1}`;
    sheetOrder.push(sheetId);
    sheetIdByName[name] = sheetId;

    const range = XLSX.utils.decode_range(sheet['!ref']);
    const cellData: Record<number, Record<number, ICellData>> = {};
    let maxRow = 0;
    let maxCol = 0;

    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[addr];
        if (!cell) continue;

        const cd: ICellData = {};
        // SheetJS cell value lives in .v; the type code in .t (n/s/b/d/e).
        // For our level of fidelity we treat everything as the value with no
        // explicit type tag — Univer infers from the JS value.
        if (cell.v !== undefined && cell.v !== null) {
          if (cell.v instanceof Date) {
            cd.v = cell.v.toISOString();
          } else if (
            typeof cell.v === 'number' ||
            typeof cell.v === 'string' ||
            typeof cell.v === 'boolean'
          ) {
            cd.v = cell.v;
          } else {
            cd.v = String(cell.v);
          }
        }
        // Cached formulas come through as cell.f without the leading '='.
        if (typeof cell.f === 'string' && cell.f.length > 0) {
          cd.f = cell.f.startsWith('=') ? cell.f : `=${cell.f}`;
        }
        const url = cell.l?.Target;
        if (typeof url === 'string' && url) {
          const display = typeof cd.v === 'string' && cd.v.length > 0 ? cd.v : String(cd.v ?? '');
          if (display.length > 0) cd.p = buildHyperlinkBody(display, url, nextHyperlinkId());
        }
        if (typeof cell.z === 'string' && cell.z.length > 0) {
          const styleId = internStyle({ n: { pattern: cell.z } });
          if (styleId) cd.s = styleId;
        }
        if (cd.v !== undefined || cd.f) {
          cellData[r] ??= {};
          cellData[r][c] = cd;
          if (r > maxRow) maxRow = r;
          if (c > maxCol) maxCol = c;
        }
      }
    }

    const styleRefs = cellStylesBySheetName.refsBySheetName[name] ?? {};
    for (const [rKey, cols] of Object.entries(styleRefs)) {
      const r = Number(rKey);
      for (const [cKey, styleName] of Object.entries(cols)) {
        const c = Number(cKey);
        const parsedStyle = cellStylesBySheetName.stylesByName[styleName];
        if (!parsedStyle) continue;
        const existing = cellData[r]?.[c];
        const currentStyle =
          typeof existing?.s === 'string'
            ? ((styles[existing.s] ?? undefined) as IStyleData | undefined)
            : (existing?.s as IStyleData | undefined);
        const merged = { ...(currentStyle ?? {}), ...parsedStyle };
        const styleId = internStyle(merged);
        if (!styleId) continue;
        cellData[r] ??= {};
        cellData[r][c] = { ...(existing ?? {}), s: styleId };
      }
    }

    const mergeData: IRange[] = (sheet['!merges'] ?? []).map((m) => ({
      startRow: m.s.r,
      startColumn: m.s.c,
      endRow: m.e.r,
      endColumn: m.e.c,
    }));

    sheets[sheetId] = {
      id: sheetId,
      name,
      cellData,
      mergeData,
      ...(dimensionsBySheetName.columnsBySheetName[name]
        ? { columnData: dimensionsBySheetName.columnsBySheetName[name] }
        : {}),
      ...(dimensionsBySheetName.rowsBySheetName[name]
        ? { rowData: dimensionsBySheetName.rowsBySheetName[name] }
        : {}),
      rowCount: Math.max(INITIAL_ROWS, maxRow + 1),
      columnCount: Math.max(INITIAL_COLUMNS, maxCol + 1),
      ...(freezeBySheetName[name] ? { freeze: freezeBySheetName[name] } : {}),
    };
  }

  resources = mergeNotesFromOds(wb, sheetIdByName, resources);
  resources = mergeDefinedNamesFromOds(wb, resources);

  if (sheetOrder.length === 0) {
    sheetOrder.push('sheet-1');
    sheets['sheet-1'] = {
      id: 'sheet-1',
      name: 'Sheet1',
      cellData: {},
      rowCount: INITIAL_ROWS,
      columnCount: INITIAL_COLUMNS,
    };
  }

  return {
    id,
    rev: 1,
    name: 'Untitled',
    appVersion: UNIVER_VERSION,
    locale: LocaleType.EN_US,
    styles,
    sheetOrder,
    sheets,
    ...(resources ? { resources } : {}),
  };
}

export async function workbookDataToOds(data: IWorkbookData): Promise<Blob> {
  const wb = XLSX.utils.book_new();
  writeDefinedNamesToOds(wb, data);
  const pxToPoints = (px: number) => Math.max(0, (px * 72) / 96);
  const resolveStyle = (s: string | IStyleData | undefined): IStyleData | undefined => {
    if (!s) return undefined;
    if (typeof s === 'string') return (data.styles?.[s] ?? undefined) as IStyleData | undefined;
    return s;
  };

  for (const sheetId of data.sheetOrder) {
    const wsd = data.sheets[sheetId];
    if (!wsd) continue;
    const sheet: XLSX.WorkSheet = {};

    const cellData = (wsd.cellData ?? {}) as Record<string, Record<string, SnapshotCell>>;
    let maxRow = 0;
    let maxCol = 0;
    for (const rKey of Object.keys(cellData)) {
      const r = Number(rKey);
      const row = cellData[rKey];
      for (const cKey of Object.keys(row)) {
        const c = Number(cKey);
        const cell = row[cKey];
        const addr = XLSX.utils.encode_cell({ r, c });
        const out: XLSX.CellObject = { t: 's', v: '' };
        if (cell.f) {
          out.f = cell.f.startsWith('=') ? cell.f.slice(1) : cell.f;
          if (cell.v !== undefined && cell.v !== null) {
            out.v = cell.v;
            out.t = typeof cell.v === 'number' ? 'n' : typeof cell.v === 'boolean' ? 'b' : 's';
          }
        } else if (cell.v !== undefined && cell.v !== null) {
          out.v = cell.v;
          out.t = typeof cell.v === 'number' ? 'n' : typeof cell.v === 'boolean' ? 'b' : 's';
        } else {
          continue;
        }
        const hyperlink = getHyperlinkFromCell(cell);
        if (hyperlink?.url) out.l = { Target: hyperlink.url };
        const style = resolveStyle(cell.s);
        if (style?.n?.pattern) out.z = style.n.pattern;
        sheet[addr] = out;
        if (r > maxRow) maxRow = r;
        if (c > maxCol) maxCol = c;
      }
    }

    if (Array.isArray(wsd.mergeData) && wsd.mergeData.length) {
      sheet['!merges'] = wsd.mergeData.map((m) => ({
        s: { r: m.startRow, c: m.startColumn },
        e: { r: m.endRow, c: m.endColumn },
      }));
    }

    const columnData = (wsd.columnData ?? {}) as Record<string, { w?: number; hd?: number }>;
    const colEntries = Object.entries(columnData)
      .filter(([, meta]) => (typeof meta?.w === 'number' && meta.w > 0) || meta?.hd === 1)
      .sort((a, b) => Number(a[0]) - Number(b[0]));
    if (colEntries.length > 0) {
      const cols: Array<{ wpx?: number; wch?: number; hidden?: boolean }> = [];
      for (const [cKey, meta] of colEntries) {
        const c = Number(cKey);
        cols[c] = {
          ...(typeof meta?.w === 'number' && meta.w > 0
            ? { wpx: meta.w, wch: Math.max(0, meta.w / 7) }
            : {}),
          ...(meta?.hd === 1 ? { hidden: true } : {}),
        };
        if (c > maxCol) maxCol = c;
      }
      sheet['!cols'] = cols;
    }

    const rowData = (wsd.rowData ?? {}) as Record<string, { h?: number; hd?: number }>;
    const rowEntries = Object.entries(rowData)
      .filter(([, meta]) => (typeof meta?.h === 'number' && meta.h > 0) || meta?.hd === 1)
      .sort((a, b) => Number(a[0]) - Number(b[0]));
    if (rowEntries.length > 0) {
      const rows: Array<{ hpx?: number; hpt?: number; hidden?: boolean }> = [];
      for (const [rKey, meta] of rowEntries) {
        const r = Number(rKey);
        rows[r] = {
          ...(typeof meta?.h === 'number' && meta.h > 0
            ? { hpx: meta.h, hpt: pxToPoints(meta.h) }
            : {}),
          ...(meta?.hd === 1 ? { hidden: true } : {}),
        };
        if (r > maxRow) maxRow = r;
      }
      sheet['!rows'] = rows;
    }

    sheet['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: maxRow, c: maxCol },
    });

    XLSX.utils.book_append_sheet(wb, sheet, wsd.name ?? sheetId);
  }

  writeNotesToOds(wb, data);

  const out = XLSX.write(wb, { type: 'array', bookType: 'ods' }) as ArrayBuffer;
  const withStyles = patchOdsContentStyles(out, data);
  const withHidden = patchOdsHiddenDimensions(withStyles, data);
  const patched = patchOdsSettings(withHidden, data);
  return new Blob([patched], {
    type: 'application/vnd.oasis.opendocument.spreadsheet',
  });
}

/**
 * Render the first sheet of the workbook as a CSV or TSV string. Both formats
 * are flat — they don't carry multi-sheet, styles, formulas (we emit cached
 * values), or merges. We export the active sheet only, matching what
 * LibreOffice / Excel do for CSV.
 */
export async function workbookDataToDelimited(
  data: IWorkbookData,
  format: 'csv' | 'tsv' | 'psv',
): Promise<Blob> {
  const firstId = data.sheetOrder[0];
  const wsd = data.sheets[firstId];
  if (!wsd) return new Blob([''], { type: 'text/plain' });

  const cellData = (wsd.cellData ?? {}) as Record<
    string,
    Record<string, { v?: string | number | boolean; f?: string }>
  >;
  let maxRow = 0;
  let maxCol = 0;
  for (const rKey of Object.keys(cellData)) {
    const r = Number(rKey);
    if (r > maxRow) maxRow = r;
    for (const cKey of Object.keys(cellData[rKey])) {
      const c = Number(cKey);
      if (c > maxCol) maxCol = c;
    }
  }

  const sep = format === 'tsv' ? '\t' : format === 'psv' ? '|' : ',';
  const lines: string[] = [];
  for (let r = 0; r <= maxRow; r++) {
    const row = cellData[r];
    const cells: string[] = [];
    for (let c = 0; c <= maxCol; c++) {
      const cell = row?.[c];
      const raw = cell?.v;
      cells.push(escapeField(raw, sep));
    }
    lines.push(cells.join(sep));
  }

  const text = lines.join('\r\n');
  return new Blob([text], {
    type: format === 'csv' ? 'text/csv;charset=utf-8' : 'text/tab-separated-values;charset=utf-8',
  });
}

function escapeField(v: unknown, sep: string): string {
  if (v === undefined || v === null) return '';
  const s = typeof v === 'string' ? v : String(v);
  // CSV (and TSV with the same rules) require quoting when the field
  // contains the separator, a quote, or a line break.
  if (s.includes(sep) || s.includes('"') || s.includes('\n') || s.includes('\r')) {
    return `"${s.replace(/"/g, '""')}"`;
  }
  return s;
}
