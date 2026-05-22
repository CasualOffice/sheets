import { useEffect, useRef, useState } from 'react';
import { useUniverAPI } from '../use-univer';
import { useUI } from '../use-ui';
import { useBusy } from '../busy-context';
import { useActiveCellState, type HAlign, type VAlign } from '../hooks/useActiveCellState';
import { Icon } from './Icon';
import {
  NUMBER_FORMATS,
  adjustFontSize,
  copy,
  cut,
  decreaseDecimal,
  increaseDecimal,
  paste,
  redo,
  setAlignment,
  setBorders,
  setFillColor,
  setFontColor,
  setFontFamily,
  setFontSize,
  setNumberFormat,
  setNumberFormatByKey,
  setVerticalAlignment,
  startFormatPainter,
  toggleBold,
  toggleItalic,
  toggleMerge,
  toggleStrikethrough,
  toggleUnderline,
  toggleWrap,
  undo,
  type BorderChoice,
  type NumberFormatKey,
} from './home-tab-actions';
import {
  TABLE_THEMES,
  applyAutoFunction,
  formatAsTable,
  insertComment,
  insertHyperlink,
  sortRange,
  toggleCommentPanel,
  toggleFilter,
  type TableThemeId,
} from './tab-actions';
import {
  BigToolbarButton,
  BordersControl,
  RibbonGroup,
  ToolbarButton,
  ToolbarColorButton,
  ToolbarDropdown,
  ToolbarSelect,
} from './RibbonControls';
import { DEFAULT_BORDER_COLOR } from './home-tab-actions';

/**
 * Fixed single-row toolbar — every group lays out its controls in one
 * horizontal row, so all icons share a baseline. No 2-row groups → no
 * visual asymmetry. Less-common features live in the top MenuBar.
 */

const FONT_FAMILIES = [
  'Calibri',
  'Arial',
  'Helvetica',
  'Inter',
  'Times New Roman',
  'Georgia',
  'Verdana',
  'Courier New',
  'JetBrains Mono',
];
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48];

const NUMBER_FORMAT_OPTIONS: { value: NumberFormatKey; label: string }[] = [
  { value: 'general', label: 'General' },
  { value: 'integer', label: 'Number (no decimals)' },
  { value: 'number', label: 'Number (2 decimals)' },
  { value: 'currency', label: 'Currency' },
  { value: 'accounting', label: 'Accounting' },
  { value: 'percent', label: 'Percent' },
  { value: 'date', label: 'Date (yyyy-mm-dd)' },
  { value: 'time', label: 'Time (hh:mm:ss)' },
  { value: 'scientific', label: 'Scientific' },
  { value: 'text', label: 'Text' },
];

function detectFormatKey(pattern: string): NumberFormatKey {
  const patterns: Record<NumberFormatKey, string> = {
    general: '',
    number: '#,##0.00',
    integer: '#,##0',
    currency: '"$"#,##0.00',
    accounting: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
    percent: '0.00%',
    date: 'yyyy-mm-dd',
    time: 'hh:mm:ss',
    scientific: '0.00E+00',
    text: '@',
  };
  for (const k of Object.keys(patterns) as NumberFormatKey[]) if (patterns[k] === pattern) return k;
  return 'general';
}

/**
 * Track whether the toolbar's horizontal scroller has more content to
 * either side. Drives the left/right overflow chevrons so users on a
 * narrow window can see that more buttons exist (instead of guessing
 * from a hidden scrollbar).
 */
function useScrollOverflow(ref: React.RefObject<HTMLElement>) {
  const [state, setState] = useState({ left: false, right: false });
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const update = () => {
      // 2 px slack so subpixel rounding doesn't leave the right chevron
      // visible at the actual scroll-end.
      const left = el.scrollLeft > 0;
      const right = el.scrollLeft + el.clientWidth < el.scrollWidth - 2;
      setState((prev) => (prev.left === left && prev.right === right ? prev : { left, right }));
    };
    update();
    el.addEventListener('scroll', update, { passive: true });
    const ro = new ResizeObserver(update);
    ro.observe(el);
    window.addEventListener('resize', update);
    return () => {
      el.removeEventListener('scroll', update);
      ro.disconnect();
      window.removeEventListener('resize', update);
    };
  }, [ref]);
  return state;
}

export function Toolbar() {
  const api = useUniverAPI();
  const ui = useUI();
  const busy = useBusy();
  const state = useActiveCellState();
  const ready = Boolean(api) && state.ready;

  const innerRef = useRef<HTMLDivElement>(null);
  const overflow = useScrollOverflow(innerRef);
  const scrollBy = (dx: number) =>
    innerRef.current?.scrollBy({ left: dx, behavior: 'smooth' });

  return (
    <nav className="toolbar" data-testid="toolbar" aria-label="Toolbar">
      {overflow.left && (
        <button
          type="button"
          className="toolbar__overflow toolbar__overflow--left"
          data-testid="toolbar-overflow-left"
          aria-label="Scroll toolbar left"
          tabIndex={-1}
          onClick={() => scrollBy(-240)}
        >
          <Icon name="chevron_left" size="sm" />
        </button>
      )}
      {overflow.right && (
        <button
          type="button"
          className="toolbar__overflow toolbar__overflow--right"
          data-testid="toolbar-overflow-right"
          aria-label="Scroll toolbar right — more buttons available"
          tabIndex={-1}
          onClick={() => scrollBy(240)}
        >
          <Icon name="chevron_right" size="sm" />
        </button>
      )}
      <div className="toolbar__inner" ref={innerRef}>
        {/* ── History — narrow group, just the two arrows. */}
        <RibbonGroup
          label="History"
          row1={<ToolbarButton id="undo" label="Undo (Ctrl+Z)" icon="undo" disabled={!ready} onClick={() => api && undo(api)} />}
          row2={<ToolbarButton id="redo" label="Redo (Ctrl+Y)" icon="redo" disabled={!ready} onClick={() => api && redo(api)} />}
        />

        {/* ── Clipboard — Paste is the BIG primary button (Excel
              canonical). Cut/Copy/Painter as small buttons stacked
              on the right. */}
        <RibbonGroup
          label="Clipboard"
          lead={
            <BigToolbarButton
              id="paste"
              label="Paste (Ctrl+V)"
              icon="content_paste"
              disabled={!ready}
              onClick={() => api && paste(api)}
            />
          }
          row1={<>
            <ToolbarButton id="cut" label="Cut (Ctrl+X)" icon="content_cut" disabled={!ready} onClick={() => api && cut(api)} />
            <ToolbarButton id="copy" label="Copy (Ctrl+C)" icon="content_copy" disabled={!ready} onClick={() => api && copy(api)} />
          </>}
          row2={
            <ToolbarButton
              id="format-painter"
              label={state.isFormatPainterActive ? 'Format Painter (armed)' : 'Format Painter'}
              icon="format_paint"
              pressed={state.isFormatPainterActive}
              disabled={!ready}
              onClick={() => api && startFormatPainter(api)}
            />
          }
        />

        {/* ── Font — Excel's layout exactly:
              Row 1: Font name dropdown + size dropdown + A↑ / A↓ steppers.
              Row 2: B I U S + Font color + Fill + Borders.
            Putting the wide dropdowns on row 1 and ALL small icons on
            row 2 keeps both rows close in width — no empty space under
            the dropdowns. */}
        <RibbonGroup
          label="Font"
          row1={<>
            <ToolbarSelect
              id="font-family"
              label="Font family"
              value={state.fontFamily || 'Calibri'}
              options={FONT_FAMILIES.map((f) => ({ value: f, label: f }))}
              width={130}
              disabled={!ready}
              onChange={(v) => api && setFontFamily(api, v)}
            />
            <ToolbarSelect
              id="font-size"
              label="Font size"
              value={String(state.fontSize || 11)}
              options={FONT_SIZES.map((s) => ({ value: String(s), label: String(s) }))}
              width={56}
              disabled={!ready}
              onChange={(v) => api && setFontSize(api, Number(v))}
            />
            <ToolbarButton id="font-size-up" label="Increase font size" icon="text_increase" disabled={!ready} onClick={() => api && adjustFontSize(api, +1)} />
            <ToolbarButton id="font-size-down" label="Decrease font size" icon="text_decrease" disabled={!ready} onClick={() => api && adjustFontSize(api, -1)} />
          </>}
          row2={<>
            <ToolbarButton id="bold" label="Bold (Ctrl+B)" icon="format_bold" pressed={state.isBold} disabled={!ready} onClick={() => api && toggleBold(api, state.isBold)} />
            <ToolbarButton id="italic" label="Italic (Ctrl+I)" icon="format_italic" pressed={state.isItalic} disabled={!ready} onClick={() => api && toggleItalic(api, state.isItalic)} />
            <ToolbarButton id="underline" label="Underline (Ctrl+U)" icon="format_underlined" pressed={state.isUnderline} disabled={!ready} onClick={() => api && toggleUnderline(api, state.isUnderline)} />
            <ToolbarButton id="strikethrough" label="Strikethrough" icon="format_strikethrough" pressed={state.isStrike} disabled={!ready} onClick={() => api && toggleStrikethrough(api, state.isStrike)} />
            <ToolbarColorButton
              id="font-color"
              label="Font color"
              icon="format_color_text"
              value={state.fontColor}
              defaultColor="#000000"
              disabled={!ready}
              onChange={(c) => api && setFontColor(api, c)}
            />
            <ToolbarColorButton
              id="fill-color"
              label="Fill / highlight color"
              icon="format_color_fill"
              value={state.fillColor}
              defaultColor="#ffeb3b"
              disabled={!ready}
              onChange={(c) => api && setFillColor(api, c)}
            />
            <BordersControl
              label="Borders"
              icon="border_all"
              disabled={!ready}
              defaultColor={DEFAULT_BORDER_COLOR}
              items={[
                { id: 'all', label: 'All borders', icon: 'border_all' },
                { id: 'outside', label: 'Outside borders', icon: 'border_outer' },
                { id: 'top', label: 'Top border', icon: 'border_top' },
                { id: 'bottom', label: 'Bottom border', icon: 'border_bottom' },
                { id: 'left', label: 'Left border', icon: 'border_left' },
                { id: 'right', label: 'Right border', icon: 'border_right' },
                { id: 'none', label: 'No border', icon: 'border_clear' },
              ]}
              onDefault={(color, weight) => api && setBorders(api, 'all', color, weight)}
              onChoose={(choice, color, weight) =>
                api && setBorders(api, choice as BorderChoice, color, weight)
              }
            />
          </>}
        />

        {/* ── Alignment — H aligns + Wrap on row 1, V aligns + Merge on row 2. */}
        <RibbonGroup
          label="Alignment"
          row1={<>
            <ToolbarButton id="align-left" label="Align left" icon="format_align_left" pressed={state.align === 'left'} disabled={!ready} onClick={() => api && setAlignment(api, 'left')} />
            <ToolbarButton id="align-center" label="Center" icon="format_align_center" pressed={state.align === 'center'} disabled={!ready} onClick={() => api && setAlignment(api, 'center')} />
            <ToolbarButton id="align-right" label="Align right" icon="format_align_right" pressed={state.align === ('right' as HAlign)} disabled={!ready} onClick={() => api && setAlignment(api, 'right')} />
            <ToolbarButton id="wrap-text" label="Wrap text" icon="wrap_text" pressed={state.isWrapped} disabled={!ready} onClick={() => api && toggleWrap(api, state.isWrapped)} />
          </>}
          row2={<>
            <ToolbarButton id="align-top" label="Align top" icon="vertical_align_top" pressed={state.vAlign === 'top'} disabled={!ready} onClick={() => api && setVerticalAlignment(api, 'top')} />
            <ToolbarButton id="align-middle" label="Align middle" icon="vertical_align_center" pressed={state.vAlign === 'middle'} disabled={!ready} onClick={() => api && setVerticalAlignment(api, 'middle')} />
            <ToolbarButton id="align-bottom" label="Align bottom" icon="vertical_align_bottom" pressed={state.vAlign === ('bottom' as VAlign)} disabled={!ready} onClick={() => api && setVerticalAlignment(api, 'bottom')} />
            <ToolbarButton
              id="merge-cells"
              label={state.isMerged ? 'Unmerge cells' : 'Merge & Center'}
              icon={state.isMerged ? 'call_split' : 'cell_merge'}
              pressed={state.isMerged}
              disabled={!ready || (!state.isMerged && !state.isMultiCell)}
              onClick={() => api && toggleMerge(api, state.isMerged)}
            />
          </>}
        />

        {/* ── Number — wide format dropdown row 1, quick toggles row 2. */}
        <RibbonGroup
          label="Number"
          row1={
            <ToolbarSelect
              id="num-format"
              label="Number format"
              value={detectFormatKey(state.numberFormat)}
              options={NUMBER_FORMAT_OPTIONS}
              width={108}
              disabled={!ready}
              onChange={(v) => api && setNumberFormatByKey(api, v as NumberFormatKey)}
            />
          }
          row2={<>
            <ToolbarButton id="numfmt-currency" label="Currency" icon="attach_money" pressed={state.numberFormat === NUMBER_FORMATS.currency} disabled={!ready} onClick={() => api && setNumberFormat(api, NUMBER_FORMATS.currency)} />
            <ToolbarButton id="numfmt-percent" label="Percent" icon="percent" pressed={state.numberFormat === NUMBER_FORMATS.percent} disabled={!ready} onClick={() => api && setNumberFormat(api, NUMBER_FORMATS.percent)} />
            <ToolbarButton id="decimal-up" label="Increase decimals" icon="add" disabled={!ready} onClick={() => api && increaseDecimal(api)} />
            <ToolbarButton id="decimal-down" label="Decrease decimals" icon="remove" disabled={!ready} onClick={() => api && decreaseDecimal(api)} />
          </>}
        />

        {/* ── Insert — AutoSum + Chart + PivotTable + Format as Table on row 1.
              Sort & Filter + Link + Comment on row 2.
              Chart and Pivot dispatch a CustomEvent that MenuBar listens
              for to open the existing InsertChartDialog / InsertPivotDialog
              — keeps dialog state in one place. */}
        <RibbonGroup
          label="Insert"
          row1={<>
            <ToolbarDropdown
              id="auto-sum"
              label="AutoSum"
              icon="functions"
              disabled={!ready}
              items={[
                { id: 'SUM', label: 'Sum', icon: 'functions' },
                { id: 'AVERAGE', label: 'Average', icon: 'bar_chart' },
                { id: 'COUNT', label: 'Count', icon: 'numbers' },
                { id: 'MAX', label: 'Max', icon: 'arrow_upward' },
                { id: 'MIN', label: 'Min', icon: 'arrow_downward' },
              ]}
              onDefault={() => api && applyAutoFunction(api, 'SUM')}
              onChoose={(id) => api && applyAutoFunction(api, id as 'SUM' | 'AVERAGE' | 'COUNT' | 'MIN' | 'MAX')}
            />
            <ToolbarButton
              id="insert-chart"
              label="Insert chart"
              icon="bar_chart"
              disabled={!ready}
              onClick={() => document.dispatchEvent(new CustomEvent('casual-open-insert-chart'))}
            />
            <ToolbarButton
              id="insert-pivot"
              label="Insert PivotTable"
              icon="pivot_table_chart"
              disabled={!ready}
              onClick={() => document.dispatchEvent(new CustomEvent('casual-open-insert-pivot'))}
            />
            <ToolbarDropdown
              id="format-as-table"
              label="Format as Table"
              icon="table_rows"
              disabled={!ready}
              items={[
                { id: 'plain', label: 'Plain (no style)', icon: 'grid_on' },
                ...TABLE_THEMES.map((t) => ({
                  id: t.id,
                  label: t.label,
                  icon: 'table_chart',
                })),
              ]}
              onDefault={() => {
                if (!api) return;
                void busy.runBusy('Creating table…', () => formatAsTable(api, 'table-default-0'));
              }}
              onChoose={(id) => {
                if (!api) return;
                void busy.runBusy('Creating table…', () =>
                  id === 'plain'
                    ? formatAsTable(api, undefined)
                    : formatAsTable(api, id as TableThemeId),
                );
              }}
            />
          </>}
          row2={<>
            <ToolbarDropdown
              id="sort-filter"
              label="Sort & Filter"
              icon="filter_list"
              disabled={!ready}
              items={[
                { id: 'sort-asc',     label: 'Sort A → Z',          icon: 'arrow_downward' },
                { id: 'sort-desc',    label: 'Sort Z → A',          icon: 'arrow_upward' },
                { id: 'toggle-filter', label: 'Toggle filter',      icon: 'filter_alt' },
              ]}
              onDefault={() => api && state.isMultiCell && sortRange(api, true)}
              onChoose={(id) => {
                if (!api) return;
                if (id === 'sort-asc') sortRange(api, true);
                else if (id === 'sort-desc') sortRange(api, false);
                else if (id === 'toggle-filter') void toggleFilter(api);
              }}
            />
            <ToolbarButton id="insert-hyperlink" label="Insert hyperlink (Ctrl+K)" icon="link" disabled={!ready} onClick={() => api && insertHyperlink(api)} />
            <ToolbarButton id="insert-comment" label="Insert comment" icon="add_comment" disabled={!ready} onClick={() => api && insertComment(api)} />
          </>}
        />

        {/* ── Panels — quick-toggle for the side-panel surfaces. Each
              button stays pressed while its panel is open so the user
              can see which panels are active at a glance. */}
        <RibbonGroup
          label="Panels"
          row1={<>
            <ToolbarButton
              id="panel-tables"
              label={ui.tablesPanelVisible ? 'Hide Tables panel' : 'Tables panel'}
              icon="table_view"
              pressed={ui.tablesPanelVisible}
              onClick={ui.toggleTablesPanel}
            />
            <ToolbarButton
              id="panel-charts"
              label={ui.chartsPanelVisible ? 'Hide Charts panel' : 'Charts panel'}
              icon="analytics"
              pressed={ui.chartsPanelVisible}
              onClick={ui.toggleChartsPanel}
            />
          </>}
          row2={<>
            <ToolbarButton
              id="panel-outline"
              label={ui.outlinePanelVisible ? 'Hide Outline panel' : 'Outline panel'}
              icon="format_list_bulleted"
              pressed={ui.outlinePanelVisible}
              onClick={ui.toggleOutlinePanel}
            />
            <ToolbarButton
              id="panel-comments"
              label="Comments panel"
              icon="forum"
              onClick={() => api && toggleCommentPanel(api)}
            />
          </>}
        />
      </div>
    </nav>
  );
}
