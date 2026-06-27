import { useMemo, useState } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { Dialog } from './Dialog';
import {
  buildDelimiterMask,
  hasActiveDelimiter,
  splitPreview,
  type DelimiterOptions,
} from './split-text';

/**
 * Text to Columns (Excel's Data → Text to Columns, delimited mode). Splits a
 * single column of text into several columns on the chosen delimiters, with a
 * live preview. The split itself runs through the Univer fork's
 * `sheet.command.split-text-to-columns` (inserts columns as needed, undoable);
 * this dialog just gathers the options + previews using the same regex.
 */

type Props = {
  api: FUniver;
  onClose: () => void;
};

const SPLIT_CMD = 'sheet.command.split-text-to-columns';
const PREVIEW_ROWS = 6;

export function TextToColumnsDialog({ api, onClose }: Props) {
  const wb = api.getActiveWorkbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ws = wb?.getActiveSheet() as any;
  const range = ws?.getActiveRange?.();

  const width: number = range?.getWidth?.() ?? 0;
  const singleColumn = width === 1 && !!range;

  const samples: string[] = useMemo(() => {
    if (!singleColumn) return [];
    const vals = (range.getValues?.() ?? []) as unknown[][];
    return vals
      .slice(0, PREVIEW_ROWS)
      .map((row) => {
        const v = row?.[0];
        return v == null ? '' : String(v);
      })
      .filter((s, i, arr) => i < PREVIEW_ROWS && (s !== '' || i < arr.length));
  }, [range, singleColumn]);

  const [opts, setOpts] = useState<DelimiterOptions>({
    tab: false,
    comma: true,
    semicolon: false,
    space: false,
    custom: '',
  });
  const [treatAsOne, setTreatAsOne] = useState(false);

  const preview = useMemo(
    () => splitPreview(samples, opts, treatAsOne),
    [samples, opts, treatAsOne],
  );
  const previewCols = preview.reduce((m, r) => Math.max(m, r.length), 0);
  const canSplit = singleColumn && hasActiveDelimiter(opts);

  const finish = () => {
    if (!canSplit || !range) return;
    const startRow: number = range.getRow();
    const startColumn: number = range.getColumn();
    const endRow: number = startRow + range.getHeight() - 1;
    void api.executeCommand(SPLIT_CMD, {
      unitId: wb!.getId(),
      subUnitId: ws.getSheetId(),
      range: { startRow, startColumn, endRow, endColumn: startColumn },
      delimiter: buildDelimiterMask(opts),
      customDelimiter: opts.custom ? opts.custom[0] : undefined,
      treatMultipleDelimitersAsOne: treatAsOne,
    });
    onClose();
  };

  const checkbox = (key: 'tab' | 'comma' | 'semicolon' | 'space', label: string) => (
    <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 13 }}>
      <input
        type="checkbox"
        data-testid={`ttc-${key}`}
        checked={opts[key]}
        onChange={(e) => setOpts((o) => ({ ...o, [key]: e.target.checked }))}
      />
      {label}
    </label>
  );

  return (
    <Dialog
      title="Text to Columns"
      onClose={onClose}
      data-testid="text-to-columns-dialog"
      footer={
        <>
          <button
            type="button"
            className="btn-secondary"
            data-testid="ttc-cancel"
            onClick={onClose}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="ttc-finish"
            onClick={finish}
            disabled={!canSplit}
          >
            Finish
          </button>
        </>
      }
    >
      {!singleColumn ? (
        <div data-testid="ttc-single-col-notice" style={{ fontSize: 13 }}>
          Select a single column of cells to split.
        </div>
      ) : (
        <>
          <div style={{ fontSize: 12, color: 'var(--cs-chrome-muted, #8a8886)', marginBottom: 6 }}>
            Delimiters
          </div>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px 16px', marginBottom: 8 }}>
            {checkbox('tab', 'Tab')}
            {checkbox('semicolon', 'Semicolon')}
            {checkbox('comma', 'Comma')}
            {checkbox('space', 'Space')}
            <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 13 }}>
              <input
                type="checkbox"
                data-testid="ttc-other"
                checked={!!opts.custom}
                onChange={(e) =>
                  setOpts((o) => ({ ...o, custom: e.target.checked ? o.custom || '|' : '' }))
                }
              />
              Other:
              <input
                type="text"
                data-testid="ttc-other-input"
                value={opts.custom ?? ''}
                maxLength={1}
                onChange={(e) => setOpts((o) => ({ ...o, custom: e.target.value.slice(0, 1) }))}
                style={{ width: 32, textAlign: 'center' }}
              />
            </label>
          </div>
          <label
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: 6,
              fontSize: 13,
              marginBottom: 10,
            }}
          >
            <input
              type="checkbox"
              data-testid="ttc-collapse"
              checked={treatAsOne}
              onChange={(e) => setTreatAsOne(e.target.checked)}
            />
            Treat consecutive delimiters as one
          </label>

          <div style={{ fontSize: 12, color: 'var(--cs-chrome-muted, #8a8886)', marginBottom: 4 }}>
            Data preview
          </div>
          <div
            style={{
              overflowX: 'auto',
              border: '1px solid var(--cs-chrome-border, #e6e9ee)',
              borderRadius: 6,
            }}
          >
            <table
              data-testid="ttc-preview"
              style={{ borderCollapse: 'collapse', fontSize: 12, width: '100%' }}
            >
              <tbody>
                {preview.map((row, ri) => (
                  <tr key={ri}>
                    {Array.from({ length: previewCols }, (_, ci) => (
                      <td
                        key={ci}
                        style={{
                          border: '1px solid var(--cs-chrome-border, #eef1f5)',
                          padding: '2px 6px',
                          whiteSpace: 'nowrap',
                          maxWidth: 140,
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                        }}
                      >
                        {row[ci] ?? ''}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ fontSize: 11, color: 'var(--cs-chrome-muted, #8a8886)', marginTop: 6 }}>
            The split overwrites cells to the right of the selection.
          </div>
        </>
      )}
    </Dialog>
  );
}
