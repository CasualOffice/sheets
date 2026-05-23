import { useState } from 'react';
import { Dialog } from './Dialog';
import type { PrintMarginPreset, PrintOptions, PrintOrientation } from './print';

type Props = {
  initial: PrintOptions;
  onCancel: () => void;
  onPrint: (options: PrintOptions) => void;
};

const MARGIN_LABELS: Record<PrintMarginPreset, string> = {
  narrow: 'Narrow (6 mm)',
  normal: 'Normal (18 mm)',
  wide: 'Wide (25 mm)',
};

/**
 * Page Setup dialog — surfaces before File → Print / Ctrl+P. Sets paper
 * orientation and margins, then defers to the browser's native print dialog
 * for everything else (paper size, headers/footers, printer choice). Mirrors
 * Excel's "Page Setup → Print" flow without trying to clone the full panel.
 *
 * Settings persist in localStorage between sessions; this dialog reads them
 * from `loadPrintOptions()` and writes them back via `savePrintOptions()`
 * when the user clicks Print.
 */
export function PageSetupDialog({ initial, onCancel, onPrint }: Props) {
  const [orientation, setOrientation] = useState<PrintOrientation>(initial.orientation);
  const [margins, setMargins] = useState<PrintMarginPreset>(initial.margins);
  const [printArea, setPrintArea] = useState<string>(initial.printArea ?? '');

  const confirm = () =>
    onPrint({
      orientation,
      margins,
      printArea: printArea.trim() || null,
    });

  return (
    <Dialog
      title="Page Setup"
      onClose={onCancel}
      data-testid="page-setup-dialog"
      footer={
        <>
          <button
            type="button"
            className="btn-secondary"
            data-testid="page-setup-cancel"
            onClick={onCancel}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="page-setup-print"
            onClick={confirm}
          >
            Print
          </button>
        </>
      }
    >
      <div className="page-setup">
        <fieldset className="page-setup__group">
          <legend className="page-setup__legend">Orientation</legend>
          <div className="page-setup__segment" role="radiogroup" aria-label="Orientation">
            <label
              className={`page-setup__option${orientation === 'portrait' ? ' page-setup__option--active' : ''}`}
              data-testid="page-setup-orientation-portrait"
            >
              <input
                type="radio"
                name="orientation"
                value="portrait"
                checked={orientation === 'portrait'}
                onChange={() => setOrientation('portrait')}
              />
              <span>Portrait</span>
            </label>
            <label
              className={`page-setup__option${orientation === 'landscape' ? ' page-setup__option--active' : ''}`}
              data-testid="page-setup-orientation-landscape"
            >
              <input
                type="radio"
                name="orientation"
                value="landscape"
                checked={orientation === 'landscape'}
                onChange={() => setOrientation('landscape')}
              />
              <span>Landscape</span>
            </label>
          </div>
        </fieldset>

        <fieldset className="page-setup__group">
          <legend className="page-setup__legend">Margins</legend>
          <select
            className="page-setup__select"
            value={margins}
            data-testid="page-setup-margins"
            onChange={(e) => setMargins(e.target.value as PrintMarginPreset)}
          >
            {(Object.keys(MARGIN_LABELS) as PrintMarginPreset[]).map((m) => (
              <option key={m} value={m}>
                {MARGIN_LABELS[m]}
              </option>
            ))}
          </select>
        </fieldset>

        <fieldset className="page-setup__group">
          <legend className="page-setup__legend">Print area</legend>
          <input
            type="text"
            className="page-setup__input"
            data-testid="page-setup-print-area"
            value={printArea}
            placeholder="e.g. A1:D20 — leave blank to print the used range"
            onChange={(e) => setPrintArea(e.target.value)}
            spellCheck={false}
          />
        </fieldset>

        <p className="page-setup__note">
          Paper size, headers, and footers come from your browser's print
          dialog after you click Print.
        </p>
      </div>
    </Dialog>
  );
}
