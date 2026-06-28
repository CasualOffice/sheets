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

import { Dialog } from '../shell/Dialog';
import type { DrillDownResult } from './drill-down';

type Props = {
  result: DrillDownResult;
  onClose: () => void;
};

/**
 * Modal that lists the source records contributing to a pivot result
 * cell. Excel ships these to a new worksheet; the popup model is
 * lower friction for one-off inspection and avoids sheet sprawl. The
 * scrollable body caps at a fixed height so the dialog stays usable
 * for thousands of contributing rows.
 */
export function DrillDownDialog({ result, onClose }: Props) {
  return (
    <Dialog
      title={`Drill down · ${result.summary}`}
      onClose={onClose}
      data-testid="drill-down-dialog"
      footer={
        <button
          type="button"
          className="btn-primary"
          data-testid="drill-down-close"
          onClick={onClose}
        >
          Close
        </button>
      }
    >
      <div className="drill-down" data-testid="drill-down-body">
        {result.rows.length === 0 ? (
          <p className="drill-down__empty">
            No source rows contributed to this cell. Either the
            aggregation found zero matches, or the source range no
            longer overlaps the recorded extent.
          </p>
        ) : (
          <div className="drill-down__scroller">
            <table className="drill-down__table">
              <thead>
                <tr>
                  {result.headers.map((h, i) => (
                    <th key={i}>{h || `Column ${i + 1}`}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {result.rows.map((row, ri) => (
                  <tr key={ri} data-testid="drill-down-row">
                    {row.map((v, ci) => (
                      <td key={ci}>{v == null ? '' : String(v)}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </Dialog>
  );
}
