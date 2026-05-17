import { useEffect, useState } from 'react';
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { useLoading } from '../loading-context';
import { loadSpreadsheetFile } from '../shell/file-actions';

const SUPPORTED_EXTENSIONS = ['xlsx', 'ods', 'csv', 'tsv', 'tab'];

/**
 * Returns whether the user is currently dragging files over the window. The
 * caller renders an overlay accordingly. The hook also intercepts the actual
 * `drop` event on `window` and routes the first supported file through the
 * shared open flow.
 *
 * Implementation note: `dragenter` / `dragleave` fire for every nested
 * element, so we use a counter rather than a boolean to know when the user
 * has actually left the window. Without the counter the overlay flickers as
 * the user moves between child elements.
 */
export function useFileDrop(): boolean {
  const api = useUniverAPI();
  const workbook = useWorkbook();
  const loading = useLoading();
  const [dragging, setDragging] = useState(false);

  useEffect(() => {
    let depth = 0;
    const isFileDrag = (e: DragEvent) =>
      Array.from(e.dataTransfer?.types ?? []).includes('Files');

    const onEnter = (e: DragEvent) => {
      if (!isFileDrag(e)) return;
      e.preventDefault();
      depth += 1;
      if (depth === 1) setDragging(true);
    };
    const onOver = (e: DragEvent) => {
      if (!isFileDrag(e)) return;
      // Required for `drop` to fire at all.
      e.preventDefault();
      if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
    };
    const onLeave = (e: DragEvent) => {
      if (!isFileDrag(e)) return;
      depth = Math.max(0, depth - 1);
      if (depth === 0) setDragging(false);
    };
    const onDrop = async (e: DragEvent) => {
      if (!isFileDrag(e)) return;
      e.preventDefault();
      depth = 0;
      setDragging(false);
      const file = pickSupportedFile(e.dataTransfer?.files);
      if (!file) {
        window.alert(
          `Drop a spreadsheet file (${SUPPORTED_EXTENSIONS.map((x) => `.${x}`).join(', ')}).`,
        );
        return;
      }
      loading.set({ fileName: file.name, sizeBytes: file.size, phase: 'reading' });
      try {
        await loadSpreadsheetFile(file, api, workbook.replaceWorkbook, (phase) =>
          loading.set({ phase }),
        );
        requestAnimationFrame(() => loading.set(null));
      } catch (err) {
        const msg = err instanceof Error ? `${err.name}: ${err.message}` : String(err);
        console.error('[drop] failed', err);
        loading.set({
          fileName: file.name,
          sizeBytes: file.size,
          phase: 'reading',
          error: msg,
        });
      }
    };

    window.addEventListener('dragenter', onEnter);
    window.addEventListener('dragover', onOver);
    window.addEventListener('dragleave', onLeave);
    window.addEventListener('drop', onDrop);
    return () => {
      window.removeEventListener('dragenter', onEnter);
      window.removeEventListener('dragover', onOver);
      window.removeEventListener('dragleave', onLeave);
      window.removeEventListener('drop', onDrop);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api, workbook.replaceWorkbook]);

  return dragging;
}

function pickSupportedFile(list: FileList | undefined | null): File | null {
  if (!list) return null;
  for (const file of Array.from(list)) {
    const ext = file.name.split('.').pop()?.toLowerCase() ?? '';
    if (SUPPORTED_EXTENSIONS.includes(ext)) return file;
  }
  return null;
}
