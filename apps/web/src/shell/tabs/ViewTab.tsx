import { useEffect, useState } from 'react';
import { useUniverAPI } from '../../use-univer';
import { useActiveCellState } from '../../hooks/useActiveCellState';
import {
  freezeAtSelection,
  freezeFirstColumn,
  freezeFirstRow,
  setZoom,
  toggleGridlines,
  unfreezePanes,
} from '../tab-actions';
import { RibbonGroup, ToolbarButton } from '../RibbonControls';

export function ViewTab() {
  const api = useUniverAPI();
  const { ready } = useActiveCellState();
  const enabled = Boolean(api) && ready;

  // Read the current showGridlines flag off the snapshot.
  const [gridlinesShown, setGridlinesShown] = useState(true);
  useEffect(() => {
    if (!api) return;
    const refresh = () => {
      const wb = api.getActiveWorkbook();
      const sheet = wb?.getActiveSheet();
      if (!sheet) return;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const flag = (sheet as any).getWorksheet?.()?.getConfig?.()?.showGridlines;
      setGridlinesShown(flag !== 0);
    };
    refresh();
    const d = api.addEvent(api.Event.CommandExecuted, (e) => {
      const id = (e as { id?: string }).id ?? '';
      if (id.includes('gridlines') || id === 'sheet.operation.set-selections') refresh();
    });
    return () => d.dispose();
  }, [api]);

  return (
    <>
      <RibbonGroup label="Freeze">
        <ToolbarButton
          id="freeze-first-row"
          label="Freeze top row"
          icon="border_horizontal"
          disabled={!enabled}
          onClick={() => api && freezeFirstRow(api)}
        />
        <ToolbarButton
          id="freeze-first-col"
          label="Freeze first column"
          icon="border_vertical"
          disabled={!enabled}
          onClick={() => api && freezeFirstColumn(api)}
        />
        <ToolbarButton
          id="freeze-panes"
          label="Freeze panes (at selection)"
          icon="grid_4x4"
          disabled={!enabled}
          onClick={() => api && freezeAtSelection(api)}
        />
        <ToolbarButton
          id="unfreeze-panes"
          label="Unfreeze"
          icon="grid_off"
          disabled={!enabled}
          onClick={() => api && unfreezePanes(api)}
        />
      </RibbonGroup>

      <RibbonGroup label="Show">
        <ToolbarButton
          id="toggle-gridlines"
          label="Gridlines"
          icon="grid_on"
          pressed={gridlinesShown}
          disabled={!enabled}
          onClick={() => api && toggleGridlines(api, gridlinesShown)}
        />
      </RibbonGroup>

      <RibbonGroup label="Zoom">
        <ToolbarButton
          id="zoom-out"
          label="Zoom out"
          icon="zoom_out"
          disabled={!enabled}
          onClick={() => api && setZoom(api, 0.75)}
        />
        <ToolbarButton
          id="zoom-100"
          label="Zoom 100%"
          icon="100_percent"
          disabled={!enabled}
          onClick={() => api && setZoom(api, 1)}
        />
        <ToolbarButton
          id="zoom-in"
          label="Zoom in"
          icon="zoom_in"
          disabled={!enabled}
          onClick={() => api && setZoom(api, 1.5)}
        />
      </RibbonGroup>
    </>
  );
}
