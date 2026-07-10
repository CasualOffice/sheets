/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License").
 */

/**
 * Panel store for the SDK chrome's side-panel rail (Tables/Charts/Comments/…).
 *
 * A single `openPanelId` is the whole state — that IS the mutex: opening one
 * panel closes any other, exactly like the standalone app's PanelMutex, but
 * without the app-only `ui-context`. The rail toggles panels through here and
 * the panel host renders whichever id is open. Kept deliberately tiny so the
 * chrome stays a thin shell around Univer.
 */
import { createContext, useContext, useMemo, useState, type ReactNode } from 'react';

export interface PanelStore {
  /** The id of the currently-open panel, or null when the rail is all-closed. */
  openPanelId: string | null;
  /** Open `id`, replacing whatever was open (mutex). */
  open: (id: string) => void;
  /** Close the open panel. */
  close: () => void;
  /** Toggle `id`: open it if closed, close it if it's the open one. */
  toggle: (id: string) => void;
}

const PanelContext = createContext<PanelStore | null>(null);

export function PanelProvider({ children }: { children: ReactNode }) {
  const [openPanelId, setOpenPanelId] = useState<string | null>(null);
  const store = useMemo<PanelStore>(
    () => ({
      openPanelId,
      open: (id) => setOpenPanelId(id),
      close: () => setOpenPanelId(null),
      toggle: (id) => setOpenPanelId((cur) => (cur === id ? null : id)),
    }),
    [openPanelId],
  );
  return <PanelContext.Provider value={store}>{children}</PanelContext.Provider>;
}

/** Read the panel store. Returns a no-op store when used outside a provider. */
export function usePanels(): PanelStore {
  return (
    useContext(PanelContext) ?? {
      openPanelId: null,
      open: () => {},
      close: () => {},
      toggle: () => {},
    }
  );
}
