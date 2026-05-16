import { useEffect, useRef, useState } from 'react';
import { Icon } from './Icon';
import { PropertiesDialog } from './PropertiesDialog';
import { openXlsx, pickXlsxFile, saveAsXlsx } from './file-actions';
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { emptyWorkbook } from '../snapshot';

/**
 * Office 365-style File dropdown.
 */
export function FileMenu({
  anchorRef,
  onClose,
}: {
  anchorRef: React.RefObject<HTMLElement | null>;
  onClose: () => void;
}) {
  const api = useUniverAPI();
  const workbook = useWorkbook();

  const [showProperties, setShowProperties] = useState(false);
  const [busy, setBusy] = useState<null | 'opening' | 'saving'>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ top: number; left: number }>({ top: 0, left: 0 });

  useEffect(() => {
    const rect = anchorRef.current?.getBoundingClientRect();
    if (rect) setPos({ top: rect.bottom + 2, left: rect.left });
  }, [anchorRef]);

  useEffect(() => {
    const onDocClick = (e: MouseEvent) => {
      if (showProperties) return;
      if (!menuRef.current?.contains(e.target as Node) && !anchorRef.current?.contains(e.target as Node)) {
        onClose();
      }
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && !showProperties) onClose();
    };
    document.addEventListener('mousedown', onDocClick);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('mousedown', onDocClick);
      document.removeEventListener('keydown', onKey);
    };
  }, [onClose, anchorRef, showProperties]);

  const handleOpen = async () => {
    setBusy('opening');
    try {
      const file = await pickXlsxFile();
      if (!file) return;
      const data = await openXlsx(file);
      workbook.replaceWorkbook(data);
    } catch (err) {
      console.error('Open failed:', err);
      alert(`Could not open file: ${(err as Error).message}`);
    } finally {
      setBusy(null);
      onClose();
    }
  };

  const handleSaveAs = async () => {
    if (!api) return;
    setBusy('saving');
    try {
      await saveAsXlsx(api, workbook.snapshot.name || 'workbook');
    } catch (err) {
      console.error('Save failed:', err);
      alert(`Could not save file: ${(err as Error).message}`);
    } finally {
      setBusy(null);
      onClose();
    }
  };

  const handleNew = () => {
    workbook.replaceWorkbook(emptyWorkbook());
    onClose();
  };

  return (
    <>
      <div
        ref={menuRef}
        className="menu"
        data-testid="file-menu"
        role="menu"
        style={{ top: pos.top, left: pos.left }}
      >
        <MenuItem
          icon="add"
          label="New"
          shortcut="Ctrl+N"
          testid="file-menu-new"
          onClick={handleNew}
        />
        <MenuItem
          icon="folder_open"
          label={busy === 'opening' ? 'Opening…' : 'Open'}
          shortcut="Ctrl+O"
          testid="file-menu-open"
          disabled={busy !== null}
          onClick={handleOpen}
        />
        <MenuItem
          icon="save"
          label={busy === 'saving' ? 'Saving…' : 'Save As'}
          shortcut="Ctrl+Shift+S"
          testid="file-menu-save-as"
          disabled={busy !== null || !api}
          onClick={handleSaveAs}
        />
        <div className="menu__divider" />
        <MenuItem
          icon="info"
          label="Properties"
          testid="file-menu-properties"
          onClick={() => setShowProperties(true)}
        />
        <div className="menu__divider" />
        <MenuItem icon="close" label="Close" disabled />
      </div>

      {showProperties && (
        <PropertiesDialog
          onClose={() => {
            setShowProperties(false);
            onClose();
          }}
        />
      )}
    </>
  );
}

function MenuItem({
  icon,
  label,
  shortcut,
  onClick,
  disabled,
  testid,
}: {
  icon: string;
  label: string;
  shortcut?: string;
  onClick?: () => void;
  disabled?: boolean;
  testid?: string;
}) {
  return (
    <button
      type="button"
      className="menu__item"
      role="menuitem"
      data-testid={testid}
      disabled={disabled}
      onClick={onClick}
    >
      <Icon name={icon} size="sm" className="menu__item-icon" />
      <span>{label}</span>
      {shortcut && <span className="menu__item-shortcut">{shortcut}</span>}
    </button>
  );
}
