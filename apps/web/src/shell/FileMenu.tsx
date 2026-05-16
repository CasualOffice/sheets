import { useEffect, useRef, useState } from 'react';
import { Icon } from './Icon';
import { PropertiesDialog } from './PropertiesDialog';

/**
 * Office 365-style File dropdown (not the full backstage view).
 * The "File" tab in the ribbon is rendered by Ribbon.tsx; this component
 * owns the dropdown panel that opens beneath it.
 */
export function FileMenu({
  anchorRef,
  onClose,
}: {
  anchorRef: React.RefObject<HTMLElement | null>;
  onClose: () => void;
}) {
  const [showProperties, setShowProperties] = useState(false);
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

  return (
    <>
      <div
        ref={menuRef}
        className="menu"
        data-testid="file-menu"
        role="menu"
        style={{ top: pos.top, left: pos.left }}
      >
        <MenuItem icon="add" label="New" disabled shortcut="Ctrl+N" />
        <MenuItem icon="folder_open" label="Open" disabled shortcut="Ctrl+O" />
        <MenuItem icon="save" label="Save As" disabled shortcut="Ctrl+Shift+S" />
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
