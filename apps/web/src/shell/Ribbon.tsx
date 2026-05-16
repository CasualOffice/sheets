import { useRef, useState, type ReactNode } from 'react';
import { useUniverAPI } from '../use-univer';
import { FileMenu } from './FileMenu';
import { useActiveCellState, type HAlign } from '../hooks/useActiveCellState';
import {
  NUMBER_FORMATS,
  setAlignment,
  setNumberFormat,
  toggleBold,
  toggleItalic,
  toggleMerge,
  toggleUnderline,
} from './home-tab-actions';
import { Icon } from './Icon';

const TABS = ['Home', 'Insert', 'Formulas', 'Data', 'Review', 'View'] as const;
type Tab = (typeof TABS)[number];

export function Ribbon() {
  const [active, setActive] = useState<Tab>('Home');
  const [fileMenuOpen, setFileMenuOpen] = useState(false);
  const fileBtnRef = useRef<HTMLButtonElement>(null);

  return (
    <nav className="ribbon" data-testid="ribbon" aria-label="Ribbon">
      <div className="ribbon__tabs" role="tablist">
        <button
          ref={fileBtnRef}
          type="button"
          className="ribbon__tab ribbon__tab--file"
          data-testid="ribbon-tab-file"
          aria-haspopup="menu"
          aria-expanded={fileMenuOpen}
          onClick={() => setFileMenuOpen((v) => !v)}
        >
          File
        </button>
        {fileMenuOpen && (
          <FileMenu anchorRef={fileBtnRef} onClose={() => setFileMenuOpen(false)} />
        )}
        {TABS.map((tab) => (
          <button
            type="button"
            key={tab}
            role="tab"
            aria-selected={active === tab}
            className="ribbon__tab"
            data-testid={`ribbon-tab-${tab.toLowerCase()}`}
            onClick={() => setActive(tab)}
          >
            {tab}
          </button>
        ))}
      </div>

      <div
        className="ribbon__body"
        role="tabpanel"
        data-testid={`ribbon-body-${active.toLowerCase()}`}
      >
        {active === 'Home' ? (
          <HomeTab />
        ) : (
          <span className="ribbon__empty" data-testid="ribbon-empty">
            {active} tab — coming soon
          </span>
        )}
      </div>
    </nav>
  );
}

function HomeTab() {
  const api = useUniverAPI();
  const state = useActiveCellState();
  const ready = Boolean(api) && state.ready;

  return (
    <>
      <RibbonGroup label="Clipboard">
        <ToolbarButton id="paste" label="Paste" icon="content_paste" disabled />
        <ToolbarButton id="cut" label="Cut" icon="content_cut" disabled />
        <ToolbarButton id="copy" label="Copy" icon="content_copy" disabled />
      </RibbonGroup>

      <RibbonGroup label="Font">
        <ToolbarButton
          id="bold"
          label="Bold (Ctrl+B)"
          icon="format_bold"
          pressed={state.isBold}
          disabled={!ready}
          onClick={() => api && toggleBold(api, state.isBold)}
        />
        <ToolbarButton
          id="italic"
          label="Italic (Ctrl+I)"
          icon="format_italic"
          pressed={state.isItalic}
          disabled={!ready}
          onClick={() => api && toggleItalic(api, state.isItalic)}
        />
        <ToolbarButton
          id="underline"
          label="Underline (Ctrl+U)"
          icon="format_underlined"
          pressed={state.isUnderline}
          disabled={!ready}
          onClick={() => api && toggleUnderline(api, state.isUnderline)}
        />
      </RibbonGroup>

      <RibbonGroup label="Alignment">
        <ToolbarButton
          id="align-left"
          label="Align left"
          icon="format_align_left"
          pressed={state.align === 'left'}
          disabled={!ready}
          onClick={() => api && setAlignment(api, 'left')}
        />
        <ToolbarButton
          id="align-center"
          label="Center"
          icon="format_align_center"
          pressed={state.align === 'center'}
          disabled={!ready}
          onClick={() => api && setAlignment(api, 'center')}
        />
        <ToolbarButton
          id="align-right"
          label="Align right"
          icon="format_align_right"
          pressed={state.align === ('right' as HAlign)}
          disabled={!ready}
          onClick={() => api && setAlignment(api, 'right')}
        />
        <ToolbarButton
          id="merge-cells"
          label={state.isMerged ? 'Unmerge cells' : 'Merge & Center'}
          icon={state.isMerged ? 'call_split' : 'cell_merge'}
          pressed={state.isMerged}
          disabled={!ready || (!state.isMerged && !state.isMultiCell)}
          onClick={() => api && toggleMerge(api, state.isMerged)}
        />
      </RibbonGroup>

      <RibbonGroup label="Number">
        <ToolbarButton
          id="numfmt-currency"
          label="Currency"
          icon="attach_money"
          pressed={state.numberFormat === NUMBER_FORMATS.currency}
          disabled={!ready}
          onClick={() => api && setNumberFormat(api, NUMBER_FORMATS.currency)}
        />
        <ToolbarButton
          id="numfmt-percent"
          label="Percent"
          icon="percent"
          pressed={state.numberFormat === NUMBER_FORMATS.percent}
          disabled={!ready}
          onClick={() => api && setNumberFormat(api, NUMBER_FORMATS.percent)}
        />
      </RibbonGroup>
    </>
  );
}

function RibbonGroup({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="ribbon__group" data-testid={`ribbon-group-${label.toLowerCase()}`}>
      <div className="ribbon__group-body">{children}</div>
      <div className="ribbon__group-label">{label}</div>
    </div>
  );
}

type ToolbarButtonProps = {
  id: string;
  label: string;
  icon: string;
  pressed?: boolean;
  disabled?: boolean;
  onClick?: () => void;
};

function ToolbarButton({ id, label, icon, pressed, disabled, onClick }: ToolbarButtonProps) {
  return (
    <button
      type="button"
      className="btn btn--icon"
      data-testid={`ribbon-btn-${id}`}
      aria-label={label}
      aria-pressed={pressed}
      title={label}
      disabled={disabled}
      onClick={onClick}
    >
      <Icon name={icon} size="sm" />
    </button>
  );
}
