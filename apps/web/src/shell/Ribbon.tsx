import { useRef, useState } from 'react';
import { FileMenu } from './FileMenu';
import { HomeTab } from './tabs/HomeTab';
import { InsertTab } from './tabs/InsertTab';
import { FormulasTab } from './tabs/FormulasTab';
import { DataTab } from './tabs/DataTab';
import { ViewTab } from './tabs/ViewTab';

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
        {active === 'Home' && <HomeTab />}
        {active === 'Insert' && <InsertTab />}
        {active === 'Formulas' && <FormulasTab />}
        {active === 'Data' && <DataTab />}
        {active === 'View' && <ViewTab />}
        {active === 'Review' && (
          <span className="ribbon__empty" data-testid="ribbon-empty">
            {active} tab — coming soon
          </span>
        )}
      </div>
    </nav>
  );
}
