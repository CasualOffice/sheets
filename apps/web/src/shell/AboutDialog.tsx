import { Dialog } from './Dialog';

type Props = { onClose: () => void };

// `__APP_VERSION__` is replaced at build time by vite.config.ts (`define`).
// Falls back to `dev` when running outside the bundler (e.g. tests).
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const APP_VERSION: string = (globalThis as any).__APP_VERSION__ ?? 'dev';
// `__COLLAB_BUILD__` is `true` when built with VITE_COLLAB_ENABLED=1
// (the Docker image). The Pages bundle ships with it `false`.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const COLLAB_BUILD: boolean = Boolean((globalThis as any).__COLLAB_BUILD__);

export function AboutDialog({ onClose }: Props) {
  return (
    <Dialog
      title="About Casual Sheets"
      onClose={onClose}
      data-testid="about-dialog"
      footer={
        <button
          type="button"
          className="btn-primary"
          data-testid="about-close"
          onClick={onClose}
        >
          Close
        </button>
      }
    >
      <div className="about">
        <img
          src={`${import.meta.env.BASE_URL}brand.svg`}
          alt=""
          width={56}
          height={56}
          className="about__icon"
        />
        <h3 className="about__title">Casual Sheets</h3>
        <p className="about__tagline">
          A web spreadsheet that feels like Excel, built on{' '}
          <a href="https://github.com/dream-num/univer" target="_blank" rel="noreferrer">
            Univer OSS
          </a>
          .
        </p>
        <dl className="about__facts">
          <dt>Version</dt>
          <dd data-testid="about-version">{APP_VERSION}</dd>
          <dt>Edition</dt>
          <dd data-testid="about-edition">
            {COLLAB_BUILD ? 'Self-hosted (co-editing enabled)' : 'Single-user (hosted demo)'}
          </dd>
          <dt>Source</dt>
          <dd>
            <a
              href="https://github.com/CasualOffice/sheets"
              target="_blank"
              rel="noreferrer"
            >
              github.com/CasualOffice/sheets
            </a>
          </dd>
          <dt>Self-host</dt>
          <dd>
            <code>docker run -p 3000:3000 schnsrw/casual-sheets</code> —{' '}
            <a href="https://casualoffice.org/#work" target="_blank" rel="noreferrer">guide</a>
          </dd>
          <dt>Engine</dt>
          <dd>Univer 0.22.1 — 478 formula functions, ExcelJS for xlsx I/O</dd>
          <dt>License</dt>
          <dd>Apache-2.0</dd>
        </dl>
      </div>
    </Dialog>
  );
}
