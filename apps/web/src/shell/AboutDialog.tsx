import { Dialog } from './Dialog';

type Props = { onClose: () => void };

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
          A web spreadsheet, built on{' '}
          <a href="https://github.com/dream-num/univer" target="_blank" rel="noreferrer">
            Univer OSS
          </a>{' '}
          (Apache-2.0).
        </p>
        <dl className="about__facts">
          <dt>Source</dt>
          <dd>
            <a
              href="https://github.com/schnsrw/sheets"
              target="_blank"
              rel="noreferrer"
            >
              github.com/schnsrw/sheets
            </a>
          </dd>
          <dt>Engine</dt>
          <dd>Univer 0.22.1 — 478 formula functions</dd>
          <dt>License</dt>
          <dd>Apache-2.0 (vendored Univer); project licence pending</dd>
        </dl>
      </div>
    </Dialog>
  );
}
