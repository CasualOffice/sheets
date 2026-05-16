type Props = { filename: string };

export function TitleBar({ filename }: Props) {
  return (
    <header className="titlebar" data-testid="titlebar" role="banner">
      <a
        className="titlebar__brand"
        href="/"
        aria-label="Casual Sheets — home"
      >
        <img
          src="/brand.svg"
          alt=""
          className="titlebar__brand-icon"
          width={28}
          height={28}
        />
        <span className="titlebar__brand-name">Casual Sheets</span>
      </a>
      <span className="titlebar__divider" aria-hidden="true" />
      <span className="titlebar__filename" data-testid="titlebar-filename">
        {filename}
      </span>
      <span className="titlebar__spacer" />
    </header>
  );
}
