import { Icon } from './Icon';

type Props = { filename: string };

export function TitleBar({ filename }: Props) {
  return (
    <header className="titlebar" data-testid="titlebar" role="banner">
      <span className="titlebar__brand">
        <Icon name="grid_view" size="sm" className="titlebar__brand-icon" />
        <span>Sheet</span>
      </span>
      <span className="titlebar__divider" aria-hidden="true" />
      <span className="titlebar__filename" data-testid="titlebar-filename">
        {filename}
      </span>
      <span className="titlebar__spacer" />
    </header>
  );
}
