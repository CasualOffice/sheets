/**
 * Pure WS-URL composition for attachCollab. Kept in its own module — free of
 * any `@univerjs` / `@hocuspocus` imports — so it's unit-testable under
 * `node:test` without dragging in the Univer ESM graph.
 */

/** Build the room WS URL: `<server>?room=<id>[&p=<pw>]&role=<role>`. */
export function buildWsUrl(
  server: string,
  room: string,
  role: 'view' | 'write',
  password?: string,
): string {
  const sep = server.includes('?') ? '&' : '?';
  return (
    `${server}${sep}room=${encodeURIComponent(room)}` +
    `${password ? `&p=${encodeURIComponent(password)}` : ''}` +
    `&role=${encodeURIComponent(role)}`
  );
}
