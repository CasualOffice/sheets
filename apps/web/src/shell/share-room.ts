/**
 * Create a new co-edit room on the server and navigate the current tab
 * into it. Best-effort copies the resulting URL to the clipboard so the
 * user can share immediately.
 *
 * Server-side allocation (POST /api/rooms) keeps the room registry in
 * sync from the start — alternatively we could mint a random id
 * client-side and have the server lazily create on first WS connect,
 * but the registry would miss the room until then.
 */
export async function startCoEditRoom(): Promise<void> {
  let roomId: string;
  try {
    const res = await fetch('/api/rooms', { method: 'POST' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const body = (await res.json()) as { roomId: string };
    roomId = body.roomId;
  } catch (err) {
    console.error('[share-room] failed to allocate room', err);
    window.alert(
      'Could not start a co-edit room. The server may be unreachable — co-editing only works on the self-hosted Docker build.',
    );
    return;
  }

  const url = `${window.location.origin}/r/${roomId}`;
  // Clipboard write can reject under permissions; downgrade to a soft
  // notice rather than blocking the join.
  try {
    await navigator.clipboard.writeText(url);
  } catch {
    /* user can still copy from the address bar after navigation */
  }
  window.location.href = url;
}
