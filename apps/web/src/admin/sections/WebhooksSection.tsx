import { useState } from 'react';
import type { AdminConfig, WebhookEvent, WebhookSubscription } from '../types';
import { ALL_WEBHOOK_EVENTS } from '../types';
import { SectionShell } from '../SectionShell';

interface Props {
  config: AdminConfig;
  save: (patch: Partial<AdminConfig>) => Promise<AdminConfig>;
}

const blankSub = (): WebhookSubscription => ({
  name: '',
  url: '',
  events: [],
  secret: '',
  enabled: true,
});

export function WebhooksSection({ config, save }: Props) {
  const [subs, setSubs] = useState<WebhookSubscription[]>(
    config.webhooks.length > 0 ? config.webhooks : [blankSub()],
  );

  const updateSub = (i: number, patch: Partial<WebhookSubscription>) => {
    setSubs(subs.map((s, idx) => (idx === i ? { ...s, ...patch } : s)));
  };
  const removeSub = (i: number) => setSubs(subs.filter((_, idx) => idx !== i));
  const addSub = () => setSubs([...subs, blankSub()]);
  const toggleEvent = (i: number, ev: WebhookEvent) => {
    const cur = subs[i].events;
    const next = cur.includes(ev) ? cur.filter((e) => e !== ev) : [...cur, ev];
    updateSub(i, { events: next });
  };

  const submit = async () => {
    // Drop blank rows (no url) so the operator doesn't have to.
    const cleaned = subs.filter((s) => s.url.trim() !== '');
    await save({ webhooks: cleaned });
    setSubs(cleaned.length > 0 ? cleaned : [blankSub()]);
  };

  return (
    <SectionShell
      title="Webhooks"
      description="HTTP POSTs to operator-configured URLs when server-side events fire. HMAC-SHA256 signed when a secret is set."
      onSubmit={submit}
      aside={
        <>
          <h4>Signature verification</h4>
          <p>
            Receivers compute <code>hmac-sha256(secret, raw_body)</code> and
            compare against the <code>X-Casual-Signature: sha256=&lt;hex&gt;</code>
            header. Constant-time compare is mandatory — see the
            <a href="/docs/sheets/customization/" target="_blank" rel="noopener"> customization docs</a> for a Node + Python verifier example.
          </p>
          <h4>Empty event list</h4>
          <p>
            Leaving every checkbox unchecked subscribes to <em>all</em> events.
            That's the cheapest "send me everything" mode for audit-log
            collectors.
          </p>
          <h4>Retry</h4>
          <p>
            Single retry after 5 s on non-2xx / network error. v0.2 ships a
            proper queue with exponential back-off + dead-letter store.
          </p>
        </>
      }
    >
      <div className="admin-webhooks">
        {subs.map((sub, i) => (
          <fieldset key={i} className="admin-webhook">
            <div className="admin-webhook__head">
              <label className="admin-field admin-field--check">
                <input
                  type="checkbox"
                  checked={sub.enabled}
                  onChange={(e) => updateSub(i, { enabled: e.target.checked })}
                />
                <span>Enabled</span>
              </label>
              {subs.length > 1 && (
                <button
                  type="button"
                  className="admin-btn admin-btn--ghost admin-btn--sm"
                  onClick={() => removeSub(i)}
                >
                  Remove
                </button>
              )}
            </div>
            <label className="admin-field">
              <span>Name</span>
              <input value={sub.name} onChange={(e) => updateSub(i, { name: e.target.value })} placeholder="audit-log" />
            </label>
            <label className="admin-field">
              <span>URL</span>
              <input value={sub.url} onChange={(e) => updateSub(i, { url: e.target.value })} placeholder="https://example.com/hooks/casual" />
            </label>
            <label className="admin-field">
              <span>Signing secret <small>(optional)</small></span>
              <input type="password" value={sub.secret} onChange={(e) => updateSub(i, { secret: e.target.value })} autoComplete="new-password" />
            </label>
            <div className="admin-field">
              <span>Events <small>(none checked = subscribed to all)</small></span>
              <div className="admin-event-grid">
                {ALL_WEBHOOK_EVENTS.map((ev) => (
                  <label key={ev} className="admin-event-grid__item">
                    <input
                      type="checkbox"
                      checked={sub.events.includes(ev)}
                      onChange={() => toggleEvent(i, ev)}
                    />
                    <code>{ev}</code>
                  </label>
                ))}
              </div>
            </div>
          </fieldset>
        ))}
      </div>
      <button type="button" className="admin-btn admin-btn--ghost" onClick={addSub}>
        + Add subscription
      </button>
    </SectionShell>
  );
}
