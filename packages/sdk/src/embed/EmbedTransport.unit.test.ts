/**
 * EmbedTransport contract tests. Mirrors the document/ repo's
 * EmbedTransport.test.ts — wire shape is uniform across docs and
 * sheet, only the `app` discriminator differs. When you change
 * EmbedTransport.ts here, change it there too.
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import { EmbedTransport, isCasualEnvelope } from './';
import type { CasualEnvelope } from './';

function fakeHostWindow() {
  let handler: ((ev: unknown) => void) | null = null;
  return {
    addEventListener(_type: string, h: EventListenerOrEventListenerObject) {
      handler = h as (ev: unknown) => void;
    },
    removeEventListener() {
      handler = null;
    },
    fire(ev: unknown) {
      if (handler) handler(ev);
    },
  };
}

function fakeParent() {
  const sent: Array<{ msg: unknown; origin: string; transfer?: Transferable[] }> = [];
  return {
    sent,
    postMessage(msg: unknown, origin: string, transfer?: Transferable[]) {
      sent.push({ msg, origin, transfer });
    },
  };
}

test('isCasualEnvelope accepts a well-formed sheet envelope', () => {
  assert.equal(
    isCasualEnvelope({
      type: 'casual.hello',
      app: 'sheet',
      v: 1,
      data: {},
    } satisfies CasualEnvelope),
    true,
  );
});

test('isCasualEnvelope rejects foreign types', () => {
  assert.equal(isCasualEnvelope({ type: 'other.hello', app: 'sheet', v: 1, data: {} }), false);
});

test('isCasualEnvelope rejects bad app values', () => {
  assert.equal(isCasualEnvelope({ type: 'casual.hello', app: 'pdf', v: 1, data: {} }), false);
});

test('EmbedTransport drops messages from disallowed origins', async () => {
  const host = fakeHostWindow();
  const parent = fakeParent();
  let called = false;
  const transport = new EmbedTransport({
    app: 'sheet',
    hostOrigin: 'https://drive.example',
    version: '1.0.0',
    commit: 'abc',
    capabilities: [],
    parentWindow: parent,
    hostWindow: host,
  });
  transport.on({
    onHostHello: () => {
      called = true;
    },
  });
  host.fire({
    origin: 'https://evil.example',
    data: { type: 'casual.hello', app: 'sheet', v: 1, data: { capabilities: [] } },
  });
  await new Promise((r) => setTimeout(r, 0));
  assert.equal(called, false);
});

test('EmbedTransport routes valid host.hello + replies with editor.ready', async () => {
  const host = fakeHostWindow();
  const parent = fakeParent();
  let received: unknown = null;
  const transport = new EmbedTransport({
    app: 'sheet',
    hostOrigin: 'https://drive.example',
    version: '1.0.0',
    commit: 'abc',
    capabilities: [],
    parentWindow: parent,
    hostWindow: host,
  });
  transport.on({
    onHostHello: (data) => {
      received = data;
    },
  });
  host.fire({
    origin: 'https://drive.example',
    data: {
      type: 'casual.hello',
      app: 'sheet',
      v: 1,
      data: { capabilities: ['saveDocument'] },
    },
  });
  await new Promise((r) => setTimeout(r, 0));
  assert.deepEqual(received, { capabilities: ['saveDocument'] });
  const lastSent = parent.sent[parent.sent.length - 1].msg as CasualEnvelope;
  assert.equal(lastSent.type, 'casual.ready');
  assert.equal(lastSent.app, 'sheet');
});

test('sendHello emits an editor.hello envelope', () => {
  const host = fakeHostWindow();
  const parent = fakeParent();
  const transport = new EmbedTransport({
    app: 'sheet',
    hostOrigin: 'https://drive.example',
    version: '1.2.3',
    commit: 'deadbee',
    capabilities: ['save', 'load'],
    parentWindow: parent,
    hostWindow: host,
  });
  transport.sendHello();
  assert.equal(parent.sent.length, 1);
  const env = parent.sent[0].msg as CasualEnvelope;
  assert.equal(env.type, 'casual.hello');
  assert.equal((env.data as { version: string }).version, '1.2.3');
  assert.deepEqual((env.data as { capabilities: string[] }).capabilities, ['save', 'load']);
});

test('destroy detaches the listener', () => {
  const host = fakeHostWindow();
  const transport = new EmbedTransport({
    app: 'sheet',
    hostOrigin: 'https://drive.example',
    version: '1.0.0',
    commit: 'abc',
    capabilities: [],
    parentWindow: fakeParent(),
    hostWindow: host,
  });
  let called = false;
  transport.on({
    onHostHello: () => {
      called = true;
    },
  });
  transport.destroy();
  host.fire({
    origin: 'https://drive.example',
    data: { type: 'casual.hello', app: 'sheet', v: 1, data: { capabilities: [] } },
  });
  assert.equal(called, false);
});
