import assert from 'node:assert/strict';
import test from 'node:test';

import {
  announce,
  getAnnouncement,
  resetAnnouncements,
} from '../src/a11y/liveAnnouncer';
import {
  KEYBOARD_CONNECT_EVENT,
  beginKeyboardConnection,
  cancelKeyboardConnection,
  completeKeyboardConnection,
  getPendingConnection,
  subscribeToPendingConnection,
  type KeyboardConnectDetail,
} from '../src/hooks/useKeyboardConnection';

test.beforeEach(() => {
  cancelKeyboardConnection();
  resetAnnouncements();
});

test('a keyboard connection needs two distinct nodes', () => {
  assert.equal(getPendingConnection(), null);
  assert.equal(completeKeyboardConnection('b'), null, 'completing without a source is a no-op');

  beginKeyboardConnection('a', 'App Service');
  assert.deepEqual(getPendingConnection(), { nodeId: 'a', label: 'App Service' });

  // Self-connections are silently dropped by React Flow, which would look like a
  // dead key, so the store rejects them up front and clears the pending source.
  assert.equal(completeKeyboardConnection('a'), null);
  assert.equal(getPendingConnection(), null);
});

test('completing a connection dispatches the source and target once', () => {
  const seen: KeyboardConnectDetail[] = [];
  const listener = (event: Event) => {
    seen.push((event as CustomEvent<KeyboardConnectDetail>).detail);
  };
  // Node has no DOM; the module only touches `window` at dispatch time, so a
  // bare EventTarget is enough to observe what the browser would receive.
  const bus = new EventTarget();
  const globals = globalThis as unknown as { window?: EventTarget };
  const previousWindow = globals.window;
  globals.window = bus;
  bus.addEventListener(KEYBOARD_CONNECT_EVENT, listener);

  try {
    beginKeyboardConnection('a', 'App Service');
    assert.equal(completeKeyboardConnection('b'), 'a');
    assert.deepEqual(seen, [{ source: 'a', target: 'b' }]);

    // The pending source must not survive, otherwise the next C press on any
    // node would silently create a second edge from the stale source.
    assert.equal(getPendingConnection(), null);
    assert.equal(completeKeyboardConnection('c'), null);
    assert.equal(seen.length, 1);
  } finally {
    bus.removeEventListener(KEYBOARD_CONNECT_EVENT, listener);
    if (previousWindow === undefined) delete globals.window;
    else globals.window = previousWindow;
  }
});

test('cancelling clears the pending source and notifies subscribers', () => {
  let notifications = 0;
  const unsubscribe = subscribeToPendingConnection(() => { notifications += 1; });
  try {
    beginKeyboardConnection('a', 'App Service');
    assert.equal(notifications, 1, 'arming must notify the store subscribers');

    cancelKeyboardConnection();
    assert.equal(getPendingConnection(), null);
    assert.equal(notifications, 2);

    // Cancelling when nothing is pending must not churn subscribers, otherwise
    // every stray Escape re-renders every node on the canvas.
    cancelKeyboardConnection();
    assert.equal(notifications, 2);
  } finally {
    unsubscribe();
  }
});

test('announcements carry a fresh id so repeats are re-read', () => {
  announce('Export complete.');
  const first = getAnnouncement('polite');
  assert.equal(first.text, 'Export complete.');

  announce('Export complete.');
  const second = getAnnouncement('polite');
  assert.equal(second.text, 'Export complete.');
  assert.notEqual(second.id, first.id, 'identical text must still change the key');
});

test('announcements route by politeness and ignore empty text', () => {
  announce('   ');
  assert.equal(getAnnouncement('polite').text, '');

  announce('Generation failed.', 'assertive');
  assert.equal(getAnnouncement('assertive').text, 'Generation failed.');
  assert.equal(getAnnouncement('polite').text, '', 'assertive must not leak into polite');
});
