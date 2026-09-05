import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { build } from 'esbuild';
import { chromium } from 'playwright';

const source = `
import React, { useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { useModalFocus } from './src/hooks/useModalFocus';
import { useEscapeKey } from './src/hooks/useEscapeKey';
import ModalScaffold from './src/components/ModalScaffold';

function Nested({ close }) {
  const ref = useModalFocus(true, close);
  return <section ref={ref} role="dialog" aria-modal="true" aria-label="Nested">
    <input id="nested-input" autoFocus />
    <button id="nested-close" onClick={close}>Close nested</button>
  </section>;
}
function Parent({ close }) {
  const [nested, setNested] = useState(false);
  const [busy, setBusy] = useState(false);
  const initial = useRef(null);
  const ref = useModalFocus(true, close, { initialFocusRef: initial, closeOnEscape: !busy });
  return <section ref={ref} role="dialog" aria-modal="true" aria-label="Parent">
    <button id="first" onClick={close}>Close parent</button>
    <input id="initial" ref={initial} />
    <button disabled>Disabled</button><button hidden>Hidden</button>
    <button id="nested-open" onClick={() => setNested(true)}>Open nested</button>
    <button id="busy" onClick={() => setBusy(value => !value)}>Toggle busy</button>
    <button id="last">Last</button>
    {nested && <Nested close={() => setNested(false)} />}
  </section>;
}
function Empty({ close }) {
  const ref = useModalFocus(true, close);
  return <section id="empty-dialog" ref={ref} role="dialog" aria-modal="true" aria-label="Empty">No controls</section>;
}
function Simultaneous({ close }) {
  const ref = useModalFocus(true, close);
  const [child, setChild] = useState(true);
  return <section ref={ref} role="dialog" aria-modal="true" aria-label="Simultaneous">
    <button id="simultaneous-close" onClick={close}>Close simultaneous</button>
    {child && <Nested close={() => setChild(false)} />}
  </section>;
}
function Legacy({ close }) {
  const [busy, setBusy] = useState(false);
  const [nested, setNested] = useState(false);
  const ref = useModalFocus(true, document.getElementById('outside'));
  useEscapeKey(!busy, close);
  return <section id="legacy-dialog" ref={ref} role="dialog" aria-modal="true" aria-label="Legacy">
    <button id="legacy-first">First</button>
    <button id="legacy-nested" onClick={() => setNested(true)}>Open nested</button>
    <button id="legacy-busy" onClick={() => setBusy(value => !value)}>Toggle busy</button>
    <button id="legacy-last">Last</button>
    {nested && <Nested close={() => setNested(false)} />}
  </section>;
}
function Scaffold({ close }) {
  const [nested, setNested] = useState(false);
  return <ModalScaffold isOpen onClose={close} ariaLabel="Scaffold">
    <button id="scaffold-nested" onClick={() => setNested(true)}>Open nested</button>
    {nested && <Nested close={() => setNested(false)} />}
  </ModalScaffold>;
}
function App() {
  const [dialog, setDialog] = useState('');
  const [panel, setPanel] = useState(true);
  useEscapeKey(panel, () => setPanel(false));
  return <>
    <button id="open" onClick={() => setDialog('parent')}>Open dialog</button>
    <button id="open-empty" onClick={() => setDialog('empty')}>Open empty</button>
    <button id="open-simultaneous" onClick={() => setDialog('simultaneous')}>Open simultaneous</button>
    <button id="open-legacy" onClick={() => setDialog('legacy')}>Open legacy dialog</button>
    <button id="open-scaffold" onClick={() => setDialog('scaffold')}>Open scaffold</button>
    {!dialog && <button id="open-fallback" onClick={() => setDialog('fallback')}>Open with removed opener</button>}
    <button id="outside" data-modal-focus-fallback>Outside</button>
    {panel && <aside id="panel">Ordinary panel <button id="panel-button">Panel action</button></aside>}
    {dialog === 'parent' && <Parent close={() => setDialog('')} />}
    {dialog === 'empty' && <Empty close={() => setDialog('')} />}
    {dialog === 'simultaneous' && <Simultaneous close={() => setDialog('')} />}
    {dialog === 'legacy' && <Legacy close={() => setDialog('')} />}
    {dialog === 'scaffold' && <Scaffold close={() => setDialog('')} />}
    {dialog === 'fallback' && <Empty close={() => setDialog('')} />}
  </>;
}
createRoot(document.getElementById('root')).render(<React.StrictMode><App /></React.StrictMode>);
`;

const bundle = await build({
  stdin: { contents: source, loader: 'tsx', resolveDir: process.cwd() },
  bundle: true, write: false, format: 'iife', platform: 'browser',
});
const server = createServer((request, response) => {
  if (request.url === '/app.js') {
    response.setHeader('Content-Type', 'application/javascript');
    response.end(bundle.outputFiles[0].text);
  } else {
    response.setHeader('Content-Type', 'text/html');
    response.end('<!doctype html><html><body><div id="root"></div><script src="/app.js"></script></body></html>');
  }
});
await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
let browser;
try {
  browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.goto(`http://127.0.0.1:${server.address().port}`);
  const focused = async id => {
    await page.waitForFunction(expected => document.activeElement?.id === expected, id);
  };
  await page.locator('#panel-button').focus();
  await page.keyboard.press('Shift+Tab');
  await focused('outside');
  await page.locator('#open').click();
  await focused('initial');
  await page.locator('#first').focus();
  await page.keyboard.press('Shift+Tab');
  await focused('last');
  await page.keyboard.press('Tab');
  await focused('first');
  await page.evaluate(() => document.getElementById('outside').focus());
  await focused('initial');
  await page.locator('#nested-open').click();
  await focused('nested-input');
  await page.locator('#nested-close').focus();
  await page.keyboard.press('Tab');
  await focused('nested-input');
  await page.keyboard.press('Escape');
  assert.equal(await page.getByRole('dialog').count(), 1, 'only the nested dialog closes');
  await focused('nested-open');
  assert.equal(await page.locator('#panel').count(), 1, 'dialog Escape does not close an ordinary panel');
  await page.locator('#busy').click();
  await page.keyboard.press('Escape');
  assert.equal(await page.getByRole('dialog').count(), 1, 'busy dialog keeps its focus scope');
  await page.locator('#last').focus();
  await page.keyboard.press('Tab');
  await focused('first');
  await page.locator('#busy').click();
  await page.keyboard.press('Escape');
  await focused('open');
  await page.locator('#open-empty').click();
  await focused('empty-dialog');
  await page.keyboard.press('Tab');
  await focused('empty-dialog');
  await page.keyboard.press('Shift+Tab');
  await focused('empty-dialog');
  await page.keyboard.press('Escape');
  await focused('open-empty');
  await page.locator('#open-simultaneous').click();
  await focused('nested-input');
  await page.keyboard.press('Escape');
  assert.equal(await page.getByRole('dialog').count(), 1, 'nested effect ordering preserves topmost dialog');
  await page.keyboard.press('Escape');
  await focused('open-simultaneous');
  await page.locator('#open-legacy').click();
  await focused('legacy-dialog');
  await page.locator('#legacy-first').focus();
  await page.keyboard.press('Shift+Tab');
  await focused('legacy-last');
  await page.locator('#legacy-nested').click();
  await focused('nested-input');
  await page.keyboard.press('Escape');
  await focused('legacy-nested');
  assert.equal(await page.getByRole('dialog').count(), 1, 'legacy companions never dismiss a nested dialog parent');
  await page.locator('#legacy-busy').click();
  await page.keyboard.press('Escape');
  assert.equal(await page.getByRole('dialog').count(), 1, 'legacy busy dialogs cannot dismiss an unrelated panel');
  await page.locator('#legacy-last').focus();
  await page.keyboard.press('Tab');
  await focused('legacy-first');
  await page.locator('#legacy-busy').click();
  await page.keyboard.press('Escape');
  await focused('outside');
  await page.locator('#open-scaffold').click();
  assert.equal(await page.locator('#root').getAttribute('aria-hidden'), 'true');
  await page.locator('#scaffold-nested').click();
  await focused('nested-input');
  await page.keyboard.press('Escape');
  await focused('scaffold-nested');
  await page.keyboard.press('Escape');
  await focused('open-scaffold');
  assert.equal(await page.locator('#root').getAttribute('aria-hidden'), null);
  await page.locator('#open-fallback').click();
  await focused('empty-dialog');
  await page.keyboard.press('Escape');
  await focused('outside');
  await page.keyboard.press('Escape');
  assert.equal(await page.locator('#panel').count(), 0, 'ordinary panels still handle Escape when no dialog is open');
  assert.deepEqual(errors, []);
  console.log('PASS modal focus: initial focus, Tab containment, nested/busy Escape, explicit/fallback return focus, legacy hooks, ModalScaffold, StrictMode, and ordinary panels.');
} finally {
  await browser?.close();
  await new Promise(resolve => server.close(resolve));
}
