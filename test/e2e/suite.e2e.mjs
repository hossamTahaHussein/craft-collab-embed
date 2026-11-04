// End to end: Craft Sheets and Craft Slides co-edited in two browsers through the demo server and
// the relay (protocol v3), driven by the SDK's JavaScript API, plus the relay's HTTP API.
//
//   SHEETS_DIR=../upstream/sheets/dist/web SLIDES_DIR=../upstream/slides/dist/web node --test test/e2e/suite.e2e.mjs
//
// Needs web builds of both editors and Playwright's Chromium.

import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { existsSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { chromium } from 'playwright';

const SHEETS_DIR = resolve(process.env.SHEETS_DIR || '../upstream/sheets/dist/web');
const SLIDES_DIR = resolve(process.env.SLIDES_DIR || '../upstream/slides/dist/web');
const PORT = 21000 + Math.floor(Math.random() * 1000);
const BASE = `http://127.0.0.1:${PORT}`;
const KEY = 'e2e-api-key-0123456789abcdef';
const work = mkdtempSync(join(tmpdir(), 'craft-e2e-'));
let server;
let browser;

function startServer() {
  return new Promise((resolveStart, reject) => {
    server = spawn(process.execPath, ['examples/server.mjs'], {
      env: {
        ...process.env,
        PORT: String(PORT),
        SHEETS_DIR,
        SLIDES_DIR,
        FILES_DIR: join(work, 'files'),
        RELAY_DATA_DIR: join(work, 'relay'),
        RELAY_SECRET: 'e2e-secret-0123456789abcdef',
        RELAY_API_KEY: KEY,
      },
      stdio: ['ignore', 'pipe', 'inherit'],
    });
    server.stdout.on('data', (d) => d.toString().includes('demo on') && resolveStart());
    server.on('exit', (code) => code && reject(new Error(`server exited ${code}`)));
  });
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function waitFor(fn, what, ms = 60_000) {
  const until = Date.now() + ms;
  let last;
  for (;;) {
    try {
      last = await fn();
      if (last) return last;
    } catch (e) {
      last = e;
    }
    if (Date.now() > until) throw new Error(`timed out waiting for ${what} (last: ${last?.message ?? JSON.stringify(last)})`);
    await sleep(250);
  }
}

/** `name` opens `doc` of `app` in a browser of their own; resolves once joined. */
async function enter(app, doc, name, role = 'edit') {
  const context = await browser.newContext({ viewport: { width: 1200, height: 760 } });
  const page = await context.newPage();
  page.on('pageerror', (e) => console.error(`[${name}] page error:`, e.message));
  await page.goto(`${BASE}/collab.html?app=${app}&doc=${doc}&name=${name}&role=${role}`);
  await waitFor(() => page.textContent('#status').then((t) => /synced|Saved/.test(t)), `${name} to join ${app}`);
  return page;
}

const api = (path) => fetch(`${BASE}${path}`, { headers: { Authorization: `Bearer ${KEY}` } });

before(async () => {
  assert.ok(existsSync(join(SHEETS_DIR, 'index.html')), `no Craft Sheets build at ${SHEETS_DIR}`);
  assert.ok(existsSync(join(SLIDES_DIR, 'index.html')), `no Craft Slides build at ${SLIDES_DIR}`);
  await startServer();
  browser = await chromium.launch({ args: ['--use-angle=swiftshader', '--enable-unsafe-swiftshader', '--ignore-gpu-blocklist'] });
});

after(async () => {
  await browser?.close();
  if (server && server.exitCode === null) {
    await new Promise((r) => {
      server.once('exit', r);
      server.kill();
    });
  }
  rmSync(work, { recursive: true, force: true });
});

test('Craft Sheets: an edit in one browser shows in the other, the API reads it, and undo takes back only your own', async () => {
  const ann = await enter('sheets', 'book', 'Ann');
  const bob = await enter('sheets', 'book', 'Bob');
  await ann.evaluate(() => window.editor.setRange('A1', [['Hello', 3]]));
  await waitFor(() => bob.evaluate(() => window.editor.getRange('A1:B1')).then((v) => JSON.stringify(v).includes('Hello')), 'Bob to see Ann’s cells');
  await bob.evaluate(() => window.editor.setRange('C1', [['Bob was here']]));
  await waitFor(() => ann.evaluate(() => window.editor.getRange('C1')).then((v) => JSON.stringify(v).includes('Bob was here')), 'Ann to see Bob’s cell');
  // The relay's HTTP API asks an editor for the workbook's text.
  const text = await waitFor(async () => {
    const r = await api('/api/sheets/book/text');
    if (!r.ok) throw new Error(`HTTP ${r.status}: ${await r.text()}`);
    return r.text();
  }, 'the text endpoint');
  assert.match(text, /Hello/);
  assert.match(text, /Bob was here/);
  const info = await (await api('/api/sheets/book')).json();
  assert.equal(info.people.length, 2);
  // Ann's undo takes back her edit for everyone, and leaves Bob's.
  await ann.evaluate(() => window.editor.run('edit.undo'));
  await waitFor(() => bob.evaluate(() => window.editor.getRange('A1:C1')).then((v) => !JSON.stringify(v).includes('Hello') && JSON.stringify(v).includes('Bob was here')), 'the undo to reach Bob');
  // A viewer reads but can't edit.
  const viewer = await enter('sheets', 'book', 'Vic', 'view');
  await assert.rejects(viewer.evaluate(() => window.editor.setRange('D1', [['nope']])), /view/i);
  assert.ok(!JSON.stringify(await ann.evaluate(() => window.editor.getRange('D1'))).includes('nope'));
  for (const p of [ann, bob, viewer]) await p.context().close();
});

test('Craft Slides: a new slide in one browser shows in the other, and the API renders it', async () => {
  const ann = await enter('slides', 'deck', 'Ann');
  const bob = await enter('slides', 'deck', 'Bob');
  const before = await bob.evaluate(() => window.editor.slides().then((s) => s.slides.length));
  await ann.evaluate(() => window.editor.run('slide.new'));
  await waitFor(() => bob.evaluate(() => window.editor.slides().then((s) => s.slides.length)).then((n) => n === before + 1), 'Bob to see the new slide');
  const png = await waitFor(async () => {
    const r = await api('/api/slides/deck/content?format=png');
    if (!r.ok) throw new Error(`HTTP ${r.status}: ${await r.text()}`);
    return r;
  }, 'a slide picture');
  assert.equal(png.headers.get('content-type'), 'image/png');
  const outline = await (await api('/api/slides/deck/text')).text();
  assert.equal(typeof outline, 'string');
  await assert.rejects(ann.evaluate(() => window.editor.getRange('A1')), /getRange is for Craft Sheets/);
  for (const p of [ann, bob]) await p.context().close();
});
