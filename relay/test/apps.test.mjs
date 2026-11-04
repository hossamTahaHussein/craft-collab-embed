// Protocol v3: one relay for the suite's three editors. Each app's documents are their own rooms,
// each app's selections and role commands are its own (apps.mjs), and the HTTP API and webhooks
// say which app.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { APPS, fromRoomKey, roomKey } from '../apps.mjs';
import { createRelay } from '../server.mjs';
import { issueToken } from '../token.mjs';
import { connect, tick } from './helpers.mjs';

const SECRET = 'server-test-secret-0123456789';
const KEY = 'api-key-0123456789abcdef';
const HOOK_SECRET = 'hook-secret-0123456789abcdef';
let relay;
let base;
let dir;
let hooks;
const delivered = [];

before(async () => {
  dir = mkdtempSync(join(tmpdir(), 'wc-apps-'));
  hooks = createServer((req, res) => {
    let body = '';
    req.on('data', (c) => (body += c));
    req.on('end', () => {
      delivered.push(JSON.parse(body));
      res.end('ok');
    });
  });
  await new Promise((r) => hooks.listen(0, '127.0.0.1', r));
  relay = await createRelay({
    secret: SECRET,
    dataDir: dir,
    port: 0,
    apiKey: KEY,
    apiTimeoutMs: 300,
    webhook: { url: `http://127.0.0.1:${hooks.address().port}/`, secret: HOOK_SECRET },
  });
  base = `http://127.0.0.1:${relay.port}`;
});

after(async () => {
  await relay.close();
  await new Promise((r) => hooks.close(r));
  rmSync(dir, { recursive: true, force: true });
});

const api = (path, { method = 'GET', body } = {}) =>
  fetch(`${base}${path}`, {
    method,
    headers: { Authorization: `Bearer ${KEY}`, ...(body ? { 'Content-Type': 'application/json' } : {}) },
    body: body === undefined ? undefined : JSON.stringify(body),
  });

const SHEET_SEL = { sheet: 'Sheet1', selection: { active: { row: 1, col: 1 }, ranges: [] } };
const SLIDE_SEL = { slide: 3, selection: { slide: 0, slides: [], shapes: [7] } };
const DOCS_SEL = { story: 'body', anchor: 0, head: 0 };

/** Someone in `doc` of `app` with `role`; the first editor uploads revision 0. `helloApp`: what
 * the editor says it is (the token's app unless given; `null` says nothing, as before v3). */
async function enter(app, doc, { role = 'edit', sub = role, client = `c${sub}${app}`, helloApp = app } = {}) {
  const m = await connect(`ws://127.0.0.1:${relay.port}/relay`);
  const token = issueToken(SECRET, { doc, sub, name: sub, role, ...(app === 'docs' ? {} : { app }) });
  m.send({ t: 'hello', token, client, ...(helloApp === null ? {} : { app: helloApp }) });
  return m;
}

async function opened(app, doc, opts) {
  const m = await enter(app, doc, opts);
  const w = await m.next('Welcome');
  if (w.snapshot === null) m.send({ t: 'Snapshot', rev: 0, doc: '{}' });
  return { m, w };
}

const op = (seq, cmd, sel, fields = {}) => ({ t: 'Op', client: 'x', seq, base: 0, kind: 'cmd', cmd, params: {}, sel, ctx: {}, ...fields });

test('room keys keep each app apart, and Craft Docs keeps its document ids', () => {
  assert.equal(roomKey('docs', 'report'), 'report');
  assert.notEqual(roomKey('sheets', 'report'), roomKey('slides', 'report'));
  assert.notEqual(roomKey('sheets', 'report'), 'report');
  for (const app of ['docs', 'sheets', 'slides']) assert.deepEqual(fromRoomKey(roomKey(app, 'a b')), { app, doc: 'a b' });
});

test('an editor whose app is not the token one is turned away with the reason', async () => {
  const m = await enter('sheets', 'mixed', { helloApp: 'slides' });
  const closed = await m.closed;
  assert.equal(closed.code, 4400);
  assert.match(closed.reason, /Craft Sheets, not Craft Slides/);
  // An editor from before v3 says no app: it is Craft Docs, which a Sheets token isn't.
  const old = await enter('sheets', 'mixed', { helloApp: null });
  assert.match((await old.closed).reason, /Craft Sheets, not Craft Docs/);
});

test('a Sheets and a Slides document with a Docs document id are three documents', async () => {
  const docs = await opened('docs', 'same');
  const sheets = await opened('sheets', 'same');
  const slides = await opened('slides', 'same');
  assert.equal(docs.w.app, 'docs');
  assert.equal(sheets.w.app, 'sheets');
  assert.equal(slides.w.app, 'slides');
  sheets.m.send(op(1, 'cell.set', SHEET_SEL, { params: { cell: 'A1', input: '5' } }));
  const got = await sheets.m.next('Op');
  assert.equal(got.cmd, 'cell.set');
  await tick(50);
  assert.equal(docs.m.got.some((x) => x.t === 'Op'), false, 'the Docs document has none of it');
  assert.equal(slides.m.got.some((x) => x.t === 'Op'), false);
  assert.ok(relay.loadedRooms().includes('sheets:same') && relay.loadedRooms().includes('same'));
  for (const x of [docs, sheets, slides]) x.m.ws.close();
});

test("each app's selections are checked and forwarded as that app's", async () => {
  const a = await opened('sheets', 'sel');
  const b = await opened('sheets', 'sel', { sub: 'b' });
  a.m.send({ t: 'Presence', rev: 0, sel: { ...SHEET_SEL, extra: 'dropped' } });
  const p = await b.m.next('Presence');
  assert.deepEqual(p.sel, SHEET_SEL, 'forwarded without what replicas do not read');
  // A Docs-shaped selection isn't a Sheets one: not forwarded, and an op with it isn't taken.
  a.m.send({ t: 'Presence', rev: 0, sel: DOCS_SEL });
  a.m.send(op(1, 'cell.set', DOCS_SEL));
  a.m.send(op(2, 'cell.set', SHEET_SEL));
  const o = await b.m.next('Op');
  assert.equal(o.seq, 2);
  assert.equal(b.m.got.some((x) => x.t === 'Presence'), false);
  const s1 = await opened('slides', 'sel');
  const s2 = await opened('slides', 'sel', { sub: 'b' });
  s1.m.send({ t: 'Presence', rev: 0, sel: { ...SLIDE_SEL, text: [{ shape: 7 }, 0, 2] } });
  assert.deepEqual((await s2.m.next('Presence')).sel, { ...SLIDE_SEL, text: [{ shape: 7 }, 0, 2] });
  s1.m.send({ t: 'Presence', rev: 0, sel: { slide: -1, selection: {} } });
  s1.m.send({ t: 'Presence', rev: 0, sel: { slide: 1, selection: { big: 'x'.repeat(20_000) } } });
  await tick(50);
  assert.equal(s2.m.got.some((x) => x.t === 'Presence'), false, 'a bad or oversized selection is not forwarded');
  for (const x of [a, b, s1, s2]) x.m.ws.close();
});

test("a commenter runs that app's comment commands and nothing else", async () => {
  const cases = [
    ['sheets', SHEET_SEL, 'review.newComment', 'cell.set'],
    ['slides', SLIDE_SEL, 'comment.add', 'slide.new'],
  ];
  for (const [app, sel, comment, edit] of cases) {
    const ed = await opened(app, `roles-${app}`);
    const c = await opened(app, `roles-${app}`, { role: 'comment' });
    c.m.send(op(1, edit, sel));
    assert.match((await c.m.next('Refused')).reason, /only comment/);
    c.m.send(op(2, comment, sel));
    assert.equal((await ed.m.next('Op')).cmd, comment);
    // Craft Docs' comment command is no Sheets or Slides one.
    c.m.send(op(3, 'review.reply', sel));
    assert.match((await c.m.next('Refused')).reason, /only comment/);
    ed.m.ws.close();
    c.m.ws.close();
  }
});

test("the HTTP API serves each app's documents in its own formats", async () => {
  const ed = await opened('sheets', 'book');
  assert.equal((await api('/api/docs/book')).status, 404, 'no Docs document of that id');
  const info = await (await api('/api/sheets/book')).json();
  assert.equal(info.app, 'sheets');
  const pending = api('/api/sheets/book/content');
  const ask = await ed.m.next('Export');
  assert.equal(ask.format, 'xlsx', 'a workbook by default');
  ed.m.send({ t: 'Exported', id: ask.id, ok: true, data: Buffer.from('PK').toString('base64') });
  const res = await pending;
  assert.equal(res.headers.get('content-type'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  assert.equal(res.headers.get('content-disposition'), 'attachment; filename="book.xlsx"');
  const text = api('/api/sheets/book/text');
  const t = await ed.m.next('Export');
  assert.equal(t.format, 'txt');
  ed.m.send({ t: 'Exported', id: t.id, ok: true, data: Buffer.from('Sheet1\n1,2\n').toString('base64') });
  assert.equal(await (await text).text(), 'Sheet1\n1,2\n');
  const bad = await api('/api/sheets/book/content?format=docx');
  assert.equal(bad.status, 400);
  assert.match((await bad.json()).error, /xlsx, csv, tsv, json, html, pdf/);
  const deck = await opened('slides', 'deck');
  const png = api('/api/slides/deck/content?format=png');
  const pa = await deck.m.next('Export');
  assert.equal(pa.format, 'png');
  deck.m.send({ t: 'Exported', id: pa.id, ok: true, data: Buffer.from('png').toString('base64') });
  assert.equal((await png).headers.get('content-type'), 'image/png');
  ed.m.ws.close();
  deck.m.ws.close();
});

test('webhooks say which app the document is', async () => {
  const x = await opened('slides', 'hooked');
  for (let i = 0; i < 100 && !delivered.some((d) => d.doc === 'hooked' && d.event === 'joined'); i++) await tick(20);
  const joined = delivered.find((d) => d.doc === 'hooked' && d.event === 'joined');
  assert.equal(joined?.app, 'slides');
  x.m.ws.close();
});

test("each app's command lists are read from its own files", () => {
  for (const [app, suffix] of [
    ['docs', ''],
    ['sheets', '.sheets'],
    ['slides', '.slides'],
  ]) {
    const file = (kind) => JSON.parse(readFileSync(new URL(`../${kind}-commands${suffix}.json`, import.meta.url), 'utf8'));
    assert.deepEqual([...APPS[app].comment], file('comment'), app);
    assert.deepEqual([...APPS[app].suggest], file('suggest'), app);
    for (const cmd of Object.keys(APPS[app].commentActions)) assert.ok(APPS[app].comment.has(cmd), `${app}: ${cmd} is a comment command`);
  }
  // The engines' lists (gridcraft_engine::collab / deckcraft_engine::collab::COMMENT_COMMANDS).
  assert.deepEqual([...APPS.sheets.comment], ['review.newComment', 'review.replyComment', 'review.deleteComment', 'review.resolveComment', 'review.newNote']);
  assert.deepEqual([...APPS.slides.comment], ['comment.add', 'comment.reply', 'comment.delete', 'comment.resolve']);
});
