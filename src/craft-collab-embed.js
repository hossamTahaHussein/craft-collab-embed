// craft-collab-embed: put the Craft Collab Suite's editors in any web page — Craft Docs (a word
// processor), Craft Sheets (a spreadsheet) and Craft Slides (presentations), each with real-time
// co-editing through the relay (relay/server.mjs).
//
// Each editor runs in the browser as a static WebAssembly site. Framed with `?host=parent`, it
// talks to its parent page over postMessage: the parent sends the document to open, and the
// editor hands back every file it saves instead of downloading it. This module wraps that
// protocol:
//
//   const editor = createEditor('#editor', {
//     app: 'sheets',                            // 'docs' (default), 'sheets' or 'slides'
//     src: '/craft-sheets/',                    // where that editor's static files are hosted
//     document: '/files/budget.xlsx',           // optional: URL, File, Blob, ArrayBuffer or Uint8Array
//     author: 'Jane Doe',                       // signs comments (and tracked changes)
//     onSave: async ({ name, blob }) => { ... } // store the file (Ctrl+S, File > Save, Export)
//   });
//
// or, declaratively, <craft-editor app="sheets" src="/craft-sheets/" document="/files/budget.xlsx">.
// `WordCraftEditor` and <wordcraft-editor> (1.x) are Craft Docs editors, as before.
//
// The editor must be served from the same origin as the page that embeds it: it only answers a
// same-origin parent, so on another origin saves would be dropped silently.

/** The suite's editors: each one's message prefix, default folder and default file name. */
export const APPS = {
  docs: { name: 'Craft Docs', prefix: 'wordcraft', src: '/wordcraft/', file: 'Document.docx', format: 'docx' },
  sheets: { name: 'Craft Sheets', prefix: 'gridcraft', src: '/craft-sheets/', file: 'Book.xlsx', format: 'xlsx' },
  slides: { name: 'Craft Slides', prefix: 'deckcraft', src: '/craft-slides/', file: 'Presentation.pptx', format: 'pptx' },
};

/** The page-bridge message types of an app's editor. */
function messages(prefix) {
  return {
    ready: `${prefix}:ready`,
    open: `${prefix}:open`,
    save: `${prefix}:save`,
    dirty: `${prefix}:dirty`,
    failed: `${prefix}:failed`,
    collab: `${prefix}:collab`,
    collabStatus: `${prefix}:collab.status`,
    call: `${prefix}:call`,
    result: `${prefix}:result`,
    event: `${prefix}:event`,
  };
}

/** The editor's events the page hears (`wordcraft:event`). */
const EDITOR_EVENTS = new Set(['change', 'selectionchange', 'presence', 'comment']);

const DEFAULT_CALL_TIMEOUT_MS = 30_000;

/** Relay close codes after which reconnecting can't help. */
const PERMANENT_CLOSE = {
  4400: 'the editor sent a bad client id',
  4401: 'not allowed to edit this document (the join token was refused)',
  4409: 'this document has too many people in it',
};

/** Close codes whose message is meant for the person: removed (with the relay's reason) or closed. */
const PERSON_CLOSE = {
  4403: (reason) => reason || 'Not allowed in this document',
  4410: () => 'This document was closed',
};

const DEFAULT_READY_TIMEOUT_MS = 60_000;

const MIME_BY_EXTENSION = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  docm: 'application/vnd.ms-word.document.macroEnabled.12',
  dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
  doc: 'application/msword',
  odt: 'application/vnd.oasis.opendocument.text',
  rtf: 'application/rtf',
  pdf: 'application/pdf',
  html: 'text/html',
  htm: 'text/html',
  md: 'text/markdown',
  tex: 'application/x-tex',
  txt: 'text/plain',
  png: 'image/png',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12',
  xls: 'application/vnd.ms-excel',
  ods: 'application/vnd.oasis.opendocument.spreadsheet',
  csv: 'text/csv',
  tsv: 'text/tab-separated-values',
  json: 'application/json',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  ppsx: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow',
  potx: 'application/vnd.openxmlformats-officedocument.presentationml.template',
  ppt: 'application/vnd.ms-powerpoint',
  odp: 'application/vnd.oasis.opendocument.presentation',
  deckcraft: 'application/zip',
};

/** The MIME type for a file name, by its extension. */
export function mimeType(name) {
  const extension = String(name).split('.').pop().toLowerCase();
  return MIME_BY_EXTENSION[extension] || 'application/octet-stream';
}

/** The last path segment of a URL, for naming a document opened from it. */
function nameFromUrl(url, fallback = 'Document.docx') {
  try {
    const segment = new URL(url, location.href).pathname.split('/').filter(Boolean).pop();
    return segment ? decodeURIComponent(segment) : fallback;
  } catch {
    return fallback;
  }
}

/** Bytes from a File, Blob, ArrayBuffer, typed array or URL, plus a name for them (`fallback`
 * when it has none). */
async function toDocument(source, name, fallback = 'Document.docx') {
  if (typeof source === 'string' || source instanceof URL) {
    const response = await fetch(source, { credentials: 'same-origin' });
    if (!response.ok) throw new Error(`Craft Collab: couldn't fetch ${source} (HTTP ${response.status})`);
    return { name: name || nameFromUrl(String(source), fallback), bytes: await response.arrayBuffer() };
  }
  if (typeof Blob !== 'undefined' && source instanceof Blob) {
    return { name: name || source.name || fallback, bytes: await source.arrayBuffer() };
  }
  if (source instanceof ArrayBuffer) return { name: name || fallback, bytes: source };
  if (ArrayBuffer.isView(source)) {
    // Copy, so transferring the buffer never detaches the caller's array.
    const copy = new Uint8Array(source.byteLength);
    copy.set(new Uint8Array(source.buffer, source.byteOffset, source.byteLength));
    return { name: name || fallback, bytes: copy.buffer };
  }
  throw new TypeError('Craft Collab: a document is a URL, File, Blob, ArrayBuffer or Uint8Array');
}

/** The editor page's URL with the host-mode query flags. */
function editorUrl(src, { author, webgl, collab, docName, language }) {
  const url = new URL(src, location.href);
  if (url.origin !== location.origin) {
    throw new Error(
      `Craft Collab: the editor (${url.origin}) must be served from this page's origin (${location.origin}). ` +
        'Host the editor folders on the same site, for example under /wordcraft/, /craft-sheets/ and /craft-slides/.',
    );
  }
  url.searchParams.set('host', 'parent');
  if (author) url.searchParams.set('author', author);
  if (webgl) url.searchParams.set('webgl', '');
  if (language) url.searchParams.set('lang', language);
  if (collab) {
    url.searchParams.set('collab', '1');
    if (docName) url.searchParams.set('docName', docName);
  }
  return url;
}

/**
 * An embedded editor of the suite (`options.app`). It is an EventTarget with these events, each a
 * CustomEvent whose `detail` is described below; the matching `on…` options are shortcuts for them:
 *
 * - `ready`  { version }                     the editor is running and can open documents
 * - `save`   { name, bytes, blob, saveAs }   the user saved or exported a file (see onSave)
 * - `dirty`  { dirty }                       unsaved changes appeared or went away
 * - `error`  { error }                       the editor couldn't start, or onSave failed
 * - `collab` { connected, pending, rev, people }  co-editing: connection, unsent edits or people changed
 * - `change` { rev, author }, `selectionchange`, `presence` { people }, `comment` (Docs)  what changed
 */
export class CraftEditor extends EventTarget {
  /**
   * @param {string|HTMLElement} container element (or selector) the editor fills
   * @param {object} options
   * @param {'docs'|'sheets'|'slides'} [options.app='docs'] which editor
   * @param {string} [options.src] URL of the hosted editor folder (same origin); by default
   *   /wordcraft/, /craft-sheets/ or /craft-slides/
   * @param {string|Blob|ArrayBuffer|Uint8Array} [options.document] document to open at start
   * @param {string} [options.documentName] name for `document` when it has none (sets the format)
   * @param {string} [options.author] user name for comments and tracked changes
   * @param {boolean} [options.webgl] force WebGL2 instead of WebGPU
   * @param {string} [options.language] interface language, e.g. 'ar', 'fr' (default: the browser's)
   * @param {boolean} [options.guardUnload=true] warn before leaving the page with unsaved changes
   * @param {number} [options.readyTimeout=60000] ms to wait for the editor before an `error`
   * @param {number} [options.callTimeout=30000] ms an API call (run, getText…) may take
   * @param {string} [options.title='Document editor'] accessible title of the iframe
   * @param {(file: {name: string, bytes: ArrayBuffer, blob: Blob, saveAs: boolean}) => any} [options.onSave]
   * @param {(info: {version: string}) => void} [options.onReady]
   * @param {(dirty: boolean) => void} [options.onDirtyChange]
   * @param {(error: Error) => void} [options.onError]
   * @param {{url: string, token: string | (() => string | Promise<string>)}} [options.collab]
   *   co-edit through a relay (relay/server.mjs): its WebSocket URL and a join token from your
   *   backend (a function is called again for each reconnect). Only the first person's
   *   `document` is opened; everyone else gets the shared copy.
   */
  constructor(container, options = {}) {
    super();
    const app = options.app ?? 'docs';
    if (!Object.hasOwn(APPS, app)) throw new TypeError(`Craft Collab: app must be one of ${Object.keys(APPS).join(', ')}`);
    const host = typeof container === 'string' ? document.querySelector(container) : container;
    if (!host) throw new Error(`Craft Collab: no element matches ${container}`);
    /** Which editor: 'docs', 'sheets' or 'slides'. */
    this.app = app;
    this._app = APPS[app];
    this._message = messages(this._app.prefix);
    this.options = { src: this._app.src, guardUnload: true, readyTimeout: DEFAULT_READY_TIMEOUT_MS, callTimeout: DEFAULT_CALL_TIMEOUT_MS, ...options, app };
    /** API calls waiting for their answer, by id. */
    this._calls = new Map();
    this._nextCall = 1;
    this.version = null;
    this.dirty = false;
    this.destroyed = false;

    for (const [type, callback, unwrap] of [
      ['ready', this.options.onReady, (d) => d],
      ['save', this.options.onSave, (d) => d],
      ['dirty', this.options.onDirtyChange, (d) => d.dirty],
      ['error', this.options.onError, (d) => d.error],
    ]) {
      if (typeof callback !== 'function') continue;
      this.addEventListener(type, (event) => {
        const result = callback(unwrap(event.detail));
        // A failed async onSave (the upload was refused, say) is reported as an `error`.
        if (type !== 'error' && typeof result?.catch === 'function') {
          result.catch((error) => this._emit('error', { error }));
        }
      });
    }

    const docName =
      this.options.documentName ||
      (typeof this.options.document === 'string' ? nameFromUrl(this.options.document, this._app.file) : this.options.document?.name) ||
      undefined;
    const url = editorUrl(this.options.src, { ...this.options, docName });
    this.docName = docName;
    this.origin = url.origin;

    this.iframe = document.createElement('iframe');
    this.iframe.src = url.href;
    this.iframe.title = this.options.title || `${this._app.name} editor`;
    this.iframe.allow = 'fullscreen; clipboard-read; clipboard-write; local-fonts';
    this.iframe.style.cssText = 'display:block;width:100%;height:100%;border:0;';

    this.ready = new Promise((resolve, reject) => {
      this._resolveReady = resolve;
      this._rejectReady = reject;
    });
    // A rejection nobody awaits shouldn't be reported as unhandled; `error` events still fire.
    this.ready.catch(() => {});

    this._onMessage = (event) => this._receive(event);
    window.addEventListener('message', this._onMessage);
    this._onBeforeUnload = (event) => {
      // Co-editing: the relay's autosaver saves the shared document; only edits that haven't
      // reached the relay yet would be lost.
      const unsaved = this._collab ? this._collab.pending > 0 : this.dirty;
      if (this.options.guardUnload && unsaved) {
        event.preventDefault();
        event.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', this._onBeforeUnload);

    this._readyTimer = setTimeout(() => {
      this._fail(
        new Error(
          `Craft Collab: ${this._app.name} did not start in time. Check that its files are served (with .wasm as ` +
            'application/wasm) at ' + url.pathname,
        ),
      );
    }, this.options.readyTimeout);

    host.appendChild(this.iframe);

    // API calls wait for the document they act on: the starting one opened, or, co-editing,
    // the shared one joined (its first status with a revision).
    this._opened = this.ready;
    if (this.options.collab) {
      this._collab = new CollabLink(this, this.options.collab);
      this.ready.then(() => this._collab?.connect(), () => {});
      this._opened = new Promise((resolve) => (this._joined = resolve));
    } else if (this.options.document != null) {
      const opening = this.open(this.options.document, this.options.documentName);
      opening.catch((error) => this._emit('error', { error }));
      this._opened = opening.then(() => this.ready);
    }
    this._opened.catch(() => {});
  }

  /**
   * Open a document in the editor, replacing what it shows. Waits until the editor is ready.
   * @param {string|URL|Blob|ArrayBuffer|Uint8Array} source
   * @param {string} [name] file name; its extension picks the format (default Document.docx,
   *   Book.xlsx or Presentation.pptx)
   */
  async open(source, name) {
    const [file] = await Promise.all([toDocument(source, name, this._app.file), this.ready]);
    if (this.destroyed) return;
    this.iframe.contentWindow.postMessage({ type: this._message.open, name: file.name, bytes: file.bytes }, this.origin, [file.bytes]);
  }

  /**
   * Run any command of the editor (the same names its ribbon and MCP tools use), e.g.
   * `run('format.bold', { value: true })`. Resolves with the command's result.
   * @param {string} command
   * @param {object} [params]
   */
  run(command, params = {}) {
    return this._call('run', { command, params });
  }

  /** The document's text (Craft Sheets: each sheet as CSV; Craft Slides: the outline). */
  async getText() {
    return (await this._call('getText')).text;
  }

  /**
   * The document as a file in `format`.
   * Craft Docs: docx, odt, rtf, pdf, html, md, txt. Craft Sheets: xlsx, csv, tsv, json, html, pdf.
   * Craft Slides: pptx, deckcraft, pdf, png (`options.slide`, from 0; else the one shown), txt, rtf.
   * @param {string} [format] the app's own format by default (docx, xlsx, pptx)
   * @param {{slide?: number}} [options]
   * @returns {Promise<{name: string, blob: Blob}>}
   */
  getDocument(format = this._app.format, options = {}) {
    return this._call('getDocument', { format, ...options });
  }

  /** Where the selection is (Craft Sheets: `{sheet, active, ranges}`; Craft Slides: `{slide, shapes, text}`). */
  getSelection() {
    return this._only(['sheets', 'slides'], 'getSelection', () => this._call('getSelection'));
  }

  /** Craft Sheets: the values in `range` ("A1:C3"), as rows; `sheet` defaults to the active one. */
  getRange(range, { sheet, formatted } = {}) {
    return this._only(['sheets'], 'getRange', () => this._call('getRange', { range: String(range), ...(sheet ? { sheet } : {}), ...(formatted ? { formatted: true } : {}) }));
  }

  /** Craft Sheets: write `values` (rows of values, or `=formulas`) from the top-left of `range`. */
  setRange(range, values, { sheet } = {}) {
    return this._only(['sheets'], 'setRange', () => this._call('setRange', { range: String(range), values, ...(sheet ? { sheet } : {}) }));
  }

  /** Craft Sheets: the active sheet, `{index, name}`. */
  activeSheet() {
    return this._only(['sheets'], 'activeSheet', () => this._call('activeSheet'));
  }

  /** Craft Slides: the slides, `{slides: [{index, id, title}], current}`. */
  slides() {
    return this._only(['slides'], 'slides', () => this._call('slides'));
  }

  /** Craft Slides: show slide `index` (from 0). */
  goToSlide(index) {
    return this._only(['slides'], 'goToSlide', () => this._call('goToSlide', { index: Number(index) }));
  }

  /** `fn()` when this editor is one of `apps`, else a rejection naming what it is. */
  _only(apps, method, fn) {
    if (!apps.includes(this.app)) return Promise.reject(new Error(`Craft Collab: ${method} is for ${apps.map((a) => APPS[a].name).join(' and ')}, not ${this._app.name}`));
    return fn();
  }

  /** Craft Docs: the document's structure (headings, paragraphs, tables, comments) as JSON. */
  inspect() {
    return this._call('inspect');
  }

  /** Craft Docs: type `text` at the cursor, replacing the selection. */
  async insertText(text) {
    await this._call('insertText', { text: String(text) });
  }

  /** Craft Docs: select the first occurrence of `text`. */
  async select(text) {
    await this._call('select', { text: String(text) });
  }

  /** Read only (`true`) or editable again: this person's edits are refused while read only. */
  async setReadOnly(on = true) {
    await this._call('setReadOnly', { on: Boolean(on) });
  }

  /** Send API call `method` once the editor is ready; resolves with its result. */
  async _call(method, params = {}) {
    await this._opened;
    await this.ready;
    if (this.destroyed) throw new Error('Craft Collab: the editor was destroyed');
    const id = this._nextCall++;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this._calls.delete(id);
        reject(new Error(`Craft Collab: ${method} timed out`));
      }, this.options.callTimeout);
      this._calls.set(id, { resolve, reject, timer });
      this.iframe.contentWindow.postMessage({ type: this._message.call, id, method, params }, this.origin);
    });
  }

  /** Give keyboard focus to the editor. */
  focus() {
    this.iframe.focus();
  }

  /** Remove the editor from the page and stop listening. */
  destroy() {
    if (this.destroyed) return;
    this.destroyed = true;
    clearTimeout(this._readyTimer);
    window.removeEventListener('message', this._onMessage);
    window.removeEventListener('beforeunload', this._onBeforeUnload);
    this._collab?.stop();
    this._collab = null;
    this._joined?.();
    for (const { reject, timer } of this._calls.values()) {
      clearTimeout(timer);
      reject(new Error('Craft Collab: the editor was destroyed'));
    }
    this._calls.clear();
    this.iframe.remove();
    this._rejectReady(new Error('Craft Collab: the editor was destroyed'));
  }

  _receive(event) {
    // Only our own frame, on our own origin, is heard.
    if (event.origin !== this.origin || event.source !== this.iframe.contentWindow) return;
    const data = event.data;
    if (!data || typeof data.type !== 'string') return;
    const MESSAGE = this._message;
    switch (data.type) {
      case MESSAGE.ready:
        clearTimeout(this._readyTimer);
        this.version = data.version || null;
        this._resolveReady({ version: this.version });
        this._emit('ready', { version: this.version });
        break;
      case MESSAGE.dirty:
        this.dirty = Boolean(data.dirty);
        this._emit('dirty', { dirty: this.dirty });
        break;
      case MESSAGE.save: {
        const bytes = data.bytes instanceof ArrayBuffer ? data.bytes : new Uint8Array(data.bytes).buffer;
        const name = String(data.name || this._app.file);
        const blob = new Blob([bytes], { type: mimeType(name) });
        this._emit('save', { name, bytes, blob, saveAs: Boolean(data.saveAs) });
        break;
      }
      case MESSAGE.collab:
        if (typeof data.msg === 'string') this._collab?.toRelay(data.msg);
        break;
      case MESSAGE.collabStatus:
        this._collab?.status(data);
        if (data.rev != null) this._joined?.();
        break;
      case MESSAGE.result: {
        const call = this._calls.get(data.id);
        if (!call) break;
        this._calls.delete(data.id);
        clearTimeout(call.timer);
        if (!data.ok) {
          call.reject(new Error(String(data.error || 'the call failed')));
        } else if (data.bytes != null) {
          const bytes = data.bytes instanceof ArrayBuffer ? data.bytes : new Uint8Array(data.bytes).buffer;
          const name = String(data.name || this._app.file);
          call.resolve({ name, blob: new Blob([bytes], { type: mimeType(name) }) });
        } else {
          call.resolve(data.result);
        }
        break;
      }
      case MESSAGE.event:
        if (EDITOR_EVENTS.has(data.event)) this._emit(data.event, data.detail);
        break;
      case MESSAGE.failed:
        this._fail(new Error(`${this._app.name} couldn't start: ${data.error || 'unknown error'}`));
        break;
    }
  }

  _fail(error) {
    clearTimeout(this._readyTimer);
    this._rejectReady(error);
    this._emit('error', { error });
  }

  _emit(type, detail) {
    this.dispatchEvent(new CustomEvent(type, { detail }));
  }
}

/**
 * Co-editing: one WebSocket to the relay, carrying messages between it and the editor frame.
 * The editor's replica does the real work (and resends unconfirmed edits after a reconnect);
 * this only connects, reconnects, and opens the document for the first person.
 */
class CollabLink {
  constructor(editor, { url, token, WebSocket: WS = globalThis.WebSocket, retryMs = [500, 10_000] }) {
    this.editor = editor;
    this.url = url;
    this.token = token;
    this.WS = WS;
    this.retry = retryMs;
    this.delay = retryMs[0];
    // A relay client id: letters and digits only.
    this.client = (globalThis.crypto?.randomUUID?.() || String(Math.random()).slice(2) + Date.now()).replace(/[^A-Za-z0-9]/g, '');
    this.ws = null;
    this.stopped = false;
    this.rev = null;
    this.pending = 0;
    this.people = 1;
    this.connected = false;
  }

  async connect() {
    if (this.stopped) return;
    let token;
    try {
      token = typeof this.token === 'function' ? await this.token() : this.token;
    } catch (error) {
      this.editor._emit('error', { error: new Error(`Craft Collab: couldn't get a join token: ${error.message || error}`) });
      return this.reconnectLater();
    }
    if (this.stopped) return;
    const ws = new this.WS(this.url);
    this.ws = ws;
    // Online (and the backoff reset) only once the relay welcomes this connection: one it
    // closes right away (too many people, replaced, refused) must not restart the editor's
    // offline clock.
    ws.onopen = () => {
      // Protocol v3: the editor says which app it is (a v2 relay ignores it).
      const hello = { t: 'hello', token, client: this.client, app: this.editor.app };
      if (this.rev !== null) hello.haveRev = this.rev;
      ws.send(JSON.stringify(hello));
    };
    ws.onmessage = (event) => this.fromRelay(String(event.data), ws);
    ws.onclose = (event) => {
      if (this.ws !== ws) return;
      this.ws = null;
      this.setConnected(false);
      const permanent = PERMANENT_CLOSE[event.code];
      const forPerson = PERSON_CLOSE[event.code];
      if (permanent || forPerson) {
        // A wrong app or token's own reason says more than the code's.
        const error = new Error(forPerson ? forPerson(event.reason) : `Craft Collab: ${event.code === 4400 && event.reason ? event.reason : permanent}`);
        error.code = event.code;
        this.editor._emit('error', { error });
        return;
      }
      this.reconnectLater();
    };
    ws.onerror = () => {};
  }

  reconnectLater() {
    if (this.stopped) return;
    setTimeout(() => this.connect(), this.delay);
    this.delay = Math.min(this.delay * 2, this.retry[1]);
  }

  fromRelay(text, ws) {
    // The first person (no snapshot yet) opens the document; everyone else gets the shared copy.
    if (text.startsWith('{"t":"Welcome"')) {
      if (ws && ws === this.ws && !this.connected) {
        this.delay = this.retry[0];
        this.setConnected(true);
      }
      try {
        const welcome = JSON.parse(text);
        if (welcome.snapshot === null && this.rev === null && this.editor.options.document != null) {
          this.editor.open(this.editor.options.document, this.editor.docName).catch((error) => this.editor._emit('error', { error }));
        }
      } catch {}
    }
    this.toEditor(text);
  }

  toEditor(msg) {
    if (this.editor.destroyed) return;
    this.editor.iframe.contentWindow.postMessage({ type: this.editor._message.collab, msg }, this.editor.origin);
  }

  /** A message from the editor for the relay; dropped while disconnected (the replica resends its edits). */
  toRelay(msg) {
    if (this.ws && this.ws.readyState === (this.WS.OPEN ?? 1)) this.ws.send(msg);
  }

  status({ pending, rev, people }) {
    if (Number.isInteger(rev)) this.rev = rev;
    if (Number.isInteger(pending)) this.pending = pending;
    if (Number.isInteger(people)) this.people = people;
    this.emit();
  }

  setConnected(on) {
    this.connected = on;
    // The editor goes read only after a long time offline (and back once its edits sync).
    this.toEditor(JSON.stringify({ t: 'Connection', online: on }));
    this.emit();
  }

  emit() {
    this.editor._emit('collab', { connected: this.connected, pending: this.pending, rev: this.rev, people: this.people });
  }

  stop() {
    this.stopped = true;
    const ws = this.ws;
    this.ws = null;
    ws?.close();
  }
}

/** A Craft Docs editor, as wordcraft-embed 1.x made it (`app` is always 'docs'). */
export class WordCraftEditor extends CraftEditor {
  constructor(container, options = {}) {
    super(container, { ...options, app: 'docs' });
  }
}

/** Mount an editor; the same as `new CraftEditor(container, options)` (`options.app` picks it). */
export function createEditor(container, options) {
  return new CraftEditor(container, options);
}

// Server-side rendering (Next.js, Nuxt…) imports this module where there is no DOM.
const ElementBase = typeof HTMLElement === 'undefined' ? class {} : HTMLElement;

/** The events an element passes on, as `<prefix>-<type>` CustomEvents. */
const ELEMENT_EVENTS = ['ready', 'save', 'dirty', 'error', 'collab', 'change', 'selectionchange', 'presence', 'comment'];

/**
 * <craft-editor app="sheets" src="/craft-sheets/" document="/files/budget.xlsx" author="Jane"></craft-editor>
 *
 * Attributes: app, src, document, document-name, author, webgl, no-unload-guard, title. Events
 * (CustomEvent, same `detail` as CraftEditor): craft-ready, craft-save, craft-dirty, craft-error,
 * craft-change… Methods: open(source, name), focus() and the editor's API (run, getText,
 * getDocument, getSelection, getRange, setRange, activeSheet, slides, goToSlide, inspect,
 * insertText, select, setReadOnly). Property: editor (the CraftEditor). Give the element a
 * height; it is a block that the editor fills.
 */
export class CraftEditorElement extends ElementBase {
  static get observedAttributes() {
    return ['document'];
  }

  /** The app when the element has no `app` attribute, and its events' prefix. */
  get _defaults() {
    return { app: 'docs', prefix: 'craft' };
  }

  connectedCallback() {
    if (!this.style.display) this.style.display = 'block';
    const { app, prefix } = this._defaults;
    const chosen = this.getAttribute('app') || app;
    this.editor = new CraftEditor(this, {
      app: chosen,
      src: this.getAttribute('src') || APPS[chosen]?.src,
      document: this.getAttribute('document') || undefined,
      documentName: this.getAttribute('document-name') || undefined,
      author: this.getAttribute('author') || undefined,
      webgl: this.hasAttribute('webgl'),
      guardUnload: !this.hasAttribute('no-unload-guard'),
      title: this.getAttribute('title') || undefined,
    });
    for (const type of ELEMENT_EVENTS) {
      this.editor.addEventListener(type, (event) => {
        this.dispatchEvent(new CustomEvent(`${prefix}-${type}`, { detail: event.detail, bubbles: true, composed: true }));
      });
    }
  }

  disconnectedCallback() {
    this.editor?.destroy();
    this.editor = null;
  }

  attributeChangedCallback(name, previous, value) {
    if (name === 'document' && this.editor && value && value !== previous) {
      this.editor.open(value, this.getAttribute('document-name') || undefined);
    }
  }

  open(source, name) {
    return this._editor().then((editor) => editor.open(source, name));
  }

  focus() {
    this.editor?.focus();
  }

  /** The editor, or a rejection while this element isn't in the page. */
  _editor() {
    if (this.editor) return Promise.resolve(this.editor);
    return Promise.reject(new Error(`Craft Collab: the <${this.localName || 'craft-editor'}> element is not in the page`));
  }
}

for (const method of ['run', 'getText', 'getDocument', 'getSelection', 'getRange', 'setRange', 'activeSheet', 'slides', 'goToSlide', 'inspect', 'insertText', 'select', 'setReadOnly']) {
  CraftEditorElement.prototype[method] = function (...args) {
    return this._editor().then((editor) => editor[method](...args));
  };
}

/** <wordcraft-editor> (1.x): a Craft Docs editor whose events are wordcraft-ready, wordcraft-save… */
export class WordCraftEditorElement extends CraftEditorElement {
  get _defaults() {
    return { app: 'docs', prefix: 'wordcraft' };
  }

  connectedCallback() {
    // Always Craft Docs, as in 1.x, whatever `app` says.
    this.removeAttribute?.('app');
    super.connectedCallback();
  }
}

if (typeof customElements !== 'undefined') {
  if (!customElements.get('craft-editor')) customElements.define('craft-editor', CraftEditorElement);
  if (!customElements.get('wordcraft-editor')) customElements.define('wordcraft-editor', WordCraftEditorElement);
}

if (typeof window !== 'undefined') {
  window.CraftCollab = { CraftEditor, WordCraftEditor, createEditor, mimeType, APPS };
  window.WordCraft = window.CraftCollab;
}
