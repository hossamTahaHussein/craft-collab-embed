// The suite's three editors on the relay: what each one's selections look like and which
// commands each role may run in it. Ops and snapshots stay opaque; only these differ per app.
//
//   docs    Craft Docs (WordCraft Collab): selections are story offsets (wordcraft_collab::proto::Anchor)
//   sheets  Craft Sheets: a sheet and its selection (gridcraft_collab::SheetSel)
//   slides  Craft Slides: a slide and its selection (deckcraft_collab::SlideSel)

import { readFileSync } from 'node:fs';

const list = (name) => new Set(JSON.parse(readFileSync(new URL(`./${name}`, import.meta.url), 'utf8')));
const isCount = (n) => Number.isSafeInteger(n) && n >= 0;
/** The largest story part index and offset a replica can hold (u32). */
const MAX_INDEX = 2 ** 32 - 1;
const isIndex = (n) => isCount(n) && n <= MAX_INDEX;
const isObject = (v) => !!v && typeof v === 'object' && !Array.isArray(v);
/** A selection is forwarded to everyone each time it moves: it stays small. */
const MAX_SEL_BYTES = 16 * 1024;
const small = (v) => {
  try {
    return JSON.stringify(v).length <= MAX_SEL_BYTES;
  } catch {
    return false;
  }
};

export const APPS = {
  docs: {
    name: 'Craft Docs',
    validSel: (sel) =>
      isObject(sel) &&
      (sel.story === 'body' || (isObject(sel.story) && isIndex(sel.story.part))) &&
      isIndex(sel.anchor) &&
      isIndex(sel.head),
    cleanSel: (sel) => ({ story: sel.story === 'body' ? 'body' : { part: sel.story.part }, anchor: sel.anchor, head: sel.head }),
    /** The commands a commenter may run (the engine's `outbox::COMMENT_COMMANDS`). */
    comment: list('comment-commands.json'),
    /** What a suggester may run besides comments: the commands the engine records whole as a
     * tracked change (`outbox::SUGGEST_COMMANDS`, each proved by a Reject All test). */
    suggest: list('suggest-commands.json'),
    /** The comment commands and the `commented` webhook action each one is. */
    commentActions: {
      'review.newComment': 'added',
      'review.reply': 'replied',
      'review.resolveComment': 'resolved',
      'review.editComment': 'edited',
      'review.deleteComment': 'deleted',
    },
  },
  sheets: {
    name: 'Craft Sheets',
    validSel: (sel) => isObject(sel) && typeof sel.sheet === 'string' && sel.sheet.length <= 255 && isObject(sel.selection) && small(sel),
    cleanSel: (sel) => {
      const out = { sheet: sel.sheet, selection: sel.selection };
      if (Array.isArray(sel.refs)) out.refs = sel.refs;
      return out;
    },
    /** gridcraft_engine::collab::COMMENT_COMMANDS. */
    comment: list('comment-commands.sheets.json'),
    suggest: list('suggest-commands.sheets.json'),
    commentActions: {
      'review.newComment': 'added',
      'review.newNote': 'added',
      'review.replyComment': 'replied',
      'review.resolveComment': 'resolved',
      'review.deleteComment': 'deleted',
    },
  },
  slides: {
    name: 'Craft Slides',
    validSel: (sel) => isObject(sel) && isCount(sel.slide) && isObject(sel.selection) && small(sel),
    cleanSel: (sel) => {
      const out = { slide: sel.slide, selection: sel.selection };
      if (Array.isArray(sel.text)) out.text = sel.text;
      if (Array.isArray(sel.slides)) out.slides = sel.slides;
      return out;
    },
    /** deckcraft_engine::collab::COMMENT_COMMANDS. */
    comment: list('comment-commands.slides.json'),
    suggest: list('suggest-commands.slides.json'),
    commentActions: {
      'comment.add': 'added',
      'comment.reply': 'replied',
      'comment.resolve': 'resolved',
      'comment.delete': 'deleted',
    },
  },
};

export const APP_IDS = Object.keys(APPS);

/** The room an app's document lives in. Craft Docs keeps the document id itself (as relays
 * before v3 did, so their data is found); the others prefix it with a control character no
 * document id may contain (token.mjs), so no two apps' rooms can meet. */
export function roomKey(app, doc) {
  return app === 'docs' ? doc : `\u0001${app}\u0001${doc}`;
}

/** The app and document of a room key. */
export function fromRoomKey(key) {
  const m = /^\u0001([a-z]+)\u0001(.*)$/s.exec(key);
  return m ? { app: m[1], doc: m[2] } : { app: 'docs', doc: key };
}
