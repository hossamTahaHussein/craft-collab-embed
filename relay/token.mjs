// Join tokens for the co-editing relay: a compact JWT (HS256) that says "this user may open
// this document, in this role". The host site's backend issues it with a secret it shares with
// the relay; the browser passes it along untouched.
//
//   claims: { doc, sub, name, role: 'edit' | 'suggest' | 'comment' | 'view', app?, iat, exp }
//
// `app` (protocol v3) is the editor: 'docs' (the default, as tokens before v3), 'sheets' or
// 'slides'. A document id holds no control characters.

import { createHmac, timingSafeEqual } from 'node:crypto';

export const ROLES = ['edit', 'suggest', 'comment', 'view'];
/** The suite's editors (apps.mjs). */
export const APPS = ['docs', 'sheets', 'slides'];
const MAX_LIFETIME_S = 24 * 3600;
const DEFAULT_LIFETIME_S = 12 * 3600;
const MAX_TOKEN_LENGTH = 4096;
const MIN_SECRET_LENGTH = 16;
const HEADER = Buffer.from(JSON.stringify({ alg: 'HS256', typ: 'JWT' })).toString('base64url');

export class TokenError extends Error {
  /** @param {'malformed' | 'bad-signature' | 'expired' | 'too-long-lived'} code */
  constructor(code, message) {
    super(message || code);
    this.code = code;
  }
}

const nowSeconds = () => Math.floor(Date.now() / 1000);

function sign(secret, data) {
  return createHmac('sha256', secret).update(data).digest('base64url');
}

function checkSecret(secret) {
  if (typeof secret !== 'string' || secret.length < MIN_SECRET_LENGTH) {
    throw new TokenError('malformed', `the relay secret must be at least ${MIN_SECRET_LENGTH} characters`);
  }
}

function checkClaims(c) {
  for (const key of ['doc', 'sub', 'name']) {
    if (typeof c[key] !== 'string' || c[key].length === 0 || c[key].length > 200) {
      throw new TokenError('malformed', `claim "${key}" must be a string of 1–200 characters`);
    }
  }
  // Control characters keep the apps' rooms apart (apps.mjs `roomKey`) and have no place in an id.
  if (/[\u0000-\u001f\u007f]/.test(c.doc)) throw new TokenError('malformed', 'claim "doc" must not hold control characters');
  if (!ROLES.includes(c.role)) throw new TokenError('malformed', `claim "role" must be one of ${ROLES.join(', ')}`);
  if (c.app !== undefined && !APPS.includes(c.app)) throw new TokenError('malformed', `claim "app" must be one of ${APPS.join(', ')}`);
  if (!Number.isInteger(c.iat) || !Number.isInteger(c.exp)) throw new TokenError('malformed', 'claims "iat" and "exp" must be integers');
  if (c.exp - c.iat > MAX_LIFETIME_S) throw new TokenError('too-long-lived', 'tokens may live at most 24 hours');
}

/**
 * Issue a token. `exp` defaults to 12 hours from now and may be at most 24 hours away.
 * @param {string} secret shared with the relay (16+ characters)
 * @param {{doc: string, sub: string, name: string, role: string, app?: 'docs' | 'sheets' | 'slides', exp?: number}} claims
 */
export function issueToken(secret, claims, now = nowSeconds()) {
  checkSecret(secret);
  const c = { doc: claims.doc, sub: claims.sub, name: claims.name, role: claims.role, iat: now, exp: claims.exp ?? now + DEFAULT_LIFETIME_S };
  if (claims.app !== undefined) c.app = claims.app;
  checkClaims(c);
  const payload = Buffer.from(JSON.stringify(c)).toString('base64url');
  return `${HEADER}.${payload}.${sign(secret, `${HEADER}.${payload}`)}`;
}

/**
 * Check a token and return its claims, or throw a TokenError.
 * @returns {{doc: string, sub: string, name: string, role: string, app: 'docs' | 'sheets' | 'slides', iat: number, exp: number}}
 */
export function verifyToken(secret, token, now = nowSeconds()) {
  checkSecret(secret);
  if (typeof token !== 'string' || token.length > MAX_TOKEN_LENGTH) throw new TokenError('malformed');
  const parts = token.split('.');
  if (parts.length !== 3) throw new TokenError('malformed');
  const [header, payload, signature] = parts;
  let h;
  let c;
  try {
    h = JSON.parse(Buffer.from(header, 'base64url').toString('utf8'));
    c = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8'));
  } catch {
    throw new TokenError('malformed');
  }
  if (h?.alg !== 'HS256' || !c || typeof c !== 'object') throw new TokenError('malformed');
  const want = Buffer.from(sign(secret, `${header}.${payload}`));
  const got = Buffer.from(signature);
  if (want.length !== got.length || !timingSafeEqual(want, got)) throw new TokenError('bad-signature');
  checkClaims(c);
  if (now >= c.exp) throw new TokenError('expired');
  return { doc: c.doc, sub: c.sub, name: c.name, role: c.role, app: c.app ?? 'docs', iat: c.iat, exp: c.exp };
}
