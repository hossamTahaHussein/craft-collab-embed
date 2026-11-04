#!/usr/bin/env node
// Download the Craft Collab Suite's web editors from a release, check each against the release's
// SHA256SUMS.txt and unpack them (the static sites the SDK frames):
//
//   editor/          Craft Docs     wordcraft-web-<version>.zip
//   editor-sheets/   Craft Sheets   craft-sheets-web-<version>.zip
//   editor-slides/   Craft Slides   craft-slides-web-<version>.zip
//
//   node scripts/fetch-editor.mjs                 # the version pinned in package.json ("craftVersion")
//   node scripts/fetch-editor.mjs 2.0.0           # a specific release
//   node scripts/fetch-editor.mjs latest
//   node scripts/fetch-editor.mjs 2.0.0 sheets    # one editor only (docs, sheets or slides)
//
// Needs Node 18+ and `unzip` on PATH.

import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync, renameSync, existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

// The suite's repository (GitHub redirects its old name, wordcraft-collab).
const REPO = 'hossamTahaHussein/craft-collab-suite';
const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const manifest = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));
const pinned = manifest.craftVersion ?? manifest.wordcraftVersion;
/** Each editor's release zip (its name before the version) and the folder it goes to. */
const EDITORS = {
  docs: { zip: 'wordcraft-web', dir: 'editor', name: 'Craft Docs' },
  sheets: { zip: 'craft-sheets-web', dir: 'editor-sheets', name: 'Craft Sheets' },
  slides: { zip: 'craft-slides-web', dir: 'editor-slides', name: 'Craft Slides' },
};

async function latestVersion() {
  const response = await fetch(`https://api.github.com/repos/${REPO}/releases/latest`);
  if (!response.ok) throw new Error(`GitHub API: HTTP ${response.status}`);
  return (await response.json()).tag_name.replace(/^v/, '');
}

async function download(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`${url}: HTTP ${response.status}`);
  return Buffer.from(await response.arrayBuffer());
}

let version = process.argv[2] || pinned;
if (version === 'latest') version = await latestVersion();
const only = process.argv[3];
if (only && !Object.hasOwn(EDITORS, only)) throw new Error(`no editor "${only}": one of ${Object.keys(EDITORS).join(', ')}`);
const base = `https://github.com/${REPO}/releases/download/v${version}`;
const sums = (await download(`${base}/SHA256SUMS.txt`)).toString('utf8');

for (const [app, editor] of Object.entries(EDITORS)) {
  if (only && app !== only) continue;
  const zipName = `${editor.zip}-${version}.zip`;
  console.log(`Downloading ${zipName} from ${REPO}…`);
  const zip = await download(`${base}/${zipName}`);
  const expected = sums
    .split('\n')
    .map((line) => line.trim().split(/\s+/))
    .find(([, name]) => name === zipName)?.[0];
  const actual = createHash('sha256').update(zip).digest('hex');
  if (!expected) throw new Error(`${zipName} is not listed in SHA256SUMS.txt`);
  if (expected !== actual) throw new Error(`Checksum mismatch for ${zipName}: expected ${expected}, got ${actual}`);
  console.log(`SHA-256 verified: ${actual}`);
  // Unpack beside the folder so the final rename stays on one filesystem.
  const work = mkdtempSync(join(root, `.${editor.zip}-`));
  try {
    writeFileSync(join(work, zipName), zip);
    execFileSync('unzip', ['-q', join(work, zipName), '-d', work]);
    const dir = join(root, editor.dir);
    if (existsSync(dir)) rmSync(dir, { recursive: true });
    renameSync(join(work, `${editor.zip}-${version}`), dir);
    writeFileSync(join(dir, 'VERSION'), `${version}\n`);
    console.log(`${editor.name} ${version} unpacked into ${editor.dir}/`);
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
}
