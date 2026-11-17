/**
 * node roll20/build.mjs
 *
 * Builds roll20/dist/:
 *   cfer-roll20-extension/          unpacked extension (load it in chrome://extensions)
 *   cfer-roll20-extension.zip       for the Chrome Web Store and Firefox Add-ons
 *   CferImporter.js                 the companion Mod script, ready to paste
 */
import { cpSync, rmSync, mkdirSync, copyFileSync, writeFileSync, readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';

const here = fileURLToPath(new URL('.', import.meta.url));
const dist = join(here, 'dist');
const ext = join(dist, 'cfer-roll20-extension');

rmSync(dist, { recursive: true, force: true });
mkdirSync(dist, { recursive: true });
cpSync(join(here, 'extension'), ext, { recursive: true });
cpSync(join(here, 'core'), join(ext, 'core'), { recursive: true });
// The vendor file re-exports the Foundry module's pure wall/light converter; ship a real copy.
copyFileSync(join(here, '..', 'shared', 'scripts', 'core', 'maps', 'foundry.js'), join(ext, 'core', 'vendor', 'maps-foundry.js'));

const version = JSON.parse(readFileSync(join(ext, 'manifest.json'), 'utf8')).version;
execFileSync('zip', ['-qr', join(dist, 'cfer-roll20-extension.zip'), '.'], { cwd: ext });
copyFileSync(join(here, 'mod', 'CferImporter.js'), join(dist, 'CferImporter.js'));
writeFileSync(join(dist, 'BUILD.txt'), `Cfer for Roll20 extension ${version}\nBuilt ${new Date().toISOString()}\n`);
console.log(`Built roll20/dist (extension ${version}).`);
