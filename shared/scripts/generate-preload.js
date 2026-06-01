#!/usr/bin/env node
// Called by build.sh after assembling each module:
//   node generate-preload.js <outDir> <moduleId>
//
// Reads every .html and .hbs file under <outDir>/templates/, embeds their
// contents as JSON-encoded strings, and writes <outDir>/scripts/preload-templates.js.
// That script registers all templates as Handlebars partials at module-load time
// so Foundry's loadTemplates() finds them already cached and skips the HTTP fetch.

import { readFileSync, writeFileSync, readdirSync } from 'fs';
import { join } from 'path';

const [outDir, moduleId] = process.argv.slice(2);

if (!outDir || !moduleId) {
  process.stderr.write('usage: generate-preload.js <outDir> <moduleId>\n');
  process.exit(1);
}

function walkTemplates(dir, base) {
  const results = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const rel = base ? `${base}/${entry.name}` : entry.name;
    if (entry.isDirectory()) {
      results.push(...walkTemplates(join(dir, entry.name), rel));
    } else if (/\.(html|hbs)$/.test(entry.name)) {
      results.push(rel);
    }
  }
  return results.sort();
}

const templatesDir = join(outDir, 'templates');
const files = walkTemplates(templatesDir, '');

const entries = files.map(rel => {
  const content = readFileSync(join(templatesDir, rel), 'utf8');
  const key     = `modules/${moduleId}/templates/${rel}`;
  return `  ${JSON.stringify(key)}: ${JSON.stringify(content)}`;
}).join(',\n');

const js =
`const _Hb = globalThis.Handlebars ?? foundry?.applications?.handlebars?.Handlebars;
if (_Hb?.registerPartial && _Hb?.compile) {
  const _t = {
${entries}
  };
  for (const _p in _t) {
    if (!(_p in (_Hb.partials ?? {}))) _Hb.registerPartial(_p, _Hb.compile(_t[_p]));
  }
}
`;

writeFileSync(join(outDir, 'scripts', 'preload-templates.js'), js);
process.stdout.write(`  → preload-templates.js (${files.length} template${files.length !== 1 ? 's' : ''})\n`);
