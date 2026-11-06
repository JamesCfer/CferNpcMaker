// Standardises Discord posts across the generator workflows:
// one message per generation, "**<Generator> — <Title>** (... attached)", with the result attached.
// Posting branches never sit on the response path and continue on error.
//
//   node tools/n8n/discord-post-migration.mjs <inDir> <outDir>
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const DISCORD_CREDENTIALS = { discordBotApi: { id: 'v845fBzdtlCbenJa', name: 'Discord Bot account' } };
const GUILD = { __rl: true, value: '853532551784431616', mode: 'list', cachedResultName: 'Celestia' };
const CHANNEL = { __rl: true, value: '982754114901512212', mode: 'list', cachedResultName: 'jane' };

const clone = value => JSON.parse(JSON.stringify(value));
const byName = (workflow, name) => workflow.nodes.find(node => node.name === name);
const discordNodes = workflow => workflow.nodes.filter(node => node.type === 'n8n-nodes-base.discord');

function sendNode(name, position, inputField = 'file') {
  return {
    id: `discord-${name.toLowerCase().replace(/[^a-z0-9]+/g, '-')}`,
    name,
    type: 'n8n-nodes-base.discord',
    typeVersion: 2,
    position,
    parameters: {
      resource: 'message',
      guildId: clone(GUILD),
      channelId: clone(CHANNEL),
      content: '={{ $json.content }}',
      options: {},
      files: { values: [{ inputFieldName: inputField }] },
    },
    credentials: clone(DISCORD_CREDENTIALS),
    onError: 'continueRegularOutput',
  };
}

function prepNode(name, position, jsCode) {
  return {
    id: `discord-${name.toLowerCase().replace(/[^a-z0-9]+/g, '-')}`,
    name,
    type: 'n8n-nodes-base.code',
    typeVersion: 2,
    position,
    parameters: { jsCode },
    onError: 'continueRegularOutput',
  };
}

function addMain(workflow, from, to) {
  const outputs = (workflow.connections[from] ??= { main: [[]] }).main;
  outputs[0] ??= [];
  if (!outputs[0].some(link => link.node === to)) outputs[0].push({ node: to, type: 'main', index: 0 });
}

function removeNode(workflow, name) {
  workflow.nodes = workflow.nodes.filter(node => node.name !== name);
  delete workflow.connections[name];
  for (const value of Object.values(workflow.connections)) {
    for (const output of value.main || []) {
      if (!output) continue;
      for (let i = output.length - 1; i >= 0; i--) if (output[i].node === name) output.splice(i, 1);
    }
  }
}

/** Nodes reachable from the main webhook up to any respondToWebhook, ignoring Discord branches. */
export function responsePath(workflow) {
  const webhook = workflow.nodes.find(node => node.type === 'n8n-nodes-base.webhook' && node.parameters?.httpMethod !== 'OPTIONS');
  if (!webhook) return [];
  const seen = new Set();
  const queue = [webhook.name];
  while (queue.length) {
    const name = queue.shift();
    if (seen.has(name)) continue;
    seen.add(name);
    for (const output of workflow.connections[name]?.main || []) {
      for (const link of output || []) {
        const target = byName(workflow, link.node);
        if (!target || target.type === 'n8n-nodes-base.discord' || /discord/i.test(target.name) || target.name === 'Wait') continue;
        queue.push(link.node);
      }
    }
  }
  return [...seen].sort();
}

// Shared code pieces for new prep nodes.
const SLUG = "function slugify(s){return String(s||'result').toLowerCase().replace(/[^a-z0-9._-]+/g,'-').replace(/^-+|-+$/g,'').slice(0,60)||'result';}";
const SHORT = "function short(s,n){s=String(s||'').replace(/\\s+/g,' ').trim();return s.length>n?s.slice(0,n-1).trimEnd()+'…':s;}";

const NEW_POSTS = {
  // Hero 6e NPC builder: native HDC file.
  pVuhHAnuo8nfrB1U: {
    after: 'Build Response',
    code: `${SLUG}
const d=$json||{};
if(!d.ok||!d.hdcXml)return [];
const pts=d.metadata&&d.metadata.points?' ('+d.metadata.points+' pts)':'';
const title='Hero 6e NPC — '+(d.name||'Unnamed NPC')+pts;
return [{json:{content:'**'+title+'** (HDC attached)'},binary:{file:{data:Buffer.from(String(d.hdcXml),'utf8').toString('base64'),fileName:(d.filename||slugify(d.name)+'.hdc'),mimeType:'application/xml'}}}];`,
  },
  KUcKqh0f3k8shbJV: { after: 'Parse & Assemble Item', label: 'PF2e Wiki Item' },
  TtU4JiQQdOX8UUSi: { after: 'Parse & Assemble Item', label: 'D&D 5e Wiki Item' },
  // OpenAI image generator: post the image.
  gPSWqhlldR7uQgeH: {
    after: 'Build Data URL',
    code: `${SLUG}
${SHORT}
const d=$json||{};
if(!d.ok||!d.imageUrl)return [];
let prompt='';try{prompt=$('Validate Input').first().json.prompt||'';}catch{}
const m=String(d.imageUrl).match(/^data:([^;]+);base64,(.*)$/);
if(!m)return [];
const title='OpenAI Image — '+(short(prompt,80)||'image');
return [{json:{content:'**'+title+'** (image attached)'},binary:{file:{data:m[2],fileName:slugify(prompt)+'.'+(m[1].split('/')[1]||'png'),mimeType:m[1]}}}];`,
  },
};

const WIKI_CODE = label => `${SLUG}
const d=$json||{};
const item=d.foundryItem||d.item;
if(!d.ok||!item)return [];
const type=item.type?' ('+item.type+')':'';
const title='${label} — '+(item.name||'Unnamed item')+type;
return [{json:{content:'**'+title+'** (JSON attached)'},binary:{file:{data:Buffer.from(JSON.stringify(item,null,2),'utf8').toString('base64'),fileName:slugify(item.name)+'.json',mimeType:'application/json'}}}];`;

// Label fixes in existing prep code.
const LABEL_FIXES = {
  N5e4EppZCKYZL5hU: [['`D&D 5e Store — ', '`D&D 3.5e Store — ']],
};

export function migrateDiscord(input) {
  const workflow = clone(input);
  const changes = [];
  const before = responsePath(input).join('|');
  const id = workflow.id;
  const isDev = /-\(DEV\)-/.test(workflow.name || '');
  const existing = discordNodes(workflow);

  if (workflow.name === 'Map Builder') return { workflow, changes, skipped: 'generated by tools/n8n/build-map-workflow.mjs' };
  if (isDev && !existing.length) return { workflow, changes, skipped: '-dev copy without Discord' };

  // NPC portrait generator: keep its image upload, standardise the text.
  const portrait = byName(workflow, 'Post Image to Discord');
  if (portrait) {
    portrait.parameters.content = "={{ (() => { const body = $('Webhook: /webhook/npc-image').item.json.body || {}; const npc = body.npcData || {}; const system = String(body.system || '').trim(); return '**NPC Portrait — ' + (npc.name || 'Unnamed NPC') + (system ? ' (' + system + ')' : '') + '** (image attached)'; })() }}";
    portrait.onError = 'continueRegularOutput';
    changes.push('Post Image to Discord: standard title, continues on error');
  }

  // Prep nodes: always attach a file (no inline code-block variant), fix labels.
  for (const prep of workflow.nodes.filter(node => node.type === 'n8n-nodes-base.code' && /prep for discord/i.test(node.name))) {
    let code = prep.parameters.jsCode || '';
    const inline = /\n?if\s*\(\s*body\.length\s*<=\s*MAX\s*\)\s*return[^\n]*\n/;
    if (inline.test(code)) { code = code.replace(inline, '\n'); changes.push(`${prep.name}: always attaches the file`); }
    const inlineBlock = /if\s*\(\s*body\.length\s*<=\s*MAX\s*\)\s*\{/;
    if (inlineBlock.test(code)) { code = code.replace(inlineBlock, 'if (false) {'); changes.push(`${prep.name}: always attaches the file`); }
    for (const [from, to] of LABEL_FIXES[id] || []) if (code.includes(from)) { code = code.split(from).join(to); changes.push(`${prep.name}: label ${to.trim()}`); }
    prep.parameters.jsCode = code;
    prep.onError = 'continueRegularOutput';
  }

  // "Here is file" + separate text message -> one message with the file.
  const fileNode = existing.find(node => node.parameters?.content === 'Here is file');
  if (fileNode) {
    for (const other of existing.filter(node => node !== fileNode)) {
      const waitPred = Object.entries(workflow.connections).find(([name, value]) => byName(workflow, name)?.type === 'n8n-nodes-base.wait' && JSON.stringify(value).includes(`"node":"${other.name}"`));
      removeNode(workflow, other.name);
      if (waitPred) removeNode(workflow, waitPred[0]);
      changes.push(`removed separate text message ${other.name}${waitPred ? ' and ' + waitPred[0] : ''}`);
    }
    fileNode.parameters.content = '={{ $json.content }}';
    fileNode.parameters.files = { values: [{ inputFieldName: 'file' }] };
    fileNode.onError = 'continueRegularOutput';
    changes.push(`${fileNode.name}: title and file in one message`);
  } else {
    // Text-only "Send to Discord": attach the prepared file.
    for (const node of existing.filter(node => node.name === 'Send to Discord')) {
      node.parameters.files = { values: [{ inputFieldName: 'file' }] };
      node.onError = 'continueRegularOutput';
      changes.push('Send to Discord: attaches the generated file');
    }
  }

  // Generators with no Discord post yet.
  const add = NEW_POSTS[id];
  if (add && !existing.length) {
    const anchor = byName(workflow, add.after);
    if (!anchor) throw new Error(`${workflow.name}: anchor node "${add.after}" not found`);
    const [x, y] = anchor.position || [0, 0];
    workflow.nodes.push(prepNode('Prep for Discord', [x + 220, y + 260], add.code || WIKI_CODE(add.label)));
    workflow.nodes.push(sendNode('Send to Discord', [x + 440, y + 260]));
    addMain(workflow, add.after, 'Prep for Discord');
    addMain(workflow, 'Prep for Discord', 'Send to Discord');
    changes.push(`added Prep for Discord + Send to Discord after "${add.after}"`);
  }

  const after = responsePath(workflow).join('|');
  if (before !== after) throw new Error(`${workflow.name}: response path changed`);
  for (const node of discordNodes(workflow)) if (JSON.stringify(node.credentials) !== JSON.stringify(DISCORD_CREDENTIALS)) throw new Error(`${workflow.name}: unexpected Discord credential on ${node.name}`);
  return { workflow, changes };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const [inDir, outDir] = process.argv.slice(2);
  if (!inDir || !outDir) { console.error('usage: node tools/n8n/discord-post-migration.mjs <inDir> <outDir>'); process.exit(1); }
  fs.mkdirSync(outDir, { recursive: true });
  for (const file of fs.readdirSync(inDir).filter(name => name.endsWith('.json'))) {
    const workflow = JSON.parse(fs.readFileSync(path.join(inDir, file), 'utf8'));
    const { workflow: out, changes, skipped } = migrateDiscord(workflow);
    if (skipped) { console.log(`SKIP ${workflow.id} ${workflow.name}: ${skipped}`); continue; }
    if (out.activeVersion?.nodes) { const v = migrateDiscord({ ...out.activeVersion, id: out.id, name: out.name }); out.activeVersion = { ...out.activeVersion, nodes: v.workflow.nodes, connections: v.workflow.connections }; }
    fs.writeFileSync(path.join(outDir, file), JSON.stringify(out, null, 2));
    console.log(`${changes.length ? 'OK  ' : 'SAME'} ${workflow.id} ${workflow.name}: ${changes.join('; ') || 'no change'}`);
  }
}
