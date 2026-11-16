import { describe, it, expect } from 'vitest';
import vm from 'node:vm';
import { migrateDiscord, responsePath } from '../../tools/n8n/discord-post-migration.mjs';
import { buildMapWorkflow } from '../../tools/n8n/build-map-workflow.mjs';

const DISCORD = { discordBotApi: { id: 'v845fBzdtlCbenJa', name: 'Discord Bot account' } };
const webhook = { name: 'Webhook', type: 'n8n-nodes-base.webhook', parameters: { httpMethod: 'POST' }, position: [0, 0] };
const respond = { name: 'Respond to Webhook', type: 'n8n-nodes-base.respondToWebhook', parameters: {}, position: [600, 0] };
const code = (name, jsCode = 'return [$input.first()];') => ({ name, type: 'n8n-nodes-base.code', parameters: { jsCode }, position: [300, 0] });
const discord = (name, content, files) => ({ name, type: 'n8n-nodes-base.discord', typeVersion: 2, credentials: DISCORD, parameters: { resource: 'message', content, options: {}, ...(files ? { files } : {}) }, position: [900, 0] });
const link = (...names) => Object.fromEntries(names.slice(0, -1).map((n, i) => [n, { main: [[{ node: names[i + 1], type: 'main', index: 0 }]] }]));

describe('Discord post migration', () => {
  it('attaches the prepared file to a text-only Send to Discord', () => {
    const wf = { id: 'x', name: 'PF2e Store Builder', nodes: [webhook, code('Build Response'), respond, code('Prep for Discord'), discord('Send to Discord', '={{ $json.content }}')],
      connections: { ...link('Webhook', 'Build Response', 'Respond to Webhook'), 'Build Response': { main: [[{ node: 'Respond to Webhook', type: 'main', index: 0 }, { node: 'Prep for Discord', type: 'main', index: 0 }]] }, ...link('Prep for Discord', 'Send to Discord') } };
    const { workflow } = migrateDiscord(wf);
    const send = workflow.nodes.find(n => n.name === 'Send to Discord');
    expect(send.parameters.files).toEqual({ values: [{ inputFieldName: 'file' }] });
    expect(send.onError).toBe('continueRegularOutput');
    expect(responsePath(workflow)).toEqual(responsePath(wf));
  });

  it('collapses "Here is file" plus a separate text message into one message', () => {
    const prep = code('prep for discord', "const body='x';const MAX=1900;\nif(body.length<=MAX)return [{json:{content:'inline'}}];\nreturn [{json:{content:'**X — Y** (JSON attached)'}}];");
    const wf = { id: 'y', name: 'D&D 5e NPC Builder', nodes: [webhook, respond, prep, { name: 'If4', type: 'n8n-nodes-base.if', parameters: {}, position: [0, 0] }, { name: 'Wait', type: 'n8n-nodes-base.wait', parameters: {}, position: [0, 0] }, discord('Send a message', '={{ $json.content }}'), discord('Send a message1', 'Here is file', { values: [{ inputFieldName: '=file' }] })],
      connections: { ...link('Webhook', 'Respond to Webhook'), ...link('prep for discord', 'If4'), If4: { main: [[{ node: 'Send a message1', type: 'main', index: 0 }, { node: 'Wait', type: 'main', index: 0 }]] }, ...link('Wait', 'Send a message') } };
    const { workflow } = migrateDiscord(wf);
    const sends = workflow.nodes.filter(n => n.type === 'n8n-nodes-base.discord');
    expect(sends.map(n => n.name)).toEqual(['Send a message1']);
    expect(sends[0].parameters.content).toBe('={{ $json.content }}');
    expect(sends[0].parameters.files.values[0].inputFieldName).toBe('file');
    expect(workflow.nodes.some(n => n.name === 'Wait')).toBe(false);
    expect(workflow.nodes.find(n => n.name === 'prep for discord').parameters.jsCode).not.toMatch(/body\.length<=MAX/);
  });

  it('adds a parallel, non-blocking post to a generator that has none', () => {
    const wf = { id: 'KUcKqh0f3k8shbJV', name: 'PF2e Wiki Item Converter', nodes: [webhook, code('Parse & Assemble Item'), respond], connections: link('Webhook', 'Parse & Assemble Item', 'Respond to Webhook') };
    const { workflow } = migrateDiscord(wf);
    expect(workflow.connections['Parse & Assemble Item'].main[0].map(l => l.node)).toEqual(['Respond to Webhook', 'Prep for Discord']);
    const send = workflow.nodes.find(n => n.name === 'Send to Discord');
    expect(send.onError).toBe('continueRegularOutput');
    expect(send.credentials).toEqual(DISCORD);
    expect(responsePath(workflow)).toEqual(responsePath(wf));
    const prep = workflow.nodes.find(n => n.name === 'Prep for Discord');
    const out = vm.runInNewContext(`(()=>{${prep.parameters.jsCode}})()`, { $json: { ok: true, item: { name: 'Flaming Sword', type: 'weapon' } }, Buffer });
    expect(out[0].json.content).toBe('**PF2e Wiki Item — Flaming Sword (weapon)** (JSON attached)');
    expect(out[0].binary.file.fileName).toBe('flaming-sword.json');
  });

  it('leaves -dev copies without Discord alone', () => {
    const wf = { id: 'z', name: 'Hero System 6e HDC Builder -(DEV)-', nodes: [webhook, respond], connections: link('Webhook', 'Respond to Webhook') };
    expect(migrateDiscord(wf).skipped).toMatch(/-dev/);
  });
});

describe('Map Builder Discord post', () => {
  const workflow = buildMapWorkflow();
  const prep = workflow.nodes.find(n => n.name === 'Discord: Prep Map Post');
  const run = (json, map) => vm.runInNewContext(`(()=>{${prep.parameters.jsCode}})()`, { $json: json, $: () => ({ first: () => ({ json: { _map: map } }) }), Buffer });

  it('runs in parallel with the response and continues on error', () => {
    expect(workflow.connections['Map: Build Response'].main[0].map(l => l.node)).toEqual(['Respond to Webhook', 'Discord: Prep Map Post']);
    expect(prep.onError).toBe('continueRegularOutput');
    expect(workflow.nodes.find(n => n.name === 'Discord: Send Map Post').onError).toBe('continueRegularOutput');
  });

  it('posts finals and edits with an image, nothing else', () => {
    const final = run({ ok: true, quality: 'standard', image: { b64: 'AAAA', mediaType: 'image/png' } }, { action: 'final', map: { description: 'A cave with an obelisk' } });
    expect(final[0].json.content).toBe('**AI Battle Map — A cave with an obelisk (Standard)** (image attached)');
    expect(final[0].json.content).not.toMatch(/\d+x\d+/);
    expect(final[0].binary.file.mimeType).toBe('image/png');
    const edit = run({ ok: true, patch: { b64: 'AAAA', mediaType: 'image/png' } }, { action: 'edit', change: 'Add a campfire' });
    expect(edit[0].json.content).toBe('**AI Battle Map edit — Add a campfire** (image attached)');
    expect(run({ ok: true, images: [] }, { action: 'mockups' })).toEqual([]);
  });
});
