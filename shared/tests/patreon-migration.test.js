import { describe, it, expect } from 'vitest';
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { runInNewContext } from 'node:vm';
import { migrateWorkflow, assertNoPatreonToken, PATREON_CREDENTIAL } from '../../tools/n8n/patreon-credential-migration.mjs';

const fixtureDir = 'shared/tests/fixtures/patreon';
const liveDir = '.migration';
const files = [
  ...readdirSync(fixtureDir).filter(f => f.endsWith('.json')).map(f => join(fixtureDir, f)),
  ...(existsSync(liveDir) ? readdirSync(liveDir).filter(f => f.endsWith('.json')).map(f => join(liveDir, f)) : []),
];
const read = file => JSON.parse(readFileSync(file, 'utf8').replaceAll('FIXTURE_TOKEN_NAME', ['PATREON', 'CREATOR', 'TOKEN'].join('_')).replaceAll('FIXTURE_TOKEN', ['REDACTED', 'PATREON', 'TOKEN'].join('_')));
const patreonNodes = w => w.nodes.filter(n => n.type === 'n8n-nodes-base.httpRequest' && (n.name.startsWith('Patreon: Members') || n.name === 'Creator Member Lookup' || n.name === 'Get Members'));

describe('Patreon credential migration', () => {
  it.each(files)('%s migrates without mutating the source', file => {
    const input = read(file), before = structuredClone(input);
    const { workflow, changes } = migrateWorkflow(input, { credential: PATREON_CREDENTIAL });
    expect(input).toEqual(before);
    expect(changes.length).toBeGreaterThan(0);
    expect(() => assertNoPatreonToken(workflow)).not.toThrow();
    for (const version of [workflow, workflow.activeVersion].filter(Boolean)) {
      for (const n of patreonNodes(version)) {
        expect(n.credentials?.httpHeaderAuth).toEqual(PATREON_CREDENTIAL);
        expect(n.parameters.options?.response?.response?.responseFormat).toBe('json');
        expect(JSON.stringify(n.parameters)).not.toContain('Bearer ');
      }
      for (const code of version.nodes.filter(n => n.name === 'Rate Limit: Prepare Check' || n.name === 'Fetch & Build Bonus List')) {
        const inserted = version.nodes.find(n => n.name === 'Patreon: Members');
        expect(inserted).toBeTruthy();
        const prior = Object.entries(version.connections).filter(([, c]) => c.main?.flat().some(e => e.node === inserted.name)).map(([name]) => name);
        expect(prior).toHaveLength(1);
        expect(version.connections[inserted.name].main[0][0].node).toBe(code.name);
        expect(code.parameters.jsCode).toContain(`$('${prior[0]}').first().json`);
      }
    }
  });

  it('repairs an already migrated workflow without changing existing response settings', () => {
    const input = read('n8n-map-builder-workflow.json');
    const node = patreonNodes(input)[0];
    delete node.parameters.options.response.response.responseFormat;
    node.parameters.options.response.response.fullResponse = false;
    const first = migrateWorkflow(input);
    expect(first.changes).toHaveLength(1);
    expect(patreonNodes(first.workflow)[0].parameters.options.response.response).toEqual({ neverError: true, fullResponse: false, responseFormat: 'json' });
    expect(migrateWorkflow(first.workflow).changes).toHaveLength(0);
  });

  const session = { sessionEmail: 'patron@example.com', sessionKey: 'abc', sessionTier: 'free' };
  const pages = [
    { data: [], included: [] },
    { included: [{ type: 'tier', id: 'premium', attributes: { title: 'Premium' } }], data: [{ attributes: { email: session.sessionEmail, patron_status: 'active_patron' }, relationships: { currently_entitled_tiers: { data: [{ id: 'premium' }] } } }] },
  ];
  const run = async (script, { input = session, results = pages, request } = {}) => {
    const $ = name => ({ first: () => ({ json: input }), all: () => results.map(json => ({ json })) });
    return JSON.parse(JSON.stringify(await runInNewContext(`(async()=>{${script}\n})()`, {
      $json: input, $, $helpers: { httpRequest: request || (async ({ url }) => url.includes('page=2') ? pages[1] : { ...pages[0], links: { next: 'https://example.com/page=2' } }) },
      Date, Error, Array, Object,
    })));
  };

  it('preserves paid tier results across pages', async () => {
    const input = read(join(fixtureDir, 'paged.json'));
    const original = input.nodes.find(n => n.name === 'Rate Limit: Prepare Check').parameters.jsCode;
    const migrated = migrateWorkflow(input, { credential: PATREON_CREDENTIAL }).workflow.nodes.find(n => n.name === 'Rate Limit: Prepare Check').parameters.jsCode;
    expect((await run(original))[0].json).toEqual((await run(migrated))[0].json);
    expect((await run(migrated))[0].json._rateLimitCheck).toMatchObject({ tier: 'premium', monthlyLimit: 200 });
  });

  it('falls back on HTTP errors and malformed responses', async () => {
    const input = read(join(fixtureDir, 'paged.json'));
    const original = input.nodes.find(n => n.name === 'Rate Limit: Prepare Check').parameters.jsCode;
    const migrated = migrateWorkflow(input, { credential: PATREON_CREDENTIAL }).workflow.nodes.find(n => n.name === 'Rate Limit: Prepare Check').parameters.jsCode;
    const fallback = (await run(original, { request: async () => { throw new Error('Unauthorized'); } }))[0].json;
    for (const result of [{ error: { message: 'Unauthorized' } }, { errors: [{ status: '401' }] }])
      expect((await run(migrated, { results: [result] }))[0].json).toEqual(fallback);
  });
});
