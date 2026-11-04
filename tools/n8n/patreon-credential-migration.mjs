import { readFileSync, writeFileSync, readdirSync, mkdirSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const TOKEN_NAME = ['PATREON', 'CREATOR', 'TOKEN'].join('_');
const REDACTED = ['REDACTED', 'PATREON', 'TOKEN'].join('_');
const OLD_PREFIX = '4Ydo' + 'Fngay8';
export const PATREON_CREDENTIAL = { id: 'Uz6EY2337IetoYeC', name: 'Header Auth account' };

export function patreonMembersNode(name, position, url, credential) {
  return {
    id: `patreon-members-${name.toLowerCase().replace(/[^a-z0-9]+/g, '-')}`,
    name, type: 'n8n-nodes-base.httpRequest', typeVersion: 4.2, position,
    parameters: {
      method: 'GET', url, authentication: 'genericCredentialType', genericAuthType: 'httpHeaderAuth',
      options: {
        pagination: { pagination: {
          paginationMode: 'responseContainsNextURL',
          nextURL: '={{ ($response.body.links || {}).next }}',
          paginationCompleteWhen: 'other',
          completeExpression: '={{ !($response.body.links || {}).next }}',
        } },
        response: { response: { neverError: true } },
      },
    },
    credentials: { httpHeaderAuth: credential }, onError: 'continueRegularOutput', alwaysOutputData: true,
  };
}

function predecessor(workflow, target) {
  const names = Object.entries(workflow.connections).filter(([, outputs]) =>
    Object.values(outputs).some(branches => branches.some(branch => branch.some(edge => edge.node === target)))).map(([name]) => name);
  if (names.length !== 1) throw new Error(`${target}: expected one predecessor, got ${names.length}`);
  return names[0];
}

function rewire(workflow, target, inserted) {
  for (const outputs of Object.values(workflow.connections))
    for (const branches of Object.values(outputs))
      for (const branch of branches)
        for (const edge of branch) if (edge.node === target) edge.node = inserted;
  workflow.connections[inserted] = { main: [[{ node: target, type: 'main', index: 0 }]] };
}

function migrateCode(workflow, node, credential, changes) {
  let code = node.parameters.jsCode;
  const campaign = code.match(/const PATREON_CAMPAIGN_ID\s*=\s*['"](\d+)['"]/);
  const query = code.match(/\/members\?([^'`]+)[\s\S]*?(?:;|`)/);
  if (!campaign || !query) throw new Error(`${node.name}: campaign or member query missing`);
  const url = `https://www.patreon.com/api/oauth2/v2/campaigns/${campaign[1]}/members?${query[1]}`;
  const before = predecessor(workflow, node.name);
  const name = workflow.nodes.some(n => n.name === 'Patreon: Members') ? 'Patreon: Members 2' : 'Patreon: Members';
  const inserted = patreonMembersNode(name, [node.position[0] - 220, node.position[1]], url, credential);
  code = code.replace(new RegExp(`^const ${TOKEN_NAME}\\s*=.*;\\r?\\n`, 'm'), '');
  if (node.name === 'Rate Limit: Prepare Check') {
    const start = code.indexOf('  ', code.indexOf('try{') + 4);
    const end = code.indexOf('  if(member){', start);
    if (end < 0) throw new Error(`${node.name}: member loop missing`);
    const replacement = `  const pages = $('${name}').all().map(i => i.json);\n  if (!pages.length || pages.some(p => !Array.isArray(p.data))) throw new Error('Patreon lookup failed');\n  const tierMap={};\n  let member=null;\n  for(const resp of pages){\n    for(const inc of(resp.included||[]))if(inc.type==='tier')tierMap[inc.id]=inc.attributes.title;\n    if(!member && email)member=resp.data.find(m=>(m.attributes?.email||'').toLowerCase()===email)||null;\n  }\n`;
    code = code.slice(0, start) + replacement + code.slice(end);
    code = `const __in = $('${before}').first().json;\n` + code.replace(/\$json\b/g, '__in');
  } else {
    const start = code.indexOf('  let nextLink = ', code.indexOf('try {'));
    const end = code.indexOf('} catch (e) {}', start);
    if (start < 0 || end < 0) throw new Error(`${node.name}: Patreon pagination block missing`);
    const replacement = `  const pages = $('${name}').all().map(i => i.json);\n  if (!pages.length || pages.some(p => !Array.isArray(p.data))) throw new Error('Patreon lookup failed');\n  for (const resp of pages) {\n    for (const inc of (resp.included||[])) if (inc.type==='tier') tierMap[inc.id] = inc.attributes.title;\n    for (const m of resp.data) { const email = (m.attributes?.email||'').toLowerCase(); if (email) patreonMembers[email] = m; }\n  }\n`;
    code = code.slice(0, start) + replacement + code.slice(end);
    code = `const __in = $('${before}').first().json;\n` + code.replace(/\$json\b/g, '__in');
  }
  node.parameters.jsCode = code;
  workflow.nodes.push(inserted);
  rewire(workflow, node.name, name);
  changes.push(`${node.name}: inserted credentialed ${name}, rewired ${before}, moved member paging from Code`);
  if (node.name === 'Fetch & Build Bonus List') changes.push('N8N_API_KEY in the n8n data-table calls remains hard-coded (outside Patreon scope)');
}

export function migrateWorkflow(input, { credential } = { credential: PATREON_CREDENTIAL }) {
  const workflow = structuredClone(input);
  const changes = [];
  for (const node of [...workflow.nodes]) {
    if (!JSON.stringify(node).includes(REDACTED) && !JSON.stringify(node).includes(TOKEN_NAME) && !JSON.stringify(node).includes(OLD_PREFIX)) continue;
    if (node.type === 'n8n-nodes-base.httpRequest') {
      const params = node.parameters;
      const headers = params.headerParameters?.parameters || [];
      const remaining = headers.filter(h => h.name?.toLowerCase() !== 'authorization');
      if (remaining.length === headers.length) throw new Error(`${node.name}: token is not an Authorization header`);
      params.headerParameters = { ...params.headerParameters, parameters: remaining };
      if (!remaining.length) params.sendHeaders = false;
      params.authentication = 'genericCredentialType';
      params.genericAuthType = 'httpHeaderAuth';
      node.credentials = { httpHeaderAuth: credential };
      changes.push(`${node.name}: replaced literal Authorization header with Header Auth credential`);
    } else if (node.type === 'n8n-nodes-base.code') migrateCode(workflow, node, credential, changes);
    else throw new Error(`${node.name}: unsupported token location`);
  }
  if (workflow.activeVersion?.nodes) {
    const active = migrateWorkflow(workflow.activeVersion, { credential });
    workflow.activeVersion = active.workflow;
    changes.push(...active.changes.map(change => `activeVersion: ${change}`));
  }
  return { workflow, changes };
}

export function assertNoPatreonToken(workflow) {
  const value = JSON.stringify(workflow);
  if (new RegExp(`${REDACTED}|${TOKEN_NAME}|${OLD_PREFIX}|Bearer\\s+[A-Za-z0-9_]`).test(value))
    throw new Error(`${workflow.name}: Patreon token string remains`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const [, , inDir, outDir] = process.argv;
  if (!inDir || !outDir) throw new Error('Usage: node patreon-credential-migration.mjs <inDir> <outDir>');
  mkdirSync(outDir, { recursive: true });
  for (const file of readdirSync(inDir).filter(f => f.endsWith('.json')).sort()) {
    const input = JSON.parse(readFileSync(join(inDir, file), 'utf8'));
    const { workflow, changes } = migrateWorkflow(input, { credential: PATREON_CREDENTIAL });
    if (!changes.length) continue;
    assertNoPatreonToken(workflow);
    writeFileSync(join(outDir, file), JSON.stringify(workflow, null, 2) + '\n');
    console.log(`${file} (${workflow.name}): ${changes.join('; ')}`);
  }
}
