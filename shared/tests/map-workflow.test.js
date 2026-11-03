import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { runInNewContext } from 'node:vm';
import path from 'node:path';
import { buildMapWorkflow } from '../../tools/n8n/build-map-workflow.mjs';
import { buildPrompt, finalPrompt, CHECK_PROMPT, SCALE_PROMPT, averageScale, sceneScale } from '../scripts/core/maps/index.js';

const workflow = buildMapWorkflow();
const node = name => workflow.nodes.find(n => n.name === name);
const runAll = (name, json, records = [], named = {}) => runInNewContext(`(function(){${node(name).parameters.jsCode}\n})()`, {
  $json: json,
  $input: { all: () => records.map(json => ({ json })) },
  $: key => ({ first: () => ({ json: named[key] }) }),
  Date, JSON, Number, String, Object, Array, Math, RegExp,
}).map(item => item.json);
const run = (name, json, records = [], named = {}) => runAll(name, json, records, named)[0];
const sample = {
  description: 'A torchlit tavern with a river outside', perspective: 'top-down',
  genre: 'fantasy', style: 'painted', mood: 'night', squares: [24, 18],
  dials: { detail: 4 },
};

describe('generated map workflow', () => {
  it('matches its committed artifact', () => {
    expect(JSON.parse(readFileSync(path.resolve('n8n-map-builder-workflow.json'), 'utf8'))).toEqual(workflow);
  });

  it('embeds the shared prompt text', () => {
    const scripts = workflow.nodes.filter(n => n.type === 'n8n-nodes-base.code').map(n => n.parameters.jsCode).join('\n');
    expect(scripts).toContain(CHECK_PROMPT);
    expect(scripts).toContain(SCALE_PROMPT);
    expect(scripts).toContain('The attached image is a rough layout mockup. Redraw it as a finished battle map.');
  });

  it('prepares both actions from the shared prompt functions', () => {
    const mockups = run('Map: Prepare Request', { body: { action: 'mockups', map: sample } });
    expect(mockups._map.prompt).toBe('Quick rough mockup for layout only. ' + buildPrompt(sample));
    expect(mockups._map.cost).toBe(3);
    expect(runAll('Mockups: Build Requests', {}, [], { 'Rate Limit: Validate': mockups })).toHaveLength(4);
    const final = run('Map: Prepare Request', { body: { action: 'final', map: sample, mockup: { b64: 'YWJj' } } });
    expect(final._map.prompt).toBe(finalPrompt(sample));
    expect(final._map.cost).toBe(6);
    expect(final._map.quality).toBe('standard');
    expect(final._map.mockup.mediaType).toBe('image/png');
    expect(run('Map: Prepare Request', { body: { action: 'final', map: sample, mockup: { b64: 'YWJj' }, quality: 'invalid' } }).error).toBe('BAD_REQUEST');
    expect(run('Map: Prepare Request', { body: { action: 'other', map: sample } }).error).toBe('BAD_REQUEST');
  });

  it('uses the selected final quality for cost, image resolution, retry, and scale calls', () => {
    for (const [quality, cost, resolution, scaleCalls] of [
      ['standard', 6, '2K', 2],
      ['economy', 4, '1K', 1],
    ]) {
      const prepared = run('Map: Prepare Request', { body: { action: 'final', quality, map: sample, mockup: { b64: 'YWJj' } } });
      expect(prepared._map.quality).toBe(quality);
      expect(prepared._map.cost).toBe(cost);
      const named = { 'Rate Limit: Validate': prepared, 'Final: Judge': {} };
      expect(run('Final: Build Request', {}, [], named).body.resolution).toBe(resolution);
      expect(run('Final: Build Retry', {}, [], named).body.resolution).toBe(resolution);
      const chosen = { image: { b64: 'YWJj', mediaType: 'image/png' } };
      expect(runAll('Scale: Build Requests', chosen, [], named)).toHaveLength(scaleCalls);
      const limited = run('Rate Limit: Validate', {}, [], {
        'Rate Limit: Prepare Check': { ...prepared, _rateLimitCheck: { monthlyLimit: 0, tier: 'free' } },
      });
      expect(limited.message).toContain(`A map final (${quality}) costs ${cost} uses.`);
    }
  });

  it('combines two scale estimates using the shared math', () => {
    const a = { width_m: 36, height_m: 24, references: [] };
    const b = { width_m: 42, height_m: 30, references: [] };
    const responses = [a, b].map(v => ({ choices: [{ message: { content: JSON.stringify(v) } }], usage: { cost: 0.01 } }));
    const prep = { _map: { units: 'metric', square: 2, map: sample, startedAt: Date.now() } };
    const chosen = { image: { b64: 'YWJj', mediaType: 'image/png' }, defects: { retry: false, reasons: [] }, retried: false, costUsd: 0.1 };
    const result = run('Scale: Combine', {}, responses, { 'Rate Limit: Validate': prep, 'Scale: Build Requests': { chosen } });
    const avg = averageScale(a, b);
    expect(result.measured.width_m).toBe(avg.width_m);
    expect(result.measured.height_m).toBe(avg.height_m);
    expect(result.scene).toEqual(sceneScale(avg, { system: 'metric', square: 2 }));
    expect(result.costUsd).toBeCloseTo(0.12);
  });

  it('uses a single economy scale estimate', () => {
    const estimate = { width_m: 36, height_m: 24, references: [] };
    const response = { choices: [{ message: { content: JSON.stringify(estimate) } }], usage: { cost: 0.01 } };
    const prep = { _map: { quality: 'economy', units: 'metric', square: 2, map: sample, startedAt: Date.now() } };
    const chosen = { image: { b64: 'YWJj', mediaType: 'image/png' }, defects: { retry: false, reasons: [] }, retried: false, costUsd: 0.1 };
    const result = run('Scale: Combine', {}, [response], { 'Rate Limit: Validate': prep, 'Scale: Build Requests': { chosen } });
    const avg = averageScale(estimate, undefined);
    expect(result.quality).toBe('economy');
    expect(result.measured.width_m).toBe(avg.width_m);
    expect(result.measured.height_m).toBe(avg.height_m);
    expect(result.scene).toEqual(sceneScale(avg, { system: 'metric', square: 2 }));
  });

  it('uses the built-in OpenRouter credential on every OpenRouter node and nowhere else', () => {
    const serialized = JSON.stringify(workflow);
    expect(serialized).not.toMatch(/sk-or-|Bearer\s+[A-Za-z0-9]|PATREON_CREATOR_TOKEN|4YdoFngay8/);
    const openRouterNodes = workflow.nodes.filter(n => n.parameters?.url?.includes('openrouter.ai'));
    expect(openRouterNodes.length).toBeGreaterThan(0);
    for (const n of openRouterNodes) {
      expect(n.credentials).toEqual({ openRouterApi: { id: 'gYYHf4NcuzeBjKSk', name: 'OpenRouter account' } });
      expect(n.type).toBe('n8n-nodes-base.httpRequest');
      expect(n.parameters.authentication).toBe('predefinedCredentialType');
      expect(n.parameters.nodeCredentialType).toBe('openRouterApi');
      expect(n.parameters).not.toHaveProperty('genericAuthType');
      expect(n.parameters.url).toMatch(/^https:\/\/openrouter\.ai\/api\/v1\//);
    }
    for (const n of workflow.nodes.filter(n => !openRouterNodes.includes(n))) {
      expect(n).not.toHaveProperty('credentials');
    }
  });

  it('keeps the store tier limits and charges only successful branches', () => {
    const store = JSON.parse(readFileSync(path.resolve('n8n-dnd5e-store-builder-workflow.json'), 'utf8'));
    const storeScript = store.nodes.find(n => n.name === 'Rate Limit: Prepare Check').parameters.jsCode;
    const limits = script => script.match(/const TIER_LIMITS=(\{[^;]+\});/)[1];
    expect(limits(node('Rate Limit: Prepare Check').parameters.jsCode)).toBe(limits(storeScript));
    const targets = (name, index) => workflow.connections[name]?.main?.[index]?.map(e => e.node) || [];
    expect(targets('IF: Mockups OK?', 0)).toEqual(['Increment Usage Counter']);
    expect(targets('IF: Mockups OK?', 1)).toEqual(['Respond: 502 Generation Failed']);
    expect(targets('IF: Scale OK?', 0)).toEqual(['Increment Usage Counter']);
    expect(targets('IF: Scale OK?', 1)).toEqual(['Respond: 502 Generation Failed']);
    const parents = Object.entries(workflow.connections).flatMap(([from, c]) => c.main.flat().filter(e => e.node === 'Increment Usage Counter').map(() => from));
    expect(parents.sort()).toEqual(['IF: Mockups OK?', 'IF: Scale OK?']);
    for (const name of workflow.nodes.filter(n => n.name.startsWith('OpenRouter:')).map(n => n.name)) {
      const seen = new Set();
      const walk = (at, failed = false) => {
        if (at === 'Increment Usage Counter') expect(failed).toBe(false);
        if (seen.has(at + failed)) return;
        seen.add(at + failed);
        for (const output of workflow.connections[at]?.main || []) for (const edge of output) walk(edge.node, failed || edge.node.startsWith('Respond:'));
      };
      walk(name);
    }
  });
});
