import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { runInNewContext } from 'node:vm';
import path from 'node:path';
import { buildMapWorkflow } from '../../tools/n8n/build-map-workflow.mjs';
import { buildPrompt, finalPrompt, CHECK_PROMPT, SCALE_PROMPT, LIGHTS_PROMPT, WALLS_PROMPT, averageScale, sceneScale, toFoundryWalls, toFoundryLights, gridFit, UNIT_SYSTEMS } from '../scripts/core/maps/index.js';

const workflow = buildMapWorkflow();
const node = name => workflow.nodes.find(n => n.name === name);
const runAll = (name, json, records = [], named = {}) => runInNewContext(`(function(){${node(name).parameters.jsCode}\n})()`, {
  $json: json,
  $input: { all: () => records.map(json => ({ json })), first: () => ({ binary: named._binary }) },
  $: key => ({ first: () => ({ json: Array.isArray(named[key]) ? named[key][0] : named[key] }), all: () => (Array.isArray(named[key]) ? named[key] : [named[key]]).map(json => ({ json })) }),
  Date, JSON, Number, String, Object, Array, Math, RegExp, Buffer, DataView,
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
    expect(scripts).toContain(LIGHTS_PROMPT);
    expect(scripts).toContain(WALLS_PROMPT);
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
    expect(run('Map: Prepare Request', { body: { action: 'final', map: sample, mockup: { b64: 'YWJj' }, walls: true } })._map.cost).toBe(7);
    expect(run('Map: Prepare Request', { body: { action: 'final', quality: 'economy', map: sample, mockup: { b64: 'YWJj' }, walls: true } })._map.cost).toBe(5);
    expect(run('Map: Prepare Request', { body: { action: 'final', map: { ...sample, perspective: 'isometric' }, mockup: { b64: 'YWJj' }, walls: true } })._map.cost).toBe(6);
    expect(final._map.quality).toBe('standard');
    expect(final._map.mockup.mediaType).toBe('image/png');
    expect(run('Map: Prepare Request', { body: { action: 'final', map: sample, mockup: { b64: 'YWJj' }, quality: 'invalid' } }).error).toBe('BAD_REQUEST');
    expect(run('Map: Prepare Request', { body: { action: 'other', map: sample } }).error).toBe('BAD_REQUEST');
  });

  it('uses the selected final quality for cost, image resolution, retry, and scale calls', () => {
    for (const [quality, cost, model, resolution, scaleCalls] of [
      ['standard', 6, 'google/gemini-3.1-flash-image', '2K', 2],
      ['economy', 4, 'google/gemini-3.1-flash-image', '1K', 1],
      ['draft', 2, 'black-forest-labs/flux.2-klein-4b', '2K', 1],
      ['high', 6, 'bytedance-seed/seedream-5-0-pro', '2K', 1],
    ]) {
      const prepared = run('Map: Prepare Request', { body: { action: 'final', quality, map: sample, mockup: { b64: 'YWJj' } } });
      expect(prepared._map.quality).toBe(quality);
      expect(prepared._map.cost).toBe(cost);
      const named = { 'Rate Limit: Validate': prepared, 'Final: Judge': {} };
      for (const request of ['Final: Build Request', 'Final: Build Retry']) {
        const body = run(request, {}, [], named).body;
        expect(body).toMatchObject({ model, resolution, prompt: finalPrompt(sample), aspect_ratio: prepared._map.aspect });
        expect(body.input_references).toEqual([{ type: 'image_url', image_url: { url: 'data:image/png;base64,YWJj' } }]);
      }
      const chosen = { image: { b64: 'YWJj', mediaType: 'image/png' } };
      expect(runAll('Scale: Build Requests', chosen, [], named)).toHaveLength(scaleCalls);
      const limited = run('Rate Limit: Validate', {}, [], {
        'Rate Limit: Prepare Check': { ...prepared, _rateLimitCheck: { monthlyLimit: 0, tier: 'free' } },
      });
      expect(limited.message).toContain(`A map final (${quality}) costs ${cost} uses.`);
    }
  });

  it('retries a resolution 400 once without resolution for either final attempt', () => {
    const prepared = run('Map: Prepare Request', { body: { action: 'final', quality: 'high', map: sample, mockup: { b64: 'YWJj' } } });
    for (const [request, check, fallback] of [
      ['Final: Build Request', 'Final: Check Resolution Error', 'Final: Build Without Resolution'],
      ['Final: Build Retry', 'Final: Check Retry Resolution Error', 'Final: Build Retry Without Resolution'],
    ]) {
      const named = { 'Rate Limit: Validate': prepared, 'Final: Judge': {} };
      const original = run(request, {}, [], named);
      const checked = run(check, { statusCode: 400, error: { message: 'Unsupported resolution 2K' } }, [], { [request]: original });
      expect(checked._resolutionFallback).toBe(true);
      expect(run(fallback, {}, [], { [request]: original }).body).toEqual(Object.fromEntries(Object.entries(original.body).filter(([key]) => key !== 'resolution')));
      expect(run(check, { statusCode: 400, error: { message: 'Invalid aspect ratio' } }, [], { [request]: original })._resolutionFallback).toBe(false);
    }
    expect(node('OpenRouter: Final Image (no resolution)').parameters.options.timeout).toBe(240000);
    expect(node('OpenRouter: Final Image (retry, no resolution)').parameters.options.timeout).toBe(240000);
  });

  it('combines two scale estimates using the shared math', () => {
    const a = { width_m: 36, height_m: 24, references: [] };
    const b = { width_m: 42, height_m: 30, references: [] };
    const responses = [a, b].map(v => ({ choices: [{ message: { content: JSON.stringify(v) } }], usage: { cost: 0.01 } }));
    const prep = { _map: { units: 'metric', square: 2, map: sample, startedAt: Date.now() } };
    const chosen = { image: { b64: 'YWJj', mediaType: 'image/png' }, defects: { retry: false, reasons: [] }, retried: false, costUsd: 0.1 };
    const result = run('Scale: Combine', {}, responses, { 'Rate Limit: Validate': prep, 'Scale: Build Requests': [{ chosen, purpose: 'scale' }, { chosen, purpose: 'scale' }] });
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
    const result = run('Scale: Combine', {}, [response], { 'Rate Limit: Validate': prep, 'Scale: Build Requests': [{ chosen, purpose: 'scale' }] });
    const avg = averageScale(estimate, undefined);
    expect(result.quality).toBe('economy');
    expect(result.measured.width_m).toBe(avg.width_m);
    expect(result.measured.height_m).toBe(avg.height_m);
    expect(result.scene).toEqual(sceneScale(avg, { system: 'metric', square: 2 }));
  });

  it('sorts scale, walls, and lights, and drops the surcharge if both traces fail', () => {
    const prep = { _map: { quality: 'standard', cost: 7, walls: true, scaleCalls: 1, units: 'imperial', map: sample, startedAt: Date.now() } };
    const chosen = { image: { b64: 'YWJj', mediaType: 'image/png' }, costUsd: 0.1 };
    const requests = runAll('Scale: Build Requests', chosen, [], { 'Rate Limit: Validate': prep });
    expect(requests.map(r => r.purpose)).toEqual(['scale', 'lights', 'walls']);
    const response = value => ({ choices: [{ message: { content: '```json\n' + JSON.stringify(value) + '\n```' } }], usage: { cost: 0.01 } });
    const records = [response({ width_m: 30, height_m: 20 }), response({ lights: [{ kind: 'torch', x: 500, y: 500 }, null, 'bad' ] }), response({ walls: [{ kind: 'door', x1: 0, y1: 0, x2: 100, y2: 0 }, null] })];
    const named = { 'Rate Limit: Validate': prep, 'Scale: Build Requests': requests };
    const result = run('Scale: Combine', {}, records, named);
    expect(result.trace.lights).toHaveLength(1);
    expect(result.trace.walls).toHaveLength(1);
    expect(result.chargedCost).toBe(7);
    expect(result.costUsd).toBeCloseTo(0.13);
    const limit = run('Rate Limit: Validate', {}, [], { 'Rate Limit: Prepare Check': { ...prep, _rateLimitCheck: { monthlyLimit: 6, tier: 'free' } } });
    expect(limit.message).toContain('costs 7 uses');
    expect(run('Map: Build Response', {}, [], { 'Scale: Combine': result, 'Rate Limit: Validate': prep }).usesCharged).toBe(7);
    const failed = run('Scale: Combine', {}, [records[0], {}, {}], named);
    expect(failed.chargedCost).toBe(6);
    expect(failed.warning).toContain('Walls could not be traced.');
    expect(failed.warning).toContain('Lights could not be placed.');
    expect(run('Map: Build Response', {}, [], { 'Scale: Combine': failed, 'Rate Limit: Validate': prep }).usesCharged).toBe(6);
    expect(run('Increment Usage Counter', {}, [], { 'Rate Limit: Validate': { ...prep, _rateLimitCheck: { currentCount: 3, sessionKey: 'x', currentMonth: '2026-11' } }, 'Scale: Combine': failed }).npc_count).toBe(9);
  });

  it('bakes the scene and Foundry documents using shared helpers', () => {
    const png = Buffer.alloc(32); png.set([0x89, 0x50, 0x4e, 0x47]); png.writeUInt32BE(1200, 16); png.writeUInt32BE(900, 20);
    const b64 = png.toString('base64');
    const trace = { walls: [{ kind: 'door', x1: 0, y1: 0, x2: 200, y2: 0 }, { kind: 'window', x1: 300, y1: 0, x2: 500, y2: 0 }], lights: [{ kind: 'torch', x: 500, y: 500 }] };
    for (const units of ['imperial', 'metric']) {
      const prepared = run('Map: Prepare Request', { body: { action: 'bake', image: { b64, mediaType: 'image/png', width: 1, height: 1 }, squares: [20, 15], units, trace, mood: 'night', name: 'Camp' } });
      expect(prepared._map.cost).toBe(0);
      expect(prepared._map.image.width).toBe(1200);
      expect(prepared._map.image.height).toBe(900);
      const baked = run('Bake: Prepare', prepared);
      const distance = UNIT_SYSTEMS[units].square;
      expect(baked.scene).toEqual({ name: 'Camp', width: gridFit(1, 1, [20, 15]).width, height: 1500, padding: 0, grid: { type: 1, size: 100, distance, units: UNIT_SYSTEMS[units].units }, environment: { darknessLevel: 0.7 } });
      expect(baked.walls).toEqual(toFoundryWalls(trace.walls, { squares: [20, 15], gridPx: 100 }).map(({ kind, ...wall }) => wall));
      expect(baked.lights).toEqual(toFoundryLights(trace.lights, { squares: [20, 15], gridPx: 100, feetPerSquare: distance }).map(({ kind, ...light }) => light));
      expect(baked.counts).toEqual({ walls: 0, doors: 1, windows: 1, lights: 1 });
      expect(baked.sharpen).toBe(0.6);
    }
    const resize = node('Bake: Resize Image');
    expect(resize.type).toBe('n8n-nodes-base.editImage');
    expect(resize.parameters.operation).toBe('resize');
    expect(resize.parameters.resizeOption).toBe('ignoreAspectRatio');
    expect(workflow.connections['Bake: Prepare'].main[0][0].node).toBe('Bake: Resize Image');
    const seen = new Set();
    const visit = at => { if (seen.has(at)) return; seen.add(at); for (const output of workflow.connections[at]?.main || []) for (const edge of output) visit(edge.node); };
    visit('IF: Bake Action?');
    expect(workflow.connections['IF: Bake Action?'].main[0][0].node).toBe('Bake: Prepare');
    const bakeOnly = new Set();
    const walkBake = at => { if (bakeOnly.has(at)) return; bakeOnly.add(at); for (const output of workflow.connections[at]?.main || []) for (const edge of output) walkBake(edge.node); };
    walkBake('Bake: Prepare');
    expect(bakeOnly.has('Increment Usage Counter')).toBe(false);
  });

  it('returns the resized binary image and leaves day scenes bright', async () => {
    expect(run('Map: Prepare Request', { body: { action: 'bake', image: { b64: 'YWJj', mediaType: 'image/png' }, squares: [8, 6], units: 'imperial' } }).message).toBe('Could not read the map image size.');
    const png = Buffer.alloc(32); png.set([0x89, 0x50, 0x4e, 0x47]); png.writeUInt32BE(1200, 16); png.writeUInt32BE(900, 20);
    const prepared = run('Map: Prepare Request', { body: { action: 'bake', image: { b64: png.toString('base64'), mediaType: 'image/png' }, squares: [8, 6], units: 'imperial', mood: 'day' } });
    const baked = run('Bake: Prepare', prepared);
    expect(baked.scene).not.toHaveProperty('environment');
    const output = await runInNewContext(`(async function(){${node('Bake: Respond').parameters.jsCode}\n}).call({helpers:{getBinaryDataBuffer:async()=>Buffer.from('resized')}})`, {
      $input: { first: () => ({ binary: { map: { mimeType: 'image/png' } } }) },
      $: () => ({ first: () => ({ json: baked }) }),
      Buffer,
    });
    expect(output[0].json.image).toEqual({ b64: Buffer.from('resized').toString('base64'), mediaType: 'image/png' });
    expect(output[0].json.sharpen).toBe(0);
    expect(output[0].json).not.toHaveProperty('look');
    expect(output[0].json.usesCharged).toBe(0);
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
