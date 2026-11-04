import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { runInNewContext } from 'node:vm';
import path from 'node:path';
import { buildMapWorkflow } from '../../tools/n8n/build-map-workflow.mjs';
import { buildPrompt, finalPrompt, PLAN_PROMPT, CHECK_PROMPT, SCALE_PROMPT, LIGHTS_PROMPT, WALLS_PROMPT, averageScale, sceneScale, toFoundryWalls, toFoundryLights, gridFit, UNIT_SYSTEMS, cropBox } from '../scripts/core/maps/index.js';

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
  it('prepares multi-megabyte bake and edit images without overflowing the stack', () => {
    const png = Buffer.alloc(24);
    png.set([0x89, 0x50, 0x4e, 0x47, 13, 10, 26, 10]);
    png.writeUInt32BE(2400, 16);
    png.writeUInt32BE(1792, 20);
    for (const [action, bytes] of [['bake', 6 * 1024 * 1024], ['edit', 5 * 1024 * 1024]]) {
      const image = { b64: Buffer.concat([png, Buffer.alloc(bytes)]).toString('base64'), mediaType: 'image/png' };
      const body = action === 'bake'
        ? { action, image, squares: [24, 18], units: 'imperial' }
        : { action, image, box: [0.1, 0.1, 0.3, 0.3], change: 'Add a doorway' };
      const result = run('Map: Prepare Request', { body });
      expect(result.error).toBeUndefined();
      expect(result._map).toMatchObject({ action, image: { width: 2400, height: 1792, mediaType: 'image/png' } });
      expect(result._map.image.b64).toBe(image.b64);
    }
  });
  it('looks up live Patreon tiers with a credential and preserves fallback tiers', () => {
    const patreon = node('Patreon: Members');
    expect(patreon.credentials.httpHeaderAuth).toEqual({ id: 'Uz6EY2337IetoYeC', name: 'Header Auth account' });
    expect(patreon.parameters.options.pagination.pagination.paginationMode).toBe('responseContainsNextURL');
    expect(workflow.connections['IF: Plan Action?'].main[1][0].node).toBe('Patreon: Members');
    expect(workflow.connections['Patreon: Members'].main[0][0].node).toBe('Rate Limit: Prepare Check');
    const input = { sessionEmail: 'patron@example.com', sessionKey: 'abc', sessionTier: 'free' };
    const member = { attributes: { email: input.sessionEmail, patron_status: 'active_patron' }, relationships: { currently_entitled_tiers: { data: [{ id: 'paid' }] } } };
    const named = { 'Map: Prepare Request': input, 'Patreon: Members': [{ data: [member], included: [{ type: 'tier', id: 'paid', attributes: { title: 'Premium' } }] }] };
    expect(run('Rate Limit: Prepare Check', {}, [], named)).toMatchObject({ sessionTier: 'premium', _rateLimitCheck: { tier: 'premium', monthlyLimit: 200 } });
    expect(run('Rate Limit: Prepare Check', {}, [], { ...named, 'Patreon: Members': [{ errors: [{ status: 401 }] }] })).toMatchObject({ sessionTier: 'free', _rateLimitCheck: { monthlyLimit: 3 } });
    expect(run('Rate Limit: Prepare Check', {}, [], { ...named, 'Map: Prepare Request': { ...input, sessionIsSuperuser: true } })).toMatchObject({ sessionTier: 'superuser', _rateLimitCheck: { monthlyLimit: 999999 } });
  });
  it('allows five minutes for an execution', () => {
    expect(workflow.settings).toEqual({ executionOrder: 'v1', executionTimeout: 300 });
  });

  it('matches its committed artifact', () => {
    expect(JSON.parse(readFileSync(path.resolve('n8n-map-builder-workflow.json'), 'utf8'))).toEqual(workflow);
  });

  it('embeds the shared prompt text', () => {
    const scripts = workflow.nodes.filter(n => n.type === 'n8n-nodes-base.code').map(n => n.parameters.jsCode).join('\n');
    expect(scripts).toContain(CHECK_PROMPT);
    expect(scripts).toContain(PLAN_PROMPT.split(' Return JSON only')[0]);
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

  it('builds a free plan request and parses a fenced response', () => {
    const prepared = run('Map: Prepare Request', { body: { action: 'plan', map: sample } });
    expect(prepared._map).toMatchObject({ action: 'plan', cost: 0 });
    const request = run('Plan: Build Request', prepared);
    expect(request.body).toMatchObject({ model: 'google/gemini-3.7-flash', response_format: { type: 'json_object' } });
    expect(request.body.messages[0].content).toContain(PLAN_PROMPT);
    expect(request.body.messages[0].content).toContain('including the main or largest space');
    expect(request.body.messages[0].content).toContain(sample.description);
    const plan = { areas: [{ what: 'broken portcullis', where: 'south', size: 'about 4 by 3 squares' }, { what: 'torchlit tables', where: 'centre', size: 'about 4 by 3 squares' }, { what: 'river stones', where: 'north', size: 'about 4 by 3 squares' }, { what: 'bad', where: 'above', size: 'large' }], connections: ['a passage between the tables and river'], notes: 'A ruined keep.' };
    const parsed = run('Plan: Parse', { choices: [{ message: { content: '```json\n' + JSON.stringify(plan) + '\n```' } }], usage: { cost: 0.001 } }, [], { 'Plan: Build Request': request });
    expect(parsed.plan.areas).toHaveLength(4);
    expect(parsed.plan.rawAreas).toBe(4);
    expect(parsed.plan.areas[3].where).toBe('above');
    expect(parsed.plan).not.toHaveProperty('warning');
    expect(parsed.layout).toContain('broken portcullis');
    expect(parsed.layout).toContain(sample.description);
    expect(run('Plan: Respond', parsed)).toMatchObject({ ok: true, usesCharged: 0, costUsd: 0.001 });
    const thin = run('Plan: Parse', { choices: [{ message: { content: JSON.stringify({ areas: [plan.areas[0]], connections: [] }) } }] }, [], { 'Plan: Build Request': request });
    expect(run('Plan: Respond', thin)).toMatchObject({ rawAreas: 1, warning: 'The layout plan was thin; check it before generating.' });
    expect(run('Plan: Parse', {}, [], { 'Plan: Build Request': request })).toMatchObject({ ok: false, error: 'PLAN_FAILED', message: 'The layout plan could not be made.', usesCharged: 0 });
    expect(workflow.connections['IF: Plan OK?'].main[1][0].node).toBe('Respond: 502 Plan Failed');
    expect(workflow.connections['IF: Plan Action?'].main[0][0].node).toBe('Plan: Build Request');
    const reached = new Set();
    const walk = at => { if (reached.has(at)) return; reached.add(at); for (const branch of workflow.connections[at]?.main || []) for (const edge of branch) walk(edge.node); };
    walk('Plan: Respond');
    expect(reached.has('Increment Usage Counter')).toBe(false);
  });

  it('keeps areas with varied positions and without sizes', () => {
    const prepared = run('Plan: Build Request', run('Map: Prepare Request', { body: { action: 'plan', map: sample } }));
    const positions = ['south-centre', 'south and centre', 'front', 'rear', 'left', 'right', 'whole south half', 'upper alcove', 'back', 'center'];
    const parsed = run('Plan: Parse', { choices: [{ message: { content: JSON.stringify({ areas: positions.map((where, i) => ({ what: `room ${i}`, where })) }) } }] }, [], { 'Plan: Build Request': prepared });
    expect(parsed.plan.rawAreas).toBe(10);
    expect(parsed.plan.areas.map(a => a.where)).toEqual(['south', 'south', 'south', 'north', 'west', 'east', 'south', 'upper alcove']);
    expect(parsed.layout).not.toContain('(undefined)');
  });

  it('runs one plan call before auto planned mockups and uses its layout', () => {
    const prepared = run('Map: Prepare Request', { body: { action: 'mockups', map: sample, autoPlan: true, count: 2 } });
    expect(prepared._map.autoPlan).toBe(true);
    expect(workflow.connections['IF: Auto Plan?'].main[0]).toHaveLength(1);
    expect(workflow.connections['IF: Auto Plan?'].main[0][0].node).toBe('Plan: Build Request');
    const request = run('Plan: Build Request', prepared);
    const plan = { areas: ['centre', 'north'].map(where => ({ what: where === 'centre' ? 'fire in the centre' : 'stone floor', where, size: 'about 3 by 3 squares' })), connections: [], notes: '' };
    const parsed = run('Plan: Parse', { choices: [{ message: { content: JSON.stringify(plan) } }] }, [], { 'Plan: Build Request': request });
    const mockups = runAll('Mockups: Build Requests', parsed, [], { 'Rate Limit: Validate': prepared });
    expect(mockups).toHaveLength(2);
    expect(mockups[0].body.prompt).toContain(parsed.layout);
    expect(workflow.connections['IF: Standalone Plan?'].main[1][0].node).toBe('Mockups: Build Requests');
    const collected = run('Mockups: Collect', {}, [{ data: [{ b64_json: 'YWJj' }] }], { 'Rate Limit: Validate': prepared, 'Plan: Parse': parsed });
    expect(collected.layout).toBe(parsed.layout);
  });

  it('checks an image after session validation without charging usage', () => {
    const prepared = run('Map: Prepare Request', { body: { action: 'check', image: { b64: 'YWJj', mediaType: 'image/png' }, perspective: 'isometric' } });
    expect(prepared._map).toMatchObject({ action: 'check', cost: 0, perspective: 'isometric' });
    expect(run('Map: Prepare Request', { body: { action: 'check', image: { b64: 'bad', mediaType: 'image/png' } } }).error).toBe('BAD_REQUEST');
    const rate = run('Rate Limit: Validate', {}, [], { 'Rate Limit: Prepare Check': { ...prepared, _rateLimitCheck: { monthlyLimit: 3, tier: 'free' } } });
    expect(rate._map.cost).toBe(0);
    const exhausted = run('Rate Limit: Validate', {}, [{ session_key: 'x', npc_count: 5 }], { 'Rate Limit: Prepare Check': { ...prepared, _rateLimitCheck: { monthlyLimit: 3, tier: 'free' } } });
    expect(exhausted._map.action).toBe('check');
    const request = run('Check: Build Request', prepared);
    expect(request.body.model).toBe('google/gemini-3.7-flash');
    expect(request.body.messages[0].content[0].text).toBe(CHECK_PROMPT);
    const result = run('Check: Parse', { choices: [{ message: { content: JSON.stringify({ grid: false, text: true, frame: false, margin: 0.5 }) } }], usage: { cost: 0.002 } }, [], { 'Rate Limit: Validate': prepared });
    expect(result).toMatchObject({ ok: true, defects: { retry: true, reasons: ['text'] }, costUsd: 0.002, usesCharged: 0 });
    expect(result.raw.margin).toBe(0.5);
    expect(workflow.connections['IF: Is Authorized?'].main[0][0].node).toBe('Map: Prepare Request');
    expect(workflow.connections['IF: Check Action?'].main[0][0].node).toBe('Check: Build Request');
    const reached = new Set();
    const walk = at => { if (reached.has(at)) return; reached.add(at); for (const output of workflow.connections[at]?.main || []) for (const edge of output) walk(edge.node); };
    walk('Check: Build Request');
    expect(reached.has('Increment Usage Counter')).toBe(false);
  });

  it('validates mockup counts and charges the selected count', () => {
    for (const [count, cost] of [[1, 1], [2, 2], [4, 3]]) {
      const prepared = run('Map: Prepare Request', { body: { action: 'mockups', map: sample, count } });
      expect(prepared._map).toMatchObject({ count, cost });
      const named = { 'Rate Limit: Validate': prepared };
      expect(runAll('Mockups: Build Requests', {}, [], named)).toHaveLength(count);
      const limited = run('Rate Limit: Validate', {}, [], { 'Rate Limit: Prepare Check': { ...prepared, _rateLimitCheck: { monthlyLimit: cost - 1, tier: 'free' } } });
      expect(limited.message).toContain(`costs ${cost} uses`);
      const collected = run('Mockups: Collect', {}, [{ data: [{ b64_json: 'YWJj' }] }], named);
      expect(collected.failed).toBe(count - 1);
      expect(run('Map: Build Response', {}, [], { ...named, 'Mockups: Collect': collected }).usesCharged).toBe(cost);
      expect(run('Increment Usage Counter', {}, [], { 'Rate Limit: Validate': { ...prepared, _rateLimitCheck: { currentCount: 0, sessionKey: 'x', currentMonth: '2026-11' } } }).npc_count).toBe(cost);
    }
    for (const count of [0, 3, '2', null]) expect(run('Map: Prepare Request', { body: { action: 'mockups', map: sample, count } }).error).toBe('BAD_REQUEST');
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
    expect(serialized).not.toMatch(new RegExp('sk-or-|Bearer\\s+[A-Za-z0-9]|' + ['PATREON', 'CREATOR', 'TOKEN'].join('_') + '|4Ydo' + 'Fngay8'));
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
    for (const n of workflow.nodes.filter(n => !openRouterNodes.includes(n) && n.name !== 'Patreon: Members')) {
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
    expect(parents.sort()).toEqual(['IF: Edit Image OK?', 'IF: Mockups OK?', 'IF: Retrace OK?', 'IF: Scale OK?']);
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

  it('crops edits and selects model, resolution, and cost by quality', () => {
    const png = Buffer.alloc(32); png.set([0x89, 0x50, 0x4e, 0x47]); png.writeUInt32BE(1600, 16); png.writeUInt32BE(1200, 20);
    const image = { b64: png.toString('base64'), mediaType: 'image/png' };
    const box = [0.2, 0.3, 0.8, 0.7];
    const [x,y,r,b] = cropBox(box,1600,1200,0.15);
    for (const [quality,model,cost,resolution] of [
      ['standard','google/gemini-3.1-flash-image',4,'2K'],
      ['economy','google/gemini-3.1-flash-image',4,'2K'],
      ['draft','black-forest-labs/flux.2-klein-4b',4,'2K'],
      ['high','bytedance-seed/seedream-5-0-pro',6,'2K'],
    ]) {
      const prep = run('Map: Prepare Request',{body:{action:'edit',image,box,change:'Broken tables',quality}});
      expect(prep._map).toMatchObject({cost,model,resolution,crop:{x,y,width:r-x,height:b-y},box:{x:320,y:360,width:960,height:480}});
      expect(prep._map.prompt).toContain('This is a cut-out piece of a top-down battle map. Broken tables');
    }
    const small = run('Map: Prepare Request',{body:{action:'edit',image,box:[0.3,0.3,0.5,0.5],change:'Ash',quality:'standard'}});
    expect(small._map.resolution).toBe('1K');
    expect(run('Map: Prepare Request',{body:{action:'edit',image,box:[0,0,0.01,1],change:'Ash'}}).error).toBe('BAD_REQUEST');
    expect(run('Map: Prepare Request',{body:{action:'edit',image,box,change:'x'.repeat(501)}}).error).toBe('BAD_REQUEST');
    expect(node('Edit: Crop Image').parameters).toMatchObject({operation:'crop',positionX:'={{ $json.crop.x }}',positionY:'={{ $json.crop.y }}'});
    const named = {'Rate Limit: Validate':small};
    expect(run('Edit: Parse Image',{data:[{b64_json:'YWJj',media_type:'image/png'}],usage:{cost:0.03}},[],named)).toMatchObject({ok:true,action:'edit',patch:{b64:'YWJj'},crop:small._map.crop,costUsd:0.03});
    expect(run('Edit: Parse Image',{error:'provider failed'},[],named)).toMatchObject({ok:false,usesCharged:0});
    expect(workflow.connections['IF: Edit Image OK?'].main[1][0].node).toBe('Respond: 502 Generation Failed');
  });

  it('retrace returns the same Foundry documents as bake and charges only on data', () => {
    const png = Buffer.alloc(32); png.set([0x89, 0x50, 0x4e, 0x47]); png.writeUInt32BE(1200, 16); png.writeUInt32BE(900, 20);
    const image = {b64:png.toString('base64'),mediaType:'image/png'};
    const trace = {walls:[{kind:'door',x1:0,y1:0,x2:200,y2:0}],lights:[{kind:'torch',x:500,y:500}]};
    const req = {image,squares:[12,9],units:'metric',trace};
    const baked = run('Bake: Prepare',run('Map: Prepare Request',{body:{action:'bake',...req}}));
    const prep = run('Map: Prepare Request',{body:{action:'retrace',image,squares:req.squares,units:req.units}});
    expect(prep._map.cost).toBe(1);
    const requests = runAll('Retrace: Build Requests',{},[],{'Rate Limit: Validate':prep});
    expect(requests.map(r=>r.purpose)).toEqual(['walls','lights']);
    const response = value => ({choices:[{message:{content:JSON.stringify(value)}}],usage:{cost:0.01}});
    const result = run('Retrace: Collect',{},[response({walls:trace.walls}),response({lights:trace.lights})],{'Rate Limit: Validate':prep,'Retrace: Build Requests':requests});
    expect(result).toMatchObject({ok:true,action:'retrace',walls:baked.walls,lights:baked.lights,counts:baked.counts,costUsd:0.02});
    const failed = run('Retrace: Collect',{},[{},{}],{'Rate Limit: Validate':prep,'Retrace: Build Requests':requests});
    expect(failed).toMatchObject({ok:false,usesCharged:0});
    expect(workflow.connections['IF: Retrace OK?'].main[1][0].node).toBe('Respond: 502 Generation Failed');
  });
});
