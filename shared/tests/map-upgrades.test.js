import { describe, it, expect, vi, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { sceneEnvironment, pickWeather, overlayShapes, packTrace, unpackTrace, mapRecord, isSavedMap, presetFromValues, upsertPreset, BUILT_IN_PRESETS, STYLE_FIELDS, buildMapRequest } from '../../modules/ai-battle-map-maker/scripts/map-logic.js';
import { sceneData, saveToWorld, mapSceneFolder, setSceneThumbnail, hasWeather, startProgress, registerMapSceneControls } from '../../modules/ai-battle-map-maker/scripts/compat.js';
import { BattleMapAdapter, renderMapThumbnails, zoomImage } from '../../modules/ai-battle-map-maker/scripts/adapter.js';

const markup = readFileSync(resolve('modules/ai-battle-map-maker/templates/form.html'), 'utf8');

describe('scene lighting and weather', () => {
  it('lights day and dusk maps globally so walled maps are not black', () => {
    expect(sceneEnvironment('day', 5)).toEqual({ darknessLevel: 0, globalLight: { enabled: true, darkness: { min: 0, max: 1 } } });
    expect(sceneEnvironment('dusk', 0)).toMatchObject({ darknessLevel: 0.35, globalLight: { enabled: true } });
    expect(sceneEnvironment('unknown')).toMatchObject({ darknessLevel: 0, globalLight: { enabled: true } });
  });
  it('leaves night to the placed lights, unless there are none', () => {
    expect(sceneEnvironment('night', 3)).toEqual({ darknessLevel: 0.75, globalLight: { enabled: false } });
    expect(sceneEnvironment('night', 0)).toMatchObject({ darknessLevel: 0.75, globalLight: { enabled: true } });
  });
  it('picks weather from an explicit choice or outdoor weather words', () => {
    expect(pickWeather('snow', 'a sunny beach', 'indoor')).toBe('snow');
    expect(pickWeather('none', 'a storm-lashed pier')).toBe('');
    expect(pickWeather('bogus', 'rain')).toBe('');
    expect(pickWeather('auto', 'A storm-lashed pier at night', 'outdoor')).toBe('rainStorm');
    expect(pickWeather('auto', 'A village in a blizzard', 'outdoor')).toBe('blizzard');
    expect(pickWeather('auto', 'A snowy mountain pass', 'outdoor')).toBe('snow');
    expect(pickWeather('auto', 'A misty graveyard', 'outdoor')).toBe('fog');
    expect(pickWeather('auto', 'A rainy market street', 'outdoor')).toBe('rain');
    expect(pickWeather('auto', 'A forest road in autumn', 'outdoor')).toBe('leaves');
    expect(pickWeather('auto', 'A quiet forest road', 'outdoor')).toBe('');
    expect(pickWeather('auto', 'A tavern during a storm', 'indoor')).toBe('');
    expect(pickWeather(undefined, 'A training yard, restrained decor', 'outdoor')).toBe('');
  });
  it('merges module scene settings into the server scene', () => {
    const baked = { name: 'Map', environment: { darknessLevel: 0.7 } };
    const scene = sceneData(baked, 'map.webp', { release: { generation: 13 } }, { environment: sceneEnvironment('night', 2), weather: 'rain', folder: 'f1' });
    expect(scene).toMatchObject({ name: 'Map', background: { src: 'map.webp' }, weather: 'rain', folder: 'f1', environment: { darknessLevel: 0.75, globalLight: { enabled: false } } });
    expect(sceneData(baked, 'map.webp', { release: { generation: 14 } })).not.toHaveProperty('weather');
    expect(hasWeather('rain', { weatherEffects: { rain: {} } })).toBe(true);
    expect(hasWeather('leaves', { weatherEffects: { rain: {} } })).toBe(false);
    expect(hasWeather('', {})).toBe(false);
  });
});

describe('Foundry helpers', () => {
  it('uploads into the world battle-maps folder and tolerates an existing folder', async () => {
    const picker = { createDirectory: vi.fn().mockRejectedValue(new Error('EEXIST')), upload: vi.fn().mockResolvedValue({ path: 'worlds/w/battle-maps/a.webp' }) };
    expect(await saveToWorld('file', 'w', picker)).toBe('worlds/w/battle-maps/a.webp');
    expect(picker.createDirectory).toHaveBeenCalledWith('data', 'worlds/w/battle-maps');
    await expect(saveToWorld('file', 'w', { createDirectory: vi.fn().mockRejectedValue(new Error('EACCES')) })).rejects.toThrow('EACCES');
  });
  it('finds or creates the scene folder and never throws', async () => {
    const game = { folders: [{ id: 'a', type: 'Actor', name: 'AI Battle Maps' }, { id: 's', type: 'Scene', name: 'AI Battle Maps' }] };
    expect(await mapSceneFolder('AI Battle Maps', game, {})).toBe('s');
    const create = vi.fn().mockResolvedValue({ id: 'new' });
    expect(await mapSceneFolder('AI Battle Maps', { folders: [] }, { create })).toBe('new');
    expect(create).toHaveBeenCalledWith({ name: 'AI Battle Maps', type: 'Scene' });
    expect(await mapSceneFolder('AI Battle Maps', { folders: [] }, { create: vi.fn().mockRejectedValue(new Error('no')) })).toBeNull();
  });
  it('sets a scene thumbnail when Foundry can make one', async () => {
    const scene = { createThumbnail: vi.fn().mockResolvedValue({ thumb: 'data:x' }), update: vi.fn() };
    expect(await setSceneThumbnail(scene)).toBe(true);
    expect(scene.update).toHaveBeenCalledWith({ thumb: 'data:x' });
    expect(await setSceneThumbnail({ createThumbnail: vi.fn().mockRejectedValue(new Error('no canvas')) })).toBe(false);
    expect(await setSceneThumbnail({})).toBe(false);
  });
  it('drives one progress notification and falls back quietly', () => {
    const update = vi.fn(), remove = vi.fn();
    const info = vi.fn().mockReturnValue({ update });
    const progress = startProgress('Working…', { notifications: { info, remove } });
    progress.step('Half', 0.5); progress.step('Over', 4); progress.done();
    expect(info).toHaveBeenCalledTimes(1);
    expect(info).toHaveBeenCalledWith('Working…', { progress: true, console: false });
    expect(update.mock.calls).toEqual([[{ message: 'Half', pct: 0.5 }], [{ message: 'Over', pct: 0.99 }], [{ pct: 1 }]]);
    const legacy = startProgress('Working…', { notifications: { info: vi.fn().mockReturnValue(7) } });
    expect(() => { legacy.step('x', 0.2); legacy.done(); }).not.toThrow();
    expect(() => startProgress('x', {}).done()).not.toThrow();
  });
  it('adds the undo tool between edit and re-trace', () => {
    const controls = { drawings: { name: 'drawings', tools: {} } };
    registerMapSceneControls(controls, () => {}, () => {}, () => {}, true, () => {});
    const tools = Object.values(controls['ai-battle-map'].tools);
    expect(tools.map(tool => tool.name)).toEqual(['ai-pointer', 'ai-draw-box', 'ai-edit-region', 'ai-undo-edit', 'ai-retrace-walls', 'ai-open-builder']);
    expect(tools.map(tool => tool.order)).toEqual([0, 1, 2, 3, 4, 5]);
  });
});

describe('saved maps and presets', () => {
  const trace = { walls: [{ kind: 'wall', x1: 0.4, y1: 10, x2: 500, y2: 10 }, { kind: 'door', x1: 500, y1: 10, x2: 540, y2: 10 }, { kind: 'odd', x1: 1, y1: 2, x2: 3, y2: 4 }, { x1: 'bad' }],
    lights: [{ kind: 'torch', x: 100.6, y: 200, color: '#ffaa00' }, { kind: 'lamp', x: 5, y: 6 }, null] };
  it('packs a trace small and restores it', () => {
    const packed = packTrace(trace);
    expect(packed).toEqual({ w: [[0, 0, 10, 500, 10], [1, 500, 10, 540, 10], [0, 1, 2, 3, 4]], l: [['torch', 101, 200, '#ffaa00'], ['lamp', 5, 6]] });
    expect(unpackTrace(packed)).toEqual({ walls: [{ kind: 'wall', x1: 0, y1: 10, x2: 500, y2: 10 }, { kind: 'door', x1: 500, y1: 10, x2: 540, y2: 10 }, { kind: 'wall', x1: 1, y1: 2, x2: 3, y2: 4 }],
      lights: [{ kind: 'torch', x: 101, y: 200, color: '#ffaa00' }, { kind: 'lamp', x: 5, y: 6 }] });
    expect(unpackTrace(undefined)).toEqual({ walls: [], lights: [] });
    expect(packTrace(null)).toEqual({ w: [], l: [] });
  });
  it('builds a reopenable Created Maps entry without the image', () => {
    const entry = mapRecord({ values: { description: '  A ruined keep  ', mood: 'night', walls: true, color: '4', secret: 'x', layout: '' }, path: 'worlds/w/battle-maps/a.webp',
      final: { scene: { squares: [15, 12] }, trace, warning: 'Walls could not be traced.', image: { b64: 'AAAA' } }, units: 'metric', now: 5, id: 'abc' });
    expect(entry).toMatchObject({ id: 'abc', name: 'A ruined keep', status: 'success', createdAt: 5, map: { path: 'worlds/w/battle-maps/a.webp', units: 'metric', squares: [15, 12], warning: 'Walls could not be traced.',
      values: { description: '  A ruined keep  ', mood: 'night', walls: true, color: '4' } } });
    expect(entry.map.values).not.toHaveProperty('secret');
    expect(entry.map.values).not.toHaveProperty('layout');
    expect(JSON.stringify(entry)).not.toContain('AAAA');
    expect(isSavedMap(entry)).toBe(true);
    expect(isSavedMap({ name: 'old entry', description: 'x' })).toBe(false);
  });
  it('saves, replaces and limits presets', () => {
    const preset = presetFromValues('  Night ink  ', { style: 'hand-drawn ink', mood: 'night', description: 'ignored', detail: 4 });
    expect(preset).toEqual({ name: 'Night ink', values: { style: 'hand-drawn ink', mood: 'night', detail: '4' } });
    let list = upsertPreset(null, preset);
    list = upsertPreset(list, presetFromValues('night INK', { mood: 'day' }));
    expect(list).toEqual([{ name: 'night INK', values: { mood: 'day' } }]);
    for (let i = 0; i < 25; i++) list = upsertPreset(list, presetFromValues(`p${i}`, { mood: 'day' }));
    expect(list).toHaveLength(20);
    expect(list.at(-1).name).toBe('p24');
    expect(() => upsertPreset(list, presetFromValues(' ', {}))).toThrow('name');
    expect(() => upsertPreset(list, presetFromValues(BUILT_IN_PRESETS[0].name.toUpperCase(), {}))).toThrow('built-in');
  });
  it('ships built-in presets and templates the form can apply', () => {
    const container = document.createElement('div');
    container.innerHTML = markup;
    const form = container.querySelector('form');
    for (const preset of BUILT_IN_PRESETS) {
      for (const [key, value] of Object.entries(preset.values)) {
        expect(STYLE_FIELDS).toContain(key);
        const el = form.elements[key];
        if (el?.tagName === 'SELECT') expect([...el.options].map(o => o.value)).toContain(value);
        else expect(Number(value)).toBeGreaterThanOrEqual(1);
      }
    }
    const ideas = [...form.querySelectorAll('.map-template-select option')].filter(o => o.value);
    expect(ideas.length).toBeGreaterThanOrEqual(10);
    for (const idea of ideas) {
      expect([...form.elements.genre.options].map(o => o.value)).toContain(idea.dataset.genre);
      expect(buildMapRequest({ description: idea.value, setting: idea.dataset.setting }).setting).toBe(idea.dataset.setting);
    }
    expect(markup).not.toContain('{{');
  });
  it('scales the trace to the preview and drops bad shapes', () => {
    const shapes = overlayShapes({ walls: [{ kind: 'door', x1: 0, y1: 0, x2: 1000, y2: 500 }, { x1: 1 }], lights: [{ x: 500, y: 500 }, { x: 2000, y: 5 }, { x: 10, y: 10, color: 'red' }] }, 800, 600);
    expect(shapes.walls).toEqual([{ kind: 'door', x1: 0, y1: 0, x2: 800, y2: 300 }]);
    expect(shapes.lights).toEqual([{ x: 400, y: 300, color: '#ffb35c' }, { x: 8, y: 6, color: '#ffb35c' }]);
  });
});

describe('builder form behaviour', () => {
  let form, adapter;
  beforeEach(() => {
    localStorage.clear();
    document.body.innerHTML = markup;
    form = document.body.querySelector('form');
    adapter = new BattleMapAdapter();
    adapter.app = { authenticated: true, history: [], storage: { saveHistory: vi.fn(), getKey: () => 'k' }, _renderHistory: vi.fn() };
    adapter.onFormMount(form);
  });
  const fire = (el, type = 'change') => el.dispatchEvent(new Event(type, { bubbles: true }));

  it('fills the description, genre and setting from a quick map idea', () => {
    const select = form.querySelector('.map-template-select');
    select.value = [...select.options].find(o => o.dataset.genre === 'sci-fi').value;
    fire(select);
    expect(form.elements.description.value).toContain('starship');
    expect(form.elements.genre.value).toBe('sci-fi');
    expect(form.elements.setting.value).toBe('indoor');
    expect(select.selectedIndex).toBe(0);
  });
  it('applies a built-in preset to style fields only and updates dial read-outs', () => {
    form.elements.description.value = 'Keep me';
    const select = form.querySelector('.map-preset-select');
    select.value = 'b2';
    fire(select);
    expect(form.elements.genre.value).toBe('horror');
    expect(form.elements.mood.value).toBe('night');
    expect(form.elements.wear.value).toBe('5');
    expect(form.querySelector('#map-dial-wear').closest('.map-dial').querySelector('output').textContent).toBe('5');
    expect(form.elements.description.value).toBe('Keep me');
    expect(form.querySelector('.map-preset-delete').disabled).toBe(true);
  });
  it('saves and deletes a personal preset', async () => {
    adapter.askText = vi.fn().mockResolvedValue('My look');
    form.elements.style.value = 'cartoon';
    form.querySelector('.map-preset-save').click();
    await vi.waitFor(() => expect(JSON.parse(localStorage.getItem('ai-battle-map-maker.presets'))).toHaveLength(1));
    const select = form.querySelector('.map-preset-select');
    expect(select.value).toBe('u0');
    expect(select.selectedOptions[0].textContent).toBe('My look');
    const remove = form.querySelector('.map-preset-delete');
    expect(remove.disabled).toBe(false);
    remove.click();
    expect(JSON.parse(localStorage.getItem('ai-battle-map-maker.presets'))).toEqual([]);
  });
  it('remembers the form between openings, including unchecked boxes', () => {
    form.elements.description.value = 'A flooded mine';
    form.elements.walls.checked = false;
    form.elements.detail.value = '5';
    fire(form.elements.detail, 'input');
    document.body.innerHTML = markup;
    const reopened = document.body.querySelector('form');
    const again = new BattleMapAdapter();
    again.app = adapter.app;
    again.onFormMount(reopened);
    expect(reopened.elements.description.value).toBe('A flooded mine');
    expect(reopened.elements.walls.checked).toBe(false);
    expect(reopened.elements.detail.value).toBe('5');
    expect(reopened.querySelector('#map-dial-detail').closest('.map-dial').querySelector('output').textContent).toBe('5');
  });
  it('keeps earlier mockups when a new set arrives and after the window is reopened', async () => {
    form.elements.description.value = 'First map';
    const image = { b64: 'AAAA', mediaType: 'image/png' };
    adapter.request = vi.fn().mockResolvedValueOnce({ images: [image, image], aspect: '4:3' }).mockResolvedValueOnce({ images: [image], aspect: '3:2', layout: 'A plan.' });
    await adapter.generateMockups();
    form.elements.description.value = 'Second map';
    await adapter.generateMockups();
    expect(adapter.mockups).toHaveLength(3);
    expect(adapter.mockups.map(m => !!m.earlier)).toEqual([false, true, true]);
    expect(adapter.mockups[0].map).toMatchObject({ description: 'Second map', layout: 'A plan.' });
    expect(adapter.mockups[1].map.description).toBe('First map');
    expect(form.querySelector('.map-mockup-note').hidden).toBe(false);
    expect([...form.querySelectorAll('.map-picked')].map(n => n.textContent)).toEqual(['', 'Earlier', 'Earlier']);
    document.body.innerHTML = markup;
    adapter.onFormMount(document.body.querySelector('form'));
    expect(document.body.querySelectorAll('.map-thumbnail')).toHaveLength(3);
  });
  it('plans again when the description changed but the old plan was left untouched', async () => {
    const image = { b64: 'AAAA', mediaType: 'image/png' };
    adapter.request = vi.fn(async payload => payload.action === 'plan' ? { layout: 'Plan for the mine.' } : { images: [image], aspect: '4:3' });
    form.elements.description.value = 'A flooded mine';
    await adapter.planLayout();
    await adapter.generateMockups();
    expect(adapter.request.mock.calls[1][0]).toMatchObject({ action: 'mockups', map: { layout: 'Plan for the mine.' } });
    expect(adapter.request.mock.calls[1][0]).not.toHaveProperty('autoPlan');
    form.elements.description.value = 'A sky temple';
    await adapter.generateMockups();
    expect(adapter.request.mock.calls[2][0].map).not.toHaveProperty('layout');
    expect(adapter.request.mock.calls[2][0].autoPlan).toBe(true);
    form.elements.layout.value = 'My own layout.';
    form.elements.description.value = 'A sky temple at dawn';
    await adapter.generateMockups();
    expect(adapter.request.mock.calls[3][0].map.layout).toBe('My own layout.');
    expect(JSON.parse(localStorage.getItem('ai-battle-map-maker.form')).layout).toBe('');
  });
  it('finishes an earlier mockup with the description it was drawn from', async () => {
    globalThis.game = { settings: { get: () => 'imperial' } };
    const image = { b64: 'AAAA', mediaType: 'image/png' };
    adapter.mockups = [{ ...image, aspect: '4:3', map: buildMapRequest({ description: 'New map' }) }, { ...image, aspect: '4:3', earlier: true, map: buildMapRequest({ description: 'Old map' }) }];
    adapter.selected = 1;
    form.elements.description.value = 'New map, edited';
    adapter.request = vi.fn().mockRejectedValue(new Error('stop here'));
    globalThis.ui = { notifications: { error: vi.fn() } };
    await adapter.generateFinal();
    expect(adapter.request.mock.calls[0][0]).toMatchObject({ action: 'final', map: { description: 'Old map' }, mockup: image });
    adapter.selected = 0;
    await adapter.generateFinal();
    expect(adapter.request.mock.calls[1][0].map.description).toBe('New map, edited');
  });
  it('retries a free action once after a dropped connection, never an auth failure', async () => {
    adapter.request = vi.fn().mockRejectedValueOnce(new TypeError('Failed to fetch')).mockResolvedValueOnce({ ok: true });
    expect(await adapter.requestFree({ action: 'bake' })).toEqual({ ok: true });
    expect(adapter.request).toHaveBeenCalledTimes(2);
    adapter.request = vi.fn().mockRejectedValue(new Error('A map final costs 6 uses.'));
    await expect(adapter.requestFree({ action: 'plan' })).rejects.toThrow('costs');
    expect(adapter.request).toHaveBeenCalledTimes(1);
  });
  it('marks picked and earlier thumbnails and offers zoom without nesting buttons', () => {
    const box = form.querySelector('.map-thumbnails');
    const onZoom = vi.fn();
    renderMapThumbnails(box, [{ b64: 'AAAA' }, { b64: 'AAAA', earlier: true, aspect: '16:9' }], '4:3', () => {}, { selected: 1, onZoom });
    const tiles = box.querySelectorAll('.map-thumbnail');
    expect(tiles[1].classList.contains('is-selected')).toBe(true);
    expect(tiles[1].querySelector('.map-picked').textContent).toBe('Picked');
    expect(tiles[1].style.getPropertyValue('--map-aspect')).toBe('16 / 9');
    tiles[0].click();
    expect(tiles[1].querySelector('.map-picked').textContent).toBe('Earlier');
    expect(box.querySelector('button button')).toBeNull();
    box.querySelectorAll('.map-zoom-btn')[1].click();
    expect(onZoom).toHaveBeenCalledWith(1);
  });
  it('opens and closes the zoom overlay', () => {
    const overlay = zoomImage('data:image/png;base64,AAAA', 'Mockup 1');
    expect(document.body.contains(overlay)).toBe(true);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    expect(document.body.contains(overlay)).toBe(false);
    const second = zoomImage('data:image/png;base64,AAAA');
    second.click();
    expect(document.body.contains(second)).toBe(false);
  });
});
