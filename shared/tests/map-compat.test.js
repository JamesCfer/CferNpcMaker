import { describe, it, expect, vi } from 'vitest';
import { filePicker, directoryExists, uploadMap, sceneData, wallData, lightData, loadMapTemplates, createMapScene, viewMapScene, notifyMapCreated, registerMapSceneControls } from '../../modules/ai-battle-map-maker/scripts/compat.js';
import { drawingBox, rememberCreatedDrawing, selectRegionDrawing, sceneTraceOptions } from '../../modules/ai-battle-map-maker/scripts/map-logic.js';
import { registerSidebar } from '../scripts/core/sidebar.js';

describe('map Foundry compatibility', () => {
  it('chooses namespaced and legacy pickers and preserves uploaded paths', async () => {
    const legacy = { upload: vi.fn().mockResolvedValue({ path: 'legacy/map.png' }) };
    const modern = { upload: vi.fn().mockResolvedValue({ path: 'modern/map.png' }) };
    expect(filePicker({}, legacy)).toBe(legacy);
    expect(filePicker({ applications: { apps: { FilePicker: { implementation: modern } } } }, legacy)).toBe(modern);
    expect(await uploadMap(modern, 'worlds/x/battle-maps', 'file')).toBe('modern/map.png');
    expect(modern.upload).toHaveBeenCalledWith('data', 'worlds/x/battle-maps', 'file', {});
    expect(directoryExists(new Error('Directory already exists'))).toBe(true);
    expect(directoryExists(new Error('EEXIST'))).toBe(true);
    await expect(uploadMap({ upload: async () => ({}) }, 'x', 'file')).rejects.toThrow('path');
  });
  it('adapts the server document data for both generations', () => {
    const baked = { name: 'Map', grid: { type: 1, size: 100, distance: 5, units: 'ft' }, padding: 0, environment: { darknessLevel: 0.7 } };
    for (const version of [13, 14]) expect(sceneData(baked, 'map.webp', { release: { generation: version } })).toMatchObject({ ...baked, background: { src: 'map.webp' } });
    expect(wallData({ c: [0, 0, 1, 1], door: 1, move: 20, sight: 20, light: 20, sound: 20, extra: true })).not.toHaveProperty('extra');
    expect(lightData({ x: 1, y: 2, config: { bright: 3, dim: 4, color: '#fff', alpha: 0.5, animation: {}, extra: true } }).config).not.toHaveProperty('extra');
  });
  it('loads templates from the namespaced or legacy API', () => {
    const modern = vi.fn(); const legacy = vi.fn();
    loadMapTemplates(['a'], { applications: { handlebars: { loadTemplates: modern } } }, legacy);
    loadMapTemplates(['b'], {}, legacy);
    expect(modern).toHaveBeenCalledWith(['a']);
    expect(legacy).toHaveBeenCalledWith(['b']);
  });
  it('creates and opens scenes through available document APIs', async () => {
    const create = vi.fn().mockResolvedValue({ view: vi.fn().mockResolvedValue(true) });
    const scene = await createMapScene({ name: 'Map' }, { Scene: { documentClass: { create } } });
    expect(create).toHaveBeenCalledWith({ name: 'Map' });
    expect(await viewMapScene(scene)).toBe(true);
    const notify = vi.fn();
    notifyMapCreated('ready', { notifications: { success: notify } });
    expect(notify).toHaveBeenCalledWith('ready');
  });
  it('registers both scene directory render and header hooks', () => {
    const hooks = [];
    const previous = globalThis.Hooks;
    globalThis.Hooks = { on: name => hooks.push(name) };
    try { registerSidebar('ai-battle-map-maker', () => {}, { directories: ['scenes'] }); }
    finally { globalThis.Hooks = previous; }
    expect(hooks).toContain('renderSceneDirectory');
    expect(hooks).toContain('getHeaderControlsSceneDirectory');
    expect(hooks).toContain('getHeaderControlsApplicationV2');
  });
  it('registers GM controls in object and array control shapes', () => {
    const edit = vi.fn(), retrace = vi.fn(), open = vi.fn();
    const object = { drawings: { name: 'drawings', order: 10, tools: { rect: { name: 'rect' } } },
      walls: { name: 'walls', order: 20, tools: {} } };
    expect(registerMapSceneControls(object, edit, retrace, open, true)).toBe(true);
    const group = object['ai-battle-map'];
    const previousCanvas = globalThis.canvas, previousUi = globalThis.ui;
    const activate = vi.fn(), controlActivate = vi.fn(), info = vi.fn();
    globalThis.canvas = { drawings: { activate } };
    globalThis.ui = { controls: { activate: controlActivate }, notifications: { info } };
    try {
      expect(group.onChange).toBeTypeOf('function');
      expect(group.onToolChange).toBeUndefined();
      group.onChange({}, false);
      group.onChange({}, true);
      expect(activate).not.toHaveBeenCalled();
      expect(controlActivate).not.toHaveBeenCalled();
      group.tools['ai-draw-box'].onChange(true);
      expect(controlActivate).toHaveBeenCalledWith({ control: 'drawings', tool: 'rect' });
      expect(info).toHaveBeenCalledWith('Drag a box over the area you want to change, then click Edit region.');
    } finally { globalThis.canvas = previousCanvas; globalThis.ui = previousUi; }
    expect(group).toMatchObject({ name: 'ai-battle-map', title: 'AI Battle Map', icon: 'fa-solid fa-map',
      layer: 'drawings', order: 21, activeTool: 'ai-pointer', visible: true });
    expect(Object.keys(group.tools)).toEqual(['ai-pointer', 'ai-draw-box', 'ai-edit-region', 'ai-retrace-walls', 'ai-open-builder']);
    expect(Object.values(group.tools).map(tool => tool.order)).toEqual([0, 1, 2, 3, 4]);
    expect(group.tools['ai-pointer']).toMatchObject({ name: 'ai-pointer', title: 'AI Battle Map tools', icon: 'fa-solid fa-map' });
    expect(group.tools['ai-draw-box']).toMatchObject({ name: 'ai-draw-box', title: 'Draw edit box', icon: 'fa-regular fa-square', button: true });
    expect(group.tools['ai-edit-region']).toMatchObject({ title: 'Edit region (4 uses)', icon: 'fa-solid fa-wand-magic-sparkles', button: true });
    expect(group.tools['ai-retrace-walls']).toMatchObject({ title: 'Re-trace walls (1 use)', icon: 'fa-solid fa-draw-polygon', button: true });
    expect(group.tools['ai-open-builder']).toMatchObject({ title: 'Open AI Battle Map Maker', icon: 'fa-solid fa-map-location-dot', button: true });
    expect(group.tools['ai-pointer']).not.toHaveProperty('button');
    for (const [name, handler] of [['ai-edit-region', edit], ['ai-retrace-walls', retrace], ['ai-open-builder', open]]) {
      group.tools[name].onChange(false);
      expect(handler).not.toHaveBeenCalled();
      group.tools[name].onChange(true);
      expect(handler).toHaveBeenCalledTimes(1);
    }
    expect(Object.keys(object.drawings.tools)).toEqual(['rect', 'ai-edit-region']);
    expect(object.drawings.tools['ai-edit-region']).toMatchObject({ button: true, order: 1 });
    expect(object.walls.tools).toEqual({});

    const array = [{ name: 'drawings', tools: [] }, { name: 'walls', tools: [] }];
    expect(registerMapSceneControls(array, edit, retrace, open, true)).toBe(true);
    expect(array[2]).toMatchObject({ name: 'ai-battle-map', layer: 'drawings', activeTool: 'ai-pointer', order: 2 });
    expect(array[2].tools.map(tool => tool.name)).toEqual(Object.keys(group.tools));
    expect(array[0].tools.map(tool => tool.name)).toEqual(['ai-edit-region']);
    expect(array[1].tools).toEqual([]);
    expect(registerMapSceneControls(array, edit, retrace, open, true)).toBe(true);
    expect(array).toHaveLength(3);
    expect(array[0].tools).toHaveLength(1);
    expect(registerMapSceneControls({}, edit, retrace, open, false)).toBe(false);
  });
  it('uses a controlled rectangle, then the last created rectangle, then the newest owned rectangle', () => {
    const scene = { id: 'scene', drawings: { contents: [
      { id: 'old', author: 'gm', shape: { type: 'r' }, sort: 1 },
      { id: 'new', author: 'gm', shape: { type: 'r' }, sort: 2 },
      { id: 'other', author: 'other', shape: { type: 'r' }, sort: 99 },
    ] } };
    const recent = new Map([['scene', 'old']]);
    const controlled = { document: { id: 'selected', shape: { type: 'r' } } };
    expect(selectRegionDrawing(scene, [controlled], 'gm', recent)).toBe(controlled);
    expect(selectRegionDrawing(scene, [], 'gm', recent)?.id).toBe('old');
    recent.set('scene', 'gone');
    expect(selectRegionDrawing(scene, [], 'gm', recent)?.id).toBe('new');
    expect(selectRegionDrawing({ id: 'empty', drawings: { contents: [] } }, [], 'gm', recent)).toBeNull();
    rememberCreatedDrawing({ id: 'old', parent: scene, shape: { type: 'r' } }, 'gm', 'gm');
    expect(selectRegionDrawing(scene, [], 'gm')?.id).toBe('old');
  });
  it('maps drawing bounds and scene grid to request values', () => {
    const scene={width:1000,height:800,grid:{size:100,units:'m'}};
    expect(drawingBox({document:{x:200,y:150,shape:{type:'r',width:500,height:300}}},scene,{sceneX:100,sceneY:50})).toEqual([0.1,0.125,0.6,0.5]);
    expect(drawingBox({document:{x:0,y:0,shape:{type:'r',width:300,height:300}}},scene,{sceneX:100,sceneY:50})).toEqual([0,0,0.2,0.3125]);
    expect(drawingBox({document:{x:0,y:0,shape:{type:'e',width:300,height:300}}},scene,{})).toBeNull();
    expect(sceneTraceOptions(scene)).toEqual({squares:[10,8],units:'metric'});
  });
});
