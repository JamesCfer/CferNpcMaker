import { describe, it, expect, vi } from 'vitest';
import { filePicker, directoryExists, uploadMap, sceneData, wallData, lightData, loadMapTemplates, createMapScene, viewMapScene, notifyMapCreated } from '../../modules/ai-battle-map-maker/scripts/compat.js';
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
});
