/**
 * Runs in the Roll20 page itself (MAIN world), where Roll20's own objects live.
 * The extension's panel talks to it with window.postMessage. It never sees the
 * Patreon key and makes no network calls.
 *
 * Writing uses Roll20's in-page campaign model (window.Campaign, the same
 * objects the editor uses and that VTT Enhancement Suite relies on). Roll20
 * does not document these, so every call is guarded, and the panel falls back
 * to the companion Mod script or a download when something is missing.
 */
(() => {
  if (window.__cferBridge) return;
  window.__cferBridge = true;

  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const campaign = () => window.Campaign || null;

  function probe() {
    const c = campaign();
    return {
      campaign: !!c,
      characters: !!c?.characters?.create,
      handouts: !!c?.handouts?.create,
      chat: !!(window.d20?.textchat?.doChatInput || chatInput()),
      gm: window.is_gm === true || window.is_gm === 'true' || !!document.querySelector('#gm-toolbar, .gm-only, [data-gm="true"]'),
      jumpgate: !!document.querySelector('#jumpgate, .jumpgate, [data-jumpgate]') || /jumpgate/i.test(document.body?.className || ''),
    };
  }

  /* ── chat ─────────────────────────────────────────── */

  function chatInput() {
    return document.querySelector('#textchat-input textarea, textarea#textchat-input, [data-testid="chat-input"] textarea');
  }

  async function sendChat(lines) {
    for (const line of lines) {
      if (window.d20?.textchat?.doChatInput) window.d20.textchat.doChatInput(line);
      else {
        const ta = chatInput();
        if (!ta) throw new Error('Could not find the Roll20 chat box.');
        ta.value = line;
        ta.dispatchEvent(new Event('input', { bubbles: true }));
        const btn = document.querySelector('#textchat-input button, #textchat-input .btn, [data-testid="chat-send"]');
        if (btn) btn.click(); else ta.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, which: 13, bubbles: true }));
      }
      await sleep(150);
    }
    return { sent: lines.length };
  }

  /** Resolve when a chat message containing `text` shows up, or time out. */
  function waitForChat(text, timeoutMs = 4000) {
    const log = document.querySelector('#textchat .content, #textchat, [data-testid="chat-log"]');
    if (!log) return Promise.resolve(false);
    if (log.textContent.includes(text)) return Promise.resolve(true);
    return new Promise(resolve => {
      const obs = new MutationObserver(() => { if (log.textContent.includes(text)) { obs.disconnect(); resolve(true); } });
      obs.observe(log, { childList: true, subtree: true, characterData: true });
      setTimeout(() => { obs.disconnect(); resolve(false); }, timeoutMs);
    });
  }

  /* ── characters ───────────────────────────────────── */

  function listCharacters() {
    const c = campaign();
    if (!c?.characters) return [];
    return c.characters.models.map(m => ({ id: m.id, name: m.get('name') })).sort((a, b) => String(a.name).localeCompare(b.name));
  }

  function findAttr(char, name) {
    return char.attribs?.models?.find(a => a.get('name') === name) || null;
  }

  async function setAttr(char, name, value) {
    let current = value, max;
    if (value && typeof value === 'object') ({ current, max } = value);
    const data = { current: String(current ?? '') };
    if (max !== undefined) data.max = String(max);
    const existing = findAttr(char, name);
    if (existing) existing.save(data);
    else char.attribs.create({ name, ...data });
  }

  async function writeAll(char, attrs) {
    let n = 0;
    for (const [k, v] of Object.entries(attrs || {})) {
      await setAttr(char, k, v);
      if (++n % 50 === 0) await sleep(50); // let Firebase breathe on big statblocks
    }
    return n;
  }

  /* ── D&D 2024 ("Beacon") sheet: everything lives in the `store` attribute ── */

  const uuid = () => (crypto.randomUUID ? crypto.randomUUID() : 'xxxxxxxx-xxxx-4xxx-8xxx-xxxxxxxxxxxx'.replace(/x/g, () => (Math.random() * 16 | 0).toString(16)));
  const shortId = () => Array.from({ length: 9 }, () => ROW.charAt(1 + Math.floor(Math.random() * 62))).join('').replace(/[-_]/g, 'x');

  async function waitForStore(char, ms = 8000) {
    for (let t = 0; t < ms; t += 200) {
      const store = findAttr(char, 'store')?.get('current');
      if (store && typeof store === 'object' && store.integrants) {
        await sleep(1500); // let the sheet finish its own first save
        return findAttr(char, 'store').get('current');
      }
      if (findAttr(char, 'sheet_type') || findAttr(char, 'npc_name') && !findAttr(char, 'appState')) return null; // a classic sheet
      await sleep(200);
    }
    return null;
  }

  /** Merge a mapper's beacon block into the sheet's store and save it. */
  function applyBeacon(char, store, beacon) {
    const next = JSON.parse(JSON.stringify(store));
    const ints = next.integrants.integrants;
    const refs = {};
    const now = Date.now();
    const existing = type => Object.values(ints).filter(i => i.type === type);
    for (const raw of beacon.integrants) if (raw._ref) refs[raw._ref] = raw._id || uuid();
    const children = {};
    for (const raw of beacon.integrants) {
      const { _ref, _parentRef, ...fields } = raw;
      // The sheet starts with a "Darkvision 0" sense and a "None" language: reuse or drop them.
      if (fields.type === 'Sense') { const old = existing('Sense').find(i => i.name === fields.name); if (old) { old.valueFormula = fields.valueFormula; continue; } }
      if (fields.type === 'Language') for (const old of existing('Language').filter(i => i.name === 'None')) delete ints[old._id];
      const id = fields._id || (_ref && refs[_ref]) || uuid();
      ints[id] = { _enabled: true, builderDisplayName: '', childIDs: '[]', label: '', parentID: _parentRef ? refs[_parentRef] : '', source: '',
        ...fields, _id: id, createdTime: now, shortID: shortId() };
      if (_parentRef) (children[refs[_parentRef]] ||= []).push(id);
    }
    for (const [parent, kids] of Object.entries(children)) if (ints[parent]) ints[parent].childIDs = JSON.stringify(kids);
    const order = JSON.parse(next.features?.speciesTraitsDisplayOrder || '[]');
    next.features = { ...next.features, speciesTraitsDisplayOrder: JSON.stringify([...order, ...(beacon.traitRefs || []).map(r => refs[r])]) };
    next.npc = { ...next.npc, ...beacon.store.npc };
    next.character = { ...next.character, ...beacon.store.character, createdWithBuilder: false };
    next.about = { ...next.about, characteristics: { ...(next.about?.characteristics || {}), ...beacon.store.about.characteristics } };
    // Show it as an NPC stat block with rolls whispered to the GM, as the sheet's own NPC button does.
    next.settings = { ...next.settings, layoutState: 'Stat Block', rolls: { ...(next.settings?.rolls || {}), privacy: 'gm' } };
    next.npcEdit = { ...(next.npcEdit || {}), cameFromView: 'Stat Block' };
    findAttr(char, 'store').save({ current: next });
    const app = findAttr(char, 'appState');
    if (app) app.save({ current: 'npc' }); else char.attribs.create({ name: 'appState', current: 'npc' });
    char.save({ character_type: 'npc' });
  }

  function reopenSheet(char) {
    try { char.view?.$el?.dialog?.('close'); } catch (_) { /* optional */ }
    setTimeout(() => { try { char.view?.showDialog?.(); } catch (_) { /* optional */ } }, 300);
  }

  async function createCharacter(plan) {
    const c = campaign();
    if (!c?.characters?.create) throw new Error('This Roll20 page does not allow in-page imports.');
    const char = c.characters.create({ name: plan.character.name, inplayerjournals: '', controlledby: '', archived: false });
    for (let i = 0; i < 40 && !char.attribs; i++) await sleep(50);
    if (!char.attribs) throw new Error('Roll20 did not finish creating the character.');
    if (typeof char.updateBlobs === 'function') char.updateBlobs({ bio: plan.character.bio || '', gmnotes: plan.character.gmnotes || '' });
    if (plan.beacon) {
      // Opening the sheet makes it set itself up; a 2024 sheet then has a `store`.
      try { char.view?.showDialog?.(); } catch (_) { /* sheet UI is optional */ }
      const store = await waitForStore(char);
      if (store) {
        applyBeacon(char, store, plan.beacon);
        reopenSheet(char);
        return { id: char.id, name: plan.character.name, mode: 'sheet2024' };
      }
    }
    await writeAll(char, plan.attributes);
    let mode = 'direct';
    if (plan.drop) {
      await writeAll(char, plan.drop.attributes);
      await setAttr(char, plan.drop.trigger.name, plan.drop.trigger.value);
      // Opening the sheet lets its workers run the compendium-drop import.
      try { char.view?.showDialog?.(); } catch (_) { /* sheet UI is optional */ }
      await sleep(plan.drop.timeoutMs || 5000);
      const built = findAttr(char, plan.drop.verify)?.get('current');
      if (built) mode = 'sheet';
      else if (plan.fallback) await writeAll(char, plan.fallback.attributes);
    }
    return { id: char.id, name: plan.character.name, mode };
  }

  async function give({ recipe, characterId }) {
    const char = campaign()?.characters?.get(characterId);
    if (!char) throw new Error('That character is no longer in this game.');
    const store = findAttr(char, 'store')?.get('current');
    if (store && typeof store === 'object' && store.integrants) {
      // D&D 2024 sheet: NPCs list gear as text; PCs need the item added from the handout by hand.
      if (char.get('character_type') !== 'npc') throw new Error(`${char.get('name')} uses the D&D 2024 sheet. Add ${recipe.name} from its handout by hand for now.`);
      const next = JSON.parse(JSON.stringify(store));
      next.npc = { ...next.npc, gear: [next.npc?.gear, recipe.name].filter(Boolean).join(', ') };
      findAttr(char, 'store').save({ current: next });
      return { name: char.get('name') };
    }
    if (recipe.drop) {
      await writeAll(char, recipe.drop.attributes);
      const before = char.attribs.length;
      await setAttr(char, recipe.drop.trigger.name, recipe.drop.trigger.value);
      try { char.view?.showDialog?.(); } catch (_) { /* optional */ }
      await sleep(4000);
      if (char.attribs.length <= before + 1 && recipe.rows) await writeAll(char, recipe.rows);
    } else if (recipe.rowsBySheetType) {
      const kind = findAttr(char, 'sheet_type')?.get('current') === 'npc' ? 'npc' : 'pc';
      const spec = recipe.rowsBySheetType[kind];
      const row = `repeating_${spec.section}_${rowId()}_`;
      await writeAll(char, Object.fromEntries(Object.entries(spec.fields).map(([k, v]) => [row + k, v])));
    } else if (recipe.rows) await writeAll(char, recipe.rows);
    return { name: char.get('name') };
  }

  const ROW = '-0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijklmnopqrstuvwxyz';
  function rowId() {
    let t = Date.now(), s = '';
    for (let i = 0; i < 8; i++) { s = ROW.charAt(t % 64) + s; t = Math.floor(t / 64); }
    for (let i = 0; i < 12; i++) s += ROW.charAt(Math.floor(Math.random() * 64));
    return s.replace(/_/g, 'Z');
  }

  /* ── handouts ─────────────────────────────────────── */

  async function createHandouts(plan) {
    const c = campaign();
    if (!c?.handouts?.create) throw new Error('This Roll20 page does not allow in-page imports.');
    const made = [];
    for (const h of plan.handouts || []) {
      const ho = c.handouts.create({ name: h.name, inplayerjournals: '', archived: false });
      for (let i = 0; i < 40 && !ho.id; i++) await sleep(50);
      if (typeof ho.updateBlobs === 'function') ho.updateBlobs({ notes: h.notes || '', gmnotes: h.gmnotes || '' });
      made.push({ id: ho.id, name: h.name, give: h.give || null });
      await sleep(80);
    }
    return { handouts: made };
  }

  /* ── battle map image ─────────────────────────────── */

  /** Drop the image file on the tabletop as if the GM dragged it there. */
  /**
   * Drop the map image onto the tabletop the way a user drag would, then wait
   * for Roll20 to upload it and move the new graphic to the map layer at
   * 70px per square, top-left at the page origin. The page grows to fit.
   */
  async function dropImage({ dataUrl, filename, squares }) {
    const page = window.Campaign?.activePage?.();
    const known = new Set(page?.thegraphics?.models?.map(g => g.id) || []);
    const blob = await (await fetch(dataUrl)).blob();
    const file = new File([blob], filename, { type: blob.type });
    const target = document.querySelector('#editor-wrapper, #finalcanvas, #babylonCanvas, canvas');
    if (!target) throw new Error('Could not find the tabletop.');
    const rect = target.getBoundingClientRect();
    const dt = new DataTransfer();
    dt.items.add(file);
    const at = { clientX: rect.left + rect.width / 2, clientY: rect.top + rect.height / 2, bubbles: true, cancelable: true, dataTransfer: dt };
    for (const type of ['dragenter', 'dragover', 'drop']) target.dispatchEvent(new DragEvent(type, at));
    if (!page?.thegraphics || !Array.isArray(squares)) return { dropped: true, placed: false };
    let graphic = null;
    for (let t = 0; t < 45000 && !graphic; t += 500) {
      await sleep(500);
      graphic = page.thegraphics.models.find(g => !known.has(g.id) && g.get('imgsrc'));
    }
    if (!graphic) return { dropped: true, placed: false };
    const [w, h] = squares.map(n => Math.max(1, Math.round(Number(n) || 1)));
    const width = w * 70, height = h * 70;
    graphic.save({ layer: 'map', width, height, left: width / 2, top: height / 2 });
    const pageSize = { width: Math.max(Number(page.get('width')) || 0, w), height: Math.max(Number(page.get('height')) || 0, h) };
    if (pageSize.width !== Number(page.get('width')) || pageSize.height !== Number(page.get('height'))) page.save(pageSize);
    return { dropped: true, placed: true, graphicId: graphic.id };
  }

  const ops = {
    probe, listCharacters, give, dropImage,
    sendChat: a => sendChat(a.lines || []),
    waitForChat: a => waitForChat(a.text, a.timeoutMs),
    createCharacter: a => createCharacter(a.plan),
    createHandouts: a => createHandouts(a.plan),
  };

  window.addEventListener('message', async ev => {
    const msg = ev.data;
    if (ev.source !== window || !msg || msg.cfer !== 'req' || !ops[msg.op]) return;
    try {
      const result = await ops[msg.op](msg.args);
      window.postMessage({ cfer: 'res', id: msg.id, ok: true, result }, '*');
    } catch (err) {
      window.postMessage({ cfer: 'res', id: msg.id, ok: false, error: err?.message || String(err) }, '*');
    }
  });
})();
