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

  async function createCharacter(plan) {
    const c = campaign();
    if (!c?.characters?.create) throw new Error('This Roll20 page does not allow in-page imports.');
    const char = c.characters.create({ name: plan.character.name, inplayerjournals: '', controlledby: '', archived: false });
    for (let i = 0; i < 40 && !char.attribs; i++) await sleep(50);
    if (!char.attribs) throw new Error('Roll20 did not finish creating the character.');
    if (typeof char.updateBlobs === 'function') char.updateBlobs({ bio: plan.character.bio || '', gmnotes: plan.character.gmnotes || '' });
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
  async function dropImage({ dataUrl, filename }) {
    const blob = await (await fetch(dataUrl)).blob();
    const file = new File([blob], filename, { type: blob.type });
    const target = document.querySelector('#editor-wrapper, #finalcanvas, #babylonCanvas, canvas');
    if (!target) throw new Error('Could not find the tabletop.');
    const rect = target.getBoundingClientRect();
    const dt = new DataTransfer();
    dt.items.add(file);
    const at = { clientX: rect.left + rect.width / 2, clientY: rect.top + rect.height / 2, bubbles: true, cancelable: true, dataTransfer: dt };
    for (const type of ['dragenter', 'dragover', 'drop']) target.dispatchEvent(new DragEvent(type, at));
    return { dropped: true };
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
