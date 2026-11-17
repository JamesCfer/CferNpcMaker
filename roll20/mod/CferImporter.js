/* Cfer Importer: companion Mod (API) script for the "Cfer Tools for Roll20" browser extension.
 *
 * The extension generates NPCs, items, stores, settlements and battle maps with
 * the Cfer AI tools and sends the result to this script in chat as
 *   !cfer-import <id> <part>/<total> <base64url data>
 * This script rebuilds the import plan and creates the characters, handouts,
 * walls, doors and lights it describes. It makes no network calls and holds no
 * paid logic: generation and Patreon sign-in happen in the extension.
 *
 * Commands (GM only):
 *   !cfer-help                      show help
 *   !cfer-give <key>                add an imported item to the selected tokens' characters
 *   !cfer-import ...                sent by the extension; not typed by hand
 *
 * License: MIT. https://github.com/JamesCfer/CferNpcMaker
 */
/* global on, log, sendChat, createObj, findObjs, getObj, playerIsGM, state */
var CferImporter = CferImporter || (function () {
  'use strict';

  var VERSION = '1.0.0';
  var SCRIPT = 'Cfer Importer';
  var B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
  var pending = {};

  /* ── helpers ─────────────────────────────────────────── */

  function decodeBase64Url(s) {
    var bytes = [], buf = 0, bits = 0, i, v;
    for (i = 0; i < s.length; i++) {
      v = B64.indexOf(s.charAt(i));
      if (v < 0) continue;
      buf = ((buf << 6) | v) & 0xffffff; bits += 6;
      if (bits >= 8) { bits -= 8; bytes.push((buf >> bits) & 255); }
    }
    // UTF-8 decode
    var out = '', b, c, j = 0;
    while (j < bytes.length) {
      b = bytes[j++];
      if (b < 0x80) c = b;
      else if (b < 0xe0) c = ((b & 0x1f) << 6) | (bytes[j++] & 0x3f);
      else if (b < 0xf0) c = ((b & 0x0f) << 12) | ((bytes[j++] & 0x3f) << 6) | (bytes[j++] & 0x3f);
      else c = ((b & 0x07) << 18) | ((bytes[j++] & 0x3f) << 12) | ((bytes[j++] & 0x3f) << 6) | (bytes[j++] & 0x3f);
      if (c > 0xffff) { c -= 0x10000; out += String.fromCharCode(0xd800 + (c >> 10), 0xdc00 + (c & 0x3ff)); }
      else out += String.fromCharCode(c);
    }
    return out;
  }

  function whisper(text) { sendChat(SCRIPT, '/w gm ' + text, null, { noarchive: true }); }

  function esc(s) { return String(s === undefined || s === null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  function store() {
    state.CferImporter = state.CferImporter || { version: VERSION, gives: {} };
    state.CferImporter.gives = state.CferImporter.gives || {};
    return state.CferImporter;
  }

  function setAttr(charId, name, value, withWorker) {
    var current = value, max;
    if (value && typeof value === 'object') { current = value.current; max = value.max; }
    var attr = findObjs({ type: 'attribute', characterid: charId, name: name })[0];
    if (!attr) attr = createObj('attribute', { characterid: charId, name: name, current: withWorker ? '' : String(current === undefined ? '' : current) });
    var update = { current: String(current === undefined ? '' : current) };
    if (max !== undefined) update.max = String(max);
    if (withWorker) attr.setWithWorker(update); else attr.set(update);
    return attr;
  }

  function attrValue(charId, name) {
    var attr = findObjs({ type: 'attribute', characterid: charId, name: name })[0];
    return attr ? attr.get('current') : '';
  }

  function writeAll(charId, attrs, withWorker) {
    var names = Object.keys(attrs || {});
    for (var i = 0; i < names.length; i++) setAttr(charId, names[i], attrs[names[i]], withWorker);
    return names.length;
  }

  /* ── characters ──────────────────────────────────────── */

  function applyCharacter(plan, done) {
    var c = createObj('character', { name: plan.character.name, inplayerjournals: '', controlledby: '', archived: false });
    if (!c) return done('Could not create the character.');
    if (plan.character.bio) c.set('bio', plan.character.bio);
    if (plan.character.gmnotes) c.set('gmnotes', plan.character.gmnotes);
    var id = c.id;
    writeAll(id, plan.attributes, false);
    if (!plan.drop) return done('Created ' + characterLink(c, plan));
    writeAll(id, plan.drop.attributes, false);
    setAttr(id, plan.drop.trigger.name, plan.drop.trigger.value, true);
    setTimeout(function () {
      var built = attrValue(id, plan.drop.verify);
      if (!built && plan.fallback) {
        writeAll(id, plan.fallback.attributes, true);
        return done('Created ' + characterLink(c, plan) + ' (written directly; the sheet\'s own import did not run).');
      }
      done('Created ' + characterLink(c, plan));
    }, plan.drop.timeoutMs || 5000);
  }

  function characterLink(c, plan) {
    return '<a href="http://journal.roll20.net/character/' + c.id + '">' + esc(plan.character.name) + '</a>' + (plan.summary ? ' ' + esc(plan.summary.replace(plan.character.name, '')) : '');
  }

  /* ── handouts and give menus ─────────────────────────── */

  function applyHandouts(plan, done) {
    var lines = [], gives = [];
    (plan.handouts || []).forEach(function (h) {
      var o = createObj('handout', { name: h.name, inplayerjournals: '', archived: false });
      if (!o) return;
      if (h.notes) o.set('notes', h.notes);
      if (h.gmnotes) o.set('gmnotes', h.gmnotes);
      lines.push('<a href="http://journal.roll20.net/handout/' + o.id + '">' + esc(h.name) + '</a>');
      if (h.give) {
        var key = o.id;
        store().gives[key] = h.give;
        gives.push('[' + esc(h.give.name) + '](!cfer-give ' + key + ')');
      }
    });
    var text = 'Created ' + lines.join(', ');
    if (gives.length) text += '<br>Select a token, then click to give: ' + gives.join(' ');
    done(text);
  }

  function give(key, selected) {
    var recipe = store().gives[key];
    if (!recipe) return whisper('That item is no longer available to give. Import it again.');
    var chars = (selected || []).map(function (s) { var g = getObj('graphic', s._id); return g && g.get('represents'); }).filter(Boolean);
    if (!chars.length) return whisper('Select one or more tokens linked to characters first.');
    chars.forEach(function (charId) {
      if (recipe.drop) {
        writeAll(charId, recipe.drop.attributes, false);
        var before = findObjs({ type: 'attribute', characterid: charId }).length;
        setAttr(charId, recipe.drop.trigger.name, recipe.drop.trigger.value, true);
        setTimeout(function () {
          if (findObjs({ type: 'attribute', characterid: charId }).length <= before + 1 && recipe.rows) writeAll(charId, recipe.rows, true);
        }, 4000);
      } else if (recipe.rowsBySheetType) {
        var kind = attrValue(charId, 'sheet_type') === 'npc' ? 'npc' : 'pc';
        var spec = recipe.rowsBySheetType[kind];
        var row = 'repeating_' + spec.section + '_' + newRowId() + '_';
        var attrs = {};
        Object.keys(spec.fields).forEach(function (k) { attrs[row + k] = spec.fields[k]; });
        writeAll(charId, attrs, true);
      } else if (recipe.rows) writeAll(charId, recipe.rows, true);
    });
    whisper('Gave ' + esc(recipe.name) + ' to ' + chars.length + ' character' + (chars.length === 1 ? '' : 's') + '.');
  }

  var ROW = '-0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijklmnopqrstuvwxyz';
  function newRowId() {
    var t = Date.now(), s = '', i;
    for (i = 0; i < 8; i++) { s = ROW.charAt(t % 64) + s; t = Math.floor(t / 64); }
    for (i = 0; i < 12; i++) s += ROW.charAt(Math.floor(Math.random() * 64));
    return s.replace(/_/g, 'Z');
  }

  /* ── battle maps ─────────────────────────────────────── */

  var GRID = 70;

  function applyMap(plan, selected, done) {
    var m = plan.map;
    var graphic = (selected || []).map(function (s) { return getObj('graphic', s._id); }).filter(Boolean)[0];
    if (!graphic) return done('Select the map image on the page first, then press "Add walls and lights" again.');
    var pageId = graphic.get('pageid');
    var page = getObj('page', pageId);
    var w = m.squares[0], h = m.squares[1];
    if (page) page.set({ width: w, height: h, scale_number: m.grid.distance, scale_units: m.grid.units, snapping_increment: 1,
      dynamic_lighting_enabled: true, explorer_mode: 'basic', daylight_mode_enabled: !m.darkness, daylightModeOpacity: m.darkness ? 0 : 1 });
    graphic.set({ layer: 'map', left: (w * GRID) / 2, top: (h * GRID) / 2, width: w * GRID, height: h * GRID, rotation: 0 });

    var walls = 0, doors = 0, windows = 0, lights = 0;
    (m.walls || []).forEach(function (s) {
      var x1 = s.c[0], y1 = s.c[1], x2 = s.c[2], y2 = s.c[3];
      if (s.kind === 'door' || s.kind === 'window') {
        // Doors and windows use an inverted y axis, with handles relative to the centre.
        var cx = (x1 + x2) / 2, cy = (y1 + y2) / 2, made = null;
        try {
          made = createObj(s.kind, { pageid: pageId, x: cx, y: -cy, isOpen: false, isLocked: false, isSecret: false,
            path: { handle0: { x: x1 - cx, y: -(y1 - cy) }, handle1: { x: x2 - cx, y: -(y2 - cy) } } });
        } catch (e) { made = null; }
        if (made) { if (s.kind === 'door') doors++; else windows++; return; }
        if (s.kind === 'window') return; // a window that could not be made stays open rather than blocking sight
      }
      var left = Math.min(x1, x2), top = Math.min(y1, y2);
      var pw = Math.max(1, Math.abs(x2 - x1)), ph = Math.max(1, Math.abs(y2 - y1));
      var path = createObj('path', { pageid: pageId, layer: 'walls', stroke: '#ff0000', stroke_width: 3, fill: 'transparent',
        path: JSON.stringify([['M', x1 - left, y1 - top], ['L', x2 - left, y2 - top]]),
        left: left + pw / 2, top: top + ph / 2, width: pw, height: ph });
      if (path) walls++;
    });
    (m.lights || []).forEach(function (l) {
      var o = createObj('graphic', { pageid: pageId, layer: 'walls', imgsrc: graphic.get('imgsrc'), name: 'Light: ' + l.kind,
        left: l.x, top: l.y, width: 14, height: 14, emits_bright_light: true, bright_light_distance: l.bright,
        emits_low_light: true, low_light_distance: l.dim, lightColor: l.color, has_bright_light_vision: false });
      if (o) lights++;
    });
    done('Set up ' + esc(m.name) + ': ' + w + ' x ' + h + ' squares, ' + walls + ' walls, ' + doors + ' doors, ' + windows + ' windows, ' + lights + ' lights.');
  }

  /* ── chat ────────────────────────────────────────────── */

  function applyPlan(plan, msg) {
    plan._playerid = msg.playerid;
    var done = function (text) { whisper(text); };
    if (plan.v !== 1) return done('This import was made by a newer version of the extension. Update the Cfer Importer script.');
    if (plan.kind === 'character') return applyCharacter(plan, done);
    if (plan.kind === 'handouts') return applyHandouts(plan, done);
    if (plan.kind === 'map') return applyMap(plan, msg.selected, done);
    done('Unknown import type: ' + esc(plan.kind));
  }

  function handleImport(msg, args) {
    var id = args[1], part = (args[2] || '').split('/'), data = args.slice(3).join('');
    var n = parseInt(part[0], 10), total = parseInt(part[1], 10);
    if (!id || !n || !total || n > total) return whisper('Malformed import message.');
    var p = pending[id] = pending[id] || { parts: [], got: 0, at: Date.now() };
    if (!p.parts[n - 1]) { p.parts[n - 1] = data; p.got++; }
    if (p.got < total) return;
    delete pending[id];
    var plan;
    try { plan = JSON.parse(decodeBase64Url(p.parts.join(''))); }
    catch (e) { return whisper('Could not read the import (' + esc(e.message) + ').'); }
    try { applyPlan(plan, msg); }
    catch (e2) { log(SCRIPT + ': ' + e2.stack); whisper('Import failed: ' + esc(e2.message)); }
  }

  function help() {
    whisper('<b>Cfer Importer ' + VERSION + '</b><br>Imports NPCs, items, stores, settlements and battle maps made with the '
      + 'Cfer Tools for Roll20 browser extension. Install the extension, open this game, and use the Cfer panel.<br>'
      + '<code>!cfer-give &lt;key&gt;</code> gives an imported item to the selected tokens.');
  }

  function onChat(msg) {
    if (msg.type !== 'api' || msg.content.indexOf('!cfer') !== 0) return;
    if (!playerIsGM(msg.playerid)) return;
    var args = msg.content.trim().split(/\s+/);
    if (args[0] === '!cfer-import') return handleImport(msg, args);
    if (args[0] === '!cfer-give') return give(args[1], msg.selected);
    if (args[0] === '!cfer-help') return help();
    if (args[0] === '!cfer-ping') return whisper('Cfer Importer ' + VERSION + ' is installed.');
  }

  on('ready', function () {
    store();
    on('chat:message', onChat);
    log(SCRIPT + ' ' + VERSION + ' ready');
  });

  return { VERSION: VERSION, decodeBase64Url: decodeBase64Url, applyPlan: applyPlan, give: give, onChat: onChat };
}());
