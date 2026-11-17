/**
 * Background service worker: the only part of the extension that talks to the
 * Cfer server. It holds the Patreon session key and makes every n8n call, so
 * the Roll20 page never sees the key and CORS never applies.
 */
const N8N_BASE = 'https://foundryrelay.dedicated2.com';
const AUTH_LOGIN = `${N8N_BASE}/webhook/oauth/patreon/login`;
const AUTH_POLL = `${N8N_BASE}/webhook/oauth/patreon/poll`;
const AUTH_VALIDATE = `${N8N_BASE}/webhook/oauth/patreon/validate`;
const USAGE_CHECK = `${N8N_BASE}/webhook/usage-check`;
const CLIENT = 'roll20-extension';

const api = globalThis.browser ?? globalThis.chrome;
const getKey = async () => (await api.storage.local.get('cferKey')).cferKey || '';
const setKey = key => api.storage.local.set({ cferKey: key || '' });

function usageFrom(res, body) {
  const n = v => (v === null || v === undefined || v === '' || !Number.isFinite(Number(v)) ? null : Number(v));
  const src = body?.usage && typeof body.usage === 'object' ? body.usage : body || {};
  let remaining = n(res.headers.get('X-Uses-Remaining')) ?? n(src.usesRemaining ?? src.remaining ?? src.usesLeft);
  const limit = n(res.headers.get('X-Uses-Limit')) ?? n(src.limit ?? src.usesLimit);
  const used = n(res.headers.get('X-Uses-Used')) ?? n(src.used ?? src.usesUsed);
  if (remaining === null && limit !== null && used !== null) remaining = Math.max(0, limit - used);
  return { remaining, limit, tier: src.tier || body?.tier || null, resetAt: src.resetDate || src.resetAt || body?.resetDate || null };
}

async function post(url, payload, { auth = true, timeoutMs = 150000 } = {}) {
  const key = auth ? await getKey() : '';
  if (auth && !key) return { ok: false, status: 401, error: 'Please sign in with Patreon first.' };
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...(auth ? { 'X-Builder-Key': key } : {}), 'X-Foundry-Origin': CLIENT },
      body: JSON.stringify({ ...payload, client: 'roll20' }),
      signal: ctrl.signal,
    });
    const text = await res.text();
    let body = null;
    try { body = JSON.parse(text); } catch (_) { /* XML or empty */ }
    const usage = usageFrom(res, body);
    if (usage.remaining !== null || usage.limit !== null) await api.storage.local.set({ cferUsage: { ...usage, at: Date.now() } });
    if (res.status === 401 || res.status === 403) {
      await setKey('');
      return { ok: false, status: res.status, error: body?.message || 'Your Patreon session expired. Please sign in again.', usage };
    }
    if (res.status === 429) return { ok: false, status: 429, error: body?.message || 'You have used all your uses for this month.', usage };
    if (!res.ok || body?.ok === false) return { ok: false, status: res.status, error: body?.message || body?.error || `The server returned ${res.status}.`, usage };
    return { ok: true, status: res.status, body, text: body ? null : text, usage };
  } catch (err) {
    return { ok: false, status: 0, error: err.name === 'AbortError' ? `The request timed out after ${Math.round(timeoutMs / 1000)} seconds.` : `Could not reach the Cfer server (${err.message}).` };
  } finally { clearTimeout(timer); }
}

async function signIn() {
  const nonce = [...crypto.getRandomValues(new Uint8Array(12))].map(b => b.toString(16).padStart(2, '0')).join('');
  const tab = await api.tabs.create({ url: `${AUTH_LOGIN}?origin=${encodeURIComponent('https://app.roll20.net')}&nonce=${nonce}` });
  const deadline = Date.now() + 5 * 60 * 1000;
  while (Date.now() < deadline) {
    await new Promise(r => setTimeout(r, 2500));
    try {
      const res = await fetch(AUTH_POLL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ nonce }) });
      if (!res.ok) continue;
      const data = await res.json();
      if (data.ok && data.key && String(data.key).length >= 32) {
        await setKey(String(data.key));
        try { await api.tabs.remove(tab.id); } catch (_) { /* user closed it */ }
        await post(USAGE_CHECK, {});
        return { ok: true };
      }
      if (data.error && data.error !== 'not_found') return { ok: false, error: data.error };
    } catch (_) { /* keep polling */ }
  }
  return { ok: false, error: 'Sign-in timed out. Please try again.' };
}

async function status() {
  const key = await getKey();
  const { cferUsage } = await api.storage.local.get('cferUsage');
  return { signedIn: !!key, usage: cferUsage || null };
}

const handlers = {
  status,
  signIn,
  signOut: async () => { await setKey(''); await api.storage.local.remove('cferUsage'); return { ok: true }; },
  validate: async () => post(AUTH_VALIDATE, { validate: true }),
  usage: async () => post(USAGE_CHECK, {}),
  generate: async ({ url, payload, timeoutMs }) => {
    if (!String(url).startsWith(`${N8N_BASE}/webhook/`)) return { ok: false, error: 'Unknown endpoint.' };
    return post(url, payload, { timeoutMs });
  },
  download: async ({ url, filename }) => {
    const id = await api.downloads.download({ url, filename, saveAs: false });
    return { ok: true, id };
  },
  openPatreon: async () => { await api.tabs.create({ url: 'https://www.patreon.com/cw/CelestiaTools' }); return { ok: true }; },
};

api.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  const fn = handlers[msg?.type];
  if (!fn) return false;
  Promise.resolve(fn(msg)).then(sendResponse, err => sendResponse({ ok: false, error: err?.message || String(err) }));
  return true;
});
