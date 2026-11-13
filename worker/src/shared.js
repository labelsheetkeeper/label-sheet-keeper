// Helpers used by every part of the Worker.

// Only the panel's own address may call the Worker from a browser.
export const ALLOWED_ORIGINS = ['https://labelsheetkeeper.github.io'];

export function corsHeaders(origin) {
  const h = { Vary: 'Origin' };
  if (ALLOWED_ORIGINS.includes(origin)) {
    h['Access-Control-Allow-Origin'] = origin;
    h['Access-Control-Allow-Methods'] = 'GET, POST, DELETE, OPTIONS';
    h['Access-Control-Allow-Headers'] = 'Content-Type, Authorization, X-File-Name';
    h['Access-Control-Max-Age'] = '86400';
  }
  return h;
}

export function json(body, status, origin) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', ...corsHeaders(origin) }
  });
}

export function randomId(bytes) {
  const a = new Uint8Array(bytes);
  crypto.getRandomValues(a);
  return btoa(String.fromCharCode(...a)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

export async function sha256(text) {
  const d = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return [...new Uint8Array(d)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

export function sameText(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length) return false;
  let x = 0;
  for (let i = 0; i < a.length; i++) x |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return x === 0;
}

// Counts one more of something for today (UTC) and says how many that makes.
// who: "*" for everyone, or any short label such as a one-day hash of an internet address.
export const COUNT_SQL = 'INSERT INTO usage (day, who, n) VALUES (?, ?, 1) ON CONFLICT (day, who) DO UPDATE SET n = n + 1 RETURNING n';
export const today = () => new Date().toISOString().slice(0, 10);
