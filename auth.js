/*
  The sign-in window. The panel opens this page in a Word dialog; it sends the
  window to Microsoft's sign-in page, Microsoft sends it back here with a
  one-time code, and this page trades the code for an ID token (proof of who
  signed in) and hands that to the panel. No password ever reaches this page.

  This is the standard OAuth "authorization code with PKCE" flow for pages
  with no server secret: only the window that started the sign-in holds the
  verifier needed to use the code.
*/
(function () {
  'use strict';
  var CLIENT = (window.LSK_CONFIG && window.LSK_CONFIG.msClientId) || '';
  var MS = 'https://login.microsoftonline.com/common/oauth2/v2.0';
  var HERE = location.origin + location.pathname;      // exactly as registered with Microsoft
  var KEEP = 'lsk.auth.tmp';
  var SCOPE = 'openid profile email';
  var q = new URLSearchParams(location.search);

  function say(t) { var m = document.getElementById('msg'); if (m) m.textContent = t; }
  function b64url(bytes) { return window.btoa(String.fromCharCode.apply(null, bytes)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, ''); }
  function random(n) { var a = new Uint8Array(n); window.crypto.getRandomValues(a); return b64url(a); }
  function challengeOf(verifier) {
    return window.crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)).then(function (d) { return b64url(new Uint8Array(d)); });
  }
  // Kept only while the window is at Microsoft's page.
  function keep(v) {
    var t = v ? JSON.stringify(v) : null;
    [window.sessionStorage, window.localStorage].forEach(function (s) { try { if (t) s.setItem(KEEP, t); else s.removeItem(KEEP); } catch (e) { /* blocked */ } });
  }
  function kept() {
    var out = null;
    [window.sessionStorage, window.localStorage].forEach(function (s) { try { if (!out) out = JSON.parse(s.getItem(KEEP) || 'null'); } catch (e) { /* blocked */ } });
    return out && Date.now() - out.at < 15 * 60 * 1000 ? out : null;
  }
  function tell(msg) {
    var send = function () { try { Office.context.ui.messageParent(JSON.stringify(msg)); } catch (e) { say('Close this window and try again from the panel.'); } };
    if (window.Office && Office.onReady) Office.onReady(send); else send();
  }
  function fail(why, shown) { say(shown || 'Signing in didn’t finish. Close this window and try again.'); tell({ type: 'signin-error', error: why }); }
  function claims(token) {
    try { return JSON.parse(decodeURIComponent(escape(window.atob(token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/'))))); } catch (e) { return null; }
  }

  function begin() {
    if (!CLIENT) { fail('not set up', 'Signing in isn’t set up yet.'); return; }
    var v = { verifier: random(48), state: random(16), nonce: random(16), at: Date.now() };
    keep(v);
    // Office finishes getting ready first, so it still knows this window when Microsoft sends it back.
    var ready = new Promise(function (resolve) { if (window.Office && Office.onReady) Office.onReady(resolve); else resolve(); });
    Promise.all([challengeOf(v.verifier), ready]).then(function (got) {
      location.replace(MS + '/authorize?' + new URLSearchParams({
        client_id: CLIENT, response_type: 'code', redirect_uri: HERE, response_mode: 'query', scope: SCOPE,
        state: v.state, nonce: v.nonce, code_challenge: got[0], code_challenge_method: 'S256', prompt: 'select_account'
      }).toString());
    }, function () { fail('crypto'); });
  }

  function finish() {
    var v = kept();
    keep(null);
    if (q.get('error')) {
      var cancelled = q.get('error') === 'access_denied' || /AADSTS65004|user_cancel/i.test(q.get('error_description') || q.get('error_subcode') || '');
      fail(cancelled ? 'cancelled' : q.get('error'), cancelled ? 'You didn’t sign in. Close this window.' : undefined);
      return;
    }
    if (!v || v.state !== q.get('state')) { fail('state'); return; }
    say('Signing you in…');
    fetch(MS + '/token', {
      method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ client_id: CLIENT, grant_type: 'authorization_code', code: q.get('code'), redirect_uri: HERE, code_verifier: v.verifier, scope: SCOPE }).toString()
    }).then(function (r) { return r.json(); }).then(function (j) {
      var c = j && j.id_token ? claims(j.id_token) : null;
      if (!c || c.nonce !== v.nonce) { fail((j && j.error) || 'token'); return; }
      say('Signed in. This window will close.');
      tell({ type: 'signin', idToken: j.id_token });
    }).catch(function () { fail('network', 'Microsoft couldn’t be reached. Close this window and try again.'); });
  }

  function go() { if (q.get('code') || q.get('error')) finish(); else begin(); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', go); else go();
})();
