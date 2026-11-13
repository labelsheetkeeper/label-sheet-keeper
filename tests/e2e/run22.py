"""Signing in and sync in the panel (Settings › Sync), and the sign-in window.
The Label Sheet Keeper server and Microsoft are stood in for here; the real
server code is checked in tests/worker-test.mjs and tests/sync-test.mjs."""
import subprocess, time, sys, json, os, base64, hashlib
from urllib.parse import urlparse, parse_qs
from playwright.sync_api import sync_playwright

S = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # tests/
REPO = os.path.dirname(S)
OUT = os.path.join(S, 'out')
os.makedirs(OUT, exist_ok=True)
MOCK = open(S + '/mock-office.js').read()
PORT = 8800 + int(os.environ.get('LSK_PORT_SHIFT', '0'))   # shift lets several runs share one machine
BASE = 'http://127.0.0.1:%d' % PORT
API = 'https://api.labelsheetkeeper.workers.dev'
CLIENT = '11111111-2222-3333-4444-555555555555'
CONFIG = "window.LSK_CONFIG = { msClientId: '%s' };" % CLIENT
OFFICE_STUB = "window.SENT = []; window.Office = { onReady: function (f) { setTimeout(f, 0); }, context: { ui: { messageParent: function (m) { window.SENT.push(JSON.parse(m)); } } } };"
results = []

def check(cond, name, detail=''):
    results.append((bool(cond), name, detail))
    print(('PASS ' if cond else 'FAIL ') + name + ('' if cond else '  -- ' + str(detail)))

def b64u(b): return base64.urlsafe_b64encode(b).rstrip(b'=').decode()
def id_token(extra=None):
    """Looks like a Microsoft ID token. The stand-in server doesn't check the signature."""
    claims = {'aud': CLIENT, 'name': 'Pat Example', 'preferred_username': 'pat@example.com'}
    claims.update(extra or {})
    return b64u(b'{"alg":"RS256"}') + '.' + b64u(json.dumps(claims).encode()) + '.' + b64u(b'sig')

class Account:
    """Stands in for the server's account and sync addresses."""
    def __init__(self):
        self.records = {}; self.rev = 0; self.sessions = set(); self.n = 0
        self.auth = []; self.syncs = []; self.signouts = 0; self.deleted = 0; self.down = False; self.made = False
    def put(self, c, rid, v, m=None, d=0):
        self.rev += 1
        self.records[c + '/' + rid] = {'c': c, 'id': rid, 'm': m or int(time.time() * 1000), 'd': d, 'v': None if d else v, 'r': self.rev}
    def handle(self, route):
        req = route.request
        cors = {'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': '*', 'Access-Control-Allow-Methods': '*'}
        if req.method == 'OPTIONS': return route.fulfill(status=204, headers=cors)
        path = urlparse(req.url).path
        say = lambda status, body: route.fulfill(status=status, headers=cors, json=body)
        if self.down: return route.abort()
        if path == '/auth/microsoft':
            self.auth.append(json.loads(req.post_data or '{}'))
            self.n += 1; tok = 'session-%d-%s' % (self.n, 'x' * 40); self.sessions.add(tok)
            new = not self.made; self.made = True
            return say(201, {'session': tok, 'isNew': new})
        tok = (req.headers.get('authorization') or '').replace('Bearer ', '')
        if tok not in self.sessions: return say(401, {'error': 'signed out'})
        if path == '/sync':
            body = json.loads(req.post_data or '{}'); self.syncs.append(body)
            mine = set()
            for p in body.get('push', []):
                k = p['c'] + '/' + p['id']; have = self.records.get(k)
                if not have or p['m'] > have['m']:
                    self.rev += 1; self.records[k] = {'c': p['c'], 'id': p['id'], 'm': p['m'], 'd': p.get('d', 0), 'v': p.get('v'), 'r': self.rev}; mine.add(k)
            pull = [r for k, r in sorted(self.records.items(), key=lambda kv: kv[1]['r']) if r['r'] > body.get('since', 0) and k not in mine]
            return say(200, {'pull': pull, 'next': self.rev, 'more': False})
        if path == '/signout':
            self.signouts += 1; self.sessions.discard(tok); return say(200, {'ok': True})
        if path == '/account' and req.method == 'DELETE':
            self.deleted += 1; self.records = {}; self.sessions = set(); self.made = False; return say(200, {'ok': True})
        return say(404, {'error': 'not found'})

srv = subprocess.Popen(['python3', '-m', 'http.server', str(PORT), '--bind', '127.0.0.1'], cwd=REPO, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
time.sleep(1)

def wait(pg, js, timeout=7000, arg=None):
    try:
        pg.wait_for_function(js, arg=arg, timeout=timeout); return True
    except Exception:
        return False
def say(pg): return pg.inner_text('#say')
def status(pg): return pg.inner_text('#sync-status')
def book(pg): return pg.evaluate("JSON.parse(localStorage.getItem('lsk.addressBook') || '[]').map(e => e.name)")
def open_settings(pg): pg.evaluate("document.getElementById('settings').open = true")
def panel(ctx, acct, configured=True, query=''):
    pg = ctx.new_page()
    errs = []
    pg.on('pageerror', lambda e: errs.append(str(e)))
    pg.route('**/office.js', lambda r: r.fulfill(body=MOCK, content_type='application/javascript'))
    if configured: pg.route('**/lsk-config.js*', lambda r: r.fulfill(body=CONFIG, content_type='application/javascript'))
    pg.goto(BASE + '/taskpane.html' + query)
    wait(pg, "window.LSK_TEST && LSK_TEST.st().kind !== 'loading'", 9000)
    return pg, errs
def sign_in(pg, extra=None):
    pg.click('#sync-on')
    wait(pg, "MOCK.dialog && /auth\\.html/.test(MOCK.dialog.url) && !MOCK.dialog.closed")
    pg.evaluate("t => MOCK.dialog.send({ type: 'signin', idToken: t })", id_token(extra))

try:
  with sync_playwright() as p:
    b = p.chromium.launch()

    # ================= Not set up: nothing shows, nothing is sent =================
    ctx0 = b.new_context(viewport={'width': 340, 'height': 1000})
    acct0 = Account(); ctx0.route(API + '/**', acct0.handle)
    pg, errs = panel(ctx0, acct0, configured=False)
    open_settings(pg)
    check(pg.is_hidden('#sync-block') and pg.evaluate("LSK_TEST.acct()") is None, 'until the app is registered with Microsoft, Settings shows no Sync section')
    check(not acct0.auth and not acct0.syncs, 'and nothing is sent to the server')
    check(not errs, 'no script errors without sign-in set up', errs)
    ctx0.close()

    # ================= Turning on sync =================
    ctx = b.new_context(viewport={'width': 340, 'height': 1000})
    acct = Account(); ctx.route(API + '/**', acct.handle)
    ctx.add_init_script("""if (!localStorage.getItem('seeded')) { localStorage.setItem('seeded', '1');
      localStorage.setItem('lsk.addressBook', JSON.stringify([{ id: 'a1', name: 'Ann Lee', lines: ['Ann Lee', '8 Maple Way', 'Boston, MA 02108'] }]));
      localStorage.setItem('lsk.geoKey', 'my-own-key'); }""")
    pg, errs = panel(ctx, acct)
    open_settings(pg)
    check(pg.is_visible('#sync-block') and pg.is_visible('#sync-on') and pg.is_hidden('#sync-now') and pg.is_hidden('#sync-off') and pg.is_hidden('#sync-delete'), 'with sign-in set up, Settings has a Sync section with Turn on sync')
    check('Sign in with Microsoft to keep your address book and settings the same on every computer' in status(pg), 'it says what signing in is for before asking', status(pg))
    check(not acct.auth and not acct.syncs, 'nothing is sent before signing in')
    pg.locator('#sync-block').screenshot(path=OUT + '/s22-sync-off.png')

    pg.click('#sync-on')
    check(wait(pg, "MOCK.dialog && /\\/auth\\.html\\?t=\\d+$/.test(MOCK.dialog.url)"), 'Turn on sync opens the sign-in window', pg.evaluate("MOCK.dialog && MOCK.dialog.url"))
    pg.evaluate("MOCK.dialog.handlers.evt({ error: 12006 })")
    pg.wait_for_timeout(300)
    check(pg.is_enabled('#sync-on') and pg.is_visible('#sync-on') and 'warn' not in (pg.get_attribute('#say', 'class') or ''), 'closing that window without signing in just leaves sync off')
    pg.click('#sync-on')
    wait(pg, "MOCK.dialog && !MOCK.dialog.closed")
    pg.evaluate("MOCK.dialog.send({ type: 'signin-error', error: 'state' })")
    check(wait(pg, "/didn’t finish/.test(document.getElementById('say').textContent)") and pg.evaluate("MOCK.dialog.closed"), 'a sign-in that goes wrong says so and closes the window', say(pg))

    sign_in(pg)
    check(wait(pg, "/Syncing as pat@example\\.com\\. Up to date at/.test(document.getElementById('sync-status').textContent)"), 'after signing in it says who is syncing and when it was last up to date', status(pg))
    check(pg.evaluate("MOCK.dialog.closed") and len(acct.auth) == 1 and acct.auth[0].get('id_token', '').count('.') == 2, 'the sign-in window closes and the server gets Microsoft’s proof of who signed in')
    check('book/a1' in acct.records and json.loads(acct.records['book/a1']['v'])['name'] == 'Ann Lee', 'this Mac’s address book goes to the account', list(acct.records))
    check(not any('geoKey' in k for k in acct.records) and 'my-own-key' not in json.dumps(acct.records), 'the Geoapify key isn’t sent')
    check('Sync is on. This Mac’s address book and settings are now in your account.' in say(pg), 'and the panel says sync is on', say(pg))
    check(pg.is_hidden('#sync-on') and pg.is_visible('#sync-now') and pg.is_visible('#sync-off') and pg.is_visible('#sync-delete'), 'the section now offers Sync now, Turn off and Delete my account')
    pg.locator('#sync-block').screenshot(path=OUT + '/s22-sync-on.png')
    keys = pg.evaluate("Object.keys(LSK_TEST.backupKeys())")
    check(not any(k.startswith('lsk.acct') or k.startswith('lsk.auth') for k in keys) and 'lsk.addressBook' in keys, 'the sign-in is never put in a backup', keys)

    # ================= Changes go both ways =================
    n = len(acct.syncs)
    pg.uncheck('#show-saved')
    check(wait(pg, "true") and (lambda: [pg.wait_for_timeout(200) for _ in range(20) if 'set/showSaved' not in acct.records] and 'set/showSaved' in acct.records)(), 'a setting changed here reaches the account a moment later, by itself', list(acct.records))
    check(acct.records.get('set/showSaved', {}).get('v') == 'off', 'with its new value', acct.records.get('set/showSaved'))
    acct.put('book', 'r1', json.dumps({'id': 'r1', 'name': 'Rae Other', 'lines': ['Rae Other', '2 Far Rd', 'Salem, MA 01970']}))
    acct.put('set', 'homeCountry', 'CA')
    pg.click('#sync-now')
    check(wait(pg, "JSON.parse(localStorage.getItem('lsk.addressBook')).length === 2"), 'Sync now brings in an address added on another computer', book(pg))
    check(wait(pg, "document.getElementById('home-country').value === 'CA'"), 'and a setting changed there shows in Settings without reopening the panel', pg.input_value('#home-country'))
    check(wait(pg, "/Synced\\./.test(document.getElementById('say').textContent)"), 'it says Synced', say(pg))
    pg.click('#fill-tabs button:has-text("Address")')
    opts = pg.evaluate("Array.from(document.querySelectorAll('#addr-pick option')).map(o => o.textContent)")
    check(any('Rae Other' in o for o in opts), 'the new address is in the address list', opts)

    # Another Word window on the same Mac is already signed in.
    pg2, errs2 = panel(ctx, acct)
    open_settings(pg2)
    check(wait(pg2, "/Syncing as pat@example\\.com/.test(document.getElementById('sync-status').textContent)") and pg2.is_hidden('#sync-on'), 'another Word window on this Mac is signed in already', status(pg2))
    check(len(acct.auth) == 1, 'without signing in again')
    pg2.close()

    # ================= Trouble =================
    acct.down = True
    pg.click('#sync-now')
    check(wait(pg, "/couldn’t be reached/.test(document.getElementById('sync-status').textContent) && /warn/.test(document.getElementById('sync-status').className)"), 'when the server can’t be reached, the section says so', status(pg))
    check(book(pg) == ['Ann Lee', 'Rae Other'], 'and the address book is still here')
    acct.down = False
    pg.click('#sync-now')
    check(wait(pg, "/Up to date at/.test(document.getElementById('sync-status').textContent)"), 'and syncing recovers once it can', status(pg))

    # ================= Turning it off =================
    pg.click('#sync-off')
    check(wait(pg, "!document.getElementById('confirm').hidden") and 'stay here' in pg.inner_text('#confirm-msg'), 'Turn off asks first, and says the addresses stay', pg.inner_text('#confirm-msg'))
    pg.click('#confirm-yes')
    check(wait(pg, "/Sync is off on this Mac/.test(document.getElementById('say').textContent)") and acct.signouts == 1, 'turning off tells the server to forget this Mac')
    check(pg.is_visible('#sync-on') and pg.is_hidden('#sync-now') and pg.evaluate("localStorage.getItem('lsk.acct')") is None, 'the section is back to Turn on sync')
    check(book(pg) == ['Ann Lee', 'Rae Other'] and 'book/r1' in acct.records, 'the addresses stay on this Mac and in the account')

    # ================= Signed out by the server; deleting the account =================
    sign_in(pg)
    check(wait(pg, "/Syncing as pat@example\\.com\\. Up to date/.test(document.getElementById('sync-status').textContent)"), 'signing in again works', status(pg))
    check('Your account’s address book and settings are here now' in say(pg), 'joining an account that already has addresses says they’re here now', say(pg))
    acct.sessions.clear()
    pg.click('#sync-now')
    check(wait(pg, "/You were signed out/.test(document.getElementById('sync-status').textContent) && !document.getElementById('sync-on').hidden"), 'a Mac the server has signed out says so and offers Turn on sync', status(pg))
    sign_in(pg)
    wait(pg, "/Up to date/.test(document.getElementById('sync-status').textContent)")
    pg.click('#sync-delete')
    check(wait(pg, "!document.getElementById('confirm').hidden") and 'everything in it' in pg.inner_text('#confirm-msg') and 'stay on this Mac' in pg.inner_text('#confirm-msg'), 'Delete my account asks first and says what stays', pg.inner_text('#confirm-msg'))
    pg.click('#confirm-yes')
    check(wait(pg, "/deleted from the server/.test(document.getElementById('say').textContent)") and acct.deleted == 1 and not acct.records, 'deleting removes everything from the server', say(pg))
    check(pg.is_visible('#sync-on') and book(pg) == ['Ann Lee', 'Rae Other'], 'and leaves this Mac’s addresses alone')

    # On a document that isn't a sheet, Sync sits with Backup.
    pg3, errs3 = panel(ctx, acct, query='?blank=1')
    pg3.evaluate("document.getElementById('backup-none').open = true")
    check(pg3.is_visible('#sync-block') and pg3.inner_text('#backup-none-title') == 'Sync, back up or restore settings', 'in a document that isn’t a sheet, Sync is offered next to Backup', pg3.inner_text('#backup-none-title'))
    pg3.close()
    check(not errs, 'no script errors in the panel', errs)

    # ================= The sign-in window =================
    def auth_page(config=CONFIG, at_microsoft=None, token=None):
        pv = ctx.new_page()
        seen = {'authorize': None, 'token': None}
        pv.route('**/office.js', lambda r: r.fulfill(body=OFFICE_STUB, content_type='application/javascript'))
        pv.route('**/lsk-config.js*', lambda r: r.fulfill(body=config, content_type='application/javascript'))
        def authorize(route):
            q = {k: v[0] for k, v in parse_qs(urlparse(route.request.url).query).items()}
            seen['authorize'] = q
            back = at_microsoft(q) if at_microsoft else q['redirect_uri'] + '?code=one-time-code&state=' + q['state']
            route.fulfill(status=302, headers={'Location': back})
        def tok(route):
            form = {k: v[0] for k, v in parse_qs(route.request.post_data or '').items()}
            seen['token'] = form
            body = token(seen['authorize'], form) if token else {'id_token': id_token({'nonce': seen['authorize']['nonce']})}
            route.fulfill(headers={'Access-Control-Allow-Origin': '*'}, json=body)
        pv.route('https://login.microsoftonline.com/common/oauth2/v2.0/authorize*', authorize)
        pv.route('https://login.microsoftonline.com/common/oauth2/v2.0/token', tok)
        pv.goto(BASE + '/auth.html?t=123')
        wait(pv, "window.SENT && window.SENT.length > 0", 6000)
        return pv, seen

    pv, seen = auth_page()
    a, t = seen['authorize'], seen['token']
    sent = pv.evaluate("window.SENT")
    check(a and a['client_id'] == CLIENT and a['response_type'] == 'code' and a['redirect_uri'] == BASE + '/auth.html' and set(a['scope'].split()) == {'openid', 'profile', 'email'}, 'the sign-in window goes to Microsoft for this app, asking only who you are', a)
    check(a and a['code_challenge_method'] == 'S256' and a['prompt'] == 'select_account' and len(a['state']) >= 16 and len(a['nonce']) >= 16, 'with the protections for apps that hold no secret', a)
    check(t and t['grant_type'] == 'authorization_code' and t['code'] == 'one-time-code' and t['redirect_uri'] == BASE + '/auth.html' and 'client_secret' not in t, 'it trades Microsoft’s one-time code for the proof, with no secret', t)
    check(t and b64u(hashlib.sha256(t['code_verifier'].encode()).digest()) == a['code_challenge'], 'only the window that started the sign-in can use the code')
    check(len(sent) == 1 and sent[0]['type'] == 'signin' and sent[0]['idToken'].count('.') == 2, 'and hands the proof to the panel', sent)
    check(pv.evaluate("sessionStorage.getItem('lsk.auth.tmp') === null && localStorage.getItem('lsk.auth.tmp') === null"), 'nothing from the sign-in is left behind in the window')
    check('Signed in' in pv.inner_text('#msg'), 'the window says it worked', pv.inner_text('#msg'))
    pv.close()

    pv, seen = auth_page(at_microsoft=lambda q: q['redirect_uri'] + '?code=one-time-code&state=someone-elses')
    check(pv.evaluate("window.SENT")[0] == {'type': 'signin-error', 'error': 'state'} and seen['token'] is None, 'an answer that doesn’t belong to this sign-in is refused, and the code isn’t used', pv.evaluate("window.SENT"))
    pv.close()
    pv, seen = auth_page(at_microsoft=lambda q: q['redirect_uri'] + '?error=access_denied&error_description=The+user+cancelled&state=' + q['state'])
    check(pv.evaluate("window.SENT")[0] == {'type': 'signin-error', 'error': 'cancelled'} and 'didn’t sign in' in pv.inner_text('#msg'), 'cancelling at Microsoft is passed on as cancelled', pv.evaluate("window.SENT"))
    pv.close()
    pv, seen = auth_page(token=lambda a, f: {'id_token': id_token({'nonce': 'not-the-one-sent'})})
    check(pv.evaluate("window.SENT")[0]['type'] == 'signin-error', 'proof that doesn’t match this sign-in is refused', pv.evaluate("window.SENT"))
    pv.close()
    pv, seen = auth_page(token=lambda a, f: {'error': 'invalid_grant'})
    check(pv.evaluate("window.SENT")[0] == {'type': 'signin-error', 'error': 'invalid_grant'}, 'Microsoft refusing the code is passed on', pv.evaluate("window.SENT"))
    pv.close()
    pv, seen = auth_page(config="window.LSK_CONFIG = { msClientId: '' };")
    check(pv.evaluate("window.SENT")[0] == {'type': 'signin-error', 'error': 'not set up'} and seen['authorize'] is None and 'isn’t set up' in pv.inner_text('#msg'), 'without the app’s Microsoft ID the window says sign-in isn’t set up and goes nowhere', pv.evaluate("window.SENT"))
    pv.close()
    b.close()
finally:
    srv.kill()

failed = [r for r in results if not r[0]]
print('%d checks, %d failed' % (len(results), len(failed)))
for f in failed:
    print('::error title=run22::' + f[1])
sys.exit(1 if failed else 0)
