import subprocess, time, sys, json, os
from urllib.parse import unquote
from playwright.sync_api import sync_playwright

S = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # tests/
REPO = os.path.dirname(S)
OUT = os.path.join(S, 'out')
os.makedirs(OUT, exist_ok=True)
MOCK = open(S + '/mock-office.js').read()
PORT = 8797 + int(os.environ.get('LSK_PORT_SHIFT', '0'))   # shift lets several runs share one machine
API = 'https://api.labelsheetkeeper.workers.dev'
PDF = S + '/fixtures/sample-return-label.pdf'
PNG = S + '/fixtures/mark-ship-good.png'
results = []

def check(cond, name, detail=''):
    results.append((bool(cond), name, detail))
    print(('PASS ' if cond else 'FAIL ') + name + ('' if cond else '  -- ' + str(detail)))

srv = subprocess.Popen(['python3', '-m', 'http.server', str(PORT), '--bind', '127.0.0.1'], cwd=REPO, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
time.sleep(1)

CORS = {'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-File-Name', 'Access-Control-Allow-Methods': 'GET, POST, DELETE, OPTIONS'}

class FakeWorker:
    """Stands in for the Cloudflare Worker (worker/src/index.js) and records each request."""
    def __init__(self, box=None, key=None):
        self.box, self.key, self.items, self.log, self.n = box, key, [], [], 0
        self.unlinked = False
        self.offline = False

    def add(self, path, type_, name):
        data = open(path, 'rb').read()
        self.n += 1
        self.items.append({'id': 'item%012d' % self.n, 'type': type_, 'name': name, 'size': len(data), 'at': int(time.time() * 1000), 'bytes': data})

    def handle(self, route):
        req = route.request
        path = req.url[len(API):].split('?')[0]
        parts = [p for p in path.split('/') if p]
        self.log.append((req.method, path, req.headers.get('authorization', '')))
        if self.offline: return route.abort()
        if req.method == 'OPTIONS': return route.fulfill(status=204, headers=CORS)
        if parts == ['pair'] and req.method == 'POST':
            self.box, self.key, self.unlinked = 'Bx' * 11, 'Kx' * 16, False
            return route.fulfill(status=201, headers=CORS, json={'box': self.box, 'key': self.key, 'send': API + '/drop/' + self.box})
        if parts[:1] == ['drop'] and req.method == 'POST':
            if len(parts) != 2 or parts[1] != self.box or self.unlinked:
                return route.fulfill(status=404, headers=CORS, json={'error': 'This sending link isn’t active. Link your phone again from the panel’s Settings.'})
            data = req.post_data_buffer or b''
            self.n += 1
            self.items.append({'id': 'item%012d' % self.n, 'type': req.headers.get('content-type', ''), 'name': unquote(req.headers.get('x-file-name', '')),
                               'size': len(data), 'at': int(time.time() * 1000), 'bytes': data})
            return route.fulfill(status=201, headers=CORS, json={'ok': True, 'message': 'Sent to Label Sheet Keeper.'})
        if parts[:1] == ['box'] and len(parts) in (2, 3):
            if self.unlinked or parts[1] != self.box or req.headers.get('authorization') != 'Bearer ' + str(self.key):
                return route.fulfill(status=401, headers=CORS, json={'error': 'not linked'})
            if len(parts) == 2 and req.method == 'GET':
                return route.fulfill(status=200, headers=CORS, json={'items': [{k: it[k] for k in ('id', 'type', 'name', 'size', 'at')} for it in self.items]})
            if len(parts) == 2 and req.method == 'DELETE':
                self.items, self.unlinked = [], True
                return route.fulfill(status=200, headers=CORS, json={'ok': True})
            it = next((x for x in self.items if x['id'] == parts[2]), None)
            if req.method == 'GET':
                if not it: return route.fulfill(status=404, headers=CORS, json={'error': 'This label is no longer waiting.'})
                if it.get('arriving'): return route.fulfill(status=409, headers=CORS, json={'error': 'This label is still on its way. Try again in a moment.'})
                return route.fulfill(status=200, headers=dict(CORS, **{'Content-Type': it['type']}), body=it['bytes'])
            if req.method == 'DELETE':
                self.items = [x for x in self.items if x['id'] != parts[2]]
                return route.fulfill(status=200, headers=CORS, json={'ok': True, 'waiting': len(self.items)})
        return route.fulfill(status=404, headers=CORS, json={'error': 'not found'})

    def calls(self, method, path):
        return [l for l in self.log if l[0] == method and l[1] == path]

def watch_errors(pg):
    errs = []
    pg.on('pageerror', lambda e: errs.append(str(e)))
    pg.on('console', lambda m: errs.append('console: ' + m.text) if m.type == 'error' and 'Failed to load resource' not in m.text else None)
    return errs

def new_page(b, query, worker):
    pg = b.new_page(viewport={'width': 340, 'height': 1000})
    errs = watch_errors(pg)
    pg.route(API + '/**', worker.handle)
    pg.route('**/office.js', lambda r: r.fulfill(body=MOCK, content_type='application/javascript'))
    pg.goto('http://127.0.0.1:%d/taskpane.html?%s' % (PORT, query))
    return pg, errs

def wait(pg, js, timeout=8000, arg=None):
    try:
        pg.wait_for_function(js, arg=arg, timeout=timeout); return True
    except Exception as e:
        return False

def say(pg):
    return pg.inner_text('#say') if pg.is_visible('#say') else ''

def open_settings(pg):
    pg.evaluate("document.getElementById('settings').open = true")

try:
  with sync_playwright() as p:
    b = p.chromium.launch()

    # ================= A: linking a phone on a shipping sheet =================
    fw = FakeWorker()
    pg, errs = new_page(b, 'fmt=shipping-2-sideways', fw)
    wait(pg, "!document.getElementById('v-sheet').hidden")
    open_settings(pg)
    check(pg.is_visible('#phone-link') and not pg.is_visible('#phone-unlink') and not pg.is_visible('#phone-qr-box'), 'Settings offers to link a phone')
    check(not fw.log, 'nothing is asked of the server before a phone is linked', fw.log)
    pg.click('#phone-link')
    check(wait(pg, "!document.getElementById('phone-qr-box').hidden && !!document.getElementById('phone-qr').getAttribute('data-url')"), 'linking shows a QR code')
    url = pg.get_attribute('#phone-qr', 'data-url')
    check(url == 'http://127.0.0.1:%d/phone.html#%s' % (PORT, fw.box), 'the QR code opens the phone page with the mailbox', url)
    dark = pg.evaluate("(() => { const c = document.getElementById('phone-qr'), d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data; let k = 0; for (let i = 0; i < d.length; i += 4) if (d[i] < 128) k++; return k; })()")
    check(dark > 1000, 'the QR code is drawn', dark)
    link = pg.evaluate("JSON.parse(localStorage.getItem('lsk.phone'))")
    check(link and link['box'] == fw.box and link['key'] == fw.key, 'the link is kept on this Mac', link)
    check('lsk.phone' not in pg.evaluate("LSK_TEST.backupKeys(true)"), 'the link’s read key stays out of backups')
    check(pg.is_visible('#phone-unlink') and not pg.is_visible('#phone-link') and 'Linked' in pg.inner_text('#phone-status'), 'Settings shows the phone is linked', pg.inner_text('#phone-status'))
    check('Scan the QR code' in say(pg), 'the panel says to scan it', say(pg))
    pg.click('#phone-show')
    check(not pg.is_visible('#phone-qr-box') and pg.inner_text('#phone-show') == 'Show the QR code', 'the QR code can be hidden')
    pg.click('#phone-show')
    check(wait(pg, "!document.getElementById('phone-qr-box').hidden"), 'and shown again')
    pg.screenshot(path=OUT + '/v18-linked.png', full_page=True)

    # ================= B: one label arrives =================
    pg.evaluate("LSK_TEST.checkPhone()")
    check(wait(pg, "LSK_TEST.phone.checking === false") and pg.is_hidden('#phone-in'), 'no note while nothing is waiting')
    fw.add(PDF, 'application/pdf', 'Amazon return.pdf')
    pg.evaluate("LSK_TEST.checkPhone()")
    check(wait(pg, "!document.getElementById('phone-in').hidden"), 'a label sent from the phone shows up')
    check(pg.inner_text('#phone-in-msg') == '1 label from your phone.' and pg.inner_text('#phone-in-use') == 'Use it', 'the note counts it', pg.inner_text('#phone-in-msg'))
    lists = fw.calls('GET', '/box/' + fw.box)
    check(lists and all(l[2] == 'Bearer ' + fw.key for l in lists), 'checking uses the read key', lists[-1:] if lists else lists)
    item = fw.items[0]['id']
    pg.click('#phone-in-use')
    check(wait(pg, "!!LSK_TEST.ship.src && !document.getElementById('crop').hidden", 20000), 'Use it opens the label to crop')
    check(pg.evaluate("LSK_TEST.ship.name") == 'Amazon return.pdf', 'with its name', pg.evaluate("LSK_TEST.ship.name"))
    check(wait(pg, "document.getElementById('phone-in').hidden"), 'the note goes away')
    check(fw.calls('GET', '/box/%s/%s' % (fw.box, item)) and fw.calls('DELETE', '/box/%s/%s' % (fw.box, item)) and not fw.items, 'the label is fetched, then deleted from the server')
    pg.click('#ship-place')
    check(wait(pg, "document.getElementById('say').textContent.indexOf('Placed on') >= 0", 10000), 'and goes on the sheet like any label file', say(pg))

    # ================= C: several labels arrive =================
    fw.add(PDF, 'application/pdf', 'Return 1.pdf')
    fw.add(PNG, 'image/png', 'Screenshot.png')
    pg.evaluate("LSK_TEST.checkPhone()")
    check(wait(pg, "document.getElementById('phone-in-msg').textContent === '2 labels from your phone.'") and pg.inner_text('#phone-in-use') == 'Use them', 'two labels are counted', pg.inner_text('#phone-in-msg'))
    pg.click('#phone-in-use')
    check(wait(pg, "document.getElementById('wait-count').textContent === '2'", 25000), 'Use them puts both on the waiting list', pg.inner_text('#wait-count') if pg.is_visible('#wait-count') else say(pg))
    names = pg.evaluate("Array.from(document.querySelectorAll('#wait-list li span')).map(s => s.textContent)")
    check(any('Return 1.pdf' in n for n in names) and any('Screenshot.png' in n for n in names), 'by name', names)
    check(not fw.items and len([l for l in fw.log if l[0] == 'DELETE']) == 3, 'both are deleted from the server', fw.log)

    # ================= D: a label that expired before it was used =================
    fw.add(PDF, 'application/pdf', 'Old.pdf')
    pg.evaluate("LSK_TEST.checkPhone()")
    wait(pg, "!document.getElementById('phone-in').hidden")
    fw.items = []          # gone from the server meanwhile
    pg.evaluate("document.getElementById('say').textContent = ''")
    pg.click('#phone-in-use')
    check(wait(pg, "document.getElementById('say').textContent.indexOf('no longer waiting') >= 0"), 'a label that’s gone says so', say(pg))
    check(wait(pg, "document.getElementById('phone-in').hidden"), 'and leaves the note')
    check(fw.calls('DELETE', '/box/%s/item%012d' % (fw.box, fw.n)), 'and is taken off the server’s list too')

    # ================= D2: a label just sent, not everywhere yet =================
    fw.add(PDF, 'application/pdf', 'Fresh.pdf')
    fw.items[-1]['arriving'] = True
    pg.evaluate("LSK_TEST.checkPhone()")
    wait(pg, "!document.getElementById('phone-in').hidden")
    pg.evaluate("document.getElementById('say').textContent = ''")
    pg.click('#phone-in-use')
    check(wait(pg, "document.getElementById('say').textContent.indexOf('still on its way') >= 0"), 'a label still on its way says to try again', say(pg))
    check(pg.is_visible('#phone-in') and len(fw.items) == 1 and not fw.calls('DELETE', '/box/%s/%s' % (fw.box, fw.items[0]['id'])), 'and stays waiting, not deleted')
    fw.items[-1]['arriving'] = False
    pg.click('#phone-in-use')
    check(wait(pg, "!!LSK_TEST.ship.src && LSK_TEST.ship.name === 'Fresh.pdf'", 15000) and not fw.items, 'once it arrives, Use it works')

    # ================= E: unlinking =================
    open_settings(pg)
    box = fw.box
    pg.click('#phone-unlink')
    check(wait(pg, "!document.getElementById('confirm').hidden") and 'labels not picked up yet are deleted' in pg.inner_text('#confirm-msg'), 'unlinking asks first')
    pg.click('#confirm-yes')
    check(wait(pg, "localStorage.getItem('lsk.phone') === null"), 'unlinking forgets the link')
    check(fw.calls('DELETE', '/box/' + box), 'and deletes the mailbox on the server')
    check(pg.is_visible('#phone-link') and not pg.is_visible('#phone-unlink') and not pg.is_visible('#phone-qr-box'), 'Settings offers to link again')
    check('unlinked' in say(pg), 'the panel confirms', say(pg))
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (A–E)', errs)
    pg.close()

    # ================= F: an address sheet, then the link stops working =================
    fw2 = FakeWorker('Cx' * 11, 'Lx' * 16)
    fw2.add(PDF, 'application/pdf', 'Return.pdf')
    pg, errs = new_page(b, 'fmt=avery-5160', fw2)
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.evaluate("l => localStorage.setItem('lsk.phone', JSON.stringify(l))", {'box': fw2.box, 'key': fw2.key})
    pg.evaluate("LSK_TEST.checkPhone()")
    check(wait(pg, "!document.getElementById('phone-in').hidden"), 'address sheets show waiting phone labels too')
    check(pg.inner_text('#phone-in-msg') == '1 label from your phone is waiting. Open a shipping-label sheet to use it.' and not pg.is_visible('#phone-in-use'), 'but say where to use them', pg.inner_text('#phone-in-msg'))
    fw2.unlinked = True
    pg.evaluate("LSK_TEST.checkPhone()")
    check(wait(pg, "localStorage.getItem('lsk.phone') === null"), 'a link the server no longer knows is forgotten')
    open_settings(pg)
    check(pg.is_hidden('#phone-in') and 'unlinked' in pg.inner_text('#phone-status') and pg.is_visible('#phone-link'), 'and Settings says to link again', pg.inner_text('#phone-status'))
    fw2.offline = True
    pg.evaluate("document.getElementById('say').textContent = ''")
    pg.click('#phone-link')
    check(wait(pg, "document.getElementById('say').textContent.indexOf('Couldn’t reach the Label Sheet Keeper server') >= 0"), 'no connection while linking says so', say(pg))
    check(pg.is_visible('#phone-link') and not pg.is_disabled('#phone-link'), 'and the button can be tried again')
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (F)', errs)
    pg.close()

    # ================= G: the phone page =================
    fw3 = FakeWorker('Dx' * 11, 'Mx' * 16)
    ctx = b.new_context(viewport={'width': 390, 'height': 844})
    pp = ctx.new_page()
    perrs = watch_errors(pp)
    pp.route(API + '/**', fw3.handle)
    pp.goto('http://127.0.0.1:%d/phone.html#%s' % (PORT, fw3.box))
    check(pp.is_visible('#pick-btn') and pp.is_hidden('#nolink'), 'the phone page opened from the QR code is ready to send')
    check(pp.input_value('#addr') == API + '/drop/' + fw3.box + '?reply=text', 'it shows the Shortcut’s sending address', pp.input_value('#addr'))
    pp.screenshot(path=OUT + '/v18-phone.png', full_page=True)
    pp.set_input_files('#pick', PDF)
    check(wait(pp, "document.getElementById('status').textContent.indexOf('Sent.') === 0"), 'choosing a label sends it', pp.inner_text('#status'))
    pdf_bytes = open(PDF, 'rb').read()
    check(len(fw3.items) == 1 and fw3.items[0]['bytes'] == pdf_bytes and fw3.items[0]['name'] == 'sample-return-label.pdf', 'the whole file arrives with its name', [(i['name'], i['size']) for i in fw3.items])
    pp.set_input_files('#pick', [PDF, PNG])
    check(wait(pp, "document.getElementById('status').textContent.indexOf('Sent 2 labels') === 0"), 'several labels can be sent at once', pp.inner_text('#status'))
    check(len(fw3.items) == 3, 'all of them arrive', len(fw3.items))
    pp.goto('http://127.0.0.1:%d/phone.html' % PORT)
    check(pp.is_visible('#pick-btn') and pp.evaluate("location.hash") == '#' + fw3.box, 'the page remembers the mailbox when opened again (Home Screen)')
    fw3.unlinked = True
    pp.set_input_files('#pick', PDF)
    check(wait(pp, "document.getElementById('status').textContent.indexOf('Link your phone again') >= 0") and 'warn' in pp.get_attribute('#status', 'class'), 'an unlinked phone is told what to do', pp.inner_text('#status'))
    check(not [e for e in perrs if 'favicon' not in e], 'no script errors (phone page)', perrs)
    ctx.close()

    ctx = b.new_context(viewport={'width': 390, 'height': 844})
    pp = ctx.new_page()
    pp.goto('http://127.0.0.1:%d/phone.html' % PORT)
    check(pp.is_visible('#nolink') and pp.is_hidden('#ready') and pp.is_hidden('#shortcut'), 'without a link the page says to scan the QR code in the panel')
    ctx.close()
    b.close()
finally:
    srv.terminate()

fails = [r for r in results if not r[0]]
print('\n%d checks, %d failed' % (len(results), len(fails)))
for f in fails: print('  FAILED:', f[1], '--', f[2])
if os.environ.get('GITHUB_ACTIONS'):
    esc = lambda s: str(s).replace('%', '%25').replace('\r', '%0D').replace('\n', '%0A')
    for f in fails: print('::error title=%s::%s' % (os.path.basename(__file__), esc(f[1])))
sys.exit(1 if fails else 0)
