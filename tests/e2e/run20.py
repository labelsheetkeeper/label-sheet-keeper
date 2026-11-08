import subprocess, time, sys, json, os
from urllib.parse import urlparse, parse_qs
from playwright.sync_api import sync_playwright

S = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # tests/
REPO = os.path.dirname(S)
OUT = os.path.join(S, 'out')
os.makedirs(OUT, exist_ok=True)
MOCK = open(S + '/mock-office.js').read()
PORT = 8794
API = 'https://api.labelsheetkeeper.workers.dev'
results = []

def check(cond, name, detail=''):
    results.append((bool(cond), name, detail))
    print(('PASS ' if cond else 'FAIL ') + name + ('' if cond else '  -- ' + str(detail)))

srv = subprocess.Popen(['python3', '-m', 'http.server', str(PORT), '--bind', '127.0.0.1'], cwd=REPO, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
time.sleep(1)

def geo(h, s, c, z):
    return {'housenumber': h, 'street': s, 'city': c, 'state_code': 'MA', 'postcode': z, 'country_code': 'us', 'address_line1': h + ' ' + s, 'address_line2': '%s, MA %s' % (c, z)}

class Server:
    """Stands in for the Worker's /suggest, and for Photon. Can hold answers back."""
    def __init__(self):
        self.asked, self.photon, self.held, self.hold, self.status = [], [], [], False, 200
    def answer(self, route):
        text = parse_qs(urlparse(route.request.url).query).get('text', [''])[0]
        cors = {'Access-Control-Allow-Origin': '*'}
        if self.status != 200:
            return route.fulfill(status=self.status, headers=cors, json={'error': 'Address suggestions are used up for today.'})
        if 'salem' in text.lower():
            res = [geo('8', 'Maplewood Ave', 'Salem', '01970')]
        elif 'pine' in text.lower() or '5 pi' in text.lower():
            res = [geo('5', 'Pine Rd', 'Salem', '01970')]
        else:
            res = [geo('8', 'Maple Way', 'Boston', '02108'), geo('8', 'Maplewood Ave', 'Salem', '01970'), geo('28', 'Maple St', 'Quincy', '02169')]
        route.fulfill(headers=cors, json={'results': res})
    def handle(self, route):
        self.asked.append(route.request.url)
        if self.hold: self.held.append(route)
        else: self.answer(route)
    def release(self):
        while self.held: self.answer(self.held.pop(0))
    def photon_handle(self, route):
        self.photon.append(route.request.url)
        route.fulfill(json={'type': 'FeatureCollection', 'features': [
            {'type': 'Feature', 'properties': {'type': 'house', 'housenumber': '9', 'street': 'Elm Street', 'city': 'Salem', 'state': 'Massachusetts', 'postcode': '01970', 'countrycode': 'US'}}]})

def new_page(b, query, server):
    pg = b.new_page(viewport={'width': 340, 'height': 1000})
    errs = []
    pg.on('pageerror', lambda e: errs.append(str(e)))
    pg.route(API + '/**', server.handle)
    pg.route('https://photon.komoot.io/**', server.photon_handle)
    pg.route('**/office.js', lambda r: r.fulfill(body=MOCK, content_type='application/javascript'))
    pg.goto('http://127.0.0.1:%d/taskpane.html?%s' % (PORT, query))
    return pg, errs

def wait(pg, js, timeout=6000, arg=None):
    try:
        pg.wait_for_function(js, arg=arg, timeout=timeout); return True
    except Exception:
        return False

SHOWN = "() => !document.getElementById('asugg').hidden"
def items(pg):
    return pg.evaluate("Array.from(document.querySelectorAll('#asugg-list button')).map(b => b.firstChild.textContent)")

def put(pg, sel, value, caret_line=None):
    """Sets an address box and puts the caret at the end of a line, as if typed there."""
    pg.evaluate("""([sel, v, line]) => { const t = document.querySelector(sel); t.focus(); t.value = v;
      let pos = v.length; if (line !== null) { const ls = v.split('\\n'); pos = ls.slice(0, line + 1).join('\\n').length; }
      t.setSelectionRange(pos, pos); t.dispatchEvent(new Event('input', {bubbles: true})); }""", [sel, value, caret_line])

try:
  with sync_playwright() as p:
    b = p.chromium.launch()
    srv1 = Server()
    pg, errs = new_page(b, 'fmt=avery-5160', srv1)
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.click('#fill-tabs button:has-text("Address")')

    # ================= A: the Address box =================
    pg.click('#addr-text')
    pg.keyboard.type('Jane Doe')
    pg.keyboard.press('Enter')
    pg.keyboard.type('8 M')
    pg.wait_for_timeout(400)
    check(pg.is_hidden('#asugg') and not srv1.asked, 'nothing is asked before a house number and two letters')
    pg.keyboard.type('a')
    check(wait(pg, SHOWN), 'suggestions appear under the Address box at "8 Ma"')
    check(items(pg) == ['8 Maple Way', '8 Maplewood Ave', '28 Maple St'], 'with the matches', items(pg))
    check(srv1.asked and srv1.asked[-1].startswith(API + '/suggest?text=8%20Ma'), 'from the shared suggestions (no key needed)', srv1.asked)
    box = pg.evaluate("(() => { const a = document.getElementById('addr-text').getBoundingClientRect(), d = document.getElementById('asugg').getBoundingClientRect(); return [Math.round(a.bottom), Math.round(d.top), Math.round(a.width), Math.round(d.width)]; })()")
    check(abs(box[1] - box[0]) <= 4 and box[2] == box[3], 'right under the box, as wide as it', box)
    check('Powered by Geoapify' in pg.inner_text('#asugg-credit') and 'OpenStreetMap' in pg.inner_text('#asugg-credit') and not pg.query_selector('#asugg-credit a'), 'credited in plain text (no links out of Word)', pg.inner_text('#asugg-credit'))
    pg.screenshot(path=OUT + '/v20-address.png')

    # Typing on narrows what's showing at once, before the server answers.
    srv1.hold = True
    n = len(srv1.asked)
    pg.keyboard.type('p')
    check(wait(pg, "document.querySelectorAll('#asugg-list button').length === 2", 400), 'the next letter narrows the list at once', items(pg))
    check(items(pg) == ['8 Maple Way', '8 Maplewood Ave'], 'to what still fits ("8 Map" isn’t 28 Maple St)', items(pg))
    pg.wait_for_timeout(400)
    check(len(srv1.asked) == n and items(pg) == ['8 Maple Way', '8 Maplewood Ave'], 'the first answer already had every match (fewer than 5), so no new lookup is spent', [len(srv1.asked), n])
    srv1.hold = False; srv1.release()
    pg.wait_for_timeout(200)

    pg.keyboard.press('ArrowDown')
    check(pg.get_attribute('#asugg-list button:nth-child(1)', 'aria-selected') == 'true', 'the arrow keys choose one')
    pg.keyboard.press('Enter')
    check(pg.input_value('#addr-text') == 'Jane Doe\n8 Maple Way\nBoston, MA 02108', 'Enter fills the street line and adds the city line', pg.input_value('#addr-text'))
    check(pg.evaluate("document.getElementById('addr-text').selectionStart") == len('Jane Doe\n8 Maple Way') and pg.is_hidden('#asugg'), 'the cursor waits at the end of the street, list closed')
    n = len(srv1.asked)
    pg.keyboard.type(', Apt 4')
    pg.wait_for_timeout(400)
    check(pg.is_hidden('#asugg') and len(srv1.asked) == n, 'adding ", Apt 4" after picking stays quiet')
    check(pg.input_value('#addr-text') == 'Jane Doe\n8 Maple Way, Apt 4\nBoston, MA 02108', 'and the apartment goes on the street line', pg.input_value('#addr-text'))

    # A city line already there is replaced, not doubled.
    put(pg, '#addr-text', 'Ann Lee\n8 Ma\nSalem, MA 01970', caret_line=1)
    check(wait(pg, SHOWN) and items(pg) == ['8 Maplewood Ave'], 'the city line under the street narrows the search', [srv1.asked[-1], items(pg)])
    pg.click('#asugg-list button')
    check(pg.input_value('#addr-text') == 'Ann Lee\n8 Maplewood Ave\nSalem, MA 01970', 'picking replaces the street and city lines, without doubling them', pg.input_value('#addr-text'))

    # Enter with nothing chosen is just a new line; Escape closes.
    put(pg, '#addr-text', 'Bob Ray\n8 Ma')
    wait(pg, SHOWN)
    pg.keyboard.press('Enter')
    check(pg.input_value('#addr-text') == 'Bob Ray\n8 Ma\n' and wait(pg, "document.getElementById('asugg').hidden"), 'Enter without choosing is a new line, as usual', repr(pg.input_value('#addr-text')))
    pg.keyboard.press('Backspace')
    check(wait(pg, SHOWN), 'back on the street line, suggestions return')
    pg.keyboard.press('Escape')
    check(pg.is_hidden('#asugg'), 'Escape closes them')
    put(pg, '#addr-text', 'Acme Inc\nSuite 200')
    pg.wait_for_timeout(300)
    check(pg.is_hidden('#asugg'), 'lines that don’t start with a house number don’t ask')

    # ================= B: the contact editor =================
    put(pg, '#addr-text', 'Sam Lee\n5 Pine Rd\nSalem, MA 01970')
    pg.click('#addr-save')
    pg.evaluate("document.querySelector('#imp').closest('details').open = true")
    pg.click('#book-list .addr')
    wait(pg, "!!document.getElementById('book-edit-2')")
    pg.click('#book-edit-2')
    pg.keyboard.type('5 Pi')
    check(wait(pg, SHOWN) and items(pg) == ['5 Pine Rd'], 'the contact editor’s second address suggests too', items(pg))
    pg.keyboard.press('Escape')
    check(pg.is_hidden('#asugg') and pg.is_visible('#book-edit-2'), 'Escape closes the list without closing the editor')
    pg.keyboard.press('Escape')
    check(wait(pg, "!document.getElementById('book-edit-2')"), 'and a second Escape still closes the editor as before')

    # ================= C: turned off =================
    pg.evaluate("document.getElementById('settings').open = true")
    pg.uncheck('#suggest-on')
    n = len(srv1.asked)
    put(pg, '#addr-text', 'Off Test\n8 Maple')
    pg.wait_for_timeout(400)
    check(pg.is_hidden('#asugg') and len(srv1.asked) == n, 'turned off in Settings, nothing is suggested or sent')
    check('as you type' in pg.inner_text('label:has(#suggest-on)'), 'the setting covers the address boxes too', pg.inner_text('label:has(#suggest-on)'))
    pg.check('#suggest-on')
    check(not errs, 'no script errors (address sheet)', errs)
    pg.close()

    # ================= D: Ship to and From on shipping sheets; used up -> Photon =================
    srv2 = Server()
    pg, errs = new_page(b, 'fmt=shipping-2-sideways', srv2)
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.click('#fill-tabs button:has-text("Address")')
    put(pg, '#to-addr', 'Sam Lee\n8 Ma')
    check(wait(pg, SHOWN) and pg.evaluate("Math.abs(document.getElementById('asugg').getBoundingClientRect().top - document.getElementById('to-addr').getBoundingClientRect().bottom) <= 4"), 'Ship to suggests, under its own box')
    pg.keyboard.press('ArrowDown'); pg.keyboard.press('Tab')
    check(pg.input_value('#to-addr') == 'Sam Lee\n8 Maple Way\nBoston, MA 02108', 'Tab picks too', pg.input_value('#to-addr'))
    srv2.status = 429
    put(pg, '#from-addr', 'Me\n9 Elm S')
    check(wait(pg, "() => !document.getElementById('asugg').hidden && document.getElementById('asugg-credit').textContent.indexOf('Photon') === 0"), 'when the shared ones are used up, Photon answers (From box)', pg.inner_text('#asugg-credit') if pg.is_visible('#asugg') else '')
    n = len(srv2.asked)
    put(pg, '#from-addr', 'Me\n9 Elm St')
    pg.wait_for_timeout(500)
    check(len(srv2.asked) == n and pg.is_visible('#asugg'), 'and the server isn’t asked again', [len(srv2.asked), n, len(srv2.photon)])
    check('used up for today' in pg.inner_text('#asugg-credit'), 'the credit line says why Photon is answering', pg.inner_text('#asugg-credit'))
    check(not errs, 'no script errors (shipping sheet)', errs)
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
