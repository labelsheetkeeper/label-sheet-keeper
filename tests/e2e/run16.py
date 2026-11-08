import subprocess, time, sys, json, os
from playwright.sync_api import sync_playwright

S = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # tests/
REPO = os.path.dirname(S)
OUT = os.path.join(S, 'out')
os.makedirs(OUT, exist_ok=True)
MOCK = open(S + '/mock-office.js').read()
PORT = 8795
results = []

def check(cond, name, detail=''):
    results.append((bool(cond), name, detail))
    print(('PASS ' if cond else 'FAIL ') + name + ('' if cond else '  -- ' + str(detail)))

srv = subprocess.Popen(['python3', '-m', 'http.server', str(PORT), '--bind', '127.0.0.1'], cwd=REPO, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
time.sleep(1)

GEO_CALLS = []
def fake_services(page):
    def zippo(route):
        url = route.request.url
        if url.endswith('/us/02108'):
            route.fulfill(json={'post code': '02108', 'places': [{'place name': 'Boston', 'state abbreviation': 'MA'}]})
        elif url.endswith('/ca/M5V'):
            route.fulfill(json={'post code': 'M5V', 'places': [{'place name': 'Toronto (Downtown West)', 'state abbreviation': 'ON'}]})
        else:
            route.fulfill(status=404, json={})
    def geo(route):
        url = route.request.url
        GEO_CALLS.append(url)
        if 'apiKey=badkeybadkeybadkey' in url:
            route.fulfill(status=401, json={'message': 'Invalid apiKey'}); return
        route.fulfill(json={'results': [
            {'housenumber': '8', 'street': 'Maple Way', 'city': 'Boston', 'state_code': 'MA', 'postcode': '02108', 'country_code': 'us', 'address_line1': '8 Maple Way', 'address_line2': 'Boston, MA 02108, United States of America'},
            {'housenumber': '290', 'street': 'Bremner Boulevard', 'city': 'Toronto', 'state_code': 'ON', 'postcode': 'M5V 3L9', 'country_code': 'ca', 'address_line1': '290 Bremner Boulevard', 'address_line2': 'Toronto, ON M5V 3L9, Canada'}]})
    def photon(route):
        url = route.request.url
        GEO_CALLS.append(url)
        route.fulfill(json={'type': 'FeatureCollection', 'features': [
            {'type': 'Feature', 'properties': {'type': 'street', 'name': 'Maple Way', 'city': 'Boston', 'state': 'Massachusetts', 'postcode': '02108', 'countrycode': 'US'}},
            {'type': 'Feature', 'properties': {'type': 'house', 'housenumber': '8', 'street': 'Maple Way', 'city': 'Boston', 'state': 'Massachusetts', 'postcode': '02108', 'countrycode': 'US'}},
            {'type': 'Feature', 'properties': {'type': 'house', 'housenumber': '8', 'street': 'Maple Rue', 'city': 'Montréal', 'state': 'Québec', 'postcode': 'h2x1y4', 'countrycode': 'CA'}},
            {'type': 'Feature', 'properties': {'type': 'house', 'housenumber': '8', 'street': 'Maplestraße', 'city': 'Berlin', 'state': 'Berlin', 'postcode': '10115', 'countrycode': 'DE'}}]})
    page.route('https://photon.komoot.io/**', photon)
    page.route('https://api.zippopotam.us/**', zippo)
    page.route('https://api.geoapify.com/**', geo)
    # The Label Sheet Keeper server without a Geoapify key set up: suggestions fall back to Photon.
    page.route('https://api.labelsheetkeeper.workers.dev/**', lambda r: r.fulfill(status=503, headers={'Access-Control-Allow-Origin': '*'}, json={'error': 'not set up'}))

def new_page(b, query, scheme='light'):
    pg = b.new_page(viewport={'width': 340, 'height': 1000}, color_scheme=scheme)
    fake_services(pg)
    errs = []
    pg.on('pageerror', lambda e: errs.append(str(e)))
    pg.on('console', lambda m: errs.append('console: ' + m.text) if m.type == 'error' else None)
    pg.route('**/office.js', lambda r: r.fulfill(body=MOCK, content_type='application/javascript'))
    pg.goto('http://127.0.0.1:%d/taskpane.html?%s' % (PORT, query))
    return pg, errs

def wait(pg, js, timeout=8000):
    try:
        pg.wait_for_function(js, timeout=timeout); return True
    except Exception as e:
        return False

def cell_text(pg, n):
    return pg.evaluate('n => MOCK.cellText(MOCK.labelCell(n))', n)

def used(pg, n):
    return pg.evaluate("n => MOCK.labelCell(n).ccs.some(c => c.tag === 'lsk-used')", n)

def say(pg):
    return pg.inner_text('#say') if pg.is_visible('#say') else ''

def select_label(pg, n):
    pg.evaluate("n => { const S = LSK_SHEET, f = LSK_FORMATS.getFormat(MOCK.props[S.PROP]) || LSK_TEST.st().fmt; const c = S.cellOf(f, n); MOCK.selection = { row: c.row, cell: c.cell }; }", n)

def refresh(pg):
    pg.click('#refresh-btn'); wait(pg, "document.getElementById('say').textContent.indexOf('Up to date') >= 0")


REG = """(entries) => { const now = Date.now(), r = {}; entries.forEach(e => { r[e.id] = Object.assign({fmtId:'avery-5160', fmtName:'Address labels, 30 per sheet', file: e.url ? e.url.split('/').pop() : '', open: 5, ready: 0, trips: 0, at: now}, e); }); localStorage.setItem('lsk.sheets', JSON.stringify(r)); }"""
try:
  with sync_playwright() as p:
    b = p.chromium.launch()

    # ================= A: the sheet list cleans itself up =================
    pg, errs = new_page(b, 'fmt=avery-5160')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.evaluate("e => (" + REG + ")(e)", [
        {'id': 'old-unsaved', 'url': '', 'seen': int(time.time() * 1000) - 20 * 60000},
        {'id': 'legacy-unsaved', 'url': '', 'at': int(time.time() * 1000) - 2 * 86400000},
        {'id': 'fresh-unsaved', 'url': '', 'seen': int(time.time() * 1000) - 60000},
        {'id': 'saved-missing', 'url': '/x/missing labels.docx'},
        {'id': 'saved-ok', 'url': '/x/ok labels.docx'}])
    refresh(pg)
    reg = pg.evaluate("Object.keys(JSON.parse(localStorage.getItem('lsk.sheets')))")
    check('old-unsaved' not in reg and 'legacy-unsaved' not in reg and 'fresh-unsaved' in reg and 'saved-missing' in reg, 'closed unsaved sheets drop off the list by themselves', reg)
    me = pg.evaluate("MOCK.settings.lskSheetId")
    check(me in reg and pg.evaluate("id => JSON.parse(localStorage.getItem('lsk.sheets'))[id].seen > Date.now() - 5000", me), 'this sheet checks in')
    pg.click('#sheets-box summary', position={'x': 30, 'y': 8})
    items = pg.evaluate("Array.from(document.querySelectorAll('#sheets-list li')).map(li => [li.textContent, !!Array.from(li.querySelectorAll('button')).find(x => x.textContent === 'Open' && !x.hidden)])")
    fresh = [i for i in items if 'not saved yet' in i[0]]
    check(len(fresh) == 1 and not fresh[0][1], 'an open unsaved sheet shows without an Open button', items)
    pg.evaluate("Array.from(document.querySelectorAll('#sheets-list li')).find(li => li.textContent.indexOf('missing labels') >= 0).querySelector('button').click()")
    check(wait(pg, "document.getElementById('say').textContent.indexOf('couldn’t find “missing labels.docx”') >= 0"), 'a sheet Word can’t find is explained', say(pg))
    check('saved-missing' not in pg.evaluate("Object.keys(JSON.parse(localStorage.getItem('lsk.sheets')))"), 'and taken off the list')
    # Fill your other sheets first skips one that can't be opened
    pg.evaluate("e => (" + REG + ")(e)", [
        {'id': 'gone', 'url': '/x/missing old.docx', 'open': 12, 'trips': 1},
        {'id': 'ok2', 'url': '/x/ok2.docx', 'open': 5, 'trips': 1}])
    pg.click('#fill-tabs button:has-text("Table")'); pg.click('#table-open'); wait(pg, "!!MOCK.dialog")
    rows = [['Person %d' % i, '', '%d Main St' % i, '', 'Boston', 'MA', '2101'] for i in range(33)]
    pg.evaluate("rows => MOCK.dialog.send({type:'fill', rows})", rows)
    wait(pg, "!document.getElementById('confirm').hidden")
    pg.click('#confirm-yes')
    check(wait(pg, "MOCK.opened.indexOf('/x/ok2.docx') >= 0"), 'the next sheet is opened when one can’t be', pg.evaluate("MOCK.opened"))
    job = pg.evaluate("JSON.parse(localStorage.getItem('lsk.fillJob'))")
    check(job and job['next'] == 'ok2' and 'gone' in job['done'] and len(job['items']) == 3, 'the list goes to the sheet that opened', job and {k: job[k] for k in ('next', 'done')})
    check('opened “ok2.docx”' in say(pg), 'message names the sheet that opened', say(pg))
    # Declining Word's question keeps the sheet on the list
    pg.evaluate("e => (" + REG + ")(e)", [{'id': 'declined', 'url': '/x/declined labels.docx'}])
    refresh(pg); pg.click('#sheets-box summary', position={'x': 30, 'y': 8})
    pg.evaluate("Array.from(document.querySelectorAll('#sheets-list li')).find(li => li.textContent.indexOf('declined labels') >= 0).querySelector('button').click()")
    check(wait(pg, "document.getElementById('say').textContent.indexOf('didn’t open “declined labels.docx”') >= 0") and 'declined' in pg.evaluate("Object.keys(JSON.parse(localStorage.getItem('lsk.sheets')))"), 'a sheet that wasn’t opened (you said no) stays on the list', say(pg))
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (A)', errs)
    pg.close()

    pg, errs = new_page(b, 'fmt=avery-5160&unsaved=1')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    me = pg.evaluate("MOCK.settings.lskSheetId")
    check(me in pg.evaluate("Object.keys(JSON.parse(localStorage.getItem('lsk.sheets')||'{}'))"), 'an unsaved sheet is listed while open')
    pg.evaluate("MOCK.url = '/x/Saved later.docx'")
    check(wait(pg, "id => (JSON.parse(localStorage.getItem('lsk.sheets'))[id] || {}).url === '/x/Saved later.docx'", 9000) if False else wait(pg, "(JSON.parse(localStorage.getItem('lsk.sheets'))[MOCK.settings.lskSheetId] || {}).url === '/x/Saved later.docx'", 9000), 'a sheet saved during the session gets its address in the list', pg.evaluate("[MOCK.url, MOCK.settings.lskSheetId, localStorage.getItem('lsk.sheets'), document.getElementById('say').textContent]"))
    pg.evaluate("window.dispatchEvent(new Event('pagehide'))")
    check(me in pg.evaluate("Object.keys(JSON.parse(localStorage.getItem('lsk.sheets')||'{}'))"), 'and stays listed when it closes')
    pg.close()
    pg, errs = new_page(b, 'fmt=avery-5160&unsaved=1')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    me = pg.evaluate("MOCK.settings.lskSheetId")
    pg.evaluate("window.dispatchEvent(new Event('pagehide'))")
    check(me not in pg.evaluate("Object.keys(JSON.parse(localStorage.getItem('lsk.sheets')||'{}'))"), 'an unsaved sheet leaves the list when its panel closes')
    pg.close()

    # ================= B: editing saved addresses =================
    pg, errs = new_page(b, 'fmt=avery-5160')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.evaluate("""localStorage.setItem('lsk.addressBook', JSON.stringify([{id:'a', name:'Ann Lee', lines:['Ann Lee','5 Pine Rd','Salem, MA 01970']}, {id:'b', name:'Bob Roe', lines:['Bob Roe','9 Oak Ave','Boston, MA 02108']}]))""")
    pg.click('#fill-tabs button:has-text("Address")')
    pg.click('#pane-addr details.inner summary', position={'x': 30, 'y': 8})
    pg.click('#book-list li:first-child .addr')
    check(pg.is_visible('#book-edit') and pg.input_value('#book-edit') == 'Ann Lee\n5 Pine Rd\nSalem, MA 01970', 'clicking an address opens it for editing')
    pg.fill('#book-edit', 'Ann Lee-Park\n5 Pine Road\nSalem, MA 01970')
    refresh(pg)
    check(pg.is_visible('#book-edit') and 'Pine Road' in pg.input_value('#book-edit'), 'a refresh keeps what’s being typed')
    pg.click('#book-list .book-edit button:has-text("Save")')
    bk = pg.evaluate("JSON.parse(localStorage.getItem('lsk.addressBook'))")
    check(bk[0]['name'] == 'Ann Lee-Park' and bk[0]['lines'][1] == '5 Pine Road' and len(bk) == 2 and not pg.is_visible('#book-edit'), 'edited address saved', bk)
    check('Saved the changes to Ann Lee-Park' in say(pg), 'save message', say(pg))
    pg.click('#book-list li:nth-child(2) .addr')
    pg.fill('#book-edit', 'Ann Lee-Park\n5 Pine Road\nSalem, MA 01970')
    pg.click('#book-list .book-edit button:has-text("Save")')
    bk = pg.evaluate("JSON.parse(localStorage.getItem('lsk.addressBook'))")
    check(len(bk) == 1 and 'now one' in say(pg), 'editing into a duplicate merges the two', [bk, say(pg)])
    pg.click('#book-list li:first-child .addr'); pg.fill('#book-edit', ''); pg.click('#book-list .book-edit button:has-text("Save")')
    check('click × to delete' in say(pg) and len(pg.evaluate("JSON.parse(localStorage.getItem('lsk.addressBook'))")) == 1, 'an empty edit is refused', say(pg))
    pg.click('#book-list .book-edit button:has-text("Cancel")')
    check(not pg.is_visible('#book-edit'), 'cancel closes the editor')
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (B)', errs)
    pg.close()

    # ================= C: undo an older print =================
    pg, errs = new_page(b, 'fmt=avery-5160')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.evaluate("localStorage.setItem('lsk.stock.avery-5160', JSON.stringify({on:5, pack:10, warn:2, packs:1, link:''}))")
    pg.evaluate("MOCK.setLabelText(1, 'Ann Lee\\n5 Pine Rd\\nSalem, MA 01970', 12); const c = MOCK.labelCell(1); c.paras[0].font.bold = true; c.paras.forEach(p => p.alignment = 'Centered'); c.paras[2].leftIndent = 9;")
    refresh(pg); pg.click('#printed-btn')
    wait(pg, "(MOCK.settings.lskTrips||0) === 1")
    check(pg.evaluate("JSON.parse(localStorage.getItem('lsk.stock.avery-5160')).on") == 4, 'first print took a sheet off the stock')
    pg.evaluate("MOCK.setLabelText(2, 'Bob Roe\\n9 Oak Ave\\nBoston, MA 02108', 10)"); refresh(pg); pg.click('#printed-btn')
    wait(pg, "(MOCK.settings.lskTrips||0) === 2")
    h = pg.evaluate("MOCK.settings.lskHistory")
    check(len(h) == 2 and h[1]['labels']['1'][0]['b'] and h[1]['stock'], 'history keeps each print’s text', h[1])
    pg.click('details:has(#hist) summary', position={'x': 30, 'y': 8})
    undos = pg.evaluate("Array.from(document.querySelectorAll('#hist li')).map(li => !!li.querySelector('button'))")
    check(undos == [True, True], 'both prints can be undone', undos)
    pg.evaluate("document.querySelectorAll('#hist li')[1].querySelector('button').click()")
    check(wait(pg, "!document.getElementById('confirm').hidden") and 'Label 1 becomes open again with its text back' in pg.inner_text('#confirm-msg'), 'undo asks first', pg.inner_text('#confirm-msg') if pg.is_visible('#confirm') else '')
    pg.click('#confirm-yes')
    check(wait(pg, "!LSK_TEST.st().used[1] && LSK_TEST.st().ready[1]"), 'label 1 is open again with its text')
    c1 = pg.evaluate("MOCK.labelCell(1).paras.map(p => [p.text, p.font.size, p.font.bold, p.alignment, p.leftIndent])")
    check([x[0] for x in c1] == ['Ann Lee', '5 Pine Rd', 'Salem, MA 01970'] and c1[0][2] and c1[0][1] == 12 and c1[0][3] == 'Centered' and c1[2][4] == 9, 'its lines, size, bold, alignment and indent come back', c1)
    check(pg.evaluate("MOCK.settings.lskTrips") == 1 and len(pg.evaluate("MOCK.settings.lskHistory")) == 1 and pg.evaluate("LSK_TEST.st().used[2] === true"), 'trip count and history updated; the other print stays', pg.evaluate("MOCK.settings.lskTrips"))
    check(pg.evaluate("JSON.parse(localStorage.getItem('lsk.stock.avery-5160')).on") == 5, 'undoing the first trip puts the sheet back in stock')
    check('Label 1 is open again with its text back' in say(pg), 'undo message', say(pg))
    # A print from before 1.6 comes back empty
    pg.evaluate("const h = MOCK.settings.lskHistory; delete h[0].labels;")
    pg.evaluate("window.LSK_TEST")  # no-op
    refresh(pg)
    pg.evaluate("document.querySelectorAll('#hist li')[0].querySelector('button').click()")
    wait(pg, "!document.getElementById('confirm').hidden")
    m = pg.inner_text('#confirm-msg')
    check('text back' in m, 'the latest print still has its exact copy', m)
    pg.click('#confirm-no')
    pg.click('#reset-btn'); pg.click('#confirm-yes'); wait(pg, "Object.keys(LSK_TEST.st().used).length === 0")
    check(pg.evaluate("Array.from(document.querySelectorAll('#hist li button')).length") == 0, 'no undo after the sheet is reset')
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (C)', errs)
    pg.close()

    # A legacy entry with no saved text
    pg, errs = new_page(b, 'fmt=avery-5160')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.evaluate("MOCK.setLabelText(3, 'Old One\\n1 A St\\nSalem, MA 01970', 10)"); refresh(pg); pg.click('#printed-btn'); wait(pg, "(MOCK.settings.lskTrips||0) === 1")
    pg.evaluate("MOCK.setLabelText(4, 'Newer\\n2 B St\\nSalem, MA 01970', 10)"); refresh(pg); pg.click('#printed-btn'); wait(pg, "(MOCK.settings.lskTrips||0) === 2")
    pg.evaluate("delete MOCK.settings.lskHistory[1].labels"); refresh(pg)
    pg.click('details:has(#hist) summary', position={'x': 30, 'y': 8})
    pg.evaluate("document.querySelectorAll('#hist li')[1].querySelector('button').click()")
    wait(pg, "!document.getElementById('confirm').hidden")
    check('comes back empty' in pg.inner_text('#confirm-msg'), 'an old print without saved text says it comes back empty', pg.inner_text('#confirm-msg'))
    pg.click('#confirm-yes')
    check(wait(pg, "!LSK_TEST.st().used[3] && !LSK_TEST.st().ready[3]"), 'label 3 open and empty')
    pg.screenshot(path=OUT + '/v16.png', full_page=True)
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (C2)', errs)
    pg.close()
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
