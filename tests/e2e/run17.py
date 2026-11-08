import subprocess, time, sys, json, os
from playwright.sync_api import sync_playwright

S = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # tests/
REPO = os.path.dirname(S)
OUT = os.path.join(S, 'out')
os.makedirs(OUT, exist_ok=True)
MOCK = open(S + '/mock-office.js').read()
PORT = 8796
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



def scan(pg, fmt, file):
    pg.evaluate("id => { localStorage.setItem('lsk.testCheck.' + id, 'pending'); localStorage.setItem('lsk.testNudge.' + id, JSON.stringify({x:0, y:0})); localStorage.removeItem('lsk.nudge.' + id); }", fmt)
    refresh(pg)
    pg.evaluate("document.getElementById('say').textContent = ''")
    pg.set_input_files('#scan-file', file)
    wait(pg, "!document.getElementById('scan-result').hidden || document.getElementById('say').textContent.length > 20", 20000)
    return (pg.inner_text('#scan-msg') if pg.is_visible('#scan-result') else ''), say(pg)

try:
  with sync_playwright() as p:
    b = p.chromium.launch()

    # ================= A: scans with the alignment squares =================
    pg, errs = new_page(b, 'fmt=shipping-2-sideways')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.evaluate("localStorage.setItem('lsk.testCheck.shipping-2-sideways', 'pending')"); refresh(pg)
    check(pg.is_visible('#scan-label'), 'sideways sheets offer a scan after a test page')
    m, s = scan(pg, 'shipping-2-sideways', S + '/fixtures/mark-ship-cw.png')
    check('.04 in too far right' in m and '.06 in too low' in m, 'a sideways test page scanned turned is measured', [m, s])
    pg.click('#scan-save')
    nz = pg.evaluate("JSON.parse(localStorage.getItem('lsk.nudge.shipping-2-sideways'))")
    check(nz == {'x': -0.04, 'y': -0.06}, 'its adjustment is saved', nz)
    m, s = scan(pg, 'shipping-2-sideways', S + '/fixtures/mark-ship-good.png')
    check('No adjustment needed' in m, 'a sideways page that printed right needs nothing', [m, s])
    m, s = scan(pg, 'shipping-2-sideways', S + '/fixtures/scan-off.png')
    check('alignment squares weren’t found' in s and not m, 'an old test page on a sideways size asks for a new one', [m, s])
    pg.click('#settings summary', position={'x': 30, 'y': 8})
    check(not pg.evaluate("document.getElementById('scan-label2').hidden"), 'Use a scan is in Settings for sideways sizes too')
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (A)', errs)
    pg.close()

    pg, errs = new_page(b, 'fmt=avery-5160')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    m, s = scan(pg, 'avery-5160', S + '/fixtures/mark-5160-upside.png')
    check('.02 in too far left' in m and '.04 in too low' in m, 'an upside-down scan is measured', [m, s])
    m, s = scan(pg, 'avery-5160', S + '/fixtures/scan-off.png')
    check('.05 in too far right' in m and '.03 in too high' in m, 'test pages from before the squares still work', [m, s])
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (A2)', errs)
    pg.close()

    # ================= B: recent addresses =================
    pg, errs = new_page(b, 'fmt=avery-5160')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.evaluate("""localStorage.setItem('lsk.addressBook', JSON.stringify([{id:'a', name:'Ann Lee', lines:['Ann Lee','5 Pine Rd','Salem, MA 01970']}]));
      localStorage.setItem('lsk.recent', JSON.stringify(Array.from({length: 10}, (_, i) => ({lines: ['Old ' + i, i + ' Elm St', 'Boston, MA 02108'], at: 1}))));""")
    pg.click('#fill-tabs button:has-text("Address")')
    pg.fill('#addr-text', 'Zed New\n1 Main St\nSalem, MA 01970'); pg.click('#addr-one')
    wait(pg, "Object.values(LSK_TEST.st().texts).some(t => t.indexOf('Zed New') === 0)")
    r = pg.evaluate("JSON.parse(localStorage.getItem('lsk.recent'))")
    check(len(r) == 10 and r[0]['lines'][0] == 'Zed New' and r[-1]['lines'][0] == 'Old 8', 'a placed address goes to the top of Recent; only ten are kept', [x['lines'][0] for x in r])
    groups = pg.evaluate("Array.from(document.querySelectorAll('#addr-pick optgroup')).map(g => [g.label, g.children.length])")
    check(groups == [['Recent', 10], ['Address book', 1]], 'one dropdown with Recent and Address book', groups)
    pg.select_option('#addr-pick', 'recent:3')
    check(pg.input_value('#addr-text') == 'Old 2\n2 Elm St\nBoston, MA 02108' and pg.input_value('#addr-pick') == '', 'choosing a recent address fills it in', pg.input_value('#addr-text'))
    pg.click('#addr-one'); wait(pg, "JSON.parse(localStorage.getItem('lsk.recent'))[0].lines[0] === 'Old 2'")
    r = pg.evaluate("JSON.parse(localStorage.getItem('lsk.recent')).map(x => x.lines[0])")
    check(r[0] == 'Old 2' and r.count('Old 2') == 1 and len(r) == 10, 'using one again moves it up without a copy', r)
    pg.select_option('#addr-pick', 'a')
    check(pg.input_value('#addr-text').startswith('Ann Lee'), 'address book entries still work')
    pg.click('#fill-tabs button:has-text("Table")'); pg.click('#table-open'); wait(pg, "!!MOCK.dialog")
    pg.evaluate("rows => MOCK.dialog.send({type:'fill', rows})", [['Table Person', '', '1 A St', '', 'Boston', 'MA', '02108']])
    wait(pg, "Object.values(LSK_TEST.st().texts).some(t => t.indexOf('Table Person') === 0)")
    check('Table Person' not in str(pg.evaluate("localStorage.getItem('lsk.recent')")), 'table fills don’t crowd Recent')
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (B)', errs)
    pg.close()

    pg, errs = new_page(b, 'fmt=shipping-2-sideways')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.click('#fill-tabs button:has-text("Address")')
    pg.fill('#from-addr', 'Me\n1 My St\nBoston, MA 02101'); pg.fill('#to-addr', 'Returns Center\n100 Warehouse Rd\nCity, KY 40000')
    pg.click('#ship-addr-place')
    wait(pg, "MOCK.cellText(MOCK.labelCell(1)).indexOf('SHIP TO') >= 0")
    r = pg.evaluate("JSON.parse(localStorage.getItem('lsk.recent'))")
    check(r and r[0]['lines'] == ['Returns Center', '100 Warehouse Rd', 'City, KY 40000'], 'Ship to addresses go to Recent', r)
    pg.select_option('#to-pick', 'recent:0')
    check(pg.input_value('#to-addr').startswith('Returns Center'), 'and can be picked for Ship to')
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (B2)', errs)
    pg.close()

    # ================= C: address book groups =================
    pg, errs = new_page(b, 'fmt=avery-5160')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.evaluate("""localStorage.setItem('lsk.addressBook', JSON.stringify([{id:'a', name:'Ann Lee', lines:['Ann Lee','5 Pine Rd','Salem, MA 01970']}, {id:'b', name:'Bob Roe', lines:['Bob Roe','9 Oak Ave','Boston, MA 02108']}, {id:'c', name:'Cy Fox', lines:['Cy Fox','2 Elm St','Lynn, MA 01901']}]))""")
    pg.click('#fill-tabs button:has-text("Address")')
    pg.click('#pane-addr details.inner summary', position={'x': 30, 'y': 8})
    check(not pg.is_visible('#group-row'), 'no group filter until there are groups')
    pg.click('#book-list li:nth-child(1) .addr')
    pg.fill('#book-edit-groups', 'Family, work ,family'); pg.click('#book-list .book-edit button:has-text("Save")')
    pg.click('#book-list li:nth-child(3) .addr')
    pg.fill('#book-edit-groups', 'FAMILY'); pg.press('#book-edit-groups', 'Enter')
    bk = pg.evaluate("JSON.parse(localStorage.getItem('lsk.addressBook'))")
    check(bk[0].get('groups') == ['Family', 'work'] and bk[2].get('groups') == ['Family'] and 'groups' not in bk[1], 'groups saved, without repeats, in one spelling', bk)
    opts = pg.evaluate("Array.from(document.querySelectorAll('#group-pick option')).map(o => o.textContent)")
    check(pg.is_visible('#group-row') and opts == ['All addresses', 'Family (2)', 'work (1)'], 'group filter lists each group with its count', opts)
    check(pg.evaluate("Array.from(document.querySelectorAll('#group-names option')).map(o => o.value)") == ['Family', 'work'], 'group names are suggested while typing')
    pg.select_option('#group-pick', 'Family')
    shown = pg.evaluate("Array.from(document.querySelectorAll('#book-list .addr')).map(b => b.textContent.split(',')[0])")
    check(shown == ['Ann Lee', 'Cy Fox'] and pg.is_visible('#group-fill'), 'choosing a group shows only its addresses', shown)
    pg.click('#group-fill')
    check(wait(pg, "['Ann Lee','Cy Fox'].every(n => Object.values(LSK_TEST.st().texts).some(t => t.indexOf(n) === 0))"), 'the group goes on the open labels', say(pg))
    check('with Family' in say(pg), 'message names the group', say(pg))
    pg.click('#book-list li:nth-child(1) .addr'); pg.fill('#book-edit-groups', ''); pg.click('#book-list .book-edit button:has-text("Save")')
    pg.click('#book-list li:nth-child(1) .addr'); pg.fill('#book-edit-groups', ''); pg.click('#book-list .book-edit button:has-text("Save")')
    check(not pg.is_visible('#group-row') and pg.evaluate("document.querySelectorAll('#book-list .addr').length") == 3, 'when the last group empties, the filter goes and all addresses show')
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (C)', errs)
    pg.close()

    tp = b.new_page(viewport={'width': 900, 'height': 640})
    tp.route('**/office.js', lambda r: r.fulfill(body='', content_type='application/javascript'))
    tp.goto('http://127.0.0.1:%d/table.html' % PORT)
    tp.evaluate("""localStorage.setItem('lsk.tableInfo', JSON.stringify({open: 5, labels: [1,2,3,4,5], name: 'Address labels, 30 per sheet', per: 30})); localStorage.setItem('lsk.tableDraft','[]');
      localStorage.setItem('lsk.addressBook', JSON.stringify([{id:'a', name:'Ann Lee', lines:['Ann Lee','5 Pine Rd','Salem, MA 01970'], groups:['Family']}, {id:'b', name:'Bob Roe', lines:['Bob Roe','9 Oak Ave','Boston, MA 02108']}, {id:'c', name:'Cy Fox', lines:['Cy Fox','2 Elm St','Lynn, MA 01901'], groups:['family','Work']}]))""")
    tp.reload(); tp.wait_for_timeout(300)
    opts = tp.evaluate("Array.from(document.querySelectorAll('#book option')).map(o => o.textContent)")
    check('Add group: Family (2)' in opts and 'Add group: Work (1)' in opts, 'the table window can add a whole group', opts)
    tp.select_option('#book', 'group:family')
    rows = tp.evaluate("LSK_TABLE.readRows()")
    check([r[0] for r in rows] == ['Ann Lee', 'Cy Fox'], 'adding a group adds its addresses', rows)
    tp.close()

    # ================= D: labels saved counter =================
    pg, errs = new_page(b, 'fmt=avery-5160')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.evaluate("localStorage.removeItem('lsk.savedLabels')")
    pg.evaluate("MOCK.setLabelText(1, 'A\\n1 St', 10)"); refresh(pg); pg.click('#printed-btn'); wait(pg, "(MOCK.settings.lskTrips||0) === 1")
    check(pg.evaluate("localStorage.getItem('lsk.savedLabels')") in (None, '0'), 'a fresh sheet saves nothing', pg.evaluate("localStorage.getItem('lsk.savedLabels')"))
    pg.evaluate("MOCK.setLabelText(2, 'B\\n2 St', 10); MOCK.setLabelText(3, 'C\\n3 St', 10)"); refresh(pg); pg.click('#printed-btn'); wait(pg, "(MOCK.settings.lskTrips||0) === 2")
    check(pg.evaluate("localStorage.getItem('lsk.savedLabels')") == '2', 'labels printed on a used sheet are counted', pg.evaluate("localStorage.getItem('lsk.savedLabels')"))
    t = pg.inner_text('#tally .saved') if pg.query_selector('#tally .saved') else ''
    check(t == '2 labels saved', 'the tally shows the count', t)
    pg.click('#settings summary', position={'x': 30, 'y': 8})
    pg.uncheck('#show-saved')
    check(not pg.query_selector('#tally .saved'), 'turning it off hides the count')
    pg.evaluate("MOCK.setLabelText(4, 'D\\n4 St', 10)"); refresh(pg); pg.click('#printed-btn'); wait(pg, "(MOCK.settings.lskTrips||0) === 3")
    check(pg.evaluate("localStorage.getItem('lsk.savedLabels')") == '3', 'the count keeps going while hidden')
    pg.check('#show-saved')
    t = pg.inner_text('#tally .saved') if pg.query_selector('#tally .saved') else ''
    check(t == '3 labels saved', 'turning it back on shows the right number', t)
    pg.click('#undo-btn') if pg.is_visible('#undo-btn') else None
    if pg.is_visible('#confirm'): pg.click('#confirm-yes')
    wait(pg, "localStorage.getItem('lsk.savedLabels') === '2'")
    check(pg.evaluate("localStorage.getItem('lsk.savedLabels')") == '2', 'undoing a print takes its labels off the count', pg.evaluate("localStorage.getItem('lsk.savedLabels')"))
    pg.click('details:has(#hist) summary', position={'x': 30, 'y': 8})
    pg.evaluate("document.querySelectorAll('#hist li')[0].querySelector('button').click()")
    wait(pg, "!document.getElementById('confirm').hidden"); pg.click('#confirm-yes')
    check(wait(pg, "localStorage.getItem('lsk.savedLabels') === '0'"), 'and so does undoing from Print history', pg.evaluate("localStorage.getItem('lsk.savedLabels')"))
    sv = pg.evaluate("document.querySelector('#tally .saved') && document.querySelector('#tally .saved').title")
    check(sv and 'partly used sheets' in sv, 'the count explains itself on hover', sv)
    pg.screenshot(path=OUT + '/v17.png', full_page=True)
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (D)', errs)
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
