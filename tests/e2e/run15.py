import subprocess, time, sys, json, os
from playwright.sync_api import sync_playwright

S = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # tests/
REPO = os.path.dirname(S)
OUT = os.path.join(S, 'out')
os.makedirs(OUT, exist_ok=True)
MOCK = open(S + '/mock-office.js').read()
PORT = 8791 + int(os.environ.get('LSK_PORT_SHIFT', '0'))   # shift lets several runs share one machine
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


def open_count(pg):
    return pg.evaluate("(() => { let n=0; const s=LSK_TEST.st(); for (let i=1;i<=30;i++) if(!s.used[i]&&!s.ready[i]) n++; return n; })()")

try:
  with sync_playwright() as p:
    b = p.chromium.launch()

    # ================= A: label stock and reorder =================
    pg, errs = new_page(b, 'fmt=avery-5160')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.click('#stock-box summary')
    check(pg.is_visible('#stock-on'), 'clicking the summary opens Label stock')
    tip = pg.get_attribute('#stock-info', 'data-tip')
    check('Avery 5160 for a laser printer or 8160 for an inkjet' in tip, 'stock tip names the laser and inkjet numbers', tip)
    pg.hover('#stock-info')
    check(wait(pg, "!document.getElementById('tipbox').hidden && document.getElementById('tipbox').textContent.indexOf('Avery 5160') >= 0"), 'pointing at (i) shows its explanation')
    box = pg.evaluate("(() => { const r = document.getElementById('tipbox').getBoundingClientRect(); return [r.left, r.right, document.documentElement.clientWidth]; })()")
    check(box[0] >= 8 and box[1] <= box[2] - 7, 'explanation box stays inside the panel', box)
    pg.mouse.move(5, 5)
    check(wait(pg, "document.getElementById('tipbox').hidden"), 'and hides when the pointer leaves')
    pg.click('#stock-info')
    check(pg.is_visible('#tipbox') and pg.is_visible('#stock-on'), 'clicking (i) shows it without closing the section')
    pg.mouse.move(5, 5); pg.wait_for_timeout(100)
    check(pg.is_visible('#tipbox'), 'a clicked explanation stays while the pointer moves away')
    pg.mouse.click(5, 5)
    check(wait(pg, "document.getElementById('tipbox').hidden"), 'clicking elsewhere hides it')
    pg.fill('#stock-on', '3'); pg.fill('#stock-pack', '10'); pg.fill('#stock-warn', '2'); pg.fill('#stock-packs', '2')
    pg.fill('#stock-link', 'https://www.amazon.com/Avery-Address-Labels-Printers-5160/dp/B00004Z5SP?th=1')
    pg.dispatch_event('#stock-link', 'input')
    check('adds the labels to your Amazon cart' in pg.inner_text('#stock-link-msg'), 'Amazon link recognized for add-to-cart', pg.inner_text('#stock-link-msg'))
    pg.fill('#stock-link', 'not a link'); pg.dispatch_event('#stock-link', 'input')
    check('starting with https://' in pg.inner_text('#stock-link-msg'), 'bad link explained', pg.inner_text('#stock-link-msg'))
    pg.click('#stock-save')
    check('doesn’t look right' in say(pg) and not pg.evaluate("localStorage.getItem('lsk.stock.avery-5160')"), 'bad link not saved', say(pg))
    pg.fill('#stock-link', 'https://www.amazon.com/Avery-Address-Labels-Printers-5160/dp/B00004Z5SP?th=1')
    pg.click('#stock-save')
    sk = pg.evaluate("JSON.parse(localStorage.getItem('lsk.stock.avery-5160'))")
    check(sk == {'on': 3, 'pack': 10, 'warn': 2, 'packs': 2, 'link': 'https://www.amazon.com/Avery-Address-Labels-Printers-5160/dp/B00004Z5SP?th=1'}, 'stock saved', sk)
    check('(3 sheets)' in pg.inner_text('#stock-box summary'), 'stock count in the heading', pg.inner_text('#stock-box summary'))
    pg.evaluate("MOCK.setLabelText(1, 'A Person\\n1 Elm St\\nBoston, MA 02108', 10)"); refresh(pg)
    pg.click('#printed-btn')
    check(wait(pg, "JSON.parse(localStorage.getItem('lsk.stock.avery-5160')).on === 2"), 'first trip takes a sheet off the count')
    check(wait(pg, "!document.getElementById('stock-low').hidden") and '2 sheets of these labels left' in pg.inner_text('#stock-low-msg'), 'low-stock reminder', pg.inner_text('#stock-low-msg') if pg.is_visible('#stock-low') else '')
    pg.click('#undo-btn')
    check(wait(pg, "JSON.parse(localStorage.getItem('lsk.stock.avery-5160')).on === 3 && document.getElementById('stock-low').hidden"), 'undo puts the sheet back')
    pg.click('#printed-btn')
    wait(pg, "JSON.parse(localStorage.getItem('lsk.stock.avery-5160')).on === 2")
    pg.evaluate("MOCK.setLabelText(2, 'B Person\\n2 Elm St\\nBoston, MA 02108', 10)"); refresh(pg)
    pg.click('#printed-btn')
    check(wait(pg, "(MOCK.settings.lskTrips||0) === 2") and pg.evaluate("JSON.parse(localStorage.getItem('lsk.stock.avery-5160')).on") == 2, 'second trip of the same sheet doesn’t count again')
    pg.click('#stock-low-order')
    check(pg.evaluate("MOCK.browser[MOCK.browser.length-1]") == 'https://www.amazon.com/gp/aws/cart/add.html?ASIN.1=B00004Z5SP&Quantity.1=2', 'reorder opens the Amazon cart with 2 packs', pg.evaluate("MOCK.browser"))
    check('add 2 packs to your cart' in say(pg), 'reorder message', say(pg))
    pg.click('#stock-low-got')
    check(wait(pg, "JSON.parse(localStorage.getItem('lsk.stock.avery-5160')).on === 12 && document.getElementById('stock-low').hidden"), 'a new pack adds its sheets and clears the reminder')
    pg.fill('#stock-link', 'https://www.walmart.com/ip/Avery-Easy-Peel-Address-Labels/12345678'); pg.click('#stock-save'); pg.click('#stock-order')
    check(pg.evaluate("MOCK.browser[MOCK.browser.length-1]") == 'https://www.walmart.com/sc/cart/addToCart?items=12345678%7C2', 'Walmart reorder goes to the cart', pg.evaluate("MOCK.browser"))
    pg.fill('#stock-link', 'https://www.staples.com/avery-5160/product_123'); pg.click('#stock-save'); pg.click('#stock-order')
    check(pg.evaluate("MOCK.browser[MOCK.browser.length-1]") == 'https://www.staples.com/avery-5160/product_123' and 'product page at Staples' in say(pg), 'other stores open the product page', say(pg))
    pg.click('#reset-btn'); pg.click('#confirm-yes')
    wait(pg, "(MOCK.settings.lskTrips||0) === 0")
    pg.evaluate("MOCK.setLabelText(3, 'C Person\\n3 Elm St\\nBoston, MA 02108', 10)"); refresh(pg)
    pg.click('#printed-btn')
    check(wait(pg, "JSON.parse(localStorage.getItem('lsk.stock.avery-5160')).on === 11"), 'a fresh sheet after Reset counts again')
    pg.click('#stock-off')
    sk = pg.evaluate("JSON.parse(localStorage.getItem('lsk.stock.avery-5160'))")
    check('on' not in sk and sk['link'].startswith('https://www.staples.com'), 'stop counting keeps the link', sk)
    check(not pg.is_visible('#stock-low') and pg.is_visible('#stock-order'), 'no reminder when not counting, reorder still there')
    pg.fill('#stock-link', ''); pg.fill('#stock-pack', ''); pg.click('#stock-save')
    check(pg.evaluate("localStorage.getItem('lsk.stock.avery-5160')") is None and 'Cleared' in say(pg), 'saving an empty form clears the size’s stock', say(pg))

    # ================= B: fill your other sheets first =================
    pg.evaluate("""() => { const r = JSON.parse(localStorage.getItem('lsk.sheets')||'{}'); r.other = {id:'other', fmtId:'avery-5160', fmtName:'Address labels, 30 per sheet', file:'Old labels.docx', url:'/x/Old labels.docx', open:12, ready:0, trips:1, at:Date.now()};
      r.full = {id:'full', fmtId:'avery-5160', fmtName:'Address labels, 30 per sheet', file:'Worn out.docx', url:'/x/Worn out.docx', open:20, ready:0, trips:3, at:Date.now()}; localStorage.setItem('lsk.sheets', JSON.stringify(r)); }""")
    pg.click('#fill-tabs button:has-text("Table")'); pg.click('#table-open'); wait(pg, "!!MOCK.dialog")
    n_open = open_count(pg)
    rows = [['Person %d' % i, '', '%d Main St' % i, '', 'Boston', 'MA', '2101'] for i in range(n_open + 5)]
    pg.evaluate("rows => MOCK.dialog.send({type:'fill', rows})", rows)
    check(wait(pg, "!document.getElementById('confirm').hidden"), 'overflow question asked')
    msg = pg.inner_text('#confirm-msg')
    check('Your other sheet of this size has 12 open labels' in msg and pg.inner_text('#confirm-yes') == 'Fill my other sheets first' and pg.inner_text('#confirm-alt') == 'Make new sheets', 'offers your other part-used sheet first (not one at its trip limit)', [msg, pg.inner_text('#confirm-yes'), pg.inner_text('#confirm-alt')])
    created_before = pg.evaluate("MOCK.created.length")
    pg.click('#confirm-yes')
    check(wait(pg, "MOCK.opened.length === 1") and pg.evaluate("MOCK.opened[0]") == '/x/Old labels.docx', 'the other sheet is opened', pg.evaluate("MOCK.opened"))
    job = pg.evaluate("JSON.parse(localStorage.getItem('lsk.fillJob'))")
    check(job and len(job['items']) == 5 and job['next'] == 'other' and job['items'][0][0]['t'] == 'Person %d' % n_open, 'the other 5 wait for that sheet', job and {k: job[k] for k in ('next', 'sheets')})
    check(wait(pg, "open => (() => { let n=0; const s=LSK_TEST.st(); for (let i=1;i<=30;i++) if(!s.used[i]&&!s.ready[i]) n++; return n; })() === 0", 0) or open_count(pg) == 0, 'this sheet filled first')
    check(pg.evaluate("MOCK.created.length") == created_before, 'no new sheet made')
    check('its panel fills them' in say(pg), 'message says what happens next', say(pg))
    job_json = pg.evaluate("localStorage.getItem('lsk.fillJob')")
    reg_json = pg.evaluate("localStorage.getItem('lsk.sheets')")
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (A, B)', errs)
    pg.close()

    # The other sheet's panel picks up the list
    pg, errs = new_page(b, 'fmt=avery-5160&url=/x/Old%20labels.docx')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.evaluate("([j, r]) => { MOCK.settings.lskSheetId = 'other'; localStorage.setItem('lsk.fillJob', j); localStorage.setItem('lsk.sheets', r); }", [job_json, reg_json])
    pg.fill('#first-n', '18'); pg.click('#first-btn'); wait(pg, "MOCK.labelCell(18).ccs.length === 1")
    pg.evaluate("LSK_TEST.checkFillJob()")
    check(wait(pg, "MOCK.cellText(MOCK.labelCell(23)).indexOf('Person') === 0"), 'other sheet fills its open labels from the list', [cell_text(pg, n) for n in (19, 23, 24)])
    check(cell_text(pg, 24) == '' and pg.evaluate("localStorage.getItem('lsk.fillJob')") is None, 'list finished and cleared', [cell_text(pg, 24), pg.evaluate("localStorage.getItem('lsk.fillJob')")])
    check('That was the last of them' in say(pg), 'done message', say(pg))
    # More than fits, no sheets left: the rest go on a new sheet
    items = [[{'t': 'Extra %d' % i, 'size': 10}] for i in range(9)]
    pg.evaluate("items => { localStorage.setItem('lsk.fillJob', JSON.stringify({id:'j2', fmtId:'avery-5160', at:Date.now(), items, sheets:['other'], next:'other', done:[]})); }", items)
    pg.evaluate("LSK_TEST.checkFillJob()")
    check(wait(pg, "MOCK.created.length === 1", 10000), 'leftovers go on a new sheet')
    check(open_count(pg) == 0 and 'Filled labels 24–30' in say(pg) and 'Made 1 new sheet for the other 2 labels' in say(pg), 'this sheet’s 7 open labels filled, 2 on the new one', say(pg))
    old = pg.evaluate("JSON.stringify({id:'j3', fmtId:'avery-5160', at:Date.now() - 31*60000, items:[[{t:'Late'}]], sheets:['other'], next:'other', done:[]})")
    pg.evaluate("j => localStorage.setItem('lsk.fillJob', j)", old); pg.evaluate("LSK_TEST.checkFillJob()")
    check(pg.evaluate("localStorage.getItem('lsk.fillJob')") is None, 'an old list is dropped')
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (B2)', errs)
    pg.close()

    # Copies overflow, choosing Make new sheets
    pg, errs = new_page(b, 'fmt=avery-5160')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.evaluate("""() => { const r = {}; r.other = {id:'other', fmtId:'avery-5160', fmtName:'Address labels, 30 per sheet', file:'Old labels.docx', url:'/x/Old labels.docx', open:12, ready:0, trips:1, at:Date.now()}; localStorage.setItem('lsk.sheets', JSON.stringify(r)); }""")
    pg.evaluate("MOCK.setLabelText(1, 'Copy Me\\n1 Elm St\\nBoston, MA 02108', 10)"); select_label(pg, 1); refresh(pg)
    pg.fill('#copies', '33'); pg.click('#copy-btn')
    check(wait(pg, "!document.getElementById('confirm').hidden") and '29 copies go on this sheet' in pg.inner_text('#confirm-msg'), 'copies overflow asks too', pg.inner_text('#confirm-msg') if pg.is_visible('#confirm') else '')
    pg.click('#confirm-alt')
    check(wait(pg, "MOCK.created.length === 1", 10000) and pg.evaluate("MOCK.opened.length") == 0, 'Make new sheets skips the other sheet')
    check(open_count(pg) == 0, 'this sheet filled with copies')
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (B3)', errs)

    # ================= C: numbers, QR codes, barcodes =================
    pg.evaluate("for (let n = 1; n <= 30; n++) { const c = MOCK.labelCell(n); c.paras = [c.paras[0]]; c.paras[0].text = ''; }")
    refresh(pg)
    pg.click('#fill-tabs button:has-text("Numbers")')
    check(pg.is_visible('#pane-codes') and pg.input_value('#num-text') == 'Box {n} of {total}' and pg.input_value('#num-count') == '30', 'Numbers tab with a starting pattern and all open labels', [pg.input_value('#num-text'), pg.input_value('#num-count')])
    pg.fill('#num-count', '3')
    check('Box 1 of 3' in pg.inner_text('#code-msg') or 'label 1' in pg.inner_text('#code-msg'), 'preview describes the first label', pg.inner_text('#code-msg'))
    pg.click('#num-fill')
    check(wait(pg, "MOCK.cellText(MOCK.labelCell(3)) === 'Box 3 of 3'"), 'numbered labels', [cell_text(pg, n) for n in (1, 2, 3)])
    check(pg.evaluate("MOCK.labelCell(1).paras[0].font.size") > 14 and pg.evaluate("MOCK.labelCell(1).paras[0].font.bold"), 'numbers alone are big and bold', pg.evaluate("MOCK.labelCell(1).paras[0].font.size"))
    check(pg.input_value('#num-start') == '4', 'numbering goes on from 4')
    pg.click('#code-kind button[data-kind="qr"]')
    pg.fill('#code-data', 'https://example.com/box/{n}'); pg.fill('#num-count', '2'); pg.dispatch_event('#num-count', 'input')
    check(wait(pg, "!!window.qrcode"), 'QR maker loaded when needed')
    pg.click('#num-fill')
    check(wait(pg, "MOCK.labelCell(5).pics === 1"), 'QR codes placed', [pg.evaluate("MOCK.labelCell(%d).pics" % n) for n in (4, 5)])
    q = pg.evaluate("(() => { const c = MOCK.labelCell(4); return { texts: c.paras.map(p => p.text), w: c.paras[0].pic && c.paras[0].pic.width, h: c.paras[0].pic && c.paras[0].pic.height, size: c.paras[1].font.size, al: c.paras[0].alignment, alt: c.paras[0].pic.altTextDescription }; })()")
    check(q['texts'] == ['', 'Box 4 of 5'] and abs(q['w'] - q['h']) < 0.01 and 40 < q['w'] < 60 and q['size'] == 10 and q['al'] == 'Centered' and q['alt'] == 'QR code: https://example.com/box/4', 'QR picture above its caption, with alt text, square, about 0.7 in', q)
    # The QR holds the right text: decode it with the same library's matrix
    same = pg.evaluate("""async () => { const b64 = MOCK.labelCell(5).paras[0].pic.b64; const img = new Image(); img.src = 'data:image/png;base64,' + b64; await img.decode();
      const c = document.createElement('canvas'); c.width = img.width; c.height = img.height; const x = c.getContext('2d'); x.drawImage(img, 0, 0);
      const m = LSK_TOOLS.qrMatrix('https://example.com/box/5'), cells = m.length + 4, px = img.width / cells; let bad = 0;
      for (let r = 0; r < m.length; r++) for (let k = 0; k < m.length; k++) { const d = x.getImageData((k + 2.5) * px, (r + 2.5) * px, 1, 1).data[0] < 128; if (d !== m[r][k]) bad++; }
      return bad; }""")
    check(same == 0, 'QR picture matches the code for its own number', same)
    pg.click('#code-kind button[data-kind="bar"]')
    pg.fill('#code-data', 'BIN-{n}'); pg.select_option('#num-pad', '3'); pg.fill('#num-text', 'Bin {n}'); pg.fill('#num-count', '1'); pg.dispatch_event('#num-count', 'input')
    pg.click('#num-fill')
    check(wait(pg, "MOCK.labelCell(6).pics === 1"), 'barcode placed')
    bar = pg.evaluate("""async () => { const c = MOCK.labelCell(6), pic = c.paras[0].pic; const img = new Image(); img.src = 'data:image/png;base64,' + pic.b64; await img.decode();
      const cv = document.createElement('canvas'); cv.width = img.width; cv.height = img.height; const x = cv.getContext('2d'); x.drawImage(img, 0, 0);
      const row = x.getImageData(0, Math.floor(img.height / 2), img.width, 1).data; let bits = ''; for (let i = 0; i < img.width; i += 4) bits += row[i * 4] < 128 ? '1' : '0';
      return { bits: bits.replace(/^0+|0+$/g, ''), want: LSK_TOOLS.code128('BIN-006'), text: c.paras[1].text, w: pic.width, h: pic.height }; }""")
    check(bar['bits'] == bar['want'] and bar['text'] == 'Bin 006', 'barcode bars are Code 128 for BIN-006 with the padded caption', {k: bar[k] for k in ('text', 'w', 'h')})
    check(bar['w'] <= 175 and 25 < bar['h'] <= 40, 'barcode fits the label width', bar)
    pg.fill('#code-data', 'Ünïcode ✓'); pg.dispatch_event('#code-data', 'input')
    check('only letters, numbers' in pg.inner_text('#code-msg'), 'barcode explains characters it can’t hold', pg.inner_text('#code-msg'))
    pg.screenshot(path=OUT + '/v15-codes.png', full_page=True)
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (C)', errs)
    pg.close()

    pg, errs = new_page(b, 'fmt=avery-5167')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.click('#fill-tabs button:has-text("Numbers")'); pg.click('#code-kind button[data-kind="qr"]')
    check(wait(pg, "document.getElementById('code-msg').textContent.indexOf('too small for a QR code') >= 0"), 'small labels: QR codes not offered at an unreadable size', pg.inner_text('#code-msg'))
    pg.close()

    # ================= D: lining up from a scan =================
    pg, errs = new_page(b, 'fmt=avery-5160')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.evaluate("localStorage.setItem('lsk.testCheck.avery-5160', 'pending'); localStorage.setItem('lsk.testNudge.avery-5160', JSON.stringify({x:0, y:0}))"); refresh(pg)
    check(pg.is_visible('#test-check') and pg.is_visible('#scan-label'), 'test check offers a scan')
    pg.set_input_files('#scan-file', S + '/fixtures/scan-scaled.png')
    check(wait(pg, "document.getElementById('say').textContent.indexOf('94% of full size') >= 0"), 'scaled print explained', say(pg))
    pg.set_input_files('#scan-file', S + '/fixtures/scan-off.png')
    check(wait(pg, "!document.getElementById('scan-result').hidden"), 'scan measured', say(pg))
    m = pg.inner_text('#scan-msg')
    check('.05 in too far right' in m and '.03 in too high' in m and 'move right -0.05 in, move down 0.03 in' in m, 'scan result in words', m)
    pg.click('#scan-save')
    check(pg.evaluate("localStorage.getItem('lsk.nudge.avery-5160')") == '{"x":-0.05,"y":0.03}' and not pg.is_visible('#test-check'), 'adjustment saved from the scan', pg.evaluate("localStorage.getItem('lsk.nudge.avery-5160')"))
    pg.evaluate("localStorage.setItem('lsk.testCheck.avery-5160', 'pending'); localStorage.setItem('lsk.testNudge.avery-5160', JSON.stringify({x:0, y:0}))"); refresh(pg)
    pg.set_input_files('#scan-file', S + '/fixtures/scan-good.png')
    check(wait(pg, "!document.getElementById('scan-result').hidden") and 'No adjustment needed' in pg.inner_text('#scan-msg') and pg.is_visible('#scan-done'), 'a good test page needs nothing', pg.inner_text('#scan-msg') if pg.is_visible('#scan-msg') else say(pg))
    pg.click('#scan-done')
    check(not pg.is_visible('#test-check') and not pg.is_visible('#scan-result'), 'done closes the test check')
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (D)', errs)
    pg.close()

    # ================= E: several return labels =================
    pg, errs = new_page(b, 'fmt=shipping-2-sideways')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    check(not pg.evaluate("document.getElementById('scan-label2').hidden"), 'scan adjustment offered for sideways sheets too')
    pg.set_input_files('#ship-file', [S + '/fixtures/sample-return-label.pdf', S + '/fixtures/two-labels.pdf', S + '/fixtures/sample-return-label.pdf'])
    check(wait(pg, "document.getElementById('wait-count').textContent === '3' && document.getElementById('say').textContent.indexOf('Added 3 labels') >= 0", 30000), 'three files wait in line', say(pg))
    check(pg.inner_text('#wait-place') == 'Put the next 2 on this sheet', 'button says how many go on', pg.inner_text('#wait-place'))
    pg.click('#wait-place')
    check(wait(pg, "MOCK.labelCell(1).pics === 1 && MOCK.labelCell(2).pics === 1"), 'two labels placed on the sheet')
    check(pg.inner_text('#wait-count') == '1' and '1 more label is waiting' in say(pg), 'one still waiting', say(pg))
    pg.click('#wait-place')
    check('Print this sheet first' in say(pg), 'asks to print before the next sheet', say(pg))
    pg.click('#printed-btn'); wait(pg, "(MOCK.settings.lskTrips||0) === 1")
    pg.click('#wait-place')
    check(wait(pg, "!document.getElementById('confirm').hidden") and 'Loaded a fresh one' in pg.inner_text('#confirm-msg'), 'asks about a fresh sheet')
    pg.click('#confirm-yes')
    check(wait(pg, "MOCK.labelCell(1).pics === 1 && (MOCK.settings.lskTrips||0) === 0 && document.getElementById('waiting').hidden"), 'fresh sheet reset and the last label placed')
    pg.set_input_files('#ship-file', S + '/fixtures/two-labels.pdf')
    check(wait(pg, "!document.getElementById('pg-all').hidden", 20000), 'every-page button for a multi-page PDF')
    pg.click('#pg-all')
    check(wait(pg, "document.getElementById('wait-count').textContent === '2'", 20000) and 'from pages 1–2' in say(pg), 'both pages’ labels wait in line', say(pg))
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (E)', errs)
    pg.screenshot(path=OUT + '/v15-waiting.png', full_page=True)
    pg.close()

    # ================= F: backup and restore =================
    pg, errs = new_page(b, 'fmt=avery-5160')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.evaluate("""() => { localStorage.setItem('lsk.fromAddress', 'Me\\n1 My St\\nBoston, MA 02101'); localStorage.setItem('lsk.addressBook', JSON.stringify([{id:'a', name:'Ann', lines:['Ann','1 A St','Salem, MA 01970']}]));
      localStorage.setItem('lsk.nudge.avery-5160', JSON.stringify({x:0.03, y:0})); localStorage.setItem('lsk.geoKey', '0123456789abcdef0123456789abcdef');
      localStorage.setItem('lsk.stock.avery-5160', JSON.stringify({on:4, pack:10, warn:2, packs:1, link:'https://www.amazon.com/dp/B00004Z5SP'})); localStorage.setItem('lsk.tableDraft', '[1]'); }""")
    refresh(pg)
    pg.click('#settings summary')
    check(pg.is_visible('#backup-make') and pg.evaluate("document.getElementById('backup-block').parentNode.id") == 'backup-sheet-slot', 'backup tools inside Settings on a sheet')
    check(pg.is_visible('#backup-key-row'), 'key option shown when a key is saved')
    pg.click('#backup-make')
    check(wait(pg, "MOCK.created.length === 1 && MOCK.created[0].opened"), 'backup document made and opened')
    bk = pg.evaluate("""async () => { const z = await JSZip.loadAsync(MOCK.created[0].b64, {base64: true}); const item = await z.file('customXml/item1.xml').async('string');
      const doc = await z.file('word/document.xml').async('string'); const props = await z.file('docProps/custom.xml').async('string');
      const m = /<lskBackup[^>]*>([^<]+)</.exec(item); const data = JSON.parse(decodeURIComponent(escape(atob(m[1])))).data;
      return { item, keys: Object.keys(data).sort(), doc, props, we: !!z.file('word/webextensions/webextension1.xml'), compat: !!z.file('word/settings.xml') && (await z.file('word/settings.xml').async('string')).indexOf('compatibilityMode') > 0 }; }""")
    check('lsk.addressBook' in bk['keys'] and 'lsk.stock.avery-5160' in bk['keys'] and 'lsk.geoKey' not in bk['keys'] and 'lsk.tableDraft' not in bk['keys'] and 'lsk.addinRef' not in bk['keys'] and 'lsk.sheets' not in bk['keys'], 'backup has settings, not the key or passing things', bk['keys'])
    check('Address book: 1 address' in bk['doc'] and 'Printer adjustments: 1 label size' in bk['doc'] and 'Geoapify key: not included' in bk['doc'] and 'LabelSheetBackup' in bk['props'] and bk['we'] and bk['compat'], 'backup document describes itself and opens the panel', bk['doc'][:300])
    open(OUT + '/backup.docx', 'wb').write(__import__('base64').b64decode(pg.evaluate("MOCK.created[0].b64")))
    item_xml = bk['item']
    pg.check('#backup-key'); pg.click('#backup-make')
    wait(pg, "MOCK.created.length === 2")
    keys2 = pg.evaluate("""async () => { const z = await JSZip.loadAsync(MOCK.created[1].b64, {base64: true}); const item = await z.file('customXml/item1.xml').async('string');
      return Object.keys(JSON.parse(decodeURIComponent(escape(atob(/<lskBackup[^>]*>([^<]+)</.exec(item)[1])))).data); }""")
    check('lsk.geoKey' in keys2, 'key included when asked')
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (F1)', errs)
    pg.close()

    # Restore from the file, in a blank document on a "new Mac"
    pg, errs = new_page(b, 'blank=1')
    wait(pg, "!document.getElementById('v-none').hidden")
    check(pg.is_visible('#backup-none'), 'backup box in a blank document too')
    pg.click('#backup-none summary')
    check(pg.is_visible('#backup-make') and not pg.is_visible('.backup-head'), 'backup tools in their own box')
    pg.evaluate("localStorage.setItem('lsk.design', 'logo'); localStorage.setItem('lsk.geoKey', 'mykeymykeymykeymykey')")
    pg.set_input_files('#restore-file', OUT + '/backup.docx')
    check(wait(pg, "!document.getElementById('confirm').hidden") and 'replace the Label Sheet Keeper settings' in pg.inner_text('#confirm-msg'), 'restore asks first', pg.inner_text('#confirm-msg') if pg.is_visible('#confirm') else say(pg))
    pg.click('#confirm-yes')
    st_ = pg.evaluate("({ from: localStorage.getItem('lsk.fromAddress'), book: localStorage.getItem('lsk.addressBook'), design: localStorage.getItem('lsk.design'), key: localStorage.getItem('lsk.geoKey'), nudge: localStorage.getItem('lsk.nudge.avery-5160') })")
    check(st_['from'].startswith('Me') and 'Ann' in st_['book'] and st_['nudge'] == '{"x":0.03,"y":0}', 'settings restored', st_)
    check(st_['design'] is None and st_['key'] == 'mykeymykeymykeymykey', 'old settings replaced, own key kept when the backup has none', st_)
    pg.set_input_files('#restore-file', S + '/fixtures/list.xlsx')
    check(wait(pg, "document.getElementById('say').textContent.indexOf('isn’t a Label Sheet Keeper backup') >= 0 || document.getElementById('say').textContent.indexOf('couldn’t be read') >= 0"), 'other files are turned away', say(pg))
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (F2)', errs)
    pg.close()

    # Open the backup document itself
    pg = b.new_page(viewport={'width': 340, 'height': 900})
    errs = []
    pg.on('pageerror', lambda e: errs.append(str(e)))
    pg.add_init_script('window.MOCK_CUSTOM_XML = %s;' % json.dumps(item_xml.replace('<lskBackup', '<ns0:lskBackup xmlns:ns0="urn:lsk:backup"').replace('</lskBackup>', '</ns0:lskBackup>')))
    pg.route('**/office.js', lambda r: r.fulfill(body=MOCK, content_type='application/javascript'))
    pg.goto('http://127.0.0.1:%d/taskpane.html?backup=1' % PORT)
    check(wait(pg, "!document.getElementById('v-backup').hidden && !document.getElementById('backup-restore').hidden"), 'backup document shows the restore view')
    check('Address book: 1 address' in pg.inner_text('#backup-sum'), 'restore view lists what’s inside', pg.inner_text('#backup-sum'))
    pg.click('#backup-restore'); pg.click('#confirm-yes')
    check(pg.evaluate("localStorage.getItem('lsk.fromAddress')") and pg.is_visible('#backup-done'), 'restored from the open backup')
    check(not pg.is_visible('#backup-make'), 'no backup tools inside a backup')
    pg.screenshot(path=OUT + '/v15-backup.png', full_page=True)
    check(not errs, 'no script errors (F3)', errs)
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
