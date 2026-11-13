import subprocess, time, sys, json, os
from playwright.sync_api import sync_playwright

S = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # tests/
REPO = os.path.dirname(S)
OUT = os.path.join(S, 'out')
os.makedirs(OUT, exist_ok=True)
MOCK = open(S + '/mock-office.js').read()
PORT = 8790 + int(os.environ.get('LSK_PORT_SHIFT', '0'))   # shift lets several runs share one machine
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

try:
  with sync_playwright() as p:
    b = p.chromium.launch()

    # ================= Scenario 1: address labels, new layout =================
    pg, errs = new_page(b, 'fmt=avery-5160')
    check(wait(pg, "!document.getElementById('v-sheet').hidden"), 'sheet view shows')
    check(wait(pg, "MOCK.header.indexOf('▲') >= 0"), 'feed arrow added to header', pg.evaluate('MOCK.header'))
    check(wait(pg, "MOCK.fp && MOCK.fp.font.size === 1"), 'empty footer shrunk to a 1-point line', pg.evaluate("MOCK.fp && MOCK.fp.font.size"))
    check(pg.evaluate("MOCK.hp.font.size") == 6, 'header with the arrow left alone', pg.evaluate("MOCK.hp.font.size"))
    check('30' in pg.inner_text('#tally') and '0 of 3 trips' in pg.inner_text('#tally'), 'tally shows 30 open and trips', pg.inner_text('#tally'))
    check(pg.evaluate("!!MOCK.settings['Office.AutoShowTaskpaneWithDocument'] && !!MOCK.settings.lskSheetId"), 'doc tagged for auto-open and has an id')
    check('Top edge first' in pg.inner_text('#feed-text') or 'top edge first' in pg.inner_text('#feed-text'), 'feeding guide text', pg.inner_text('#feed-text'))

    pg.fill('#first-n', '3'); pg.click('#first-btn')
    check(wait(pg, "MOCK.labelCell(3).ccs.length === 1"), 'mark first 3 used')
    check(used(pg, 1) and used(pg, 2) and used(pg, 3), 'labels 1-3 locked')
    check(pg.evaluate("Object.keys(MOCK.settings.lskSpots||{}).length === 3"), 'spot history recorded for manual marks')
    pg.click('.cell[data-n="2"]')
    check(wait(pg, "MOCK.labelCell(2).ccs.length === 0"), 'click used label makes it open')

    pg.evaluate("MOCK.setLabelText(8, 'Jane Doe\\n12 Elm St\\nBoston, MA 02108', 10)")
    select_label(pg, 8); refresh(pg)
    check(pg.evaluate("LSK_TEST.st().ready[8] === true && LSK_TEST.st().here === 8"), 'typed label detected as ready, cursor in it')
    pg.evaluate("MOCK.labelCell(8).paras[2].leftIndent = 18")
    pg.fill('#copies', '2'); pg.click('#copy-btn')
    check(wait(pg, "MOCK.cellText(MOCK.labelCell(4)).indexOf('Jane Doe') === 0"), 'copy to next open labels (2 and 4)', [cell_text(pg, 2), cell_text(pg, 4)])
    check(pg.evaluate("MOCK.labelCell(2).paras.length === 3"), 'copy has the same 3 paragraphs', pg.evaluate("MOCK.labelCell(2).paras.length"))
    check(pg.evaluate("MOCK.calls.indexOf('cell.nestedTable') < 0"), 'copy puts no table inside a label', pg.evaluate("MOCK.calls"))
    check(pg.evaluate("MOCK.labelCell(4).paras[2].leftIndent === 18"), 'copy keeps the last line\'s layout', pg.evaluate("MOCK.labelCell(4).paras.map(p => p.leftIndent)"))

    # Fit text
    pg.evaluate("MOCK.setLabelText(9, 'Jane Doe Long Name Company Incorporated Worldwide\\n12345 Very Long Street Name Boulevard Suite 100\\nBoston, MA 02108\\nExtra line here\\nAnother line here', 12)")
    refresh(pg)
    check(pg.evaluate("LSK_TEST.st().over[9] === true"), 'overflow detected on label 9')
    check(pg.is_visible('#fit-warn') and 'label 9' in pg.inner_text('#fit-warn'), 'fit warning shown', pg.inner_text('#fit-warn') if pg.is_visible('#fit-warn') else 'hidden')
    pg.click('#fit-warn button')
    check(wait(pg, "!LSK_TEST.st().over[9] && MOCK.labelCell(9).paras[0].font.size < 12"), 'fit text shrinks label 9', pg.evaluate("MOCK.labelCell(9).paras.map(p=>p.font.size)"))

    # Print detection -> asks whether they printed; No leaves them ready
    time.sleep(1.2)
    ready_before = pg.evaluate("Object.keys(LSK_TEST.st().ready).length")
    pg.evaluate("MOCK.lastPrint = Date.now()")
    check(wait(pg, "!document.getElementById('confirm').hidden && /Word sent this sheet to print.*Did labels .* print\\?/.test(document.getElementById('confirm-msg').textContent)", 9000), 'print noticed: asks whether the labels printed', pg.inner_text('#confirm-msg'))
    check(pg.inner_text('#confirm-yes') == 'Yes, mark used' and pg.inner_text('#confirm-no') == 'No', 'answers are Yes, mark used and No', [pg.inner_text('#confirm-yes'), pg.inner_text('#confirm-no')])
    check(pg.evaluate("(MOCK.settings.lskTrips||0) === 0"), 'nothing is marked before the answer')
    pg.click('#confirm-no')
    time.sleep(0.6)
    check(pg.evaluate("(MOCK.settings.lskTrips||0) === 0 && Object.keys(LSK_TEST.st().ready).length === %d" % ready_before), 'No leaves the labels ready to print')
    # Yes marks them used, trips, history, undo
    time.sleep(1.2); pg.evaluate("MOCK.lastPrint = Date.now()")
    check(wait(pg, "!document.getElementById('confirm').hidden", 9000), 'the next print asks again')
    pg.click('#confirm-yes')
    check(wait(pg, "(MOCK.settings.lskTrips||0) === 1", 9000), 'Yes: trip counted')
    check(wait(pg, "Object.keys(LSK_TEST.st().ready).length === 0"), 'ready labels marked used after print', ready_before)
    check(pg.evaluate("(MOCK.settings.lskHistory||[]).length === 1"), 'history entry written')
    check(pg.evaluate("!!localStorage.getItem('lsk.last.avery-5160')"), 'last label saved for reprint')
    check(pg.is_visible('#undo-row'), 'undo offered')
    pg.click('#undo-btn')
    check(wait(pg, "(MOCK.settings.lskTrips||0) === 0 && Object.keys(LSK_TEST.st().ready).length === %d" % ready_before), 'undo restores labels and trip count')
    pg.click('#printed-btn')
    check(wait(pg, "(MOCK.settings.lskTrips||0) === 1"), 'I printed these counts a trip')
    check(not pg.is_checked('#auto-mark'), 'marking without asking is off unless chosen')

    # One-pass warning on second trip
    pg.evaluate("MOCK.setLabelText(12, 'Ann Lee\\n5 Pine Rd\\nSalem, MA 01970', 10)"); select_label(pg, 12); refresh(pg)
    check(pg.is_visible('#warn-pass'), 'Avery one-pass warning before second trip')
    pg.click('#warn-pass-hide')
    check(wait(pg, "document.getElementById('warn-pass').hidden"), 'warning can be hidden')

    # Last trip: use up the sheet
    pg.click('#settings summary'); pg.fill('#trip-limit', '2'); pg.dispatch_event('#trip-limit', 'change')
    check(wait(pg, "!document.getElementById('use-up').hidden"), 'use-up offered on last trip')
    check(not pg.is_visible('#use-up-mine'), 'return address option hidden when none saved')
    open_before = pg.evaluate("LSK_TEST.st() && 0") or pg.evaluate("(() => { let n=0; for (let i=1;i<=30;i++){ const s=LSK_TEST.st(); if(!s.used[i]&&!s.ready[i]) n++; } return n; })()")
    pg.click('#use-up-copy')
    check(wait(pg, "(() => { const s=LSK_TEST.st(); for (let i=1;i<=30;i++) if(!s.used[i]&&!s.ready[i]) return false; return true; })()"), 'use up the sheet fills every open label', open_before)
    check('2 of 2 trips' in pg.inner_text('#tally') or '1 of 2 trips' in pg.inner_text('#tally'), 'tally shows trips of limit', pg.inner_text('#tally'))
    pg.click('#printed-btn')
    check(wait(pg, "!document.getElementById('trips-out').hidden"), 'trip limit reached message')

    # Reset, then reprint last label
    pg.click('#reset-btn'); pg.click('#confirm-yes')
    check(wait(pg, "(MOCK.settings.lskTrips||0) === 0 && Object.keys(LSK_TEST.st().used).length === 0"), 'reset opens all and zeroes trips')
    check(pg.is_visible('#reprint-btn'), 'reprint offered', pg.inner_text('#reprint-btn') if pg.is_visible('#reprint-btn') else '')
    pg.click('#reprint-btn')
    check(wait(pg, "Object.keys(LSK_TEST.st().ready).length === 1"), 'reprint put the last label on an open label')

    # Damage and repair
    pg.evaluate("MOCK.setLabelText(5, 'Keep Me\\n1 Keep St\\nBoston, MA 02101', 10)")
    pg.fill('#first-n', '2'); pg.click('#first-btn'); wait(pg, "MOCK.labelCell(2).ccs.length === 1")
    refresh(pg)
    before = pg.evaluate("({used: Object.keys(LSK_TEST.st().used).length, ready: Object.keys(LSK_TEST.st().ready).length})")
    pg.evaluate("MOCK.table().rows.splice(9, 1)")   # delete the last row
    refresh(pg)
    check(wait(pg, "!document.getElementById('v-broken').hidden"), 'damaged sheet detected')
    check('Rows of labels' in pg.inner_text('#broken-msg'), 'damage explained', pg.inner_text('#broken-msg'))
    pg.click('#repair-btn')
    check(wait(pg, "!document.getElementById('v-sheet').hidden && MOCK.table().rows.length === 10"), 'repair rebuilt the table')
    after = pg.evaluate("({used: Object.keys(LSK_TEST.st().used).length, ready: Object.keys(LSK_TEST.st().ready).length})")
    check(after['used'] == before['used'], 'repair kept used labels', [before, after])
    check(after['ready'] >= before['ready'] - 0, 'repair kept ready labels', [before, after])
    check(pg.evaluate("MOCK.blocks.filter(b=>b.type==='para').length === 1 && MOCK.blocks[0].type === 'spacer'"), 'one trailing paragraph and spacer after repair', pg.evaluate("MOCK.blocks.map(b=>b.type)"))
    pg.evaluate("MOCK.labelCell(7).paras[0].level = 2")
    refresh(pg)
    check(wait(pg, "!document.getElementById('v-broken').hidden") and 'table inside' in pg.inner_text('#broken-msg'), 'table inside a label detected', pg.inner_text('#broken-msg') if pg.is_visible('#broken-msg') else '')
    pg.evaluate("MOCK.labelCell(7).paras[0].level = 1; MOCK.selHandler && MOCK.selHandler()")
    check(wait(pg, "!document.getElementById('v-sheet').hidden"), 'sheet view returns once the table is gone')
    W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    def pkgxml(body):
        return ('<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData>'
                '<w:document ' + W + '><w:body>' + body + '<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>')
    tc = lambda inner: '<w:tc><w:tcPr><w:tcW w:w="100"/></w:tcPr>' + inner + '</w:tc>'
    row = '<w:tbl><w:tblPr/><w:tr>' + tc('<w:p/>') + tc('<w:p/>') + tc('<w:p><w:r><w:t>Jane</w:t></w:r></w:p><w:p><w:r><w:t>1 Elm St</w:t></w:r></w:p>') + tc('<w:p/>') + '</w:tr></w:tbl><w:p/>'
    un = pg.evaluate("x => LSK_TOOLS.unwrapTable(x)", pkgxml(row))
    check('<w:tbl' not in un and un.count('<w:t>') == 2 and 'Jane' in un and '<w:sectPr' in un and 'tcPr' not in un, 'table around label text is removed', un[-400:])
    plain = pkgxml('<w:p><w:r><w:t>Jane</w:t></w:r></w:p><w:tbl><w:tr>' + tc('<w:p><w:r><w:t>x</w:t></w:r></w:p>') + '</w:tr></w:tbl>')
    check(pg.evaluate("x => LSK_TOOLS.unwrapTable(x) === x", plain), 'text with a table after it is left alone')
    pg.evaluate("MOCK.blocks.splice(1, 0, {type:'para', p:{text:'oops', level:0, font:{}}})")
    refresh(pg)
    check(wait(pg, "!document.getElementById('v-broken').hidden"), 'text outside labels detected')
    pg.click('#repair-btn'); wait(pg, "!document.getElementById('v-sheet').hidden")

    # Feed arrow toggle and printer adjustment
    pg.click('#settings summary') if not pg.evaluate("document.getElementById('settings').open") else None
    pg.uncheck('#feed-arrow')
    check(wait(pg, "MOCK.header === ''"), 'feed arrow removed when turned off')
    pg.check('#feed-arrow')
    check(wait(pg, "MOCK.header.indexOf('▲') >= 0"), 'feed arrow back when turned on')
    pg.fill('#nudge-x', '0.06'); pg.fill('#nudge-y', '-0.05'); pg.click('#nudge-save')
    check(pg.is_visible('#nudge-apply'), 'apply to this sheet offered')
    pg.click('#nudge-apply')
    check(wait(pg, "JSON.parse(MOCK.props.LabelSheetLayout).nx === 0.06"), 'adjustment recorded on the sheet', pg.evaluate("MOCK.props.LabelSheetLayout"))
    check(pg.evaluate("MOCK.table().indent === 86 && MOCK.blocks[0].type === 'spacer' && Number(MOCK.blocks[0].height) === 187"), 'labels moved right and up in place', pg.evaluate("[MOCK.table().indent, MOCK.blocks[0].height]"))

    # Address tab with a design, address book, import
    pg.click('#fill-tabs button:has-text("Address")')
    pg.fill('#addr-text', 'Pat Example\n8 Test Way\nBoston, MA 02108')
    pg.click('.design:has-text("Bold name")')
    pg.click('#addr-one')
    check(wait(pg, "Object.values(LSK_TEST.st().texts).some(t => t.indexOf('Pat Example') === 0)"), 'address placed on next open label')
    n_addr = pg.evaluate("Number(Object.keys(LSK_TEST.st().texts).find(k => LSK_TEST.st().texts[k].indexOf('Pat Example') === 0))")
    check(pg.evaluate("n => MOCK.labelCell(n).paras[0].font.bold === true && MOCK.labelCell(n).paras[1].font.bold === false", n_addr), 'bold name design applied')
    check(pg.evaluate("n => JSON.stringify(MOCK.labelCell(n).paras.map(p => p.text)) === JSON.stringify(['Pat Example', '8 Test Way', 'Boston, MA 02108'])", n_addr), 'address lines are separate paragraphs', pg.evaluate("n => MOCK.labelCell(n).paras.map(p => p.text)", n_addr))
    check(pg.evaluate("MOCK.calls.indexOf('cell.insertParagraph.End') < 0"), 'no paragraphs added at the end of a cell (Word for Mac joins them)')
    pg.click('#addr-save'); pg.click('#addr-mine')
    pg.click('#pane-addr details summary')
    pg.set_input_files('#book-file', S + '/fixtures/contacts.vcf')
    check(wait(pg, "!document.getElementById('imp').hidden") and pg.inner_text('#book-count') == '1', 'contacts file shows the list to choose from first')
    pg.click('#imp-all'); pg.click('#imp-go')
    check(wait(pg, "document.getElementById('book-count').textContent === '3'"), 'contacts imported into address book', pg.inner_text('#book-count'))
    pg.click('.design:has-text("Return to")')
    pg.select_option('#addr-pick', label='My return address')
    check('Pat Example' in pg.input_value('#addr-text'), 'my return address picked')
    pg.click('#addr-one')
    check(wait(pg, "Object.values(LSK_TEST.st().texts).some(t => t.indexOf('RETURN TO') === 0)"), 'return-to design placed')
    n_rt = pg.evaluate("Number(Object.keys(LSK_TEST.st().texts).find(k => LSK_TEST.st().texts[k].indexOf('RETURN TO') === 0))")
    rt = pg.evaluate("n => MOCK.labelCell(n).paras.map(p => [p.text, p.font.size, p.font.bold])", n_rt)
    check(len(rt) == 4 and rt[0][0] == 'RETURN TO' and rt[0][1] < 10 and rt[0][2] and rt[1][0] == 'Pat Example' and rt[1][1] == 10, 'return-to design: small bold heading, then the address', rt)

    # Table window: fill and spill to new sheets
    created_before = pg.evaluate("MOCK.created.length")
    pg.click('#fill-tabs button:has-text("Table")'); pg.click('#table-open')
    check(wait(pg, "!!MOCK.dialog"), 'table window opened', pg.evaluate("MOCK.dialog && MOCK.dialog.url"))
    open_now = pg.evaluate("(() => { let n=0; const s=LSK_TEST.st(); for (let i=1;i<=30;i++) if(!s.used[i]&&!s.ready[i]) n++; return n; })()")
    rows = [['Person %d' % i, '', '%d Main St' % i, '', 'Boston', 'MA', '2101'] for i in range(open_now + 5)]
    pg.evaluate("rows => MOCK.dialog.send({type:'fill', rows})", rows)
    check(wait(pg, "!document.getElementById('confirm').hidden"), 'spill-over asked', pg.inner_text('#confirm-msg') if pg.is_visible('#confirm') else '')
    pg.click('#confirm-yes')
    check(wait(pg, "MOCK.created.length === %d" % (created_before + 1), 10000), 'one new sheet made for the extra addresses')
    check(wait(pg, "(() => { const s=LSK_TEST.st(); for (let i=1;i<=30;i++) if(!s.used[i]&&!s.ready[i]) return false; return true; })()"), 'open labels filled from the table')
    check(pg.evaluate("MOCK.cellText(MOCK.labelCell(Number(Object.keys(LSK_TEST.st().texts).find(k => LSK_TEST.st().texts[k].indexOf('Person 0') === 0)))).indexOf('Boston, MA 02101') > 0"), 'table rows formatted as addresses with ZIP fixed')
    spill = pg.evaluate("""async () => { const d = MOCK.created[MOCK.created.length-1]; const z = await JSZip.loadAsync(d.b64, {base64:true});
      const doc = await z.file('word/document.xml').async('string'); const hdr = z.file('word/header1.xml') ? await z.file('word/header1.xml').async('string') : '';
      const we = !!z.file('word/webextensions/webextension1.xml'); const props = await z.file('docProps/custom.xml').async('string');
      const ftr = z.file('word/footer1.xml') ? await z.file('word/footer1.xml').async('string') : ''; const refs = /footerReference/.test(doc) && /headerReference/.test(doc);
      return { people: (doc.match(/Person \\d+/g)||[]).length, opened: d.opened, header: hdr.indexOf('▲') >= 0, we, spec: props.indexOf('LabelSheetSpec1') >= 0, nx: /&quot;nx&quot;:0.06/.test(props), footer: /w:line="20"/.test(ftr) && refs }; }""")
    check(spill['people'] == 5 and spill['opened'], 'new sheet holds the 5 extra addresses', spill)
    check(spill['header'] and spill['we'] and spill['spec'], 'new sheet has feed arrow, auto-open link and its own measurements', spill)
    check(spill['footer'], 'new sheet has a tiny footer', spill)

    # Your sheets, finish first, add a size
    reg = pg.evaluate("JSON.parse(localStorage.getItem('lsk.sheets')||'{}')")
    check(len(reg) == 1, 'this sheet is listed in your sheets', reg)
    pg.evaluate("""() => { const r = JSON.parse(localStorage.getItem('lsk.sheets')||'{}'); r.other = {id:'other', fmtId:'avery-5160', fmtName:'Address labels, 30 per sheet', file:'Old labels.docx', url:'/x/Old labels.docx', open:12, ready:0, trips:1, at:Date.now()}; localStorage.setItem('lsk.sheets', JSON.stringify(r)); }""")
    pg.click('#other-btn')
    check(wait(pg, "!document.getElementById('finish-first').hidden"), 'finish-first list shown when making another sheet')
    check('Old labels.docx' in pg.inner_text('#finish-list'), 'other sheet listed with its file', pg.inner_text('#finish-list'))
    created_before = pg.evaluate("MOCK.created.length")
    pg.click('#create-btn')
    check(wait(pg, "!document.getElementById('confirm').hidden && document.getElementById('confirm-msg').textContent.indexOf('Old labels.docx') >= 0"), 'asked to finish the old sheet first')
    pg.click('#confirm-yes')
    check(wait(pg, "MOCK.created.length === %d" % (created_before + 1)), 'new sheet made anyway')
    pg.click('#other-btn')
    pg.click('#add-size summary', position={'x': 30, 'y': 8})
    pg.fill('#sz-name', 'Test squares'); pg.fill('#sz-across', '4'); pg.fill('#sz-down', '5'); pg.fill('#sz-w', '1.5'); pg.fill('#sz-h', '1.5'); pg.fill('#sz-gapx', '0.5'); pg.fill('#sz-gapy', '0.5'); pg.fill('#sz-top', '0.75'); pg.fill('#sz-left', '0.75')
    pg.dispatch_event('#sz-left', 'input')
    check(not pg.is_disabled('#sz-save'), 'valid custom size can be saved', pg.inner_text('#sz-msg'))
    pg.fill('#sz-w', '3'); pg.dispatch_event('#sz-w', 'input')
    check(pg.is_disabled('#sz-save') and 'right edge' in pg.inner_text('#sz-msg'), 'too-wide size is caught', pg.inner_text('#sz-msg'))
    pg.fill('#sz-w', '1.5'); pg.dispatch_event('#sz-w', 'input')
    pg.click('#sz-save')
    check(wait(pg, "document.getElementById('fmt-list').textContent.indexOf('Test squares') >= 0"), 'custom size listed')
    pg.screenshot(path=OUT + '/s1-chooser.png', full_page=True)
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (scenario 1)', errs)
    pg.close()

    # ================= Scenario 2: older shipping sheet =================
    pg, errs = new_page(b, 'fmt=shipping-2-sideways&legacy=1')
    check(wait(pg, "!document.getElementById('v-sheet').hidden"), 'shipping sheet shows (older layout)')
    check(pg.evaluate("MOCK.header === ''"), 'no feed arrow on sheets without a top strip')
    check(wait(pg, "MOCK.hp && MOCK.hp.font.size === 1 && MOCK.fp && MOCK.fp.font.size === 1"), 'empty header and footer shrunk on a full-height sheet', pg.evaluate("[MOCK.hp && MOCK.hp.font.size, MOCK.fp && MOCK.fp.font.size]"))
    check(pg.is_hidden('#arrow-row'), 'arrow setting hidden for this size')
    pg.click('#side-bottom')
    check('bottom label' in pg.inner_text('#ship-target'), 'top/bottom naming', pg.inner_text('#ship-target'))
    pg.set_input_files('#ship-file', S + '/fixtures/sample-return-label.pdf')
    check(wait(pg, "document.getElementById('say').textContent.indexOf('Found the label') >= 0", 20000), 'smart crop found the label border', say(pg))
    crop = pg.evaluate("LSK_TEST.ship.crop")
    check(abs(crop['x'] / 300 - 0.75) < 0.05 and abs(crop['w'] / 300 - 4) < 0.1 and abs(crop['h'] / 300 - 6) < 0.1, 'crop matches the 4x6 label', crop)
    pg.click('#ship-place')
    check(wait(pg, "MOCK.labelCell(1).pics === 1"), 'label picture placed on label 1')
    # paste a screenshot
    pg.evaluate("""async () => { const c = document.createElement('canvas'); c.width = 400; c.height = 600; const x = c.getContext('2d'); x.fillStyle='#fff'; x.fillRect(0,0,400,600); x.lineWidth=6; x.strokeRect(40,40,320,520);
      const blob = await new Promise(r => c.toBlob(r, 'image/png')); const dt = new DataTransfer(); dt.items.add(new File([blob], 'shot.png', {type:'image/png'}));
      document.body.dispatchEvent(new ClipboardEvent('paste', {clipboardData: dt, bubbles: true})); }""")
    check(wait(pg, "document.getElementById('say').textContent.indexOf('Found the label') >= 0 && LSK_TEST.ship.src && LSK_TEST.ship.src.width === 400"), 'pasted screenshot opened and its border found', say(pg))
    pg.click('#fill-tabs button:has-text("Address")')
    pg.fill('#from-addr', 'Me\n1 My St\nBoston, MA 02101'); pg.fill('#to-addr', 'Returns Center\n100 Warehouse Rd\nCity, KY 40000')
    pg.click('#ship-addr-place')
    check(wait(pg, "MOCK.cellText(MOCK.labelCell(2)).indexOf('SHIP TO') >= 0"), 'shipping address label placed on label 2')
    sp = pg.evaluate("MOCK.labelCell(2).paras.map(p => [p.text, p.leftIndent, p.font.size])")
    check(len(sp) == 8 and sp[4][0] == 'SHIP TO' and sp[4][1] == 36 and sp[-1][0] == 'City, KY 40000' and sp[-1][1] == 36 and sp[-1][2] == 16, 'shipping label lines keep their own size and indent', sp)
    pg.evaluate("localStorage.setItem('lsk.markWithoutAsking', 'on')")
    time.sleep(1.2); pg.evaluate("MOCK.lastPrint = Date.now()")
    check(wait(pg, "(MOCK.settings.lskTrips||0) === 1", 9000), 'with marking without asking on, a print is counted straight away (shipping sheet)')
    check(pg.is_hidden('#confirm') and 'now marked used' in say(pg), 'and says so without asking', say(pg))
    check(pg.is_hidden('#use-up'), 'no use-up suggestion on shipping sheets')
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (scenario 2)', errs)
    pg.screenshot(path=OUT + '/s2-ship.png', full_page=True)
    pg.close()

    # ================= Scenario 3: blank document =================
    pg, errs = new_page(b, 'blank=1')
    check(wait(pg, "!document.getElementById('v-none').hidden"), 'blank document shows the chooser')
    pg.click('#fmt-list label:has-text("Return address labels, 80")')
    pg.click('#create-btn')
    check(wait(pg, "MOCK.created.length === 1 && MOCK.created[0].opened"), 'sheet created from blank document')
    check(wait(pg, "!document.getElementById('made').hidden"), 'confirmation shown')
    pg.click('#fmt-list label:has-text("Address labels, 30")')
    check(not [e for e in errs if 'favicon' not in e], 'no script errors (scenario 3)', errs)
    pg.close()

    # ================= Scenario 3b: table rows keep their place =================
    pg, errs = new_page(b, 'fmt=avery-5160')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.click('#fill-tabs button:has-text("Table")'); pg.click('#table-open')
    wait(pg, "!!MOCK.dialog")
    info = pg.evaluate("JSON.parse(localStorage.getItem('lsk.tableInfo'))")
    check(info and info.get('labels', [None])[0] == 1 and len(info['labels']) == 30, 'table window is told which labels are open', info)
    pg.evaluate("rows => MOCK.dialog.send({type:'fill', rows})", [['Johnny Appleseed', 'Apple', '123 Main St', '456', 'Townsville', 'NE', '01234'], ['', '', '', '', '', '', ''], ['Jane Roe', '', '9 Oak Ave', '', 'Salem', 'MA', '01970']])
    check(wait(pg, "MOCK.cellText(MOCK.labelCell(3)).indexOf('Jane Roe') === 0"), 'third row goes on label 3', [cell_text(pg, 1), cell_text(pg, 2), cell_text(pg, 3)])
    check(cell_text(pg, 2) == '', 'empty row leaves its label open', cell_text(pg, 2))
    lines = pg.evaluate("MOCK.labelCell(1).paras.map(p => p.text)")
    check(lines == ['Johnny Appleseed', 'Apple', '123 Main St, 456', 'Townsville, NE 01234'], 'name and company on their own lines', lines)
    pg.click('#settings summary')
    pg.select_option('#home-country', 'CA')
    pg.evaluate("rows => MOCK.dialog.send({type:'fill', rows})", [['', '', '', '', '', '', ''] for _ in range(3)] + [['Jane Roe', '', '123 Main St', '5', 'Toronto', 'ON', 'm5v3l9'], ['Ann Lee', '', '5 Pine Rd', '', 'Salem', 'MA', '1970']])
    # open labels are now 2, 4, 5, 6, 7...: rows 4 and 5 go on labels 6 and 7, blank rows leave 2, 4, 5 open
    check(wait(pg, "MOCK.cellText(MOCK.labelCell(7)).indexOf('Ann Lee') === 0"), 'rows after blank rows land on their labels', [cell_text(pg, n) for n in (2, 4, 5, 6, 7)])
    check(cell_text(pg, 2) == '' and cell_text(pg, 4) == '' and cell_text(pg, 5) == '', 'blank rows add nothing, not even a country line', [cell_text(pg, n) for n in (2, 4, 5)])
    check(pg.evaluate("MOCK.labelCell(6).paras.map(p => p.text)") == ['Jane Roe', '5-123 Main St', 'Toronto ON  M5V 3L9'], 'Canadian address from Canada: unit-civic, no country line', pg.evaluate("MOCK.labelCell(6).paras.map(p => p.text)"))
    check(pg.evaluate("MOCK.labelCell(7).paras.map(p => p.text)")[-1] == 'USA', 'US address from Canada gets a USA line', pg.evaluate("MOCK.labelCell(7).paras.map(p => p.text)"))
    pg.select_option('#home-country', 'US')
    pg.fill('#geo-key', 'badkeybadkeybadkey'); pg.click('#geo-save')
    check(wait(pg, "document.getElementById('geo-status').textContent.indexOf('accept that key') >= 0") and not pg.evaluate("localStorage.getItem('lsk.geoKey')"), 'a rejected key is not saved', pg.inner_text('#geo-status'))
    pg.fill('#geo-key', '0123456789abcdef0123456789abcdef'); pg.click('#geo-save')
    check(wait(pg, "localStorage.getItem('lsk.geoKey') === '0123456789abcdef0123456789abcdef'") and 'Key saved' in pg.inner_text('#geo-status'), 'a good key is saved on this Mac', pg.inner_text('#geo-status'))
    pg.click('#geo-remove')
    check(not pg.evaluate("localStorage.getItem('lsk.geoKey')"), 'key can be removed')
    pg.uncheck('#suggest-on')
    check(pg.evaluate("localStorage.getItem('lsk.suggest')") == 'off', 'suggestions can be turned off')
    pg.check('#suggest-on')
    check(pg.evaluate("localStorage.getItem('lsk.suggest')") is None, 'and on again')
    check(not [e for e in errs if 'favicon' not in e and 'status of 401' not in e], 'no script errors (scenario 3b)', errs)
    pg.close()

    # ================= Scenario 4: the table window =================
    tp = b.new_page(viewport={'width': 900, 'height': 640})
    fake_services(tp)
    terrs = []
    tp.on('pageerror', lambda e: terrs.append(str(e)))
    tp.route('**/office.js', lambda r: r.fulfill(body='', content_type='application/javascript'))
    tp.goto('http://127.0.0.1:%d/table.html' % PORT)
    tp.evaluate("localStorage.setItem('lsk.tableInfo', JSON.stringify({open: 3, labels: [2, 5, 9], name: 'Address labels, 30 per sheet', per: 30})); localStorage.setItem('lsk.tableDraft','[]');")
    tp.reload(); tp.wait_for_timeout(300)
    check(tp.evaluate("document.querySelectorAll('#rows tr.row').length") == 3 and not tp.query_selector('#rows tr.divider'), 'table starts with one row per open label')
    tp.click('#more')
    nums = tp.evaluate("Array.from(document.querySelectorAll('#rows tr.row')).slice(0, 5).map(tr => tr.firstChild.textContent)")
    check(nums == ['2', '5', '9', '1', '2'], 'first column shows the label each row goes on', nums)
    div = tp.evaluate("(() => { const d = document.querySelector('#rows tr.divider'); return d ? [d.textContent, Array.from(d.parentNode.children).indexOf(d)] : null; })()")
    check(div and 'Extra sheet 1' in div[0] and div[1] == 3, 'divider marks where rows stop fitting', div)
    tp.evaluate("""() => { const inp = document.querySelector('#rows input'); inp.focus(); const dt = new DataTransfer(); dt.setData('text/plain', 'Jane Doe\\tAcme\\t12 Elm St\\t\\tBoston\\tMA\\t02108\\nJohn Roe\\t\\t9 Oak Ave\\tApt 2\\tCambridge\\tMA\\t02139\\n');
      inp.dispatchEvent(new ClipboardEvent('paste', {clipboardData: dt, bubbles: true})); }""")
    rows = tp.evaluate("LSK_TABLE.readRows()")
    check(len(rows) == 2 and rows[1][3] == 'Apt 2' and rows[0][1] == 'Acme', 'pasting cells from Excel fills the table', rows)
    tp.set_input_files('#file', S + '/fixtures/list.xlsx')
    check(wait(tp, "LSK_TABLE.readRows().length === 4"), 'Excel file added with headings matched', tp.evaluate("LSK_TABLE.readRows()"))
    rows = tp.evaluate("LSK_TABLE.readRows()")
    check(rows[2][0] == 'Cara Diaz' and rows[2][6] == '02169' and rows[3][4] == 'Medford', 'Excel columns mapped and ZIP leading zero restored', rows[2:])
    check('1 row doesn’t fit on it' in tp.inner_text('#count'), 'count shows rows that don’t fit', tp.inner_text('#count'))
    # ZIP and postal code fill in the city
    tp.evaluate("localStorage.removeItem('lsk.geoKey')")
    tp.evaluate("() => { const r = document.querySelectorAll('#rows tr.row')[5].querySelectorAll('input'); r[0].value = 'Zip Test'; r[6].value = '02108'; r[6].dispatchEvent(new Event('input', {bubbles: true})); }")
    check(wait(tp, "(() => { const r = document.querySelectorAll('#rows tr.row')[5].querySelectorAll('input'); return r[4].value === 'Boston' && r[5].value === 'MA'; })()"), 'ZIP code fills city and state')
    tp.evaluate("() => { const r = document.querySelectorAll('#rows tr.row')[6].querySelectorAll('input'); r[0].value = 'Postal Test'; r[6].value = 'm5v3l9'; r[6].dispatchEvent(new Event('input', {bubbles: true})); }")
    check(wait(tp, "(() => { const r = document.querySelectorAll('#rows tr.row')[6].querySelectorAll('input'); return r[4].value === 'Toronto' && r[5].value === 'ON' && r[6].value === 'M5V 3L9'; })()"), 'postal code fills city and province and tidies the code',
          tp.evaluate("Array.from(document.querySelectorAll('#rows tr.row')[6].querySelectorAll('input')).map(i => i.value)"))
    # Suggestions without a key come from Photon
    tp.click('#more')
    tp.evaluate("document.querySelectorAll('#rows tr.row')[8].querySelector('input[aria-label=street]').focus()")
    tp.keyboard.type('8 Maple')
    check(wait(tp, "!document.getElementById('sugg').hidden && document.querySelectorAll('#sugg-list button').length === 2"), 'Photon suggestions appear without a key (US and Canada only, no duplicates)',
          tp.evaluate("Array.from(document.querySelectorAll('#sugg-list button')).map(b => b.textContent)"))
    check('Photon by Komoot' in tp.inner_text('#sugg') and 'OpenStreetMap' in tp.inner_text('#sugg'), 'Photon and OpenStreetMap credited')
    check(any('photon.komoot.io' in u and 'countrycode=US' in u for u in GEO_CALLS), 'Photon asked for US and Canada only')
    tp.keyboard.press('ArrowDown'); tp.keyboard.press('ArrowDown'); tp.keyboard.press('Enter')
    vals = tp.evaluate("Array.from(document.querySelectorAll('#rows tr.row')[8].querySelectorAll('input')).map(i => i.value)")
    check(vals[2] == '8 Maple Rue' and vals[4] == 'Montréal' and vals[5] == 'QC' and vals[6] == 'H2X 1Y4', 'Photon suggestion fills a Canadian row with province code and postal code', vals)
    tp.evaluate("localStorage.setItem('lsk.suggest', 'off')")
    tp.evaluate("document.querySelectorAll('#rows tr.row')[9].querySelector('input[aria-label=street]').focus()")
    tp.keyboard.type('8 Maple')
    tp.wait_for_timeout(700)
    check(tp.evaluate("document.getElementById('sugg').hidden"), 'no suggestions when turned off in Settings')
    tp.evaluate("localStorage.removeItem('lsk.suggest')")
    # Suggestions with a key come from Geoapify
    tp.evaluate("localStorage.setItem('lsk.geoKey', '0123456789abcdef0123456789abcdef')")
    tp.evaluate("document.querySelectorAll('#rows tr.row')[7].querySelector('input[aria-label=street]').focus()")
    tp.keyboard.type('8 Maple')
    check(wait(tp, "!document.getElementById('sugg').hidden && document.querySelectorAll('#sugg-list button').length === 2"), 'street suggestions appear')
    check('Powered by Geoapify' in tp.inner_text('#sugg'), 'Geoapify credit shown with its suggestions')
    tp.screenshot(path=OUT + '/s4-sugg.png')
    tp.keyboard.press('ArrowDown'); tp.keyboard.press('Enter')
    vals = tp.evaluate("Array.from(document.querySelectorAll('#rows tr.row')[7].querySelectorAll('input')).map(i => i.value)")
    check(vals[2] == '8 Maple Way' and vals[4] == 'Boston' and vals[5] == 'MA' and vals[6] == '02108' and tp.evaluate("document.getElementById('sugg').hidden"), 'choosing a suggestion fills the row', vals)
    check(tp.evaluate("document.activeElement.getAttribute('aria-label')") == 'apt', 'cursor moves to Apt/Suite after choosing')
    tp.screenshot(path=OUT + '/s4-table.png')
    check(not terrs, 'no script errors (table window)', terrs)
    tp.close()
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
