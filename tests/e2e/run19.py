import subprocess, time, sys, json, os
from playwright.sync_api import sync_playwright

S = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # tests/
REPO = os.path.dirname(S)
OUT = os.path.join(S, 'out')
os.makedirs(OUT, exist_ok=True)
MOCK = open(S + '/mock-office.js').read()
PORT = 8798
results = []

def check(cond, name, detail=''):
    results.append((bool(cond), name, detail))
    print(('PASS ' if cond else 'FAIL ') + name + ('' if cond else '  -- ' + str(detail)))

srv = subprocess.Popen(['python3', '-m', 'http.server', str(PORT), '--bind', '127.0.0.1'], cwd=REPO, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
time.sleep(1)

def watch_errors(pg):
    errs = []
    pg.on('pageerror', lambda e: errs.append(str(e)))
    return errs

def new_page(b, query):
    pg = b.new_page(viewport={'width': 340, 'height': 1000})
    errs = watch_errors(pg)
    pg.route('**/office.js', lambda r: r.fulfill(body=MOCK, content_type='application/javascript'))
    pg.goto('http://127.0.0.1:%d/taskpane.html?%s' % (PORT, query))
    return pg, errs

def wait(pg, js, timeout=8000, arg=None):
    try:
        pg.wait_for_function(js, arg=arg, timeout=timeout); return True
    except Exception:
        return False

def say(pg):
    return pg.inner_text('#say') if pg.is_visible('#say') else ''

def book(pg):
    return pg.evaluate("JSON.parse(localStorage.getItem('lsk.addressBook') || '[]')")

def label_with(pg, start):
    """Text of each paragraph on the label whose text starts with `start` (or None)."""
    return pg.evaluate("""s => { const t = LSK_TEST.st().texts; const k = Object.keys(t).find(k => t[k].indexOf(s) === 0);
      return k ? MOCK.labelCell(Number(k)).paras.map(p => p.text) : null; }""", start)

def label_num(pg, start):
    return pg.evaluate("s => { const t = LSK_TEST.st().texts; return Number(Object.keys(t).find(k => t[k].indexOf(s) === 0) || 0); }", start)

def type_address(pg, text):
    pg.fill('#addr-text', text)
    pg.dispatch_event('#addr-text', 'input')

def open_book(pg):
    pg.evaluate("document.querySelector('#imp').closest('details').open = true")

def placed(pg, start):
    return wait(pg, "s => Object.values(LSK_TEST.st().texts).some(t => t.indexOf(s) === 0)", arg=start)

try:
  with sync_playwright() as p:
    b = p.chromium.launch()

    # ================= A: a contact with a second address, two phones, two emails =================
    pg, errs = new_page(b, 'fmt=avery-5160')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.click('#fill-tabs button:has-text("Address")')
    check(pg.is_visible('#x-phone') and pg.is_visible('#x-email') and pg.is_hidden('#x-phone2-row') and pg.is_hidden('#addr-which'), 'the address box has Phone and Email fields, nothing more until needed')
    check(pg.evaluate("document.getElementById('addr-text').nextElementSibling.id") == 'addr-extras', 'they sit right under the address')
    type_address(pg, 'Jane Doe\n12 Elm St\nBoston, MA 02108')
    pg.click('#addr-save')
    check(len(book(pg)) == 1, 'address saved', book(pg))
    open_book(pg)
    pg.click('#book-list .addr')
    check(wait(pg, "!!document.getElementById('book-edit-2')"), 'the contact editor has a second address box')
    pg.fill('#book-edit-2', 'PO Box 44\nBoston, MA 02101')
    pg.fill('#book-edit-phone', '617-555-0100'); pg.fill('#book-edit-phone2', '617-555-0111')
    pg.fill('#book-edit-email', 'jane@example.com'); pg.fill('#book-edit-email2', 'jane@work.example')
    pg.click('.book-edit .btn.primary')
    e = book(pg)[0]
    check(e.get('lines2') == ['Jane Doe', 'PO Box 44', 'Boston, MA 02101'], 'a second address typed without the name gets it', e.get('lines2'))
    check(e.get('phone') == '617-555-0100' and e.get('phone2') == '617-555-0111' and e.get('email') == 'jane@example.com' and e.get('email2') == 'jane@work.example', 'phones and emails saved with the contact', e)
    jane = e['id']

    # ================= B: picking her: details come along, nothing ticked yet =================
    pg.select_option('#addr-pick', value=jane)
    check(pg.input_value('#x-phone') == '617-555-0100' and pg.input_value('#x-email2') == 'jane@work.example', 'picking a contact fills in her phones and emails')
    check(pg.is_visible('#x-phone2-row') and pg.is_visible('#x-email2-row') and pg.is_visible('#addr-which'), 'her second phone, email and address appear')
    check(not any(pg.is_checked('#x-%s-on' % k) for k in ('phone', 'phone2', 'email', 'email2')), 'phone and email start off the label (the sheet’s choice)')
    check(pg.is_hidden('#x-keep'), 'nothing to remember yet')
    pg.check('#x-phone-on')
    check(pg.is_visible('#x-sheet') and pg.is_visible('#x-contact') and pg.inner_text('#x-contact') == 'Remember for Jane Doe', 'a changed choice can be kept for the sheet or for her')
    pg.click('#addr-one')
    check(placed(pg, 'Jane Doe'), 'placed')
    lab = label_with(pg, 'Jane Doe')
    check(lab == ['Jane Doe', '12 Elm St', 'Boston, MA 02108', '617-555-0100'], 'the ticked phone goes under the address, the email doesn’t', lab)
    rec = pg.evaluate("JSON.parse(localStorage.getItem('lsk.recent'))[0]")
    check(rec['lines'] == ['Jane Doe', '12 Elm St', 'Boston, MA 02108'] and rec['x']['on']['phone'] and rec['x']['id'] == jane, 'Recent keeps the address and what went with it', rec)

    # ================= C: her second address =================
    pg.click('#addr-which-2')
    check(pg.input_value('#addr-text') == 'Jane Doe\nPO Box 44\nBoston, MA 02101' and pg.get_attribute('#addr-which-2', 'aria-checked') == 'true', 'Address 2 puts her PO box in the box')
    pg.click('#addr-one')
    check(wait(pg, "Object.values(LSK_TEST.st().texts).some(t => t.indexOf('PO Box 44') > 0)"), 'and on the label')

    # ================= D: Use for this sheet =================
    pg.click('#x-sheet')
    check(pg.evaluate("MOCK.settings.lskOnLabel") == {'phone': True, 'email': False}, 'the sheet remembers phone on, email off', pg.evaluate("MOCK.settings.lskOnLabel"))
    check(pg.is_hidden('#x-sheet') and 'Saved for this sheet' in say(pg), 'and says so', say(pg))
    type_address(pg, '')
    check(pg.input_value('#x-phone') == '' and pg.is_hidden('#addr-which'), 'clearing the address box clears what came with it')
    type_address(pg, 'Bob Ray\n9 Oak Ave\nCambridge, MA 02139')
    pg.fill('#x-phone', '978-555-0123'); pg.dispatch_event('#x-phone', 'input')
    pg.fill('#x-email', 'bob@example.com'); pg.dispatch_event('#x-email', 'input')
    check(pg.is_checked('#x-phone-on') and not pg.is_checked('#x-email-on'), 'a phone typed in is ticked because this sheet puts phones on labels')
    pg.click('#addr-save')
    bob = [x for x in book(pg) if x['name'] == 'Bob Ray']
    check(bob and bob[0].get('phone') == '978-555-0123' and bob[0].get('email') == 'bob@example.com', 'Save to address book keeps the phone and email', bob)

    # ================= E: Remember for one contact =================
    pg.select_option('#addr-pick', value=jane)
    check(pg.is_checked('#x-phone-on'), 'picking her now starts from the sheet’s choice (phone on)')
    pg.uncheck('#x-phone-on'); pg.check('#x-email-on'); pg.click('#addr-which-2')
    pg.click('#x-contact')
    e = [x for x in book(pg) if x['id'] == jane][0]
    check(e.get('show') == {'addr': 2, 'phone': False, 'phone2': False, 'email': True, 'email2': False}, 'her own choice is saved: PO box, email, no phone', e.get('show'))
    check('Saved for Jane Doe' in say(pg) and pg.is_hidden('#x-contact'), 'and says so', say(pg))
    type_address(pg, '')
    pg.select_option('#addr-pick', value=jane)
    check(pg.input_value('#addr-text').startswith('Jane Doe\nPO Box 44') and pg.is_checked('#x-email-on') and not pg.is_checked('#x-phone-on'), 'next time she comes back that way, whatever the sheet says')
    pg.fill('#x-phone2', ''); pg.dispatch_event('#x-phone2', 'input')
    pg.click('#x-contact')
    e = [x for x in book(pg) if x['id'] == jane][0]
    check('phone2' not in e, 'clearing a phone and remembering removes it from the contact', e)
    # A group fill uses each contact's own choice.
    open_book(pg)
    for who in ('Jane Doe', 'Bob Ray'):
        pg.click('#book-list .addr:text-matches("^%s")' % who)
        wait(pg, "!!document.getElementById('book-edit-groups')")
        pg.fill('#book-edit-groups', 'Family'); pg.press('#book-edit-groups', 'Enter')
    pg.select_option('#group-pick', value='Family')
    if wait(pg, "!document.getElementById('group-fill').hidden"):
        pg.click('#group-fill')
        check(wait(pg, "Object.values(LSK_TEST.st().texts).filter(t => t.indexOf('Jane Doe') === 0).length >= 3"), 'Put on open labels fills the group')
        texts = pg.evaluate("Object.values(LSK_TEST.st().texts)")
        check(any('PO Box 44' in t and 'jane@example.com' in t and '617-555-0100' not in t for t in texts) and any(t.startswith('Bob Ray') and '978-555-0123' in t for t in texts),
              'each contact as remembered: Jane’s PO box and email; Bob with the sheet’s phone', [t for t in texts if t.startswith(('Jane', 'Bob'))])
    else:
        check(False, 'group fill button shows for a group')

    # ================= F: too long: warn, then ask to shrink =================
    pg.click('#confirm-no') if pg.is_visible('#confirm') else None
    type_address(pg, '')
    type_address(pg, 'Dr. Alexandra Montgomery-Whitfield\nAcme Widgets International Inc.\n1234 Very Long Street Name Avenue\nSuite 500, Building C\nBoston, MA 02108')
    pg.fill('#x-phone', '617-555-0199'); pg.dispatch_event('#x-phone', 'input')
    pg.fill('#x-email', 'alexandra.montgomery@example.com'); pg.dispatch_event('#x-email', 'input'); pg.check('#x-email-on')
    pg.click('#addr-one')
    check(wait(pg, "!document.getElementById('confirm').hidden", 10000), 'text that may be cut off is pointed out right away')
    check('may be cut off' in pg.inner_text('#confirm-msg') and 'Shrink it to fit?' in pg.inner_text('#confirm-msg') and pg.inner_text('#confirm-yes') == 'Shrink', 'and the panel asks before shrinking', pg.inner_text('#confirm-msg'))
    n_long = label_num(pg, 'Dr. Alexandra')
    sizes0 = pg.evaluate("n => MOCK.labelCell(n).paras.map(p => p.font.size)", n_long)
    pg.click('#confirm-yes')
    check(wait(pg, "/Shrank|too much text/.test(document.getElementById('say').textContent)", 10000), 'Shrink fits it (or says it can’t at a readable size)', say(pg))
    sizes1 = pg.evaluate("n => MOCK.labelCell(n).paras.map(p => p.font.size)", n_long)
    check(('Shrank' in say(pg) and max(sizes1) < max(sizes0) and min(sizes1) >= 7) or 'too much text' in say(pg), 'never below a readable size', [sizes0, sizes1, say(pg)])
    pg.screenshot(path=OUT + '/v19-address.png', full_page=True)

    # ================= G: choosing contacts to import =================
    pg.set_input_files('#book-file', S + '/fixtures/contacts-more.vcf')
    check(wait(pg, "!document.getElementById('imp').hidden"), 'a contacts file shows the list first')
    check('(2)' in pg.inner_text('#imp-head') and '1 contact without a street address was left out' in pg.inner_text('#imp-note'), 'with how many, and who was left out', [pg.inner_text('#imp-head'), pg.inner_text('#imp-note')])
    items = pg.evaluate("Array.from(document.querySelectorAll('#imp-list li')).map(li => [li.querySelector('span').textContent, li.querySelector('small').textContent, li.querySelector('input').checked])")
    check(len(items) == 2 and not any(i[2] for i in items), 'nothing is ticked to start with', items)
    check('already saved' in items[0][1] and '2 addresses' in items[1][1] and 'sam@example.com' in items[1][1], 'each shows its address, details and whether it’s already saved', items)
    check(pg.is_disabled('#imp-go'), 'Import waits until something is ticked')
    pg.fill('#imp-filter', 'sam'); pg.dispatch_event('#imp-filter', 'input')
    check(pg.evaluate("document.querySelectorAll('#imp-list li').length") == 1, 'search narrows the list')
    pg.click('#imp-all')
    check(pg.inner_text('#imp-go') == 'Import 1 contact', 'Select all ticks only the ones shown', pg.inner_text('#imp-go'))
    pg.fill('#imp-filter', ''); pg.dispatch_event('#imp-filter', 'input')
    before = len(book(pg))
    pg.click('#imp-go')
    sam = [x for x in book(pg) if x['name'] == 'Sam Lee']
    check(len(book(pg)) == before + 1 and sam and sam[0].get('lines2') == ['Sam Lee', 'PO Box 77', 'Salem, MA 01971'] and sam[0].get('email2') == 'sam@work.example', 'only the ticked contact comes in, with both addresses and emails', sam)
    check(pg.is_hidden('#imp') and 'Imported 1 contact' in say(pg), 'the list closes and the panel says what came in', say(pg))
    pg.set_input_files('#book-file', S + '/fixtures/contacts-more.vcf')
    wait(pg, "!document.getElementById('imp').hidden")
    pg.click('#imp-list li label')   # Jane, already saved
    pg.click('#imp-go')
    e = [x for x in book(pg) if x['id'] == jane][0]
    check(e.get('phone2') == '617-555-0111' and 'details to 1 contact already saved' in say(pg), 'importing someone already saved fills in what she was missing', [e, say(pg)])
    pg.set_input_files('#book-file', S + '/fixtures/contacts-more.vcf')
    wait(pg, "!document.getElementById('imp').hidden")
    pg.click('#imp-cancel')
    check(pg.is_hidden('#imp') and 'Nothing was imported' in say(pg), 'Cancel imports nothing')

    # ================= H: the table window's Phone and Email columns =================
    created = pg.evaluate("MOCK.created.length")
    pg.evaluate("rows => MOCK.dialog = null", [])
    pg.evaluate("""() => MOCK.dialog = null""")
    rows = [['Table Person', '', '1 A St', '', 'Boston', 'MA', '02108', '617-555-0142', 'tp@example.com'],
            ['No Phone', '', '2 B St', '', 'Boston', 'MA', '02108', '', 'np@example.com']]
    pg.click('#fill-tabs button:has-text("Table")'); pg.click('#table-open')
    check(wait(pg, "!!MOCK.dialog"), 'table window opened')
    info = pg.evaluate("JSON.parse(localStorage.getItem('lsk.tableInfo'))")
    check(info.get('onLabel') == {'phone': True, 'email': False}, 'the table window gets the sheet’s phone and email choice', info)
    pg.evaluate("rows => MOCK.dialog.send({type:'fill', rows, show: {phone: true, email: false}})", rows)
    check(placed(pg, 'Table Person'), 'table rows filled')
    check(label_with(pg, 'Table Person')[-1] == '617-555-0142' and 'tp@example.com' not in ' '.join(label_with(pg, 'Table Person')), 'the table’s Phone column goes on, Email doesn’t', label_with(pg, 'Table Person'))
    check(label_with(pg, 'No Phone') == ['No Phone', '2 B St', 'Boston, MA 02108'], 'a row with no phone is just the address', label_with(pg, 'No Phone'))
    check(not errs, 'no script errors (panel)', errs)
    pg.close()

    tp = b.new_page(viewport={'width': 1100, 'height': 640})
    terrs = watch_errors(tp)
    tp.route('**/office.js', lambda r: r.fulfill(body='', content_type='application/javascript'))
    tp.goto('http://127.0.0.1:%d/table.html' % PORT)
    tp.evaluate("""() => { localStorage.setItem('lsk.tableInfo', JSON.stringify({open: 3, labels: [1, 2, 3], name: 'Address labels', per: 30, onLabel: {phone: true, email: false}}));
      localStorage.setItem('lsk.tableDraft', '[]');
      localStorage.setItem('lsk.addressBook', JSON.stringify([{id: 'j', name: 'Jane Doe', lines: ['Jane Doe', '12 Elm St', 'Boston, MA 02108'], phone: '617-555-0100', email: 'jane@example.com'}])); }""")
    tp.reload(); tp.wait_for_timeout(300)
    heads = tp.evaluate("Array.from(document.querySelectorAll('#grid th')).map(t => t.textContent)")
    check('Phone' in heads and 'Email' in heads, 'the table has Phone and Email columns', heads)
    check(tp.is_checked('#on-phone') and not tp.is_checked('#on-email'), 'its On labels boxes start from the sheet’s choice')
    tp.select_option('#book', value='j')
    rows = tp.evaluate("LSK_TABLE.readRows()")
    check(rows and rows[0][7] == '617-555-0100' and rows[0][8] == 'jane@example.com', 'a saved contact comes in with phone and email', rows)
    tp.set_input_files('#file', S + '/fixtures/contacts-more.vcf')
    check(wait(tp, "LSK_TABLE.readRows().length === 3"), 'a contacts file adds rows')
    rows = tp.evaluate("LSK_TABLE.readRows()")
    check(rows[2][0] == 'Sam Lee' and rows[2][7] == '978-555-0123' and rows[2][8] == 'sam@example.com', 'with phones and emails', rows)
    tp.screenshot(path=OUT + '/v19-table.png')
    # Address suggestions for everyone: Geoapify through the server, Photon when that's used up.
    asked = []
    server = {'status': 200}
    def worker_route(route):
        asked.append(route.request.url)
        if server['status'] != 200:
            return route.fulfill(status=server['status'], headers={'Access-Control-Allow-Origin': '*'}, json={'error': 'Address suggestions are used up for today.'})
        route.fulfill(headers={'Access-Control-Allow-Origin': '*'}, json={'results': [
            {'housenumber': '8', 'street': 'Maple Way', 'city': 'Boston', 'state_code': 'MA', 'postcode': '02108', 'country_code': 'us', 'address_line1': '8 Maple Way', 'address_line2': 'Boston, MA 02108'}]})
    photon_asked = []
    def photon_route(route):
        photon_asked.append(route.request.url)
        route.fulfill(json={'type': 'FeatureCollection', 'features': [
            {'type': 'Feature', 'properties': {'type': 'house', 'housenumber': '9', 'street': 'Elm Street', 'city': 'Salem', 'state': 'Massachusetts', 'postcode': '01970', 'countrycode': 'US'}}]})
    tp.route('https://api.labelsheetkeeper.workers.dev/**', worker_route)
    tp.route('https://photon.komoot.io/**', photon_route)
    tp.route('https://api.geoapify.com/**', lambda r: (asked.append('DIRECT ' + r.request.url), r.fulfill(status=500, json={})))
    tp.evaluate("localStorage.removeItem('lsk.geoKey')")
    tp.reload(); tp.wait_for_timeout(300)
    tp.click('#more')
    street_at = lambda i: tp.evaluate("i => { const r = document.querySelectorAll('#rows tr.row'); const el = r[i].querySelector('input[aria-label=street]'); r[i].querySelectorAll('input').forEach(x => x.value = ''); el.focus(); return r.length; }", i)
    street_at(4)
    tp.keyboard.type('8 Maple')
    check(wait(tp, "!document.getElementById('sugg').hidden && document.querySelectorAll('#sugg-list button').length === 1"), 'without a key of your own, suggestions come through the Label Sheet Keeper server')
    check(asked and asked[0].startswith('https://api.labelsheetkeeper.workers.dev/suggest?text=8%20Maple') and not any(a.startswith('DIRECT') for a in asked), 'the table window never needs Geoapify’s key itself', asked)
    check('Powered by Geoapify' in tp.inner_text('#sugg') and 'OpenStreetMap' in tp.inner_text('#sugg'), 'and credit Geoapify and OpenStreetMap', tp.inner_text('#sugg'))
    tp.keyboard.press('ArrowDown'); tp.keyboard.press('Enter')
    vals = tp.evaluate("Array.from(document.querySelectorAll('#rows tr.row')[4].querySelectorAll('input')).map(i => i.value)")
    check(vals[2] == '8 Maple Way' and vals[4] == 'Boston' and vals[5] == 'MA' and vals[6] == '02108', 'picking one fills the row', vals)
    server['status'] = 429
    street_at(5)
    tp.keyboard.type('9 Elm St')
    check(wait(tp, "!document.getElementById('sugg').hidden && document.getElementById('sugg').textContent.indexOf('Photon') >= 0"), 'when the shared ones are used up for the day, Photon takes over', tp.inner_text('#sugg') if tp.is_visible('#sugg') else '')
    n_asked = len(asked)
    tp.keyboard.press('Escape')
    street_at(6)
    tp.keyboard.type('9 Elm Ave')
    check(wait(tp, "!document.getElementById('sugg').hidden") and len(asked) == n_asked and len(photon_asked) >= 2, 'and the server isn’t asked again that session', [len(asked), n_asked, len(photon_asked)])
    check(not terrs, 'no script errors (table)', terrs)
    tp.close()

    # ================= I: Ship to on shipping sheets =================
    pg, errs = new_page(b, 'fmt=shipping-2-sideways')
    wait(pg, "!document.getElementById('v-sheet').hidden")
    pg.evaluate("""() => localStorage.setItem('lsk.addressBook', JSON.stringify([{id: 's', name: 'Sam Lee', lines: ['Sam Lee', '5 Pine Rd', 'Salem, MA 01970'], phone: '978-555-0123'}]))""")
    pg.click('#fill-tabs button:has-text("Address")')
    pg.fill('#from-addr', 'Me\n1 Home St\nBoston, MA 02108')
    wait(pg, "!!document.querySelector('#to-pick option[value=\"s\"]')")
    check(pg.evaluate("document.getElementById('to-addr').nextElementSibling.id") == 'addr-extras', 'on shipping sheets, phone and email sit under Ship to')
    pg.select_option('#to-pick', value='s')
    check(pg.input_value('#x-phone') == '978-555-0123', 'picking Ship to brings the phone')
    pg.check('#x-phone-on')
    pg.click('#ship-addr-place')
    check(wait(pg, "Object.values(LSK_TEST.st().texts).some(t => t.indexOf('978-555-0123') >= 0)", 10000), 'the ticked phone goes on the shipping label')
    check(pg.input_value('#to-addr') == '' and pg.input_value('#x-phone') == '', 'Ship to and its phone clear after placing')
    check(not errs, 'no script errors (shipping)', errs)
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
