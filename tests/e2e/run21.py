"""Label text drawn where Word puts it, address suggestions while typing on a
label in Word, and the print preview window."""
import subprocess, time, sys, json, os
from urllib.parse import urlparse, parse_qs
from playwright.sync_api import sync_playwright

S = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # tests/
REPO = os.path.dirname(S)
OUT = os.path.join(S, 'out')
os.makedirs(OUT, exist_ok=True)
MOCK = open(S + '/mock-office.js').read()
PORT = 8799
API = 'https://api.labelsheetkeeper.workers.dev'
# The preview window only needs Office for its Close button.
OFFICE_STUB = "window.Office = { onReady: function (f) { setTimeout(f, 0); }, context: { ui: { messageParent: function (m) { window.SENT = m; } } } };"
results = []

def check(cond, name, detail=''):
    results.append((bool(cond), name, detail))
    print(('PASS ' if cond else 'FAIL ') + name + ('' if cond else '  -- ' + str(detail)))

srv = subprocess.Popen(['python3', '-m', 'http.server', str(PORT), '--bind', '127.0.0.1'], cwd=REPO, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
time.sleep(1)

def geo(h, s, c, z):
    return {'housenumber': h, 'street': s, 'city': c, 'state_code': 'MA', 'postcode': z, 'country_code': 'us', 'address_line1': h + ' ' + s, 'address_line2': '%s, MA %s' % (c, z)}

class Server:
    """Stands in for the Worker's /suggest."""
    def __init__(self):
        self.asked = []
    def handle(self, route):
        text = parse_qs(urlparse(route.request.url).query).get('text', [''])[0]
        self.asked.append(text)
        cors = {'Access-Control-Allow-Origin': '*'}
        if 'salem' in text.lower():
            res = [geo('8', 'Maplewood Ave', 'Salem', '01970')]
        else:
            res = [geo('8', 'Maple Way', 'Boston', '02108'), geo('8', 'Maplewood Ave', 'Salem', '01970'), geo('28', 'Maple St', 'Quincy', '02169')]
        route.fulfill(headers=cors, json={'results': res})

def wait(pg, js, timeout=6000, arg=None):
    try:
        pg.wait_for_function(js, arg=arg, timeout=timeout); return True
    except Exception:
        return False

def say(pg): return pg.inner_text('#say')
def paras(pg, n): return pg.evaluate("n => MOCK.labelCell(n).paras.map(p => p.text)", n)
def refresh(pg):
    pg.click('#refresh-btn'); wait(pg, "document.getElementById('say').textContent.indexOf('Up to date') >= 0")
def cursor(pg, n, para):
    """Puts Word's cursor on a line of a label, and tells the panel, as Word does."""
    pg.evaluate("([n, para]) => { const S = LSK_SHEET, f = LSK_TEST.st().fmt, c = S.cellOf(f, n); MOCK.selection = { row: c.row, cell: c.cell, para: para }; MOCK.selHandler(); }", [n, para])
SHOWN = "() => !document.getElementById('dsugg').hidden"
def card(pg):
    return pg.evaluate("Array.from(document.querySelectorAll('#dsugg-list button')).map(b => b.firstChild.textContent)")
def pick(pg, label):
    pg.click('#dsugg-list button:has-text("%s")' % label)

try:
  with sync_playwright() as p:
    b = p.chromium.launch()
    ctx = b.new_context(viewport={'width': 340, 'height': 1000})
    server = Server()
    ctx.route(API + '/**', server.handle)
    pg = ctx.new_page()
    errs = []
    pg.on('pageerror', lambda e: errs.append(str(e)))
    pg.route('**/office.js', lambda r: r.fulfill(body=MOCK, content_type='application/javascript'))
    pg.goto('http://127.0.0.1:%d/taskpane.html' % PORT)
    check(wait(pg, "window.LSK_TEST && LSK_TEST.st().kind === 'sheet'", 9000), 'panel opens on a sheet')

    # ================= Label text where Word puts it =================
    pg.evaluate("MOCK.setLabelText(4, 'Jane Doe\\n12 Elm St\\nBoston, MA 02108', 10)")
    pg.evaluate("MOCK.setLabelText(5, 'Bob Ray\\n1 Oak Ave', 14); MOCK.labelCell(5).paras.forEach(p => p.alignment = 'Centered'); MOCK.labelCell(5).paras[0].font.bold = true; MOCK.labelCell(5).paras[1].font.italic = true")
    refresh(pg)
    g = pg.evaluate("""() => {
      const f = LSK_TEST.st().fmt, cell = document.querySelector('#map .cell[data-n="4"]'), inn = cell.querySelector('.lt-in');
      const c = cell.getBoundingClientRect(), i = inn.getBoundingClientRect(), map = document.getElementById('map').getBoundingClientRect();
      const first = inn.firstElementChild, fs = parseFloat(getComputedStyle(first).fontSize);
      return { dy: Math.abs((i.top + i.bottom) / 2 - (c.top + c.bottom) / 2), left: i.left - c.left, wantLeft: f.padding.left / f.label.width * c.width,
        fs: fs, wantFs: 10 / 72 / f.page.width * map.width, lines: inn.children.length, texts: Array.from(inn.children).map(d => d.textContent) };
    }""")
    check(g['dy'] < 1.5, 'the map shows label text centred top to bottom, like Word (not at the top)', g)
    check(abs(g['left'] - g['wantLeft']) < 1, 'with the label\'s side padding', g)
    check(abs(g['fs'] - g['wantFs']) < 0.15, 'and the font size to scale with the sheet', g)
    check(g['texts'] == ['Jane Doe', '12 Elm St', 'Boston, MA 02108'], 'every line shows, not just the first three', g['texts'])
    s5 = pg.evaluate("""() => Array.from(document.querySelectorAll('#map .cell[data-n="5"] .lt-in > div')).map(d => { const s = getComputedStyle(d); return [s.textAlign, s.fontWeight, s.fontStyle]; })""")
    check(s5 == [['center', '700', 'normal'], ['center', '400', 'italic']], 'centred, bold and italic lines show as they are in Word', s5)
    big = pg.evaluate("""() => {
      const f = LSK_TEST.st().fmt, m = document.getElementById('map');
      const box = document.createElement('div'); box.style.cssText = 'position:absolute;left:0;top:0;width:30%;height:9%'; m.appendChild(box);
      const lines = Array.from({ length: 9 }, (_, i) => ({ text: 'Line ' + i, size: 12 }));
      LSK_TOOLS.drawLabelText(box, lines, f);
      const over = box.querySelector('.lt-in').getBoundingClientRect().top - box.getBoundingClientRect().top;
      const top = Object.assign({}, f, { align: 'top' });
      LSK_TOOLS.drawLabelText(box, [{ text: 'One', size: 10, before: 20 }], top);
      const atTop = box.querySelector('.lt-in').getBoundingClientRect().top - box.getBoundingClientRect().top;
      const firstTop = box.querySelector('.lt-in > div').getBoundingClientRect().top - box.getBoundingClientRect().top;
      box.remove();
      return { over: over, atTop: atTop, firstTop: firstTop, want: 20 / 72 / f.page.width * m.getBoundingClientRect().width };
    }""")
    check(abs(big['over']) < 0.5, 'text too long for the label starts at its top and is cut off at the bottom, as Word does', big)
    check(abs(big['atTop']) < 0.5 and abs(big['firstTop'] - big['want']) < 0.5, 'top-aligned sheets start at the top, after the first line\'s space above', big)
    pg.locator('#map').screenshot(path=OUT + '/s21-map.png')

    # ================= Suggestions while typing on a label in Word =================
    pg.evaluate("MOCK.setLabelText(7, 'Jane Doe\\n8 Ma', 10)")
    cursor(pg, 7, 1)
    check(wait(pg, SHOWN), 'typing a house number and street on a label shows suggestions in the panel')
    check(pg.inner_text('#dsugg-title') == 'Suggestions for label 7', 'the card says which label they\'re for', pg.inner_text('#dsugg-title'))
    check(card(pg) == ['8 Maple Way', '8 Maplewood Ave', '28 Maple St'] and 'Geoapify' in pg.inner_text('#dsugg-credit'), 'with the suggested addresses and their credit', card(pg))
    check(server.asked == ['8 Ma'], 'asked once, about the line the cursor is on', server.asked)
    pg.screenshot(path=OUT + '/s21-label-suggestions.png')
    pick(pg, '8 Maple Way')
    check(wait(pg, "MOCK.labelCell(7).paras.length === 3 && MOCK.labelCell(7).paras[2].text === 'Boston, MA 02108'"), 'picking one fills the street and adds the city line on the label', paras(pg, 7))
    check(paras(pg, 7)[:2] == ['Jane Doe', '8 Maple Way'], 'the name stays, the street is completed', paras(pg, 7))
    f7 = pg.evaluate("[MOCK.labelCell(7).paras[2].font.name, MOCK.labelCell(7).paras[2].font.size]")
    check(f7 == ['Arial', 10], 'the new city line has the street line\'s font and size', f7)
    check(pg.is_hidden('#dsugg'), 'and the card goes away')
    check(wait(pg, "LSK_TEST.st().ready[7] === true"), 'the panel sees the label is ready to print')
    n_before = len(server.asked)
    cursor(pg, 7, 1); time.sleep(0.9)
    check(pg.is_hidden('#dsugg') and len(server.asked) == n_before, 'back on the street just filled, it stays quiet', server.asked)

    # A city line already under the street is replaced, not doubled.
    pg.evaluate("MOCK.setLabelText(9, 'Ann Lee\\n8 Map\\nSalem, MA', 10)")
    cursor(pg, 9, 1)
    check(wait(pg, SHOWN) and card(pg) == ['8 Maplewood Ave'], 'the city typed under the street narrows the suggestions', card(pg))
    for _ in range(20):
        if server.asked[-1] == '8 Map, Salem, MA': break
        pg.wait_for_timeout(100)
    check(server.asked[-1] == '8 Map, Salem, MA', 'the city is sent with the street once typing pauses', server.asked[-1:])
    pick(pg, '8 Maplewood Ave')
    check(wait(pg, "MOCK.labelCell(9).paras.length === 3 && MOCK.labelCell(9).paras[2].text === 'Salem, MA 01970'"), 'the city line under it is replaced', paras(pg, 9))

    # A line under the street that isn't a city stays, under the new city line.
    pg.evaluate("MOCK.setLabelText(12, 'Ray Cho\\n8 Ma\\n617-555-0100', 10)")
    cursor(pg, 12, 1)
    check(wait(pg, SHOWN), 'suggestions on label 12')
    pick(pg, '8 Maple Way')
    check(wait(pg, "MOCK.labelCell(12).paras.length === 4"), 'the city line goes in between', paras(pg, 12))
    check(paras(pg, 12) == ['Ray Cho', '8 Maple Way', 'Boston, MA 02108', '617-555-0100'], 'and the phone line stays last', paras(pg, 12))
    check(pg.evaluate("MOCK.calls.indexOf('para.insertParagraph.AfterLast') < 0"), 'no line is ever added after a label\'s last line (Word for Mac would join it on)')
    check(pg.evaluate("MOCK.calls.filter(c => c === 'select').length") >= 3, 'the cursor goes back to the end of the street line')

    # An answer that already held every match answers longer text without asking again.
    pg.evaluate("MOCK.setLabelText(11, 'Al Fox\\n8 Mapl', 10)")
    n_before = len(server.asked)
    cursor(pg, 11, 1)
    check(wait(pg, SHOWN) and card(pg) == ['8 Maple Way', '8 Maplewood Ave'], 'what still fits shows at once', card(pg))
    time.sleep(0.8)
    check(len(server.asked) == n_before, 'and no new lookup is spent on it', server.asked[n_before:])

    # Other lines stay quiet.
    cursor(pg, 11, 0)
    check(wait(pg, "() => document.getElementById('dsugg').hidden"), 'moving to the name line closes the card')
    pg.click('#map .cell[data-n="1"]')
    check(wait(pg, "LSK_TEST.st().used[1] === true"), 'label 1 marked used')
    pg.evaluate("MOCK.labelCell(1).paras[0].text = '8 Ma'")
    cursor(pg, 1, 0); time.sleep(0.9)
    check(pg.is_hidden('#dsugg'), 'no suggestions on a used label')

    # Closing the card keeps it closed on that line.
    pg.evaluate("MOCK.setLabelText(13, '8 Ma', 10)")
    cursor(pg, 13, 0)
    check(wait(pg, SHOWN), 'suggestions on label 13')
    pg.click('#dsugg-close')
    check(pg.is_hidden('#dsugg'), '× closes the card')
    cursor(pg, 13, 0); time.sleep(0.9)
    check(pg.is_hidden('#dsugg'), 'and it stays closed while typing on that line')
    pg.evaluate("MOCK.setLabelText(14, 'Kim Wu\\n8 Ma', 10)")
    cursor(pg, 14, 1)
    check(wait(pg, SHOWN) and pg.inner_text('#dsugg-title') == 'Suggestions for label 14', 'another line gets suggestions again')

    # The line changed before the click: it's left alone.
    pg.evaluate("MOCK.labelCell(14).paras[1].text = '9 Elm'")
    pick(pg, '8 Maple Way')
    check(wait(pg, "document.getElementById('say').textContent.indexOf('left as it is') >= 0"), 'a line that changed in the meantime is left alone, and the panel says so', say(pg))
    check(paras(pg, 14) == ['Kim Wu', '9 Elm'], 'the label is unchanged', paras(pg, 14))

    # Too much for the label: offers to shrink it.
    pg.evaluate("MOCK.setLabelText(16, 'Jane Doe\\nAcme Inc\\n8 Ma', 18)")
    cursor(pg, 16, 2)
    check(wait(pg, SHOWN), 'suggestions on a label with large text')
    pick(pg, '8 Maple Way')
    check(wait(pg, "!document.getElementById('confirm').hidden && /Shrink it to fit/.test(document.getElementById('confirm-msg').textContent)"), 'when the filled label won\'t fit, it offers to shrink it', pg.inner_text('#confirm-msg') if pg.is_visible('#confirm') else 'no question')
    pg.click('#confirm-no')

    # Turned off in Settings.
    pg.evaluate("document.getElementById('settings').open = true")
    pg.uncheck('#suggest-on')
    pg.evaluate("MOCK.setLabelText(17, '8 Ma', 10)")
    n_before = len(server.asked)
    cursor(pg, 17, 0); time.sleep(0.9)
    check(pg.is_hidden('#dsugg') and len(server.asked) == n_before, 'with suggestions turned off, labels get none')
    pg.check('#suggest-on')

    # ================= The preview window =================
    pg.click('#map .cell[data-n="2"]'); pg.click('#map .cell[data-n="3"]')
    check(wait(pg, "LSK_TEST.st().used[2] && LSK_TEST.st().used[3]"), 'labels 2 and 3 marked used')
    pg.evaluate("""() => { const c = MOCK.labelCell(10), base = MOCK.labelCell(4).paras[0];
      const mk = (text, size, extra) => Object.assign(JSON.parse(JSON.stringify(base)), { text: text, alignment: 'Centered', font: { name: 'Arial', size: size } }, extra || {});
      c.ccs = []; c.paras = [mk('', 10, { pic: { width: 60, height: 40, b64: MOCK.GREY_PNG } }), mk('Logo Co', 9)]; c.pics = 1; }""")
    refresh(pg)
    pg.click('#preview-btn')
    check(wait(pg, "MOCK.dialog && /preview\\.html\\?t=/.test(MOCK.dialog.url)"), 'Preview opens the preview window', pg.evaluate("MOCK.dialog && MOCK.dialog.url"))
    data = pg.evaluate("JSON.parse(localStorage.getItem('lsk.preview'))")
    labs = {l['n']: l for l in data['labels']}
    check(sorted(data['used']) == [1, 2, 3], 'the window is told which labels are used', data['used'])
    check(4 in labs and [x['text'] for x in labs[4]['specs']] == ['Jane Doe', '12 Elm St', 'Boston, MA 02108'], 'and each ready label\'s lines', labs.get(4))
    pic = labs.get(10, {}).get('specs', [{}])[0].get('pic') or {}
    check(pic.get('src', '').startswith('data:image/jpeg') and pic.get('w') == 60 and pic.get('h') == 40, 'with pictures made small and their size on the label', {k: (v[:30] if isinstance(v, str) else v) for k, v in pic.items()})
    check('lsk.preview' not in pg.evaluate("LSK_TEST.backupKeys()"), 'the preview isn\'t put in backups')

    pv = ctx.new_page()
    pv.set_viewport_size({'width': 720, 'height': 900})
    perrs = []
    pv.on('pageerror', lambda e: perrs.append(str(e)))
    pv.route('**/office.js', lambda r: r.fulfill(body=OFFICE_STUB, content_type='application/javascript'))
    pv.goto(pg.evaluate("MOCK.dialog.url"))
    check(wait(pv, "document.querySelectorAll('#page .lab').length === 30"), 'the preview draws every label\'s edge', pv.evaluate("document.querySelectorAll('#page .lab').length"))
    check(pv.evaluate("document.querySelectorAll('#page .lab.used').length") == 3, 'used labels are shaded')
    t4 = pv.evaluate("Array.from(document.querySelectorAll('#page .lab')[3].querySelectorAll('.lt-in > div')).map(d => d.textContent)")
    check(t4 == ['Jane Doe', '12 Elm St', 'Boston, MA 02108'], 'label 4\'s text is on label 4', t4)
    geo4 = pv.evaluate("""() => { const l = document.querySelectorAll('#page .lab')[3], i = l.querySelector('.lt-in'); const a = l.getBoundingClientRect(), b = i.getBoundingClientRect(); return Math.abs((a.top + a.bottom) / 2 - (b.top + b.bottom) / 2); }""")
    check(geo4 < 1.5, 'centred top to bottom, where Word prints it', geo4)
    img = pv.evaluate("(() => { const i = document.querySelectorAll('#page .lab')[9].querySelector('img'); return i ? [i.complete && i.naturalWidth > 0, Math.round(i.getBoundingClientRect().width / document.getElementById('page').getBoundingClientRect().width * 8.5 * 72)] : null; })()")
    check(img and img[0] and abs(img[1] - 60) <= 1, 'the picture shows at its size on the label', img)
    check('FEED THIS EDGE' in pv.inner_text('#page'), 'the small "top" line shows above the labels')
    check('Neither prints.' in pv.inner_text('.key'), 'the window says the edges and shading don\'t print')
    rect = pv.evaluate("(() => { const r = document.getElementById('page').getBoundingClientRect(); return [r.width, r.height, innerWidth, innerHeight]; })()")
    check(rect[1] <= rect[3] and abs(rect[0] / rect[1] - 8.5 / 11) < 0.01, 'the whole page fits the window, in proportion', rect)
    pv.screenshot(path=OUT + '/s21-preview.png')

    # Follows changes while it's open.
    pg.evaluate("MOCK.setLabelText(6, 'New Person\\n5 Pine Rd\\nSalem, MA 01970', 10)")
    refresh(pg)
    check(wait(pv, "document.querySelectorAll('#page .lab')[5].textContent.indexOf('New Person') >= 0", 6000), 'a label typed while the preview is open shows up in it')
    pv.click('#close')
    check(pv.evaluate("window.SENT") == '{"type":"close"}', 'Close asks the panel to close the window')
    pg.evaluate("MOCK.dialog.send({ type: 'close' })")
    check(pg.evaluate("MOCK.dialog.closed"), 'and the panel closes it')
    check(not perrs, 'no script errors in the preview', perrs)
    pv.close()

    pg.locator('.print-row').screenshot(path=OUT + '/s21-print-row.png')
    check(not [e for e in errs if 'favicon' not in e], 'no script errors', errs)
    b.close()
finally:
    srv.kill()

failed = [r for r in results if not r[0]]
print('%d checks, %d failed' % (len(results), len(failed)))
for f in failed:
    print('::error title=run21::' + f[1])
sys.exit(1 if failed else 0)
