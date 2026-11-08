const T = require(require('path').join(__dirname, '..', 'lsk-tools.js'));
const assert = require('assert');
let n = 0; const ok = (c, m) => { assert(c, m); n++; };
// CSV with quotes, embedded newline and comma
let rows = T.parseDelimited('Name,Street,City,State,ZIP\n"Doe, Jane","12 Elm St\nUnit 4",Boston,MA,2108\n');
ok(rows.length === 2 && rows[1][0] === 'Doe, Jane' && rows[1][1] === '12 Elm St\nUnit 4', 'csv quotes');
// Excel paste (tabs)
rows = T.parseDelimited('Jane Doe\tAcme\t12 Elm St\t\tBoston\tMA\t02108\nJohn Roe\t\t9 Oak Ave\tApt 2\tCambridge\tMA\t02139');
ok(rows.length === 2 && rows[1][3] === 'Apt 2', 'tsv');
// Header mapping: Google Contacts style
let g = T.rowsToAddresses(T.parseDelimited('First Name,Last Name,Organization Name,Address 1 - Street,Address 1 - City,Address 1 - Region,Address 1 - Postal Code,Address 1 - Country\nJane,Doe,,12 Elm St,Boston,MA,2108,United States\nNo,Address,,,,,,\n'));
ok(g.length === 2, 'google rows ' + g.length);
let lines = T.addressLines(g[0]);
ok(JSON.stringify(lines) === JSON.stringify(['Jane Doe','12 Elm St','Boston, MA 02108']), 'google lines ' + JSON.stringify(lines));
ok(T.addressLines(g[1]).length === 1, 'name-only row stays name only');
// Formatted address column
let fm = T.rowsToAddresses(T.parseDelimited('Name,Address 1 - Formatted\n"Ann Lee","5 Pine Rd\nSalem, MA 01970"\n'));
ok(JSON.stringify(T.addressLines(fm[0])) === JSON.stringify(['Ann Lee','5 Pine Rd','Salem, MA 01970']), 'formatted ' + JSON.stringify(T.addressLines(fm[0])));
// Positional (no header)
let pos = T.rowsToAddresses([['Jane Doe','','12 Elm St','Apt 3','Boston','MA','2108']]);
ok(JSON.stringify(T.addressLines(pos[0])) === JSON.stringify(['Jane Doe','12 Elm St Apt 3','Boston, MA 02108']), 'positional ' + JSON.stringify(T.addressLines(pos[0])));
// vCard (Apple style, folded line, item prefix)
let vc = T.parseVCard('BEGIN:VCARD\r\nVERSION:3.0\r\nN:Doe;Jane;;;\r\nFN:Jane Doe\r\nORG:Acme Inc;\r\nitem1.ADR;type=HOME;type=pref:;;12 Elm St\\nUnit 4;Boston;MA;02108;United\r\n  States\r\nEND:VCARD\r\nBEGIN:VCARD\r\nFN:No Addr\r\nEND:VCARD\r\n');
ok(vc.length === 1, 'vcard count ' + vc.length);
ok(JSON.stringify(vc[0].lines) === JSON.stringify(['Jane Doe','Acme Inc','12 Elm St','Unit 4','Boston, MA 02108']), 'vcard lines ' + JSON.stringify(vc[0].lines));
// lines -> fields
let f = T.linesToFields(['Jane Doe','12 Elm St Apt 3','Boston, MA 02108']);
ok(f.name === 'Jane Doe' && f.street === '12 Elm St Apt 3' && f.city === 'Boston' && f.state === 'MA' && f.zip === '02108', 'linesToFields ' + JSON.stringify(f));
// Canada
let ca = T.addressLines({ name: 'Jane Roe', street: '123 Main St', apt: '5', city: 'Toronto', state: 'on', zip: 'm5v3l9' });
ok(JSON.stringify(ca) === JSON.stringify(['Jane Roe', '5-123 Main St', 'Toronto ON  M5V 3L9', 'CANADA']), 'canada from us ' + JSON.stringify(ca));
ca = T.addressLines({ name: 'Jane Roe', street: '123 Main St', apt: 'Suite 200', city: 'Toronto', state: 'ON', zip: 'M5V 3L9' }, { home: 'CA' });
ok(JSON.stringify(ca) === JSON.stringify(['Jane Roe', '123 Main St Suite 200', 'Toronto ON  M5V 3L9']), 'canada from canada ' + JSON.stringify(ca));
let us2 = T.addressLines({ name: 'Ann Lee', street: '5 Pine Rd', city: 'Salem', state: 'MA', zip: '1970' }, { home: 'CA' });
ok(JSON.stringify(us2) === JSON.stringify(['Ann Lee', '5 Pine Rd', 'Salem, MA 01970', 'USA']), 'us from canada ' + JSON.stringify(us2));
ok(T.countryOf({ state: 'CA', zip: '90210' }) === 'US' && T.countryOf({ state: 'QC' }) === 'CA' && T.countryOf({ zip: 'H2X 1Y4' }) === 'CA', 'countryOf');
ok(T.fixPostal('k1a0b1') === 'K1A 0B1', 'fixPostal');
let cf = T.linesToFields(['Jane Roe', '5-123 Main St', 'Toronto ON  M5V 3L9', 'CANADA']);
ok(cf.city === 'Toronto' && cf.state === 'ON' && cf.zip === 'M5V 3L9' && cf.street === '5-123 Main St', 'linesToFields canada ' + JSON.stringify(cf));
ok(T.regionCode('Massachusetts') === 'MA' && T.regionCode('Québec') === 'QC' && T.regionCode('on') === 'ON' && T.regionCode('Nouvelle-Écosse') === 'NS', 'regionCode');
// Fit: fake measure = 0.5 * size per character
const measure = (text, css) => { const size = parseFloat(css.match(/([\d.]+)pt/)[1]); return text.length * size * 0.5; };
const box = { width: 160, height: 60 };
const short = [{ text: 'Jane Doe', size: 10 }, { text: '12 Elm St', size: 10 }, { text: 'Boston, MA 02108', size: 10 }];
ok(T.fits(short, box, measure), 'short fits');
const long = short.concat([{ text: 'Second line of a much longer company name here', size: 10 }, { text: 'Extra line', size: 10 }]);
ok(!T.fits(long, box, measure), 'long does not fit');
const s = T.fitScale(long, box, measure);
ok(s > 0.5 && s < 1, 'scale ' + s);
const scaled = long.map(p => ({ ...p, size: p.size * s }));
ok(T.fits(scaled, box, measure), 'scaled fits');
ok(T.fitScale([{ text: 'x'.repeat(2000), size: 10 }], box, measure) === null, 'impossible -> null');
// Border finder: 400x500 white page, 4x6 box border 3px thick at (40,40)-(240,340), text and a full-width rule below
const w = 400, h = 500, gray = new Uint8Array(w * h).fill(255);
const rect = (x0, y0, x1, y1) => { for (let y = y0; y <= y1; y++) for (let x = x0; x <= x1; x++) gray[y * w + x] = 0; };
rect(40, 40, 240, 42); rect(40, 338, 240, 340); rect(40, 40, 42, 340); rect(238, 40, 240, 340);
rect(60, 100, 220, 102); // inner rule
rect(20, 420, 380, 421); // instructions rule outside the box
for (let i = 0; i < 30; i++) rect(60 + i * 5, 200, 61 + i * 5, 260); // barcode bars
const b = T.findLabelBorder(gray, w, h);
ok(b && Math.abs(b.x - 40) <= 2 && Math.abs(b.y - 40) <= 2 && Math.abs(b.w - 201) <= 3 && Math.abs(b.h - 301) <= 3, 'border ' + JSON.stringify(b));
// No border -> null
const plain = new Uint8Array(w * h).fill(255); for (let i = 0; i < 10; i++) for (let x = 50; x < 150; x++) plain[(60 + i * 10) * w + x] = 0;
ok(T.findLabelBorder(plain, w, h) === null, 'no border -> null');

// ---------- contacts: two addresses, phones, emails ----------
const J = (x) => JSON.stringify(x);
// vCard: home and PO box addresses become one contact; the preferred phone and email go first.
let two = T.parseVCard('BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Jane Doe\r\nTEL;type=HOME:617-555-0100\r\nTEL;type=CELL;type=pref:(617) 555-0199\r\n' +
  'EMAIL;type=WORK:jane@work.example\r\nEMAIL;type=HOME;type=pref:jane@home.example\r\nEMAIL:third@example.com\r\n' +
  'ADR;type=HOME:;;12 Elm St;Boston;MA;02108;\r\nADR;type=POSTAL:;;PO Box 44;Boston;MA;02101;\r\nADR;type=OTHER:;;3 Far Rd;Salem;MA;01970;\r\nEND:VCARD\r\n' +
  'BEGIN:VCARD\r\nFN:Phone Only\r\nTEL:555-1212\r\nEND:VCARD\r\n');
ok(two.length === 1 && two.skipped === 1, 'vcard: one contact, one without an address skipped ' + two.length + '/' + two.skipped);
ok(J(two[0].lines) === J(['Jane Doe', '12 Elm St', 'Boston, MA 02108']) && J(two[0].lines2) === J(['Jane Doe', 'PO Box 44', 'Boston, MA 02101']), 'vcard: first two addresses ' + J([two[0].lines, two[0].lines2]));
ok(two[0].phone === '(617) 555-0199' && two[0].phone2 === '617-555-0100', 'vcard: preferred phone first ' + J([two[0].phone, two[0].phone2]));
ok(two[0].email === 'jane@home.example' && two[0].email2 === 'jane@work.example', 'vcard: preferred email first, two kept ' + J([two[0].email, two[0].email2]));
let v4 = T.parseVCard('BEGIN:VCARD\r\nVERSION:4.0\r\nFN:Ann\r\nTEL;VALUE=uri;PREF=1:tel:+1-617-555-0123\r\nEMAIL:mailto:ann@example.com\r\nADR:;;5 Pine Rd;Salem;MA;01970;\r\nEND:VCARD\r\n');
ok(v4[0].phone === '+1-617-555-0123' && v4[0].email === 'ann@example.com' && !v4[0].lines2, 'vcard 4: tel: and mailto: removed ' + J(v4[0]));
// Google Contacts CSV: Address 1 and 2, "Phone 1 - Value", several emails in one cell.
let gc = T.rowsToContacts(T.parseDelimited(
  'First Name,Last Name,E-mail 1 - Label,E-mail 1 - Value,Phone 1 - Label,Phone 1 - Value,Phone 2 - Value,Address 1 - Street,Address 1 - City,Address 1 - Region,Address 1 - Postal Code,Address 2 - Street,Address 2 - City,Address 2 - Region,Address 2 - Postal Code\n' +
  'Jane,Doe,* Home,jane@a.example ::: jane@b.example,Mobile,617-555-0100,617-555-0100,12 Elm St,Boston,MA,02108,PO Box 9,Boston,MA,02101\n' +
  'Ann,Lee,,ann@example.com,,,,,,,,,,,\n'));
ok(gc.length === 1 && gc.skipped === 1, 'google: one contact, one without an address skipped ' + gc.length + '/' + gc.skipped);
ok(J(gc[0].lines2) === J(['Jane Doe', 'PO Box 9', 'Boston, MA 02101']) && gc[0].email === 'jane@a.example' && gc[0].email2 === 'jane@b.example', 'google: second address, two emails from one cell ' + J(gc[0]));
ok(gc[0].phone === '617-555-0100' && gc[0].phone2 === '', 'google: the same phone twice is kept once ' + J([gc[0].phone, gc[0].phone2]));
// Outlook CSV: Business and Home sets; a person with only a home address still comes in.
let ol = T.rowsToContacts(T.parseDelimited(
  'First Name,Last Name,Company,Business Street,Business City,Business State,Business Postal Code,Home Street,Home City,Home State,Home Postal Code,Business Fax,Business Phone,Home Phone,Mobile Phone,E-mail Address,E-mail 2 Address\n' +
  'Bob,Ray,Acme,77 Main St,Lowell,MA,01852,9 Oak Ave,Cambridge,MA,02139,555-0000,978-555-0101,617-555-0102,617-555-0103,bob@acme.example,bob@home.example\n' +
  'Sue,Kim,,,,,,4 Elm Ct,Quincy,MA,02169,,,,,sue@example.com,\n'));
ok(ol.length === 2 && J(ol[0].lines) === J(['Bob Ray', 'Acme', '77 Main St', 'Lowell, MA 01852']) && J(ol[0].lines2) === J(['Bob Ray', 'Acme', '9 Oak Ave', 'Cambridge, MA 02139']), 'outlook: business and home addresses ' + J(ol[0]));
ok(ol[0].phone === '617-555-0103' && ol[0].phone2 === '617-555-0102' && ol[0].email2 === 'bob@home.example', 'outlook: mobile phone first, fax left out ' + J([ol[0].phone, ol[0].phone2]));
ok(J(ol[1].lines) === J(['Sue Kim', '4 Elm Ct', 'Quincy, MA 02169']) && !ol[1].lines2 && ol[1].email === 'sue@example.com', 'outlook: home-only contact ' + J(ol[1]));
// Spreadsheet with Phone and Email columns; rows for the table keep them too.
let sp = T.rowsToAddresses(T.parseDelimited('Name,Street,City,State,ZIP,Phone,Email\nJane Doe,12 Elm St,Boston,MA,02108,617-555-0100,jane@example.com\n'));
ok(sp[0].phone === '617-555-0100' && sp[0].email === 'jane@example.com', 'spreadsheet: phone and email columns ' + J(sp[0]));
let ps = T.rowsToAddresses([['Jane Doe', '', '12 Elm St', '', 'Boston', 'MA', '02108', '', '617-555-0100', 'jane@example.com']]);
ok(ps[0].phone === '617-555-0100' && ps[0].email === 'jane@example.com', 'no header: phone and email after the address ' + J(ps[0]));

// ---------- suggestions as you type ----------
(async () => {
  const realFetch = globalThis.fetch;
  const calls = [];
  let serverStatus = 200;
  const geo = (h, s, c) => ({ housenumber: h, street: s, city: c, state_code: 'MA', postcode: '02108', country_code: 'us', address_line1: h + ' ' + s, address_line2: c + ', MA 02108' });
  globalThis.fetch = async (u) => {
    u = String(u); calls.push(u);
    if (u.includes('labelsheetkeeper.workers.dev')) {
      if (serverStatus !== 200) return new Response('{"error":"used up"}', { status: serverStatus });
      if (u.includes('Oak')) return new Response(JSON.stringify({ results: ['1', '2', '3', '4', '5'].map((h) => geo('5' + h, 'Oak St', 'Boston')) }));
      return new Response(JSON.stringify({ results: [geo('8', 'Maple Way', 'Boston'), geo('8', 'Maplewood Ave', 'Salem'), geo('28', 'Maple St', 'Quincy')] }));
    }
    if (u.includes('photon.komoot.io')) return new Response(JSON.stringify({ features: [{ properties: { type: 'house', housenumber: '8', street: 'Elm St', city: 'Boston', state: 'Massachusetts', postcode: '02108', countrycode: 'US' } }] }));
    return new Response('{}', { status: 404 });
  };
  const notes = [];
  const S = T.makeSuggester({ key: () => '', notice: (t) => notes.push(t) });
  ok(!S.ready('8 M') && S.ready('8 Ma') && !S.ready('Maple'), 'suggestions start at a house number and two letters ("8 Ma")');
  ok(S.delay() === 250, 'and ask through the server once typing pauses a quarter second');
  let got = await S.ask('8 Ma');
  ok(got.length === 3 && got[0].source === 'geoapify' && calls[0].includes('/suggest?text=8%20Ma'), 'the first answer comes through the server ' + calls[0]);
  ok(J(S.quick('8 map').map((i) => i.street)) === J(['8 Maple Way', '8 Maplewood Ave']), 'typing on narrows it at once, by the start of each word (not 28 Maple St)');
  ok(J(S.quick('8 maplew').map((i) => i.street)) === J(['8 Maplewood Ave']) && S.quick('9 elm').length === 0, 'and only to what still fits');
  ok(J(S.quick('8 map, salem').map((i) => i.street)) === J(['8 Maplewood Ave']), 'a city typed after it narrows it too');
  const before = calls.length;
  got = await S.ask('8 ma');
  ok(calls.length === before && got.length === 3, 'the same question isn’t asked twice');
  got = await S.ask('8 Maplew');
  ok(calls.length === before && J(got.map((i) => i.street)) === J(['8 Maplewood Ave']), 'an answer that already had every match (fewer than 5) answers longer text without asking again');
  await S.ask('8 Maplew, Salem');
  ok(calls.length === before + 1, 'but a different city after it is asked about');
  await S.ask('5 Oak');
  await S.ask('5 Oak S');
  ok(calls.length === before + 3, 'and an answer with 5 (maybe more out there) is asked about again as you type');
  serverStatus = 429;
  got = await S.ask('9 Elm');
  ok(got.length === 1 && got[0].source === 'photon' && S.delay() === 300 && !S.ready('9 El'), 'when the shared ones are used up, Photon answers (a little slower, a little more typed first)');
  const n2 = calls.length;
  await S.ask('9 Elm S');
  ok(calls.length === n2 + 1 && calls[n2].includes('photon.komoot.io'), 'and the server isn’t asked again');
  ok(S.quick('8 map').length === 0, 'answers from the server aren’t mixed in with Photon’s');
  ok(/Photon by Komoot/.test(S.credit(got)) && /used up for today/.test(S.credit(got)), 'the credit says why Photon is answering: ' + S.credit(got));
  const realNow = Date.now;
  try {
    const t0 = realNow();
    Date.now = () => t0 + 26 * 3600 * 1000;
    serverStatus = 200;
    ok(S.delay() === 250 && S.source() === 'shared', 'the next day the shared suggestions are used again');
    const n3 = calls.length;
    got = await S.ask('9 Elm St');
    ok(calls[n3].includes('/suggest?') && got[0].source === 'geoapify', 'and asked first again', calls.slice(n3));
    ok(S.quick('9 elm').every((i) => i.source === 'geoapify'), 'Photon’s answers never stand in for the shared ones');
    // Busy, or no answer: Photon for a minute only.
    serverStatus = 503;
    await S.ask('5 Oak Rd');
    ok(S.source() === 'photon' && /back in a minute/.test(S.why()), 'when the server is busy, Photon answers for now: ' + S.why());
    const t1 = Date.now();
    Date.now = () => t1 + 61 * 1000;
    ok(S.source() === 'shared' && S.why() === '', 'and a minute later the shared ones are tried again');
  } finally { Date.now = realNow; }
  const K = T.makeSuggester({ key: () => '0123456789abcdef', notice: (t) => notes.push(t) });
  ok(K.delay() === 200, 'with their own key, asks after a fifth of a second');
  globalThis.fetch = async (u) => { calls.push(String(u)); return new Response('{"message":"Invalid apiKey"}', { status: 401 }); };
  let err = null;
  try { await K.ask('8 Ma'); } catch (e) { err = e; }
  ok(err && err.status === 401 && !K.keyOk() && /didn’t accept your key/.test(notes[notes.length - 1]), 'a key Geoapify refuses is set aside, and the user is told');
  globalThis.fetch = realFetch;
  console.log('all', n, 'tool checks passed');
})().catch((e) => { console.error(e); process.exit(1); });
