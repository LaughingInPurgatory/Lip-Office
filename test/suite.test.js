import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';

import { zip, unzip } from '../lib/zip.js';
import { parseMarkup, textOf, find, findAll, escapeXml } from '../shared/html.js';
import { sanitizeHtml } from '../shared/sanitize.js';
import { makeEngine, colToNum, numToCol, parseRef, toRef } from '../shared/formula.js';
import { parseCsv, toCsv } from '../shared/csv.js';
import { normalizeSheet, normalizeWord, kindOf, usedRange } from '../shared/model.js';
import { createStore, safeFileName, safeRelativePath } from '../lib/store.js';
import { htmlToDocx, docxToHtml, sheetsToXlsx, xlsxToSheets } from '../lib/ooxml.js';
import { htmlToOdt, odtToHtml, sheetsToOds, odsToSheets } from '../lib/odf.js';
import { htmlToRtf, rtfToText, htmlToText } from '../lib/rtf.js';
import { parseDocument, serializeDocument } from '../lib/convert.js';

const show = (doc, ref) => makeEngine(doc.sheets[0].cells).display(ref);
const engine = (cells) => makeEngine(cells);

// ---------------------------------------------------------------------- zip

test('zip round-trips entries, including stored (uncompressed) ones', () => {
  const entries = unzip(zip([
    { name: 'mimetype', data: 'application/x-test', store: true },
    { name: 'nested/content.xml', data: '<a>'.repeat(500) },
    { name: 'empty.txt', data: '' },
  ]));

  assert.equal(entries.get('mimetype').toString(), 'application/x-test');
  assert.equal(entries.get('nested/content.xml').toString(), '<a>'.repeat(500));
  assert.equal(entries.get('empty.txt').toString(), '');
});

test('unzip rejects something that is not an archive', () => {
  assert.throws(() => unzip(Buffer.from('this is plain text, not a zip file at all')), /not a ZIP/);
});

// --------------------------------------------------------------------- html

test('parseMarkup handles self-closing tags, attributes and entities', () => {
  const tree = parseMarkup('<w:p><w:b/><w:t xml:space="preserve">a &amp; b</w:t></w:p>');
  const paragraph = find(tree, 'w:p');
  assert.equal(textOf(find(paragraph, 'w:t')), 'a & b');
  assert.ok(find(paragraph, 'w:b'), 'self-closing tag parsed');
  assert.equal(find(paragraph, 'w:t').attrs['xml:space'], 'preserve');
});

test('parseMarkup survives unclosed and mismatched tags', () => {
  const tree = parseMarkup('<div><p>one<p>two</div><p>three');
  assert.equal(findAll(tree, 'p').length, 3);
});

test('escapeXml escapes markup characters', () => {
  assert.equal(escapeXml('a & b < c > "d"'), 'a &amp; b &lt; c &gt; &quot;d&quot;');
});

// ---------------------------------------------------------------- sanitizer

test('sanitizeHtml strips scripts, event handlers and javascript: urls', () => {
  const clean = sanitizeHtml('<p onclick="steal()">ok</p><script>evil()</script><a href="javascript:evil()">x</a>');
  assert.equal(clean, '<p>ok</p><a>x</a>');
});

test('sanitizeHtml keeps formatting and drops unsafe CSS', () => {
  const clean = sanitizeHtml('<p><b>bold</b><span style="color:#ff0000;background-image:url(http://x/y)">red</span></p>');
  assert.equal(clean, '<p><b>bold</b><span style="color:#ff0000">red</span></p>');
});

test('sanitizeHtml normalises the legacy <font color> that execCommand emits', () => {
  assert.equal(sanitizeHtml('<font color="#00FF00">hi</font>'), '<span style="color:#00ff00">hi</span>');
});

test('sanitizeHtml removes iframes and images with handlers', () => {
  assert.equal(sanitizeHtml('<iframe src="http://x"></iframe><img src=x onerror=go()>'), '');
});

// ------------------------------------------------------------------ formulas

test('colToNum / numToCol / parseRef round-trip', () => {
  assert.equal(colToNum('A'), 1);
  assert.equal(colToNum('Z'), 26);
  assert.equal(colToNum('AA'), 27);
  assert.equal(numToCol(27), 'AA');
  assert.equal(numToCol(colToNum('BA')), 'BA');
  assert.deepEqual(parseRef('c12'), { col: 3, row: 12 });
  assert.equal(toRef(3, 12), 'C12');
});

test('arithmetic respects precedence and Excel unary-minus binding', () => {
  assert.equal(engine({ A1: '=2+3*4' }).display('A1'), '14');
  assert.equal(engine({ A1: '=(2+3)*4' }).display('A1'), '20');
  assert.equal(engine({ A1: '=2^3^2' }).display('A1'), '512'); // right associative
  assert.equal(engine({ A1: '=-2^2' }).display('A1'), '4');    // (-2)^2, as Excel does
  assert.equal(engine({ A1: '=2-3-4' }).display('A1'), '-5');
  assert.equal(engine({ A1: '=50%' }).display('A1'), '0.5');
});

test('text, comparison and concatenation operators', () => {
  assert.equal(engine({ A1: '="a"&"b"&1' }).display('A1'), 'ab1');
  assert.equal(engine({ A1: '=1<2' }).display('A1'), 'TRUE');
  assert.equal(engine({ A1: '=2>=2' }).display('A1'), 'TRUE');
  assert.equal(engine({ A1: '="yes"="YES"' }).display('A1'), 'TRUE');
  assert.equal(engine({ A1: '="b"<"a"' }).display('A1'), 'FALSE');
});

test('cell values, ranges and the aggregate functions', () => {
  const cells = { A1: '1', A2: '2', A3: '3', A4: 'note', B1: '=SUM(A1:A3)' };
  assert.equal(engine(cells).display('B1'), '6');
  assert.equal(engine({ A1: '=SUM(B1:B3,"x",10)' }).display('A1'), '10'); // text ignored
  assert.equal(engine({ A1: 'SUM(A1:A4)' }).display('A1'), 'SUM(A1:A4)'); // bare text is not a formula
  assert.equal(engine({ A1: '=AVERAGE(1,2,3)' }).display('A1'), '2');
  assert.equal(engine({ A1: '=MIN(3,1,2)' }).display('A1'), '1');
  assert.equal(engine({ A1: '=MAX(3,1,2)' }).display('A1'), '3');
  assert.equal(engine({ A1: '=COUNT(1,"x",3)' }).display('A1'), '2');
  assert.equal(engine({ A1: '=COUNTA(1,"x")' }).display('A1'), '2');
  assert.equal(engine({ A1: '=SUM(A9:B10)' }).display('A1'), '0');
});

test('functions, rounding and lazy IF', () => {
  assert.equal(engine({ A1: '=IF(1>2,"yes","no")' }).display('A1'), 'no');
  assert.equal(engine({ A1: '=IF(TRUE,1,1/0)' }).display('A1'), '1');   // the error branch is never evaluated
  assert.equal(engine({ A1: '=IFERROR(1/0,"safe")' }).display('A1'), 'safe');
  assert.equal(engine({ A1: '=ROUND(2.345,2)' }).display('A1'), '2.35');
  assert.equal(engine({ A1: '=ROUND(2.5,0)' }).display('A1'), '3');
  assert.equal(engine({ A1: '=MOD(7,3)' }).display('A1'), '1');
  assert.equal(engine({ A1: '=ABS(-4)' }).display('A1'), '4');
  assert.equal(engine({ A1: '=LEN("hello")' }).display('A1'), '5');
  assert.equal(engine({ A1: '=UPPER("abc")' }).display('A1'), 'ABC');
  assert.equal(engine({ A1: '=MID("abcdef",2,3)' }).display('A1'), 'bcd');
  assert.equal(engine({ A1: '=AND(TRUE,1)' }).display('A1'), 'TRUE');
  assert.equal(engine({ A1: '=NOT(TRUE)' }).display('A1'), 'FALSE');
});

test('error codes are reported, not thrown', () => {
  assert.equal(engine({ A1: '=1/0' }).display('A1'), '#DIV/0!');
  assert.equal(engine({ A1: '=NOPE(1)' }).display('A1'), '#NAME?');
  assert.equal(engine({ A1: '=1+"abc"' }).display('A1'), '#VALUE!');
  assert.equal(engine({ A1: '=SUM(1,2' }).display('A1'), '#ERROR!');
  assert.equal(engine({ A1: '=AVERAGE(A9:A10)' }).display('A1'), '#DIV/0!');
});

test('circular references are detected instead of recursing forever', () => {
  const cells = { A1: '=B1', B1: '=A1' };
  assert.equal(engine(cells).display('A1'), '#CYCLE!');
  assert.equal(engine({ A1: '=A1' }).display('A1'), '#CYCLE!');
});

test('a literal cell is a number when it looks like one, text otherwise', () => {
  assert.equal(engine({ A1: '42' }).display('A1'), '42');
  assert.equal(engine({ A1: ' 7 ' }).display('A1'), '7');
  assert.equal(engine({ A1: '12abc' }).display('A1'), '12abc');
  assert.equal(engine({ A1: 'hello' }).display('A1'), 'hello');
  assert.equal(engine({ A1: '' }).display('A1'), '');
});

// ----------------------------------------------------------------------- csv

test('csv round-trips quotes, commas and newlines', () => {
  const rows = [['plain', 'with,comma'], ['quote "inside"', 'multi\nline'], ['', 'tail']];
  assert.deepEqual(parseCsv(toCsv(rows)), rows);
});

test('csv tolerates CRLF, a BOM and a trailing newline', () => {
  assert.deepEqual(parseCsv('\uFEFFa,b\r\n1,2\r\n'), [['a', 'b'], ['1', '2']]);
  assert.deepEqual(parseCsv(''), []);
});

// --------------------------------------------------------------- model guard

test('normalizeSheet clamps hostile input and drops bad references', () => {
  const doc = normalizeSheet({
    sheets: [{ name: '  S  ', cells: { a1: 'x', 'not-a-ref': 'y', B2: 5, C3: { evil: true }, D4: null } }],
    active: 9,
  });
  assert.equal(doc.sheets.length, 1);
  assert.deepEqual(doc.sheets[0].cells, { A1: 'x', B2: '5' });
  assert.equal(doc.sheets[0].name, 'S');
  assert.equal(doc.active, 0);
});

test('normalizeSheet always yields at least one sheet', () => {
  assert.equal(normalizeSheet({ sheets: [] }).sheets.length, 1);
  assert.equal(normalizeSheet(null).sheets[0].name, 'Sheet1');
});

test('normalizeWord falls back to a usable document', () => {
  assert.equal(normalizeWord({}).title, 'Untitled document');
  assert.equal(normalizeWord({ html: 5 }).html, '');
  assert.equal(normalizeWord({ title: '  T  ', html: '<p>x</p>' }).title, 'T');
});

test('usedRange reports the bounding box of used cells', () => {
  assert.deepEqual(usedRange({ A1: '1', C5: '2' }), { rows: 5, cols: 3 });
  assert.equal(usedRange({}), null);
});

test('kindOf maps extensions to components', () => {
  assert.equal(kindOf('a.DOCX'), 'word');
  assert.equal(kindOf('a.ods'), 'sheet');
  assert.equal(kindOf('a.csv'), 'sheet');
  assert.equal(kindOf('a.pdf'), null);
  assert.equal(kindOf('noextension'), null);
});

// --------------------------------------------------------------------- store

test('safeFileName drops any directory part instead of following it', () => {
  assert.equal(safeFileName('report.docx'), 'report.docx');
  assert.equal(safeFileName('a/b.docx'), 'b.docx');
  assert.equal(safeFileName('../../etc/passwd.docx'), 'passwd.docx');
});

test('safeFileName rejects names that could escape or confuse the store', () => {
  const bad = ['..\\evil.docx', '.hidden.docx', '', 'x'.repeat(130) + '.docx', 'notes.pdf', 'CON.docx', 'noextension'];
  for (const name of bad) {
    assert.throws(() => safeFileName(name), `should reject ${JSON.stringify(name)}`);
  }
});

test('safeFileName keeps characters people actually put in file names', () => {
  assert.equal(safeFileName('Q3 report (final) [v2].docx'), 'Q3 report (final) [v2].docx');
  assert.equal(safeFileName('  spaced.xlsx  '), 'spaced.xlsx');
});

test('safeRelativePath accepts nested documents and rejects traversal', () => {
  assert.equal(safeRelativePath('Docs/Q3 report.docx', { file: true }), 'Docs/Q3 report.docx');
  assert.equal(safeRelativePath('Docs/2026'), 'Docs/2026');
  for (const value of ['../outside.docx', 'Docs/../outside.docx', '/outside.docx', 'Docs\\outside.docx']) {
    assert.throws(() => safeRelativePath(value, { file: true }));
  }
});

test('document store lists and reads files in navigable subfolders', async (t) => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'lip-office-store-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const store = createStore(root);
  await store.ensureDirectory('Docs');
  await store.write('Docs/nested.docx', Buffer.from('word file'));

  const rootListing = await store.list();
  assert.deepEqual(rootListing.folders.map((folder) => folder.name), ['Docs']);
  assert.deepEqual(rootListing.documents, []);

  const nestedListing = await store.list('Docs');
  assert.equal(nestedListing.documents[0].path, 'Docs/nested.docx');
  assert.equal((await store.read('Docs/nested.docx')).toString(), 'word file');
});

test('document store moves files between folders and refuses to overwrite', async (t) => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'lip-office-store-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const store = createStore(root);
  await store.ensureDirectory('Docs');
  await store.ensureDirectory('Spreadsheets');
  await store.write('Docs/profile.docx', Buffer.from('profile'));

  assert.deepEqual(await store.move('Docs/profile.docx', 'Spreadsheets'), {
    source: 'Docs/profile.docx',
    path: 'Spreadsheets/profile.docx',
  });
  assert.equal((await store.read('Spreadsheets/profile.docx')).toString(), 'profile');
  await store.write('Docs/profile.docx', Buffer.from('replacement'));
  await assert.rejects(store.move('Docs/profile.docx', 'Spreadsheets'), { status: 409 });
  assert.equal((await store.read('Docs/profile.docx')).toString(), 'replacement');
});

// --------------------------------------------------------------------- docx

test('docx survives a round trip through HTML', () => {
  const html = '<h1>Title</h1><p>Hello <strong>bold</strong> and <em>italic</em> and <u>underline</u>.</p>'
    + '<ul><li>first</li><li>second</li></ul>'
    + '<p style="text-align:center">centred</p>'
    + '<table><tr><td>a</td><td>b</td></tr></table>';
  const back = docxToHtml(htmlToDocx(html, 'Title')).html;

  assert.match(back, /<h1>Title<\/h1>/);
  assert.match(back, /<strong>bold<\/strong>/);
  assert.match(back, /<em>italic<\/em>/);
  assert.match(back, /<u>underline<\/u>/);
  assert.match(back, /<li>first<\/li>/);
  assert.match(back, /text-align:center/);
  assert.match(back, /<td><p>a<\/p><\/td>/);
});

test('docx escapes markup characters in the text', () => {
  const back = docxToHtml(htmlToDocx('<p>1 &lt; 2 &amp; 3</p>')).html;
  assert.match(back, /1 &lt; 2 &amp; 3/);
});

// Browsers (and the editor's own table button) wrap rows in <tbody>, which is
// why the writers must not look for <tr> as a direct child of <table>.
test('tables wrapped in tbody survive every writer', () => {
  const html = '<table><tbody><tr><td>Metric</td><td>Value</td></tr><tr><td>Revenue</td><td>1.2M</td></tr></tbody></table>';

  const docx = docxToHtml(htmlToDocx(html)).html;
  assert.match(docx, /<td><p>Metric<\/p><\/td>/);
  assert.match(docx, /<td><p>1\.2M<\/p><\/td>/);

  const odt = odtToHtml(htmlToOdt(html)).html;
  assert.match(odt, /Metric/);
  assert.match(odt, /1\.2M/);

  const { text } = rtfToText(htmlToRtf(html));
  assert.match(text, /Metric\tValue/);
  assert.match(text, /Revenue\t1\.2M/);

  assert.match(htmlToText(html), /Metric\tValue/);
  assert.match(htmlToText(html), /Revenue\t1\.2M/);
});

test('tables without tbody still work, and thead does not leak into rows', () => {
  const flat = '<table><tr><td>a</td><td>b</td></tr></table>';
  assert.match(docxToHtml(htmlToDocx(flat)).html, /<td><p>b<\/p><\/td>/);

  const headed = '<table><thead><tr><th>H</th></tr></thead><tbody><tr><td>v</td></tr></tbody></table>';
  const back = docxToHtml(htmlToDocx(headed)).html;
  assert.match(back, /<td><p>H<\/p><\/td>/);
  assert.match(back, /<td><p>v<\/p><\/td>/);
});

test('docx carries a hyperlink through the relationship part', () => {
  const buffer = htmlToDocx('<p><a href="https://example.test/x">link</a></p>');
  const rels = unzip(buffer).get('word/_rels/document.xml.rels').toString();
  assert.match(rels, /Target="https:\/\/example\.test\/x"/);
  assert.match(unzip(buffer).get('word/document.xml').toString(), /<w:hyperlink r:id="rIdLink1">/);
});

test('docx rejects a non-Word file with a clear message', () => {
  assert.throws(() => docxToHtml(zip([{ name: 'x', data: 'y' }])), /not a Word document/);
});

// --------------------------------------------------------------------- xlsx

test('xlsx round-trips values, formulas and sheet names', () => {
  const doc = {
    sheets: [
      { name: 'Numbers', cells: { A1: '1', A2: '2', A3: '=SUM(A1:A2)', B1: 'text', C1: 'TRUE' } },
      { name: 'Second', cells: { D4: '=A1+1' } },
    ],
    active: 1,
  };
  const back = xlsxToSheets(sheetsToXlsx(doc));

  assert.equal(back.sheets.length, 2);
  assert.equal(back.sheets[0].name, 'Numbers');
  assert.equal(back.sheets[1].name, 'Second');
  assert.equal(back.sheets[0].cells.A3, '=SUM(A1:A2)');
  assert.equal(back.sheets[0].cells.B1, 'text');
  assert.equal(back.sheets[0].cells.C1, 'TRUE');
  assert.equal(back.sheets[1].cells.D4, '=A1+1');
});

test('xlsx writes the cached formula result so other apps show a value', () => {
  const buffer = sheetsToXlsx({ sheets: [{ name: 'S', cells: { A1: '2', A2: '3', A3: '=SUM(A1:A2)' } }] });
  const sheet = unzip(buffer).get('xl/worksheets/sheet1.xml').toString();
  assert.match(sheet, /<f>SUM\(A1:A2\)<\/f><v>5<\/v>/);
});

test('xlsx does not turn leading-zero strings into numbers', () => {
  const back = xlsxToSheets(sheetsToXlsx({ sheets: [{ name: 'S', cells: { A1: '007' } }] }));
  assert.equal(back.sheets[0].cells.A1, '007');
});

test('xlsx sheet names are sanitised for characters Excel rejects', () => {
  const buffer = sheetsToXlsx({ sheets: [{ name: 'a/b:c*d?e[f]g', cells: {} }, { name: '', cells: {} }] });
  const workbook = unzip(buffer).get('xl/workbook.xml').toString();
  assert.match(workbook, /name="a-b-c-d-e-f-g"/);
  assert.match(workbook, /name="Sheet2"/);
});

test('xlsx reads its own shared formulas only through the cached value', () => {
  // Hand-built workbook mirroring Excel's shared-formula output.
  const buffer = zip([
    {
      name: 'xl/workbook.xml',
      data: '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
    },
    { name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>' },
    {
      name: 'xl/worksheets/sheet1.xml',
      data: '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><f t="shared" ref="A3:A4" si="0">SUM(A1:A2)</f><v>3</v></c></row><row r="4"><c r="A4"><f t="shared" si="0"/><v>3</v></c></row></sheetData></worksheet>',
    },
  ]);

  const sheet = xlsxToSheets(buffer).sheets[0];
  assert.equal(sheet.cells.A3, '=SUM(A1:A2)');
  assert.equal(sheet.cells.A4, '3', 'shared follower keeps its cached value instead of becoming "="');
});

test('xlsx converts date-formatted serial numbers to dates', () => {
  const buffer = zip([
    { name: 'xl/workbook.xml', data: '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>' },
    { name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>' },
    { name: 'xl/styles.xml', data: '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>' },
    { name: 'xl/worksheets/sheet1.xml', data: '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>45000</v></c><c r="B1"><v>45000</v></c></row></sheetData></worksheet>' },
  ]);

  const sheet = xlsxToSheets(buffer).sheets[0];
  assert.equal(sheet.cells.A1, '2023-03-15');
  assert.equal(sheet.cells.B1, '45000', 'unstyled numbers stay numeric');
});

// ---------------------------------------------------------------------- odf

test('odt round-trips headings, emphasis, lists and tables', () => {
  const html = '<h1>Title</h1><p>Body <em>stress</em> and <strong>weight</strong>.</p>'
    + '<ul><li>one</li><li>two</li></ul>'
    + '<p style="text-align:right">right</p>'
    + '<table><tr><td>a</td><td>b</td></tr></table>';
  const back = odtToHtml(htmlToOdt(html, 'Title')).html;

  assert.match(back, /<h1>Title<\/h1>/);
  assert.match(back, /<em>stress<\/em>/);
  assert.match(back, /<strong>weight<\/strong>/);
  // ODF requires the paragraph inside a list item, so <li> keeps its <p>.
  assert.match(back, /<li><p>one<\/p><\/li>/);
  assert.match(back, /text-align:right/);
  assert.match(back, /<td><p>a<\/p><\/td>/);
});

test('odt stores the mimetype first and uncompressed, as ODF requires', () => {
  const buffer = htmlToOdt('<p>x</p>');
  assert.equal(buffer.toString('latin1', 30, 38), 'mimetype');
  const entries = unzip(buffer);
  assert.equal(entries.get('mimetype').toString(), 'application/vnd.oasis.opendocument.text');
});

test('ods round-trips values, formulas and repeated rows', () => {
  const doc = {
    sheets: [
      { name: 'First', cells: { A1: '1', A2: '2', A3: '=SUM(A1:A2)', B1: 'text' } },
      { name: 'Gap', cells: { A1: 'top', A400: 'bottom' } },
    ],
    active: 0,
  };
  const back = odsToSheets(sheetsToOds(doc));

  assert.equal(back.sheets.length, 2);
  assert.equal(back.sheets[0].name, 'First');
  assert.equal(back.sheets[0].cells.A3, '=SUM(A1:A2)');
  assert.equal(back.sheets[0].cells.B1, 'text');
  assert.equal(back.sheets[1].cells.A400, 'bottom', 'the blank run between rows survived');
  assert.equal(back.sheets[1].cells.A2, undefined);
});

test('ods collapses repeated empty cells rather than writing them out', () => {
  const buffer = sheetsToOds({ sheets: [{ name: 'S', cells: { A1: '1', CV1: '2' } }] });
  const content = unzip(buffer).get('content.xml').toString();
  assert.match(content, /table:number-columns-repeated="98"/);
});

// ---------------------------------------------------------------------- rtf

test('rtf export flattens to text that imports back cleanly', () => {
  const { text } = rtfToText(htmlToRtf('<h1>Heading</h1><p>One &amp; two</p><ul><li>item</li></ul>'));
  assert.match(text, /Heading/);
  assert.match(text, /One & two/);
  assert.match(text, /item/);
});

test('rtf import reads cp1252 escapes, unicode and paragraph breaks', () => {
  const { text } = rtfToText("{\\rtf1\\ansi{\\fonttbl{\\f0 Arial;}}\\pard Hello\\'92s \\u8364? world\\par second\\par}");
  assert.equal(text, 'Hello\u2019s \u20ac world\nsecond\n');
});

test('rtf import skips destination groups such as \\*\\generator', () => {
  const { text } = rtfToText('{\\rtf1\\ansi{\\*\\generator Microsoft Word;}\\pard kept\\par}');
  assert.equal(text, 'kept\n');
});

// ------------------------------------------------------------------ convert

test('parseDocument and serializeDocument agree across every supported format', () => {
  const wordFormats = ['docx', 'odt', 'rtf', 'txt', 'html'];
  for (const format of wordFormats) {
    const source = { title: 'Round trip', html: '<h1>Round trip</h1><p>Body <strong>text</strong>.</p>' };
    const written = serializeDocument(source, format, 'word');
    const parsed = parseDocument(`sample.${format}`, written.buffer);

    assert.equal(parsed.kind, 'word', `${format} parsed as word`);
    assert.match(parsed.doc.html, /Round trip/, `${format} kept the heading text`);
  }

  for (const format of ['xlsx', 'ods', 'csv']) {
    const source = { sheets: [{ name: 'Sheet1', cells: { A1: '4', B1: '5', C1: '=A1+B1', A2: 'label' } }], active: 0 };
    const written = serializeDocument(source, format, 'sheet');
    const parsed = parseDocument(`sample.${format}`, written.buffer);

    assert.equal(parsed.kind, 'sheet', `${format} parsed as sheet`);
    assert.equal(show(parsed.doc, 'C1'), '9', `${format} recomputed the formula`);
    assert.equal(parsed.doc.sheets[0].cells.A2, 'label', `${format} kept the text cell`);
  }
});

test('csv import maps rows and columns onto a single sheet', () => {
  const parsed = parseDocument('list.csv', Buffer.from('name,qty\nwidget,3\n'));
  assert.equal(parsed.kind, 'sheet');
  assert.equal(parsed.doc.sheets[0].cells.A1, 'name');
  assert.equal(parsed.doc.sheets[0].cells.B2, '3');
});

test('importing an unknown or legacy binary extension explains what to do', () => {
  assert.throws(() => parseDocument('notes.pdf', Buffer.from('x')), /unsupported file type/);
  assert.throws(() => parseDocument('old.doc', Buffer.from('x')), /legacy binary format.*\.docx/s);
  assert.throws(() => parseDocument('old.xls', Buffer.from('x')), /legacy binary format.*\.xlsx/s);
});

test('serializeDocument sanitises markup before it can be stored', () => {
  const written = serializeDocument({ html: '<p onclick="x">ok</p><script>bad()</script>' }, 'html', 'word');
  assert.equal(written.buffer.toString(), '<p>ok</p>');
});

test('importing an HTML document keeps its structure but drops its scripts', () => {
  const parsed = parseDocument('page.html', Buffer.from('<h2>Title</h2><p>Body</p><script>bad()</script>'));
  assert.equal(parsed.kind, 'word');
  assert.equal(parsed.doc.html, '<h2>Title</h2><p>Body</p>');
});
