#!/usr/bin/env node
// One-off: build templates/appendix-cover.docx from an RMIT student
// appendices example. Keeps the cover page and branding, and swaps the
// course-specific details for {{fields}} that the builder fills in.
// Usage: node test/make-appendix-cover.js <example.docx> <out.docx>
const fs = require('fs');
const JSZip = require('jszip');
const K = require('../lib/docx-kit.js');

const replaceParagraph = (xml, match, text) => {
  let hits = 0;
  const out = xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (p) => {
    if (!match(K.clean(K.textOf(p)))) return p;
    hits++;
    return K.setParagraphText(p, text);
  });
  if (!hits) throw new Error('No paragraph matched for: ' + text);
  return out;
};

(async () => {
  const [src, out] = process.argv.slice(2);
  const zip = await JSZip.loadAsync(fs.readFileSync(src));
  let doc = await zip.file('word/document.xml').async('string');
  const body = K.inner(doc, 'w:body');
  const items = K.splitTop(body.content);
  const firstAppendix = items.findIndex((i) => i.tag === 'w:p' && /^APPENDIX\s+\d+/i.test(K.clean(K.textOf(i.xml))));
  if (firstAppendix < 0) throw new Error('Could not find "APPENDIX 1" in the example');
  // cover = everything before the first appendix heading, minus trailing blank paragraphs
  let cover = items.slice(0, firstAppendix).filter((i) => i.tag !== 'w:sectPr');
  while (cover.length && cover[cover.length - 1].tag === 'w:p' && !K.clean(K.textOf(cover[cover.length - 1].xml))) cover.pop();
  let xml = cover.map((i) => i.xml).join('');

  xml = replaceParagraph(xml, (t) => /^\d{5}VIC|Certificate|Diploma/.test(t) && t.includes('|'), '{{qualCode}} {{qualTitle}} | {{programCode}}');
  xml = replaceParagraph(xml, (t) => /^AT\d+$/.test(t), '{{assessmentTask}}');
  xml = replaceParagraph(xml, (t) => /^[A-Z]{2,}\d{3,}[A-Z]?\s+\S/.test(t) && !/^Appendix/i.test(t), '{{unitCode}} {{unitTitle}}');
  xml = replaceParagraph(xml, (t) => /^Name the file /.test(t), 'Name the file YourName_{{unitCode}}_{{assessmentTask}}.pdf. Submit photographs separately as JPEG or PNG files where required by the assessor.');

  // Evidence table: keep the header row and one "{{appendixNumber}}" row.
  let seenTemplateRow = false;
  xml = xml.replace(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g, (tr) => {
    const cells = [...tr.matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((m) => m[0]);
    const first = K.clean(K.textOf(cells[0] || ''));
    if (!/^Appendix\s+\d+$/i.test(first)) return tr;
    if (seenTemplateRow) return '';
    seenTemplateRow = true;
    let row = tr.replace(cells[0], replaceParagraph(cells[0], () => true, '{{appendixLabel}}'));
    row = row.replace(cells[1], replaceParagraph(cells[1], (t) => t.length > 0, '{{appendixTitle}}'));
    return row;
  });
  if (!seenTemplateRow) throw new Error('Evidence table rows not found');

  doc = doc.slice(0, body.openEnd) + xml + items.filter((i) => i.tag === 'w:sectPr').map((i) => i.xml).join('') + doc.slice(body.closeStart);
  zip.file('word/document.xml', doc);

  for (const name of Object.keys(zip.files)) {
    if (!/^word\/footer\d*\.xml$/.test(name)) continue;
    let f = await zip.file(name).async('string');
    f = f.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (p) => (/\|\s*Student appendices/i.test(K.clean(K.textOf(p))) ? K.setParagraphText(p, '{{footerLine}}') : p));
    zip.file(name, f);
  }
  const core = zip.file('docProps/core.xml');
  if (core) zip.file('docProps/core.xml', (await core.async('string')).replace(/<dc:title>[\s\S]*?<\/dc:title>/, '<dc:title>Student assessment appendices</dc:title>'));
  fs.writeFileSync(out, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log('Wrote', out);
  console.log(K.textOf(xml).replace(/\n+/g, '\n'));
})().catch((e) => { console.error(e); process.exit(1); });
