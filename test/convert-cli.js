#!/usr/bin/env node
// Usage: node test/convert-cli.js <tvsc.docx> [template.docx] [out.docx]
// Runs the same converter the web page uses, from the command line.
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const C = require('../converter.js');

(async () => {
  const [src, tpl = path.join(__dirname, '../template/rmit-assessor-guide-template.docx'), outArg] = process.argv.slice(2);
  if (!src) {
    console.error('Usage: node test/convert-cli.js <tvsc.docx> [template.docx] [out.docx]');
    process.exit(1);
  }
  const tv = await JSZip.loadAsync(fs.readFileSync(src));
  const info = await C.extractInfo(tv);
  info.footerLine = C.defaultFooter(info);
  console.log('Extracted:', JSON.stringify({ ...info, headings: info.headings.length + ' headings' }, null, 2));
  const tp = await JSZip.loadAsync(fs.readFileSync(tpl));
  const { zip, report } = await C.convert(tv, tp, info);
  const out = outArg || path.join(__dirname, 'output', C.defaultFileName(info));
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log('Report:', JSON.stringify(report, null, 2));
  console.log('Wrote', out);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
