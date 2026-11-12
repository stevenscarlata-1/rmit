#!/usr/bin/env node
// Usage: node test/appendices-cli.js [--separate] [--no-cover] [--keep-format] doc1.docx doc2.docx ...
// Builds student appendices with the same code the web page uses.
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const A = require('../lib/appendices.js');

(async () => {
  const args = process.argv.slice(2);
  const flags = new Set(args.filter((a) => a.startsWith('--')));
  const inputs = args.filter((a) => !a.startsWith('--'));
  if (!inputs.length) {
    console.error('Usage: node test/appendices-cli.js [--separate] [--no-cover] [--keep-format] doc1.docx ...');
    process.exit(1);
  }
  const load = (b) => JSZip.loadAsync(b);
  const docs = [];
  for (const f of inputs) {
    const bytes = fs.readFileSync(f);
    const t = await A.readTitle(await load(bytes), path.basename(f));
    docs.push({ bytes, name: path.basename(f), title: t.title, dropTitleIndex: t.index });
  }
  const info = {
    qualCode: '22682VIC', qualTitle: 'Certificate II in Electrotechnology (Pre-vocational)', programCode: 'C2229',
    assessmentTask: 'AT2', unitCode: 'VU23735', unitTitle: 'Install sustainable extra low voltage energy supply systems',
  };
  const cover = fs.readFileSync(path.join(__dirname, '../templates/appendix-cover.docx'));
  const { files, report } = await A.build(load, cover, docs, info, {
    combine: !flags.has('--separate'), includeCover: !flags.has('--no-cover'), restyle: !flags.has('--keep-format'),
  });
  const outDir = path.join(__dirname, 'output');
  fs.mkdirSync(outDir, { recursive: true });
  for (const f of files) {
    fs.writeFileSync(path.join(outDir, f.name), await f.zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
    console.log('Wrote', path.join(outDir, f.name));
  }
  console.log(JSON.stringify(report, null, 2));
})().catch((e) => { console.error(e); process.exit(1); });
