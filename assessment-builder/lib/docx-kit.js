/*
 * docx-kit: small toolkit for merging Word (.docx) bodies into an RMIT
 * branded "shell" document, working directly on the WordprocessingML text.
 *
 * Works in the browser and in Node. Needs JSZip (passed in as zip objects).
 *
 * Main pieces:
 *   - XML helpers: splitTop, inner, textOf, setParagraphText, fillTokens
 *   - Output: an output package (the shell .docx) that several source
 *     documents can be imported into, one after another. Each import brings
 *     its numbering (bullets/lists), images, hyperlinks and styles along.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.DocxKit = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ------------------------------------------------------------ XML helpers

  function findClose(xml, tag, start) {
    const esc = tag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const rx = new RegExp(`<(/?)${esc}(?=[\\s>/])[^>]*?(/?)>`, 'g');
    rx.lastIndex = start;
    let depth = 0;
    let m;
    while ((m = rx.exec(xml))) {
      if (m[2] === '/') {
        if (depth === 0) return m.index + m[0].length;
        continue;
      }
      depth += m[1] ? -1 : 1;
      if (depth === 0) return m.index + m[0].length;
    }
    throw new Error(`Malformed XML: no closing </${tag}>`);
  }

  /** Top-level elements of an XML fragment, as {tag, xml}. */
  function splitTop(xml) {
    const out = [];
    const open = /<([A-Za-z][\w.-]*:[\w.-]+|[A-Za-z][\w.-]*)\b[^>]*?(\/?)>/g;
    let i = 0;
    while (i < xml.length) {
      open.lastIndex = i;
      const m = open.exec(xml);
      if (!m) break;
      if (m[2] === '/') {
        out.push({ tag: m[1], xml: m[0] });
        i = m.index + m[0].length;
        continue;
      }
      const end = findClose(xml, m[1], m.index);
      out.push({ tag: m[1], xml: xml.slice(m.index, end) });
      i = end;
    }
    return out;
  }

  /** First <tag>…</tag> in xml: positions plus inner content. */
  function inner(xml, tag) {
    const s = xml.search(new RegExp(`<${tag}(?=[\\s>/])`));
    if (s < 0) return null;
    const e = findClose(xml, tag, s);
    const openEnd = xml.indexOf('>', s) + 1;
    if (xml[openEnd - 2] === '/') return { start: s, end: e, openEnd, closeStart: openEnd, content: '', selfClosing: true };
    const closeStart = xml.lastIndexOf(`</${tag}>`, e);
    return { start: s, end: e, openEnd, closeStart, content: xml.slice(openEnd, closeStart) };
  }

  const decode = (s) =>
    s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
      .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n)).replace(/&amp;/g, '&');
  const encode = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const clean = (s) => String(s || '').replace(/\s+/g, ' ').trim();

  /** Visible text (paragraph ends become newlines). */
  function textOf(xml) {
    let out = '';
    const rx = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:(?:tab|br|cr)\b[^>]*\/>|<\/w:p>/g;
    let m;
    while ((m = rx.exec(xml))) {
      if (m[1] !== undefined) out += decode(m[1]);
      else if (m[0] === '</w:p>') out += '\n';
      else out += ' ';
    }
    return out;
  }

  const styleOf = (pXml) => (/<w:pStyle w:val="([^"]+)"/.exec(pXml) || [])[1] || null;

  /**
   * Replace a paragraph's text, keeping its paragraph properties and the
   * formatting of its first run. Used for filling {{fields}} that Word may
   * have split across several runs.
   */
  function setParagraphText(pXml, text) {
    const p = inner(pXml, 'w:p');
    const pPr = (/^\s*<w:pPr>[\s\S]*?<\/w:pPr>/.exec(p.content) || [''])[0];
    const firstRun = /<w:r\b[^>]*>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)?/.exec(p.content);
    const rPr = firstRun && firstRun[1] ? firstRun[1] : '';
    const lines = String(text).split('\n');
    const runs = lines
      .map((l, i) => (i ? `<w:r>${rPr}<w:br/></w:r>` : '') + (l ? `<w:r>${rPr}<w:t xml:space="preserve">${encode(l)}</w:t></w:r>` : ''))
      .join('');
    return pXml.slice(0, p.openEnd) + pPr + runs + pXml.slice(p.closeStart);
  }

  /** Fill {{field}} tokens in every paragraph of an XML part. */
  function fillTokens(xml, values) {
    return xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>|<w:p\b[^>]*\/>/g, (pXml) => {
      if (pXml.endsWith('/>') && !pXml.includes('</w:p>')) return pXml;
      const t = textOf(pXml).replace(/\n$/, '');
      if (!/\{\{\s*\w+\s*\}\}/.test(t)) return pXml;
      const filled = t.replace(/\{\{\s*(\w+)\s*\}\}/g, (_, k) => (values[k] !== undefined && values[k] !== null ? String(values[k]) : ''));
      return setParagraphText(pXml, filled);
    });
  }

  /**
   * In each table, find rows containing a {{marker}} field and repeat that
   * row once per item, filling the fields from the item.
   */
  function repeatRows(xml, marker, items) {
    return xml.replace(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g, (tr) => {
      if (!new RegExp(`\\{\\{\\s*(?:${marker})\\s*\\}\\}`).test(textOf(tr))) return tr;
      return items.map((it) => fillTokens(tr, it)).join('');
    });
  }

  // ------------------------------------------------------------ RMIT look

  const GREY_HEADER = 'E1E1E1';
  const GREY_LABEL = 'F2F2F2';

  /** Remove off-brand fonts/colours so text inherits the RMIT theme. */
  function scrubFormatting(xml) {
    return xml
      .replace(/<w:rFonts\b[^>]*\/>/g, (tag) => {
        // keep only complex-script / east-asian hints; drop explicit Latin fonts
        const keep = (tag.match(/\sw:(?:cs|cstheme|eastAsia|eastAsiaTheme|hint)="[^"]*"/g) || []).join('');
        return keep ? `<w:rFonts${keep}/>` : '';
      })
      .replace(/<w:color w:val="(?:FFFFFF|ffffff)"[^>]*\/>/g, '')
      .replace(/<w:color w:val="[0-9A-Fa-f]{6}"[^>]*w:themeColor="(?:accent\d|text2|background2)"[^>]*\/>/g, '')
      .replace(/<w:color w:val="(?:002060|1F3864|1F4E79|2F5496|4472C4|203864|0F243E|17365D|365F91|0070C0|2E74B5)"[^>]*\/>/gi, '');
  }

  /** Plain RMIT table: grid borders, grey header row, light grey label column. */
  function restyleTable(tblXml) {
    const t = inner(tblXml, 'w:tbl');
    const tblPrM = /<w:tblPr>[\s\S]*?<\/w:tblPr>/.exec(t.content);
    let tblPr = tblPrM ? tblPrM[0] : '<w:tblPr></w:tblPr>';
    const look = (/<w:tblLook [^>]*\/>/.exec(tblPr) || [''])[0];
    const bits = parseInt((/w:val="([0-9A-Fa-f]{4})"/.exec(look) || [, '04A0'])[1], 16);
    const flag = (name, bit) => (new RegExp(`w:${name}="(1|true)"`).test(look) || (!new RegExp(`w:${name}=`).test(look) && (bits & bit) !== 0));
    const firstRow = flag('firstRow', 0x20);
    const firstCol = flag('firstColumn', 0x80) && /<w:tblStyle w:val="(?!TableGrid)/.test(tblPr);
    tblPr = tblPr.replace(/<w:tblStyle w:val="[^"]*"\/>/, '<w:tblStyle w:val="TableGrid"/>');
    if (!/<w:tblStyle /.test(tblPr)) tblPr = tblPr.replace('<w:tblPr>', '<w:tblPr><w:tblStyle w:val="TableGrid"/>');
    tblPr = tblPr.replace(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/, '');
    const rest = tblPrM ? t.content.replace(tblPrM[0], '') : t.content;
    let row = 0;
    const parts = splitTop(rest).map((p) => {
      if (p.tag !== 'w:tr') return p.xml;
      const header = row++ === 0 && firstRow;
      const r = inner(p.xml, 'w:tr');
      let col = 0;
      const cells = splitTop(r.content).map((c) => {
        if (c.tag !== 'w:tc') return c.xml;
        const idx = col++;
        return restyleCell(c.xml, header ? GREY_HEADER : !header && firstCol && idx === 0 ? GREY_LABEL : null);
      });
      return p.xml.slice(0, r.openEnd) + cells.join('') + p.xml.slice(r.closeStart);
    });
    return tblXml.slice(0, t.openEnd) + tblPr + parts.join('') + tblXml.slice(t.closeStart);
  }

  function restyleCell(tcXml, fill) {
    const t = inner(tcXml, 'w:tc');
    let content = t.content;
    let tcPr = '';
    const m = /^\s*<w:tcPr>[\s\S]*?<\/w:tcPr>/.exec(content);
    if (m) { tcPr = m[0]; content = content.slice(m[0].length); }
    const ex = /<w:shd [^>]*w:fill="([0-9A-Fa-f]{6})"/.exec(tcPr);
    if (!fill && ex && ex[1].toUpperCase() !== 'FFFFFF') fill = GREY_HEADER; // coloured cells → grey
    tcPr = tcPr.replace(/<w:shd [^>]*\/>/g, '').replace(/<w:tcBorders>[\s\S]*?<\/w:tcBorders>/, '');
    if (fill) {
      if (!tcPr) tcPr = '<w:tcPr></w:tcPr>';
      const shd = `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`;
      const after = /<w:(noWrap|tcMar|textDirection|tcFitText|vAlign|hideMark|headers|cellIns|cellDel|cellMerge|tcPrChange)\b/.exec(tcPr);
      tcPr = after ? tcPr.slice(0, after.index) + shd + tcPr.slice(after.index) : tcPr.replace('</w:tcPr>', shd + '</w:tcPr>');
    }
    const parts = splitTop(content).map((p) => (p.tag === 'w:tbl' ? restyleTable(p.xml) : p.xml));
    return tcXml.slice(0, t.openEnd) + tcPr + parts.join('') + tcXml.slice(t.closeStart);
  }

  // ------------------------------------------------------------ packages

  function parseRels(xml) {
    const map = {};
    const rx = /<Relationship\b([^>]*)\/?>/g;
    let m;
    while ((m = rx.exec(xml))) {
      const g = (n) => (new RegExp(`\\b${n}="([^"]*)"`).exec(m[1]) || [])[1];
      map[g('Id')] = { type: g('Type'), target: g('Target'), mode: g('TargetMode') };
    }
    return map;
  }

  function mergeRootNamespaces(dstXml, srcXml, rootTag) {
    const rx = new RegExp(`<${rootTag}\\b[^>]*>`);
    const d = rx.exec(dstXml);
    const s = rx.exec(srcXml);
    if (!d || !s) return dstXml;
    let root = d[0];
    const have = new Set([...root.matchAll(/\sxmlns:(\w+)=/g)].map((m) => m[1]));
    const add = [...s[0].matchAll(/\sxmlns:(\w+)="([^"]*)"/g)].filter((m) => !have.has(m[1])).map((m) => ` xmlns:${m[1]}="${m[2]}"`);
    if (add.length) root = root.replace(new RegExp(`^<${rootTag}`), `<${rootTag}${add.join('')}`);
    const ig = (x) => ((/mc:Ignorable="([^"]*)"/.exec(x) || [, ''])[1]).split(/\s+/).filter(Boolean);
    const union = [...new Set([...ig(root), ...ig(s[0])])];
    if (union.length) {
      root = /mc:Ignorable=/.test(root) ? root.replace(/mc:Ignorable="[^"]*"/, `mc:Ignorable="${union.join(' ')}"`) : root.replace(/>$/, ` mc:Ignorable="${union.join(' ')}">`);
    }
    return dstXml.replace(d[0], root);
  }

  const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', emf: 'image/x-emf', wmf: 'image/x-wmf', svg: 'image/svg+xml', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff' };
  const hex8 = () => Array.from({ length: 8 }, () => '0123456789ABCDEF'[Math.floor(Math.random() * 16)]).join('');

  /**
   * An output document built from a shell .docx (header, footer, styles,
   * page setup). Call importBody() for each source, then save().
   */
  class Output {
    static async open(shellZip) {
      const o = new Output();
      o.zip = shellZip;
      o.doc = await shellZip.file('word/document.xml').async('string');
      o.rels = await shellZip.file('word/_rels/document.xml.rels').async('string');
      o.ct = await shellZip.file('[Content_Types].xml').async('string');
      o.styles = await shellZip.file('word/styles.xml').async('string');
      const nf = shellZip.file('word/numbering.xml');
      o.numbering = nf ? await nf.async('string') : null;
      o.styleIds = new Set([...o.styles.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"/g)].map((m) => m[1]));
      o.relCounter = 1;
      o.mediaCounter = 1;
      o.bookmarkOffset = 100000;
      o.warnings = [];
      const body = inner(o.doc, 'w:body');
      o.shellItems = splitTop(body.content);
      o.sectPr = (o.shellItems.filter((i) => i.tag === 'w:sectPr').pop() || {}).xml || '';
      return o;
    }

    /** Shell body content (everything except the final sectPr). */
    shellBody() {
      return this.shellItems.filter((i) => i.tag !== 'w:sectPr').map((i) => i.xml).join('');
    }

    addRel(type, target, external) {
      const id = `rIdKit${this.relCounter++}`;
      this.rels = this.rels.replace('</Relationships>', `<Relationship Id="${id}" Type="${type}" Target="${encode(target).replace(/"/g, '&quot;')}"${external ? ' TargetMode="External"' : ''}/></Relationships>`);
      return id;
    }

    ensureDefault(ext) {
      if (new RegExp(`Extension="${ext}"`, 'i').test(this.ct)) return;
      this.ct = this.ct.replace(/(<Types\b[^>]*>)/, `$1<Default Extension="${ext}" ContentType="${MIME[ext] || 'application/octet-stream'}"/>`);
    }

    mergeNumbering(srcNum) {
      if (!this.numbering) {
        const r = inner(srcNum, 'w:numbering');
        this.numbering = srcNum.slice(0, r.openEnd) + '</w:numbering>';
        this.ct = this.ct.replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>');
        this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering', 'numbering.xml');
      }
      const max = (rx) => Math.max(0, ...[...this.numbering.matchAll(rx)].map((m) => +m[1]));
      const absOff = max(/<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"/g) + 1;
      const numOff = max(/<w:num\b[^>]*w:numId="(\d+)"/g);
      const abs = [];
      const nums = [];
      for (const p of splitTop(inner(srcNum, 'w:numbering').content)) {
        if (p.tag === 'w:abstractNum') {
          abs.push(p.xml.replace(/(<w:abstractNum\b[^>]*w:abstractNumId=")(\d+)"/, (_, a, n) => `${a}${+n + absOff}"`)
            .replace(/<w:nsid w:val="[0-9A-Fa-f]{8}"\/>/, () => `<w:nsid w:val="${hex8()}"/>`)
            .replace(/<w:(?:numStyleLink|styleLink)\b[^>]*\/>/g, ''));
        } else if (p.tag === 'w:num') {
          nums.push(p.xml.replace(/(<w:num\b[^>]*w:numId=")(\d+)"/, (_, a, n) => `${a}${+n + numOff}"`)
            .replace(/(<w:abstractNumId w:val=")(\d+)"/, (_, a, n) => `${a}${+n + absOff}"`));
        }
      }
      this.numbering = mergeRootNamespaces(this.numbering, srcNum, 'w:numbering');
      const ni = inner(this.numbering, 'w:numbering');
      const cur = splitTop(ni.content);
      const order = [
        ...cur.filter((p) => p.tag === 'w:numPicBullet'),
        ...cur.filter((p) => p.tag === 'w:abstractNum'), ...abs,
        ...cur.filter((p) => p.tag === 'w:num'), ...nums,
        ...cur.filter((p) => !['w:numPicBullet', 'w:abstractNum', 'w:num'].includes(p.tag)),
      ].map((p) => (typeof p === 'string' ? p : p.xml));
      this.numbering = this.numbering.slice(0, ni.openEnd) + order.join('') + this.numbering.slice(ni.closeStart);
      return numOff;
    }

    /**
     * Bring a source document's body into this package and return its XML.
     * opts.restyle: true = RMIT look (fonts, tables); false = keep formatting.
     * opts.dropIndex: index of one body element to leave out (e.g. the title).
     */
    async importBody(srcZip, opts = {}) {
      const label = opts.label || 'document';
      const srcDoc = await srcZip.file('word/document.xml').async('string');
      this.doc = mergeRootNamespaces(this.doc, srcDoc, 'w:document');
      let items = splitTop(inner(srcDoc, 'w:body').content).filter((i) => i.tag !== 'w:sectPr');
      if (opts.dropIndex >= 0) items = items.filter((_, i) => i !== opts.dropIndex);
      let xml = items.map((i) => i.xml).join('');

      // section breaks inside the source become page breaks
      xml = xml.replace(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<w:sectPr\b[\s\S]*?<\/w:sectPr>[\s\S]*?<\/w:p>/g, (p) =>
        p.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/, '').replace(/<\/w:p>$/, '<w:r><w:br w:type="page"/></w:r></w:p>'));

      // relationships: hyperlinks and images
      const srcRels = parseRels(await srcZip.file('word/_rels/document.xml.rels').async('string'));
      const map = {};
      for (const id of new Set([...xml.matchAll(/\br:(?:id|embed|link|pict)="([^"]+)"/g)].map((m) => m[1]))) {
        const rel = srcRels[id];
        if (!rel) continue;
        const type = rel.type.split('/').pop();
        if (type === 'hyperlink') map[id] = this.addRel(rel.type, rel.target, true);
        else if (type === 'image') {
          if (rel.mode === 'External') { map[id] = this.addRel(rel.type, rel.target, true); continue; }
          const f = srcZip.file('word/' + rel.target.replace(/^\.?\//, '').replace(/^word\//, ''));
          if (!f) { this.warnings.push(`${label}: image ${rel.target} is missing.`); continue; }
          const ext = rel.target.split('.').pop().toLowerCase();
          const name = `media/kit${this.mediaCounter++}.${ext}`;
          this.zip.file('word/' + name, await f.async('uint8array'));
          this.ensureDefault(ext);
          map[id] = this.addRel(rel.type, name);
        } else {
          this.warnings.push(`${label}: an embedded ${type} (${rel.target}) could not be carried over.`);
        }
      }
      xml = xml.replace(/\b(r:(?:id|embed|link|pict))="([^"]+)"/g, (s, a, id) => (map[id] ? `${a}="${map[id]}"` : s));

      // numbering
      const nf = srcZip.file('word/numbering.xml');
      if (nf) {
        const off = this.mergeNumbering(await nf.async('string'));
        if (off) xml = xml.replace(/<w:numId w:val="(\d+)"\/>/g, (_, n) => `<w:numId w:val="${+n === 0 ? 0 : +n + off}"/>`);
        this._lastNumOffset = off;
      } else this._lastNumOffset = 0;

      // things that point into parts we do not copy
      if (/<w:footnoteReference\b|<w:endnoteReference\b/.test(xml)) {
        this.warnings.push(`${label}: footnotes/endnotes were removed (the note text is not carried over).`);
        xml = xml.replace(/<w:r\b[^>]*>(?:(?!<\/w:r>)[\s\S])*?<w:(?:footnote|endnote)Reference\b[^>]*\/>[\s\S]*?<\/w:r>/g, '');
      }
      if (/<w:comment(?:RangeStart|RangeEnd|Reference)\b/.test(xml)) {
        this.warnings.push(`${label}: Word comments were removed.`);
        xml = xml.replace(/<w:commentRange(?:Start|End)\b[^>]*\/>/g, '').replace(/<w:r\b[^>]*>(?:(?!<\/w:r>)[\s\S])*?<w:commentReference\b[^>]*\/>[\s\S]*?<\/w:r>/g, '');
      }

      // tidy ids that could collide between merged documents
      const bo = this.bookmarkOffset;
      this.bookmarkOffset += 100000;
      xml = xml
        .replace(/\sw:rsid\w*="[^"]*"/g, '').replace(/\sw14:(?:paraId|textId)="[^"]*"/g, '')
        .replace(/<w:proofErr\b[^>]*\/>/g, '').replace(/<w:lastRenderedPageBreak\/>/g, '')
        .replace(/(<w:bookmark(?:Start|End)\b[^>]*w:id=")(\d+)"/g, (_, a, n) => `${a}${+n + bo}"`)
        .replace(/(<wp:docPr\b[^>]*\sid=")(\d+)"/g, (_, a, n) => `${a}${+n + bo}"`)
        .replace(/<w:pPr>\s*<\/w:pPr>/g, '');

      // styles the content needs
      await this.importStyles(srcZip, xml, opts.restyle);

      if (opts.restyle) {
        xml = scrubFormatting(xml);
        xml = splitTop(xml).map((p) => (p.tag === 'w:tbl' ? restyleTable(p.xml) : p.xml)).join('');
      }
      return xml;
    }

    async importStyles(srcZip, xml, restyle) {
      const sf = srcZip.file('word/styles.xml');
      if (!sf) return;
      const src = await sf.async('string');
      const defs = {};
      for (const s of splitTop(inner(src, 'w:styles').content)) {
        if (s.tag !== 'w:style') continue;
        const id = (/w:styleId="([^"]+)"/.exec(s.xml) || [])[1];
        if (id) defs[id] = s.xml;
      }
      const need = new Set([...xml.matchAll(/<w:(?:pStyle|rStyle|tblStyle)\s+w:val="([^"]+)"/g)].map((m) => m[1]));
      const queue = [...need];
      while (queue.length) {
        const s = defs[queue.pop()];
        if (!s) continue;
        for (const m of s.matchAll(/<w:(?:basedOn|link|next)\s+w:val="([^"]+)"/g)) if (!need.has(m[1])) { need.add(m[1]); queue.push(m[1]); }
      }
      const add = [];
      for (const id of need) {
        if (this.styleIds.has(id) || !defs[id]) continue; // shell's own style wins
        let s = defs[id].replace(/\sw:rsid\w*="[^"]*"/g, '').replace(/<w:rsid\b[^>]*\/>/g, '');
        if (restyle) s = scrubFormatting(s);
        if (this._lastNumOffset) s = s.replace(/<w:numId w:val="(\d+)"\/>/g, (_, n) => `<w:numId w:val="${+n === 0 ? 0 : +n + this._lastNumOffset}"/>`);
        add.push(s);
        this.styleIds.add(id);
      }
      if (add.length) this.styles = this.styles.replace('</w:styles>', add.join('') + '</w:styles>');
      this.styles = mergeRootNamespaces(this.styles, src, 'w:styles');
    }

    /** Write the body and all changed parts; returns the JSZip. */
    finish(bodyXml) {
      const b = inner(this.doc, 'w:body');
      this.doc = this.doc.slice(0, b.openEnd) + bodyXml + this.sectPr + this.doc.slice(b.closeStart);
      this.zip.file('word/document.xml', this.doc);
      this.zip.file('word/_rels/document.xml.rels', this.rels);
      this.zip.file('[Content_Types].xml', this.ct);
      this.zip.file('word/styles.xml', this.styles);
      if (this.numbering) this.zip.file('word/numbering.xml', this.numbering);
      return this.zip;
    }

    /** Apply fillTokens to every header and footer part. */
    async fillHeadersFooters(values) {
      for (const name of Object.keys(this.zip.files)) {
        if (/^word\/(header|footer)\d*\.xml$/.test(name)) {
          this.zip.file(name, fillTokens(await this.zip.file(name).async('string'), values));
        }
      }
    }

    async setTitle(title) {
      const f = this.zip.file('docProps/core.xml');
      if (!f) return;
      let c = await f.async('string');
      c = /<dc:title>/.test(c) ? c.replace(/<dc:title>[\s\S]*?<\/dc:title>|<dc:title\/>/, `<dc:title>${encode(title)}</dc:title>`) : c.replace('</cp:coreProperties>', `<dc:title>${encode(title)}</dc:title></cp:coreProperties>`);
      this.zip.file('docProps/core.xml', c);
    }
  }

  /** Best guess at a document's title: Title / Heading 1 / first text. */
  async function guessTitle(srcZip, fallback) {
    const doc = await srcZip.file('word/document.xml').async('string');
    const items = splitTop(inner(doc, 'w:body').content);
    let firstText = null;
    for (let i = 0; i < items.length && i < 40; i++) {
      const it = items[i];
      if (it.tag !== 'w:p') continue;
      const t = clean(textOf(it.xml));
      if (!t) continue;
      const st = styleOf(it.xml) || '';
      if (/^(Title|Heading1|Heading 1)$/i.test(st)) return { title: t, index: i, fromHeading: true };
      if (firstText === null) {
        // A short first line that is all bold (or 14pt+) reads as a title too.
        const runs = [...it.xml.matchAll(/<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g)].map((m) => m[1]).filter((r) => /<w:t[\s>]/.test(r));
        const pBold = /<w:pPr>[\s\S]*?<w:rPr>[\s\S]*?<w:b(?:\s+w:val="(?:1|true)")?\s*\/>[\s\S]*?<\/w:pPr>/.test(it.xml);
        const titleLike = t.length <= 120 && runs.length > 0 && runs.every((r) =>
          /<w:b(?:\s+w:val="(?:1|true)")?\s*\/>/.test(r) || pBold || +((/<w:sz w:val="(\d+)"/.exec(r) || [])[1] || 0) >= 28);
        firstText = { title: t, index: i, fromHeading: titleLike };
        if (titleLike) return firstText;
      }
    }
    if (firstText && firstText.title.length <= 120) return firstText;
    return { title: fallback, index: -1, fromHeading: false };
  }

  return {
    splitTop, inner, textOf, decode, encode, clean, styleOf, findClose,
    setParagraphText, fillTokens, repeatRows, scrubFormatting, restyleTable,
    parseRels, mergeRootNamespaces, Output, guessTitle,
  };
});
