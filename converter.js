/*
 * TVSC → RMIT Assessor Guide converter.
 *
 * Takes a TVSC assessor guide (.docx) and the RMIT assessor guide template
 * (.docx), and produces a new .docx that:
 *   - keeps the RMIT letterhead, "ASSESSOR" tab, footer, page setup and
 *     cover information table from the template;
 *   - carries over EVERY section of the TVSC guide in its original order and
 *     wording (headings, paragraphs, nested lists, tables, merged cells,
 *     bold/italic/red text, hyperlinks, images);
 *   - restyles that content in the RMIT look (Calibri, black headings,
 *     grey table header rows, plain grid tables).
 *
 * Runs entirely in the browser (or Node for testing). Needs JSZip.
 * Works on the raw WordprocessingML strings so nothing is lost in a
 * parse/re-serialise round trip.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.AssessorConverter = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const GREY_HEADER = 'D9D9D9'; // RMIT template label / header cells
  const GREY_LABEL = 'F2F2F2'; // lighter grey for first-column labels
  const SHD = (fill) =>
    `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`;

  // ---------------------------------------------------------------- XML utils

  /** Split an XML fragment into its top-level elements (keeps raw text). */
  function splitTop(xml) {
    const out = [];
    const open = /<([A-Za-z][\w.-]*:[\w.-]+|[A-Za-z][\w.-]*)\b[^>]*?(\/?)>/g;
    let i = 0;
    while (i < xml.length) {
      open.lastIndex = i;
      const m = open.exec(xml);
      if (!m) break;
      const tag = m[1];
      if (m[2] === '/') {
        out.push({ tag, xml: m[0] });
        i = m.index + m[0].length;
        continue;
      }
      const end = findClose(xml, tag, m.index);
      out.push({ tag, xml: xml.slice(m.index, end) });
      i = end;
    }
    return out;
  }

  /** Index just past the element (named `tag`) that starts at `start`. */
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

  /** Inner content of the first <tag>...</tag> in xml. */
  function inner(xml, tag) {
    const s = xml.search(new RegExp(`<${tag}(?=[\\s>])`));
    if (s < 0) return null;
    const e = findClose(xml, tag, s);
    const openEnd = xml.indexOf('>', s) + 1;
    const closeStart = xml.lastIndexOf(`</${tag}>`, e);
    return { start: s, end: e, openEnd, closeStart, content: xml.slice(openEnd, closeStart) };
  }

  /** All direct children <tag> elements inside a container string. */
  function children(xml, tag) {
    return splitTop(xml).filter((c) => c.tag === tag);
  }

  const decode = (s) =>
    s
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n))
      .replace(/&amp;/g, '&');
  const encode = (s) =>
    String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  /** Plain text of a paragraph / cell / table (tabs and breaks as spaces). */
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
  const clean = (s) => s.replace(/\s+/g, ' ').trim();

  const styleOf = (pXml) => {
    const m = /<w:pStyle w:val="([^"]+)"/.exec(pXml);
    return m ? m[1] : null;
  };

  // ------------------------------------------------------------- extraction

  /**
   * Read the TVSC guide and pull out the details that go in the RMIT cover
   * table. Everything is a best guess; the UI lets the user edit it.
   */
  async function extractInfo(tvscZip) {
    const doc = await tvscZip.file('word/document.xml').async('string');
    const body = inner(doc, 'w:body').content;
    const items = splitTop(body);
    const info = {
      taskNumber: '',
      taskTotal: '',
      taskName: '',
      unitCodes: '',
      unitTitles: '',
      qualCode: '',
      qualTitle: '',
      programCode: '',
      courseCode: '',
      headings: [],
    };

    // Headings outline (for the preview) and the task title.
    for (const it of items) {
      if (it.tag !== 'w:p') continue;
      const st = styleOf(it.xml);
      const t = clean(textOf(it.xml));
      if (!t) continue;
      const lvl = st && /^Heading(\d)$/.exec(st);
      if (lvl) info.headings.push({ level: +lvl[1], text: t });
      if (lvl && lvl[1] === '1' && !info.taskName && /assessment task/i.test(t)) {
        const at = /\bAT\s*(\d+)\b/i.exec(t);
        if (at) info.taskNumber = at[1];
        info.taskName = clean(t.replace(/^assessment task\s*/i, '').replace(/^AT\s*\d+\s*[-–:]?\s*/i, ''));
      }
    }

    // Qualification and unit lines on the cover page, e.g.
    // "22682VIC Certificate II in Electrotechnology Studies".
    const coverText = textOf(items.slice(0, 12).map((i) => i.xml).join(''));
    for (const line of coverText.split('\n').map(clean)) {
      const q = /^([0-9]{5}VIC|[A-Z]{2,4}[0-9]{5}|[0-9]{5,6}[A-Z]{0,3})\s+((?:Certificate|Diploma|Advanced Diploma|Graduate|Course)\b.*)$/i.exec(line);
      if (q && !info.qualCode) {
        info.qualCode = q[1];
        info.qualTitle = q[2];
      }
    }

    // Tables: assessment details (task number) and units of competency.
    const codes = [];
    const titles = [];
    for (const it of items) {
      if (it.tag !== 'w:tbl') continue;
      const rows = children(inner(it.xml, 'w:tbl').content, 'w:tr').map((r) =>
        children(r.xml, 'w:tc').map((c) => clean(textOf(c.xml)))
      );
      if (!rows.length) continue;
      const head = rows[0].join(' ');
      for (const r of rows) {
        if (/assessment task number/i.test(r[0] || '') && r[1]) {
          const n = /(\d+)/.exec(r[1]);
          if (n && !info.taskNumber) info.taskNumber = n[1];
        }
      }
      if (/unit.*of competency|units? of competency/i.test(head) && !/pre-?requisite/i.test(head)) {
        for (const r of rows.slice(1)) {
          if (/^unit code$/i.test(r[0] || '')) continue;
          if (r[0] && /[A-Z]{3,}[0-9]{3,}/.test(r[0])) {
            codes.push(r[0]);
            titles.push((r[1] || '').replace(/\.$/, ''));
          }
        }
      }
    }
    if (!codes.length) {
      // Fall back to the cover line "UEECD0042 Solve problems in ..."
      for (const line of coverText.split('\n').map(clean)) {
        const u = /^([A-Z]{3,}[0-9]{3,}[A-Z]?)\s+(.+)$/.exec(line);
        if (u && !/^(certificate|diploma)/i.test(u[2]) && !codes.includes(u[1])) {
          codes.push(u[1]);
          titles.push(u[2]);
        }
      }
    }
    info.unitCodes = codes.join(', ');
    info.unitTitles = titles.join('; ');
    return info;
  }

  function defaultFileName(info, date = new Date()) {
    const ymd =
      date.getFullYear() +
      String(date.getMonth() + 1).padStart(2, '0') +
      String(date.getDate()).padStart(2, '0');
    const task = info.taskNumber ? `AT${info.taskNumber}` : 'AT';
    const name = (info.taskName || 'Practical')
      .replace(/[^A-Za-z0-9 ]+/g, ' ')
      .split(/\s+/)
      .filter(Boolean)
      .map((w) => w[0].toUpperCase() + w.slice(1))
      .join('');
    const unit = (info.unitCodes || 'Unit').split(/[,;\s]+/).filter(Boolean).join('-');
    return `${ymd}_${task}_${name}_${unit}_Assessor.docx`;
  }

  function defaultFooter(info, date = new Date()) {
    const month = date.toLocaleString('en-AU', { month: 'long' });
    const unit = (info.unitCodes || 'UoC code').split(/[,;\s]+/).filter(Boolean).join(', ');
    const task = info.taskNumber ? `AT${info.taskNumber}` : 'AT#';
    return `${unit} – ${task} ${info.taskName || 'Practical'} – ${month} ${date.getFullYear()}`;
  }

  // ------------------------------------------------------------- conversion

  /**
   * Restyled definitions for TVSC's paragraph styles, in the RMIT look.
   * Any other TVSC style the content needs is copied across and scrubbed.
   */
  const RMIT_STYLES = {
    Heading1: { name: 'heading 1', outline: 0, sz: 28, before: 240, after: 120 },
    Heading2: { name: 'heading 2', outline: 1, sz: 24, before: 200, after: 80 },
    Heading3: { name: 'heading 3', outline: 2, sz: 22, before: 160, after: 60 },
    Heading4: { name: 'heading 4', outline: 3, sz: 22, before: 120, after: 40, italic: true },
    Heading5: { name: 'heading 5', outline: 4, sz: 20, before: 120, after: 40 },
    Heading6: { name: 'heading 6', outline: 5, sz: 20, before: 120, after: 40, italic: true },
  };

  function headingStyle(id, d) {
    return (
      `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${d.name}"/>` +
      `<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>` +
      `<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${d.before}" w:after="${d.after}" w:line="240" w:lineRule="auto"/>` +
      `<w:outlineLvl w:val="${d.outline}"/></w:pPr>` +
      `<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:b/><w:bCs/>` +
      (d.italic ? '<w:i/><w:iCs/>' : '') +
      `<w:color w:val="000000" w:themeColor="text1"/><w:sz w:val="${d.sz}"/><w:szCs w:val="${d.sz}"/></w:rPr></w:style>`
    );
  }

  const TABLE_HEADING_STYLE =
    `<w:style w:type="paragraph" w:customStyle="1" w:styleId="TableHeading"><w:name w:val="Table Heading"/>` +
    `<w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>` +
    `<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:b/><w:bCs/>` +
    `<w:color w:val="000000" w:themeColor="text1"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>`;

  /** Remove TVSC brand colouring/fonts from run/paragraph properties. */
  function scrubFormatting(xml) {
    return (
      xml
        // white text (used on dark-blue header cells) → automatic
        .replace(/<w:color w:val="(?:FFFFFF|ffffff)"[^>]*\/>/g, '')
        // theme-coloured text (TVSC blues) → automatic; keep explicit reds etc.
        .replace(/<w:color w:val="[0-9A-Fa-f]{6}"[^>]*w:themeColor="(?:accent\d|text2|background2|hyperlink)"[^>]*\/>/g, '')
        .replace(/<w:color w:val="(?:002060|1F3864|1F4E79|2F5496|4472C4|203864|0F243E|17365D)"[^>]*\/>/gi, '')
        // TVSC heading font → body font
        .replace(/w:(ascii|hAnsi|eastAsia)Theme="majorHAnsi"/g, 'w:$1Theme="minorHAnsi"')
        .replace(/w:(ascii|hAnsi)Theme="majorEastAsia"/g, 'w:$1Theme="minorHAnsi"')
    );
  }

  /** Make a TVSC table look like an RMIT table. */
  function restyleTable(tblXml, warnings) {
    // Nested tables are handled recursively.
    const t = inner(tblXml, 'w:tbl');
    const tblPrM = /<w:tblPr>[\s\S]*?<\/w:tblPr>/.exec(t.content);
    let tblPr = tblPrM ? tblPrM[0] : '<w:tblPr></w:tblPr>';
    const look = /<w:tblLook [^>]*\/>/.exec(tblPr);
    const lookXml = look ? look[0] : '';
    const val = /w:val="([0-9A-Fa-f]{4})"/.exec(lookXml);
    const bits = val ? parseInt(val[1], 16) : 0x04a0;
    const firstRow = /w:firstRow="(1|true)"/.test(lookXml) || (!/w:firstRow=/.test(lookXml) && (bits & 0x20) !== 0);
    const firstCol = /w:firstColumn="(1|true)"/.test(lookXml) || (!/w:firstColumn=/.test(lookXml) && (bits & 0x80) !== 0);

    tblPr = tblPr.replace(/<w:tblStyle w:val="[^"]*"\/>/, '<w:tblStyle w:val="TableGrid"/>');
    if (!/<w:tblStyle /.test(tblPr)) tblPr = tblPr.replace('<w:tblPr>', '<w:tblPr><w:tblStyle w:val="TableGrid"/>');
    // explicit borders in TVSC were white-on-blue; drop them so TableGrid's black grid shows
    tblPr = tblPr.replace(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/, '');

    let rest = tblPrM ? t.content.replace(tblPrM[0], '') : t.content;
    const parts = splitTop(rest);
    let rowIndex = 0;
    const out = parts.map((p) => {
      if (p.tag !== 'w:tr') return p.xml;
      const isFirst = rowIndex++ === 0;
      return restyleRow(p.xml, isFirst && firstRow, firstCol && !(isFirst && firstRow), warnings);
    });
    return tblXml.slice(0, t.openEnd) + tblPr + out.join('') + tblXml.slice(t.closeStart);
  }

  function restyleRow(trXml, headerRow, labelCol, warnings) {
    const t = inner(trXml, 'w:tr');
    let cellIndex = 0;
    const parts = splitTop(t.content).map((p) => {
      if (p.tag !== 'w:tc') return p.xml.replace(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/, '');
      const idx = cellIndex++;
      return restyleCell(p.xml, headerRow, labelCol && idx === 0, warnings);
    });
    return trXml.slice(0, t.openEnd) + parts.join('') + trXml.slice(t.closeStart);
  }

  function restyleCell(tcXml, headerRow, labelCell, warnings) {
    const t = inner(tcXml, 'w:tc');
    let content = t.content;
    let tcPr = '';
    const m = /^\s*<w:tcPr>[\s\S]*?<\/w:tcPr>/.exec(content);
    if (m) {
      tcPr = m[0];
      content = content.slice(m[0].length);
    }
    const hasHeadingPara = /<w:pStyle w:val="(TableHeading|Heading\d)"/.test(content);
    const existing = /<w:shd [^>]*w:fill="([0-9A-Fa-f]{6}|auto)"[^>]*\/>/.exec(tcPr);
    const existingFill = existing ? existing[1].toUpperCase() : null;
    const coloured = existingFill && existingFill !== 'AUTO' && existingFill !== 'FFFFFF';

    let fill = null;
    if (headerRow || hasHeadingPara || coloured) fill = GREY_HEADER;
    else if (labelCell) fill = GREY_LABEL;
    else if (existingFill === 'FFFFFF') fill = null; // keep white explicitly as none

    tcPr = tcPr.replace(/<w:shd [^>]*\/>/g, '').replace(/<w:tcBorders>[\s\S]*?<\/w:tcBorders>/, '');
    if (fill) {
      if (!tcPr) tcPr = '<w:tcPr></w:tcPr>';
      // w:shd must come after tcW/gridSpan/vMerge/tcBorders and before noWrap/tcMar/vAlign...
      const after = /<w:(noWrap|tcMar|textDirection|tcFitText|vAlign|hideMark|headers|cellIns|cellDel|cellMerge|tcPrChange)\b/.exec(tcPr);
      if (after) tcPr = tcPr.slice(0, after.index) + SHD(fill) + tcPr.slice(after.index);
      else tcPr = tcPr.replace('</w:tcPr>', SHD(fill) + '</w:tcPr>');
    }

    // Nested tables inside cells.
    const parts = splitTop(content).map((p) => (p.tag === 'w:tbl' ? restyleTable(p.xml, warnings) : p.xml));
    return tcXml.slice(0, t.openEnd) + tcPr + parts.join('') + tcXml.slice(t.closeStart);
  }

  /** Copy numbering definitions from TVSC into the template's numbering. */
  function mergeNumbering(tplNum, tvNum) {
    const maxOf = (xml, rx) => {
      let max = 0;
      let m;
      while ((m = rx.exec(xml))) max = Math.max(max, +m[1]);
      return max;
    };
    const absOffset = tplNum ? maxOf(tplNum, /<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"/g) + 1 : 0;
    const numOffset = tplNum ? maxOf(tplNum, /<w:num\b[^>]*w:numId="(\d+)"/g) : 0;
    const tvInner = inner(tvNum, 'w:numbering').content;
    const parts = splitTop(tvInner);
    const abs = [];
    const nums = [];
    for (const p of parts) {
      if (p.tag === 'w:abstractNum') {
        let x = p.xml.replace(/(<w:abstractNum\b[^>]*w:abstractNumId=")(\d+)"/, (_, a, n) => `${a}${+n + absOffset}"`);
        // fresh nsid so Word never merges these lists with the template's
        x = x.replace(/<w:nsid w:val="[0-9A-Fa-f]{8}"\/>/, () => `<w:nsid w:val="${randomHex8()}"/>`);
        abs.push(x);
      } else if (p.tag === 'w:num') {
        nums.push(
          p.xml
            .replace(/(<w:num\b[^>]*w:numId=")(\d+)"/, (_, a, n) => `${a}${+n + numOffset}"`)
            .replace(/(<w:abstractNumId w:val=")(\d+)"/, (_, a, n) => `${a}${+n + absOffset}"`)
        );
      }
    }
    let merged = tplNum;
    if (!merged) {
      merged = tvNum.slice(0, inner(tvNum, 'w:numbering').openEnd) + '</w:numbering>';
    }
    merged = mergeRootNamespaces(merged, tvNum, 'w:numbering');
    const ni = inner(merged, 'w:numbering');
    const tplParts = splitTop(ni.content);
    const tplAbs = tplParts.filter((p) => p.tag !== 'w:num' && p.tag !== 'w:numIdMacAtCleanup').map((p) => p.xml);
    const tplNums = tplParts.filter((p) => p.tag === 'w:num').map((p) => p.xml);
    const tail = tplParts.filter((p) => p.tag === 'w:numIdMacAtCleanup').map((p) => p.xml);
    const body = [...tplAbs, ...abs, ...tplNums, ...nums, ...tail].join('');
    return { xml: merged.slice(0, ni.openEnd) + body + merged.slice(ni.closeStart), numOffset, absOffset };
  }

  function randomHex8() {
    let s = '';
    for (let i = 0; i < 8; i++) s += '0123456789ABCDEF'[Math.floor(Math.random() * 16)];
    return s;
  }

  /** Add xmlns declarations / mc:Ignorable prefixes from src root to dst root. */
  function mergeRootNamespaces(dstXml, srcXml, rootTag) {
    const rootRx = new RegExp(`<${rootTag}\\b[^>]*>`);
    const dRoot = rootRx.exec(dstXml);
    const sRoot = rootRx.exec(srcXml);
    if (!dRoot || !sRoot) return dstXml;
    let root = dRoot[0];
    const nsRx = /\sxmlns:([\w]+)="([^"]*)"/g;
    let m;
    const have = new Set();
    while ((m = nsRx.exec(root))) have.add(m[1]);
    const add = [];
    while ((m = nsRx.exec(sRoot[0]))) if (!have.has(m[1])) add.push(` xmlns:${m[1]}="${m[2]}"`);
    if (add.length) root = root.replace(new RegExp(`^<${rootTag}`), `<${rootTag}${add.join('')}`);
    const ig = (x) => (/mc:Ignorable="([^"]*)"/.exec(x) || [, ''])[1].split(/\s+/).filter(Boolean);
    const union = [...new Set([...ig(root), ...ig(sRoot[0])])];
    if (union.length) {
      if (/mc:Ignorable=/.test(root)) root = root.replace(/mc:Ignorable="[^"]*"/, `mc:Ignorable="${union.join(' ')}"`);
      else root = root.replace(/>$/, ` mc:Ignorable="${union.join(' ')}">`);
    }
    return dstXml.replace(dRoot[0], root);
  }

  /** Parse a .rels file into a map id → {type, target, mode}. */
  function parseRels(xml) {
    const map = {};
    const rx = /<Relationship\b([^>]*)\/?>/g;
    let m;
    while ((m = rx.exec(xml))) {
      const a = m[1];
      const g = (n) => (new RegExp(`\\b${n}="([^"]*)"`).exec(a) || [])[1];
      map[g('Id')] = { type: g('Type'), target: g('Target'), mode: g('TargetMode') };
    }
    return map;
  }

  /** Fill the RMIT cover table (label cell → following value cell). */
  function fillCoverTable(tblXml, info) {
    const value = {
      'task number': [info.taskNumber, info.taskTotal].every(Boolean)
        ? `${info.taskNumber} of ${info.taskTotal}`
        : info.taskNumber || '',
      'task name': info.taskName,
      'national unit/s code': info.unitCodes,
      'national unit/s title': info.unitTitles,
      'national qualification code': info.qualCode,
      'national qualification title': info.qualTitle,
      'rmit program code': info.programCode,
      'rmit course code': info.courseCode,
    };
    const t = inner(tblXml, 'w:tbl');
    const rows = splitTop(t.content).map((r) => {
      if (r.tag !== 'w:tr') return r.xml;
      const ri = inner(r.xml, 'w:tr');
      const cells = splitTop(ri.content);
      let pending = null;
      const outCells = cells.map((c) => {
        if (c.tag !== 'w:tc') return c.xml;
        const label = clean(textOf(c.xml)).toLowerCase();
        if (pending !== null) {
          const v = pending;
          pending = null;
          return setCellText(c.xml, v);
        }
        if (Object.prototype.hasOwnProperty.call(value, label)) pending = value[label] || '';
        return c.xml;
      });
      return r.xml.slice(0, ri.openEnd) + outCells.join('') + r.xml.slice(ri.closeStart);
    });
    return tblXml.slice(0, t.openEnd) + rows.join('') + tblXml.slice(t.closeStart);
  }

  function setCellText(tcXml, text) {
    const t = inner(tcXml, 'w:tc');
    const tcPr = (/^\s*<w:tcPr>[\s\S]*?<\/w:tcPr>/.exec(t.content) || [''])[0];
    const rPr = '<w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>';
    const lines = String(text || '').split(/\n/);
    const paras = lines
      .map(
        (l) =>
          `<w:p><w:pPr><w:spacing w:after="0"/>${rPr}</w:pPr>` +
          (l ? `<w:r>${rPr}<w:t xml:space="preserve">${encode(l)}</w:t></w:r>` : '') +
          `</w:p>`
      )
      .join('');
    return tcXml.slice(0, t.openEnd) + tcPr + paras + tcXml.slice(t.closeStart);
  }

  /** Replace the "UoC code – AT# Practical – Month Year" footer line. */
  function fillFooter(xml, footerLine) {
    if (!footerLine) return xml;
    const ps = [];
    const rx = /<w:p\b[\s\S]*?<\/w:p>/g;
    let m;
    while ((m = rx.exec(xml))) ps.push({ s: m.index, e: m.index + m[0].length, x: m[0] });
    for (let i = ps.length - 1; i >= 0; i--) {
      const p = ps[i];
      if (!/UoC\s*code|AT#|Month\s+Year/i.test(textOf(p.x))) continue;
      const pPr = (/<w:pPr>[\s\S]*?<\/w:pPr>/.exec(p.x) || [''])[0];
      const rPr = (/<w:r>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/.exec(p.x) || [, ''])[1];
      const np = `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${encode(footerLine)}</w:t></w:r></w:p>`;
      xml = xml.slice(0, p.s) + np + xml.slice(p.e);
    }
    return xml;
  }

  /** Optional find → replace on text (applied per text run). */
  function applyReplacements(xml, replacements) {
    if (!replacements || !replacements.length) return xml;
    return xml.replace(/(<w:t(?:\s[^>]*)?>)([^<]*)(<\/w:t>)/g, (_, a, txt, c) => {
      let s = decode(txt);
      for (const r of replacements) if (r.find) s = s.split(r.find).join(r.replace || '');
      return a + encode(s) + c;
    });
  }

  /**
   * Main entry point.
   * @param tvscZip      JSZip of the TVSC guide
   * @param templateZip  JSZip of the RMIT template
   * @param info         cover details (see extractInfo) + footerLine
   * @param opts         { replacements: [{find, replace}] }
   * @returns { zip, report }
   */
  async function convert(tvscZip, templateZip, info, opts = {}) {
    const warnings = [];
    const out = templateZip; // we modify the template package in place
    const tvDoc = await tvscZip.file('word/document.xml').async('string');
    const tplDoc = await out.file('word/document.xml').async('string');

    // ---- 1. TVSC content: everything after the cover page
    const tvBody = inner(tvDoc, 'w:body');
    let items = splitTop(tvBody.content).filter((i) => i.tag !== 'w:sectPr');
    let start = 0;
    const firstH1 = items.findIndex((i) => i.tag === 'w:p' && styleOf(i.xml) === 'Heading1');
    const firstBreak = items.findIndex((i) => i.tag === 'w:p' && /<w:sectPr\b/.test(i.xml));
    // The TVSC cover (image, qualification/unit lines, © notice, template
    // version) is everything before the first Heading 1. If the document has
    // no Heading 1, fall back to the first section break.
    if (firstH1 >= 0) start = firstH1;
    else if (firstBreak >= 0) start = firstBreak + 1;
    const dropped = items.slice(0, start);
    items = items.slice(start);
    const coverText = clean(textOf(dropped.map((d) => d.xml).join('\n')));

    // ---- 2. Relationships used by the content (hyperlinks, images)
    const tvRels = parseRels(await tvscZip.file('word/_rels/document.xml.rels').async('string'));
    let tplRelsXml = await out.file('word/_rels/document.xml.rels').async('string');
    let ctXml = await out.file('[Content_Types].xml').async('string');
    const relMap = {};
    let relCounter = 1000;
    // (section-break properties reference TVSC headers/footers, which are replaced by RMIT's)
    const contentXml = items.map((i) => i.xml.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, '')).join('');
    const usedIds = new Set();
    contentXml.replace(/\br:(?:id|embed|link|pict)="([^"]+)"/g, (_, id) => usedIds.add(id));
    for (const id of usedIds) {
      const rel = tvRels[id];
      if (!rel) continue;
      const newId = `rIdTvsc${relCounter++}`;
      const type = rel.type.split('/').pop();
      if (type === 'hyperlink') {
        tplRelsXml = tplRelsXml.replace(
          '</Relationships>',
          `<Relationship Id="${newId}" Type="${rel.type}" Target="${encode(rel.target).replace(/"/g, '&quot;')}" TargetMode="External"/></Relationships>`
        );
        relMap[id] = newId;
      } else if (type === 'image') {
        const srcPath = 'word/' + rel.target.replace(/^\.?\//, '');
        const f = tvscZip.file(srcPath);
        if (!f) {
          warnings.push(`Image ${rel.target} is missing from the TVSC file.`);
          continue;
        }
        const name = 'media/tvsc_' + rel.target.split('/').pop();
        out.file('word/' + name, await f.async('uint8array'));
        tplRelsXml = tplRelsXml.replace(
          '</Relationships>',
          `<Relationship Id="${newId}" Type="${rel.type}" Target="${name}"/></Relationships>`
        );
        const ext = name.split('.').pop().toLowerCase();
        if (!new RegExp(`Extension="${ext}"`, 'i').test(ctXml)) {
          const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', emf: 'image/x-emf', wmf: 'image/x-wmf', svg: 'image/svg+xml', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff' }[ext] || 'application/octet-stream';
          ctXml = ctXml.replace('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">', (s) => `${s}<Default Extension="${ext}" ContentType="${mime}"/>`);
        }
        relMap[id] = newId;
      } else {
        warnings.push(`Content of type "${type}" (${rel.target}) could not be carried over and was left out.`);
      }
    }

    // ---- 3. Numbering
    const tvNumF = tvscZip.file('word/numbering.xml');
    let numOffset = 0;
    if (tvNumF) {
      const tplNumF = out.file('word/numbering.xml');
      const merged = mergeNumbering(tplNumF ? await tplNumF.async('string') : null, await tvNumF.async('string'));
      numOffset = merged.numOffset;
      out.file('word/numbering.xml', merged.xml);
      if (!tplNumF) {
        ctXml = ctXml.replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>');
        tplRelsXml = tplRelsXml.replace('</Relationships>', '<Relationship Id="rIdTvscNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>');
      }
    }

    // ---- 4. Transform each content element
    let sectionBreaks = 0;
    const transformed = items.map((it) => {
      let x = it.xml;
      // section breaks inside paragraphs → page break (RMIT has one section)
      x = x.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, () => {
        sectionBreaks++;
        return '';
      });
      x = x.replace(/<w:pPr>\s*<\/w:pPr>/g, '');
      // tidy attributes that could clash with the template
      x = x.replace(/\sw:rsid\w*="[^"]*"/g, '').replace(/\sw14:(?:paraId|textId)="[^"]*"/g, '');
      x = x.replace(/<w:proofErr\b[^>]*\/>/g, '').replace(/<w:lastRenderedPageBreak\/>/g, '');
      // list numbering ids
      if (numOffset)
        x = x.replace(/<w:numId w:val="(\d+)"\/>/g, (_, n) => (+n === 0 ? `<w:numId w:val="0"/>` : `<w:numId w:val="${+n + numOffset}"/>`));
      // relationship ids
      x = x.replace(/\b(r:(?:id|embed|link|pict))="([^"]+)"/g, (s, attr, id) => (relMap[id] ? `${attr}="${relMap[id]}"` : s));
      x = scrubFormatting(x);
      if (it.tag === 'w:tbl') x = restyleTable(x, warnings);
      return x;
    });
    if (sectionBreaks) {
      // put a page break where each section break was
      let n = 0;
      for (let i = 0; i < transformed.length && n < sectionBreaks; i++) {
        if (/<w:sectPr\b/.test(items[i].xml)) {
          transformed[i] = transformed[i].replace(/<\/w:p>\s*$/, '<w:r><w:br w:type="page"/></w:r></w:p>');
          n++;
        }
      }
    }
    let content = applyReplacements(transformed.join(''), opts.replacements);

    // ---- 5. RMIT cover (title + info table + course note) from the template
    const tplBody = inner(tplDoc, 'w:body');
    const tplItems = splitTop(tplBody.content);
    const firstTbl = tplItems.findIndex((i) => i.tag === 'w:tbl');
    const cover = [];
    if (firstTbl >= 0) {
      for (let i = 0; i <= firstTbl; i++) cover.push(i === firstTbl ? fillCoverTable(tplItems[i].xml, info) : tplItems[i].xml);
      const note = tplItems[firstTbl + 1];
      if (note && note.tag === 'w:p' && /course is made up/i.test(textOf(note.xml))) cover.push(note.xml);
    } else {
      warnings.push('The RMIT template has no cover table; only the TVSC content was used.');
    }
    const finalSect = (tplItems.filter((i) => i.tag === 'w:sectPr').pop() || {}).xml || '';
    const newBody = cover.join('') + content + finalSect;
    let newDoc = tplDoc.slice(0, tplBody.openEnd) + newBody + tplDoc.slice(tplBody.closeStart);
    newDoc = mergeRootNamespaces(newDoc, tvDoc, 'w:document');
    out.file('word/document.xml', newDoc);

    // ---- 6. Styles: template styles + RMIT-look versions of TVSC styles
    let styles = await out.file('word/styles.xml').async('string');
    const tvStyles = await tvscZip.file('word/styles.xml').async('string');
    styles = mergeRootNamespaces(styles, tvStyles, 'w:styles');
    // Keep TVSC's paragraph spacing so the layout matches the original
    // (the RMIT template adds 10pt after every paragraph; TVSC uses none).
    const pPrDefault = (xml) => {
      const s = xml.search(/<w:pPrDefault\b/);
      if (s < 0) return null;
      return { s, e: findClose(xml, 'w:pPrDefault', s) };
    };
    const tvP = pPrDefault(tvStyles);
    const tpP = pPrDefault(styles);
    if (tvP && tpP) styles = styles.slice(0, tpP.s) + tvStyles.slice(tvP.s, tvP.e) + styles.slice(tpP.e);
    const tplStyleIds = new Set();
    styles.replace(/<w:style\b[^>]*w:styleId="([^"]+)"/g, (_, id) => tplStyleIds.add(id));
    const tvStyleMap = {};
    for (const s of children(inner(tvStyles, 'w:styles').content, 'w:style')) {
      const id = (/w:styleId="([^"]+)"/.exec(s.xml) || [])[1];
      if (id) tvStyleMap[id] = s.xml;
    }
    const needed = new Set();
    content.replace(/<w:(?:pStyle|rStyle|tblStyle)\s+w:val="([^"]+)"/g, (_, id) => needed.add(id));
    // follow basedOn / link chains
    const queue = [...needed];
    while (queue.length) {
      const id = queue.pop();
      const s = tvStyleMap[id];
      if (!s) continue;
      for (const m of s.matchAll(/<w:(?:basedOn|link)\s+w:val="([^"]+)"/g))
        if (!needed.has(m[1])) {
          needed.add(m[1]);
          queue.push(m[1]);
        }
    }
    const addStyles = [];
    for (const id of needed) {
      if (RMIT_STYLES[id]) {
        styles = removeStyle(styles, id);
        addStyles.push(headingStyle(id, RMIT_STYLES[id]));
      } else if (id === 'TableHeading') {
        styles = removeStyle(styles, id);
        addStyles.push(TABLE_HEADING_STYLE);
      } else if (!tplStyleIds.has(id) && tvStyleMap[id]) {
        let s = scrubFormatting(tvStyleMap[id]).replace(/\sw:rsid\w*="[^"]*"/g, '').replace(/<w:rsid\b[^>]*\/>/g, '');
        if (numOffset) s = s.replace(/<w:numId w:val="(\d+)"\/>/g, (_, n) => `<w:numId w:val="${+n === 0 ? 0 : +n + numOffset}"/>`);
        addStyles.push(s);
      }
    }
    styles = styles.replace('</w:styles>', addStyles.join('') + '</w:styles>');
    out.file('word/styles.xml', styles);

    // ---- 7. Footer line, relationships, content types, metadata
    for (const name of Object.keys(out.files)) {
      if (/^word\/footer\d*\.xml$/.test(name)) {
        out.file(name, fillFooter(await out.file(name).async('string'), info.footerLine));
      }
    }
    out.file('word/_rels/document.xml.rels', tplRelsXml);
    out.file('[Content_Types].xml', ctXml);
    const core = out.file('docProps/core.xml');
    if (core) {
      const title = `Assessor Guide – ${info.unitCodes || ''} AT${info.taskNumber || ''} ${info.taskName || ''}`.replace(/\s+/g, ' ').trim();
      let c = await core.async('string');
      c = /<dc:title>/.test(c) ? c.replace(/<dc:title>[\s\S]*?<\/dc:title>/, `<dc:title>${encode(title)}</dc:title>`) : c.replace('</cp:coreProperties>', `<dc:title>${encode(title)}</dc:title></cp:coreProperties>`);
      out.file('docProps/core.xml', c);
    }

    // ---- 8. Report: prove nothing was lost
    const srcText = clean(textOf(items.map((i) => i.xml).join('')));
    const outText = clean(textOf(content));
    const count = (x, rx) => (x.match(rx) || []).length;
    const report = {
      droppedCover: coverText,
      paragraphs: { source: count(items.map((i) => i.xml).join(''), /<w:p[\s>/]/g), output: count(content, /<w:p[\s>/]/g) },
      tables: { source: items.filter((i) => i.tag === 'w:tbl').length, output: count(content, /<w:tbl>/g) },
      words: { source: srcText.split(' ').filter(Boolean).length, output: outText.split(' ').filter(Boolean).length },
      textIdentical: !opts.replacements || !opts.replacements.length ? srcText === outText : null,
      sectionBreaks,
      warnings,
    };
    return { zip: out, report };
  }

  function removeStyle(stylesXml, id) {
    const rx = new RegExp(`<w:style\\b[^>]*w:styleId="${id}"[^>]*>`);
    const m = rx.exec(stylesXml);
    if (!m) return stylesXml;
    const end = findClose(stylesXml, 'w:style', m.index);
    return stylesXml.slice(0, m.index) + stylesXml.slice(end);
  }

  return { extractInfo, convert, defaultFileName, defaultFooter, _internal: { splitTop, textOf } };
});
