/*
 * Student appendices builder.
 *
 * Takes one or more (usually unbranded) Word documents and produces RMIT
 * branded student appendices using a cover template .docx:
 *   - the template's header ("STUDENT" tab, RMIT logo, classification) and
 *     footer (unit | task | Student appendices, CRICOS, page numbers);
 *   - optionally the template's cover page, with its {{fields}} filled and
 *     one "List of evidence" row per appendix;
 *   - each document under an "APPENDIX N: title" heading on a new page.
 *
 * Cover template fields: {{qualCode}} {{qualTitle}} {{programCode}}
 * {{courseCode}} {{assessmentTask}} {{unitCode}} {{unitTitle}} {{footerLine}}
 * and, in one table row, {{appendixLabel}} {{appendixTitle}} {{appendixNumber}}.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./docx-kit.js'));
  else root.Appendices = factory(root.DocxKit);
})(typeof self !== 'undefined' ? self : this, function (K) {
  'use strict';

  const DEFAULTS = {
    combine: true, // one file, or one file per document
    includeCover: true,
    headings: true, // add "APPENDIX N: title"
    headingWord: 'APPENDIX',
    startAt: 1,
    restyle: true, // RMIT fonts + plain grey-header tables
    pageBreaks: true, // each document starts on a new page
  };

  function headingPara(text, pageBreak) {
    return (
      `<w:p><w:pPr><w:keepNext/>${pageBreak ? '<w:pageBreakBefore/>' : ''}` +
      `<w:spacing w:before="200" w:after="140"/><w:outlineLvl w:val="0"/><w:rPr><w:rFonts w:cs="Calibri"/></w:rPr></w:pPr>` +
      `<w:r><w:rPr><w:rFonts w:cs="Calibri"/><w:b/><w:sz w:val="26"/></w:rPr><w:t xml:space="preserve">${K.encode(text)}</w:t></w:r></w:p>`
    );
  }
  const pageBreakPara = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

  /** Values for the cover and footer fields. */
  function fieldValues(info) {
    return {
      qualCode: info.qualCode || '',
      qualTitle: info.qualTitle || '',
      programCode: info.programCode || '',
      courseCode: info.courseCode || '',
      assessmentTask: info.assessmentTask || '',
      unitCode: info.unitCode || '',
      unitTitle: info.unitTitle || '',
      footerLine: info.footerLine || defaultFooter(info),
    };
  }

  function defaultFooter(info) {
    return [info.unitCode, info.assessmentTask, 'Student appendices'].filter(Boolean).join(' | ');
  }

  function ymd(d = new Date()) {
    return d.getFullYear() + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0');
  }

  function fileBase(info) {
    return [ymd(), info.unitCode, info.assessmentTask].filter(Boolean).join('_').replace(/[^\w.-]+/g, '');
  }

  function defaultFileName(info) {
    return `${fileBase(info)}_Student_Appendices.docx`;
  }

  /**
   * Read a source document's suggested title (for the appendix heading).
   * Returns { title, index, fromHeading }.
   */
  // "Appendix 3:", "APPENDIX 2 –", "Appendix 1." at the start of a title
  const PREFIX = /^\s*appendix\s*[0-9A-Z]{1,3}\s*[:\-–—.]?\s*/i;

  async function readTitle(srcZip, fileName) {
    const fallback = fileName.replace(/\.docx$/i, '').replace(/^[A-Z]{2,}\d{3,}[A-Z]?_AT\d+_/i, '')
      .replace(/_v\d+(\.\d+)*$/i, '').replace(/^appendix_?\d+_*/i, '').replace(/[_-]+/g, ' ').trim();
    const t = await K.guessTitle(srcZip, fallback);
    if (!PREFIX.test(t.title)) return t;
    // The document already numbers itself ("Appendix 1: Job cards"). Keep
    // only the title text; the builder adds its own "APPENDIX N:".
    const rest = t.title.replace(PREFIX, '').trim();
    if (rest) return { ...t, title: rest };
    // Title split over two lines ("Appendix 2:" then the title): use the
    // next non-empty paragraph and drop both.
    const doc = await srcZip.file('word/document.xml').async('string');
    const items = K.splitTop(K.inner(doc, 'w:body').content);
    for (let j = t.index + 1; j < items.length && j <= t.index + 4; j++) {
      if (items[j].tag !== 'w:p') break;
      const next = K.clean(K.textOf(items[j].xml));
      if (next) return { title: next.slice(0, 160), index: [...range(t.index, j)], fromHeading: t.fromHeading };
    }
    return { ...t, title: fallback };
  }
  function* range(a, b) { for (let i = a; i <= b; i++) yield i; }

  /**
   * Build the appendices.
   * @param loadZip   async (bytes) => JSZip  (lets this run in browser or Node)
   * @param coverBytes  the cover template .docx (ArrayBuffer / Buffer)
   * @param docs  [{ bytes, name, title, dropTitleIndex }]  in order
   * @param info  { qualCode, qualTitle, programCode, courseCode, assessmentTask, unitCode, unitTitle, footerLine }
   * @param options  see DEFAULTS
   * @returns { files: [{ name, zip }], report }
   */
  async function build(loadZip, coverBytes, docs, info, options = {}) {
    const o = { ...DEFAULTS, ...options };
    const values = fieldValues(info);
    const numbered = docs.map((d, i) => ({
      ...d,
      number: o.startAt + i,
      label: `Appendix ${o.startAt + i}`,
      heading: `${o.headingWord} ${o.startAt + i}: ${d.title}`,
    }));
    const groups = o.combine ? [numbered] : numbered.map((d) => [d]);
    const files = [];
    const report = { documents: [], warnings: [] };

    for (const group of groups) {
      const out = await K.Output.open(await loadZip(coverBytes));
      // Appendices never use the assessment guides' "Section A/B/C" heading
      // numbering, even if a custom cover template's heading styles have it.
      out.styles = out.styles.replace(/<w:style\b[^>]*w:styleId="(Heading\d|Title)"[^>]*>[\s\S]*?<\/w:style>/g, (st) => st.replace(/<w:numPr>[\s\S]*?<\/w:numPr>/g, ''));
      for (const id of [...out.numberedStyles]) if (/^(Heading\d|Title)$/.test(id)) out.numberedStyles.delete(id);
      const parts = [];
      if (o.includeCover) {
        let cover = out.shellBody();
        cover = K.repeatRows(cover, 'appendixLabel|appendixTitle|appendixNumber', group.map((d) => ({ appendixLabel: d.label, appendixTitle: d.title, appendixNumber: d.number })));
        cover = K.fillTokens(cover, values);
        parts.push(cover);
      }
      for (const d of group) {
        const first = parts.length === 0;
        const src = await loadZip(d.bytes);
        const xml = await out.importBody(src, { restyle: o.restyle, label: d.name, dropIndex: o.headings ? d.dropTitleIndex : -1 });
        if (o.headings) parts.push(headingPara(d.heading, o.pageBreaks && !first));
        else if (o.pageBreaks && !first) parts.push(pageBreakPara);
        parts.push(xml);
        const srcText = K.clean(K.textOf(xml));
        report.documents.push({ name: d.name, heading: o.headings ? d.heading : d.title, words: srcText ? srcText.split(' ').length : 0, tables: (xml.match(/<w:tbl>/g) || []).length });
      }
      await out.fillHeadersFooters(values);
      await out.setTitle(o.combine ? `Student assessment appendices – ${values.unitCode} ${values.assessmentTask}`.trim() : `${group[0].label} – ${group[0].title}`);
      const zip = out.finish(parts.join(''));
      const name = o.combine ? info.fileName || defaultFileName(info) : `${fileBase(info)}_${group[0].label.replace(/\s+/g, '')}.docx`;
      files.push({ name, zip });
      report.warnings.push(...out.warnings);
    }
    return { files, report };
  }

  return { build, readTitle, defaultFooter, defaultFileName, DEFAULTS };
});
