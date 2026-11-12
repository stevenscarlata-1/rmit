# RMIT Assessment Builder

Browser-only tool, at `https://stevenscarlata-1.github.io/rmit/assessment-builder/`. Nothing is uploaded; all conversion happens in the page.

## Student appendices (ready)

Converts one or more Word documents, usually unbranded, into RMIT branded student appendices.

- **Combine into one file**, or **convert each document separately** (downloaded as a .zip).
- **Cover page** (optional): the details table, student instructions, list of appendices (one row per document) and student declaration.
- **"APPENDIX N: title"** heading on a new page for each document. You can change the heading word and the starting number. Each document's own title, whether a Title/Heading 1 or a bold first line, becomes the heading.
- **Apply RMIT style** (RMIT fonts, plain tables with grey header rows) or **keep the original formatting** and add only the header, footer and cover.
- RMIT header (STUDENT tab, logo, classification) and footer (`unit | task | Student appendices`, CRICOS, page X of Y) on every page.
- Lists, numbering, tables, images and hyperlinks are carried over. Footnotes and Word comments are removed, and the page lists a warning when that happens.

### Cover page template

`templates/appendix-cover.docx` was made from the VU23735 AT2 student appendices by `test/make-appendix-cover.js`. To use your own cover, edit a copy in Word, keep these fields where the details should go, and choose it on the page under **Cover page template**:

`{{qualCode}} {{qualTitle}} {{programCode}} {{courseCode}} {{assessmentTask}} {{unitCode}} {{unitTitle}} {{footerLine}}`

The table row containing `{{appendixLabel}}` / `{{appendixTitle}}` (or `{{appendixNumber}}`) is repeated once per appendix. The header and footer of the cover template are used for every page.

## Assessment converter (next)

Will convert and combine documents into the RMIT v3.0 ACCESSIBLE templates in `templates/assessment/` (assessor and student versions of Knowledge quiz, Practical and Product). Being built from example documents.

## Files

| Path | What it is |
|---|---|
| `index.html` | The web page |
| `lib/docx-kit.js` | Shared Word merging toolkit (numbering, images, links, styles, fields) |
| `lib/appendices.js` | Student appendices builder |
| `templates/appendix-cover.docx` | Default appendices cover and branding |
| `templates/assessment/*.docx` | Blank RMIT v3.0 assessment templates |
| `test/appendices-cli.js` | Command-line build: `node test/appendices-cli.js [--separate] [--no-cover] [--keep-format] a.docx b.docx` |
| `test/fixtures/` | Sample unbranded documents used for testing |
| `test/check-xml.py` | Checks every XML part of an output .docx |
