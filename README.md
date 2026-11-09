# Assessor Guide Converter: TVSC to RMIT

A single web page that converts a **TVSC assessor guide (.docx)** into the **RMIT assessor guide (practical task)** template.

- **Keeps every TVSC section** in its original order and wording: assessment details, units of competency and pre-requisite tables, safety, reasonable adjustment, special consideration, preparation and briefing checklists, and each observation task (prior to / during / post assessment, scenarios, step tables, oral questions) plus the benchmark for submitted evidence. Nested lists, merged cells, bold, italic and red text, hyperlinks and images are kept too.
- **Applies RMIT branding**: the RMIT letterhead and "ASSESSOR" tab, the cover information table, the RMIT footer (copyright, CRICOS, RTO code, page numbers), Calibri text, black headings, and grid tables with grey header rows.
- **Drops the TVSC cover page** (cover image, © notice, template version), because the RMIT cover table replaces it.
- **Checks its own work**: after each conversion it reports paragraph, table and word counts for the input and output, and confirms the text is identical.
- **Runs entirely in the browser.** Documents are never uploaded anywhere.

## Using it

1. Open the hosted page (GitHub Pages).
2. Drop in the TVSC `.docx`.
3. Check the cover details it read from the guide, and fill in the RMIT program and course codes.
4. Click **Convert and download .docx**.

Optional extras on the page:

- **Find and replace**: one rule per line, for example `Resources supplied TAFE => Resources supplied by RMIT`.
- **Different template**: use another RMIT template `.docx` instead of the bundled one.

## Hosting on GitHub Pages

1. Push this repository to GitHub.
2. Go to **Settings → Pages → Build and deployment**, choose **Deploy from a branch**, then pick `main` and `/ (root)`.
3. The tool will be live at `https://<user>.github.io/<repo>/`.

> The RMIT template in `template/` is published with the page. On a free GitHub plan, Pages needs a public repository, so the template becomes publicly downloadable.

## Files

| Path | What it is |
|---|---|
| `index.html` | The web page (upload, cover details form, convert and download) |
| `converter.js` | The conversion logic. Works in the browser and in Node. |
| `template/rmit-assessor-guide-template.docx` | RMIT assessor guide (practical task) template v3.0 |
| `vendor/jszip.min.js` | JSZip 3.10.1, for reading and writing .docx files |
| `test/convert-cli.js` | Runs the same conversion from the command line |
| `test/check-xml.py` | Checks that every XML part of an output .docx is well-formed |

## Command-line testing

```bash
npm install
node test/convert-cli.js path/to/TVSC_guide.docx
python3 test/check-xml.py test/output/*.docx
```

## How the conversion works

1. Reads the TVSC body and drops everything before the first Heading 1 (the cover page).
2. Copies the remaining WordprocessingML across **unchanged in content**. Only formatting is touched:
   - Headings and "Table Heading" paragraphs are restyled to the RMIT look.
   - TVSC blue/white brand colours are removed. Red benchmark text is kept.
   - Tables switch to `Table Grid`, with grey (`D9D9D9`) header rows and light grey (`F2F2F2`) label columns.
   - List numbering definitions are merged into the template, so bullets and nesting survive.
   - Hyperlinks and images are copied with new relationship IDs.
   - Section breaks become page breaks, so the RMIT header and footer apply throughout.
3. Takes the RMIT title, cover table and course note from the template, and fills the table from the form.
4. Writes the footer line ("UoC code – AT# Practical – Month Year").
