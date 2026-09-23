# CRO Invoice Automator

A browser tool that reads a CRO invoice or cost proposal (PDF) and turns it into
a filled Word form. The goal is for the scientist who receives the invoice to
produce one document that has what finance/accounting needs (vendor, invoice
number, dates, amount, line items) and what science/operations needs (study,
project number, timing, test materials).

Everything runs in the browser. Nothing is uploaded and nothing needs installing.

## What it does

1. **Reads any CRO invoice.** A general reader pulls the invoice number, invoice
   date, due date, PO/order/agreement number, total, CRO name and line items.
2. **Reads some CROs in more detail.** SGS, TKL, IIVS and Hilltop invoices (and
   IIVS cost proposals) have dedicated readers that also get the project/study
   number, study description, payment stage, start and completion dates, and
   listed test materials.
3. **Builds a general form.** With no template, it creates an "Invoice Summary"
   .docx with a Finance / Accounting section and a Science / Operations section.
4. **Or fills your team's own templates** from a profile file (see Profiles).
5. **Takes manual entries.** Test materials and notes are typed or pasted in.
6. **Leaves you in control.** Every field can be edited before download.

## How to use

1. Open the page: https://maurwhal.github.io/CRO-invoice-automator/
2. **Details (optional):**
   1. **Company Name:** shown on the form, and your prefix on CRO agreement numbers (ACME_12345 gives project 12345).
   2. **Output Document Type:** the form's title, also used in file names (default "Invoice Summary").
   3. **CRO Name:** filled in from the PDF. Change it if needed.
   4. **Type of Study:** shown on the form, and your project code prefix on invoices (ABC finds ABC-123).
3. Drop the invoice PDF.
4. Check the fields. Anything not found is highlighted.
5. Enter test materials and notes by hand.
6. Download and check the file before you use it.

## Profiles (team templates)

A profile is a small JSON file a team keeps privately (for example on a shared
drive) that holds its own Word templates for each CRO. Load it once with
**Load profile**; the browser remembers it. When the Company Name and Output
Document Type typed in Details match the profile, a check box appears to use
the team's templates instead of the general form. Nothing about any team is
built into this code; it all comes from the profile file.

Profile format:

```json
{
  "format": "cro-invoice-automator-profile",
  "version": 1,
  "name": "ACME Summary Sheet",
  "match": { "company": "ACME", "docType": "Summary Sheet" },
  "templates": {
    "sgs":     { "name": "SGS template.docx",     "b64": "<base64 of the .docx>" },
    "tkl":     { "name": "TKL template.docx",     "b64": "..." },
    "iivs":    { "name": "IIVS template.docx",    "b64": "..." },
    "hilltop": { "name": "Hilltop template.docx", "b64": "..." }
  }
}
```

A template is filled by finding paragraphs that start with these labels:

| CRO | Labels filled |
|---|---|
| SGS | Invoice Number:, Project Name:, Panel:, Date Signed:, Total Project Cost:, Date Started:, Date for Testing Completion/Draft Report: |
| TKL | Invoice Number:, Study Name:, Panel:, Date Signed:, Total Project Cost:, Date Started:, Date for Testing Completion/Draft Report: |
| IIVS | Invoice Number:, IIVS Project No.:, Study Type:, Date Signed:, Total Project Cost:, Date Started:, Date for Testing Completion/Draft Report: |
| Hilltop | Invoice:, Date Signed:, Total Project Cost:, Date Ordered/Shipped to CRO: |

1. Test materials go at the end of the paragraph that starts with "List all".
2. The "Submitted by" name goes on the "Submitted by:" line.
3. Notes are added at the end of the form.

## Adding a detailed CRO reader

All readers live in `lib/invoice-automator.js`. To add one:

1. Get a few real invoices from the CRO (they need selectable text, not a scan).
2. In `detectLab()`, add a rule that recognises the CRO from text that is always on its invoices (for example its company name).
3. Write a `parseX(lines, today, S)` function. `lines` is the invoice text, one line per printed line. Use `find(lines, /regex/)` to pull each value, anchored on the label printed next to it (for example `/Invoice Number\s*:?\s*(\d+)/`).
4. Return `fields` (label/value pairs for templates) and `summary` (values for the general form).
5. Add the CRO to `LAB_NAMES` and to the `if / else if` chain in `parseLines()`.
6. Test with each sample invoice and check the downloaded form against the PDF.

## Privacy

Everything runs in your browser. See `PRIVACY.md`.

## Attribution

Created by Maura Lavelle. AI tools supported code editing and error checking;
the author verified results.

## Disclaimer

Provided "as is" without warranties of any kind. Users are responsible for
checking each generated form against the source invoice before using it.

## License and reuse

All rights reserved. For reuse or distribution beyond fair use, please contact the author.
