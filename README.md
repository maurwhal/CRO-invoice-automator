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
2. **Reads some CROs in more detail.** Certain organizations' invoices (and
   some cost proposals) have dedicated readers that also get the project/study
   number, study description, payment stage, start and completion dates, and
   listed test materials.
3. **Fills a standard form.** With no template, it fills the built-in Study
   Authorization Form (CRO Testing): contractor, invoice, cost, payment
   conditions, testing type, timelines, materials and sign-off.
4. **Or fills your team's own templates** from a profile file (see Profiles).
5. **Takes manual entries.** Test materials and notes are typed or pasted in.
6. **Leaves you in control.** Every field can be edited before download.

## How to use

1. Open the page: https://maurwhal.github.io/CRO-invoice-automator/
2. **Upload an invoice or cost proposal.** The PDF reader fills in the standard
   Study Authorization Form.
3. **Optional: use a template for a specific CRO.** Load your own Word template,
   or a team profile. Details (all optional):
   1. **Company Name:** your prefix on CRO agreement numbers (ACME_12345 gives project 12345).
   2. **Output Document Type:** used in file names and on the download button.
   3. **CRO Name:** filled in from the PDF. Change it if needed.
   4. **Type of Study:** your project code prefix on invoices (ABC finds ABC-123).
4. **Optional: add details the PDF reader missed.** Highlighted fields were not
   found on the PDF. Add test materials and notes.
5. Download and check the file before you use it.

## Your own template

**Use my own Word template** fills any .docx you give it. A paragraph that
starts with one of these labels gets the value written after it, in the
template's own font:

1. Contractor:, Contractor's Number:, Invoice Number:, Invoice Date:, Due Date:, Testing Type:, Project Number/Name:, Date Signed:, Total Project Cost:, Payment Conditions:, Board Approval Needed?, Date Testing Started:, Estimated Date for Testing Completion/Draft Report:, PO Number:, Company:
2. Materials being tested:, Other Notes:, Line Items: (each entry goes on its own line under the label)
3. Submitted by: (the name typed on the page, plus today's date after "Date:" on the same line)

A CRO-specific template can have that CRO's fixed details already typed in
after the labels (contractor name, phone, testing type, payment conditions).
Those lines are kept as they are, and a line that ends in "$" gets just the
amount. See the **EssGeeEss template** example on the page
(`templates/example-essgeeess.docx`), and use **Try it with your invoice** to
see it filled.

**Download blank form** gives the built-in Study Authorization Form, which
already uses these labels, to restyle in Word. Labels your template doesn't
use are skipped.

## Profiles (team templates)

A profile is a small JSON file a team keeps privately (for example on a shared
drive) that holds its Word templates, one per CRO plus an optional one for any
other CRO. Make one on the page with **Make a profile from your templates**,
then have teammates load it with **Load profile**; the browser remembers it.
When the Company Name and Output Document Type typed in Details match the
profile, a check box appears to use the team's templates. Nothing about any
team is built into this code; it all comes from the profile file.

Profile format:

```json
{
  "format": "cro-invoice-automator-profile",
  "version": 1,
  "name": "ACME Summary Sheet",
  "match": { "company": "ACME", "docType": "Summary Sheet" },
  "templates": {
    "cro:example labs": { "name": "Example Labs template.docx", "cro": "Example Labs", "b64": "<base64 of the .docx>" },
    "*":                { "name": "Any CRO template.docx", "b64": "..." }
  }
}
```

Templates are filled using the standard labels above. For CROs that have a
detailed reader, a profile template can also use that reader's own field
labels (see the `fields` returned by each reader in `lib/invoice-automator.js`).

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
