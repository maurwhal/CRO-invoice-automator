/*
 * CRO Invoice Automator: CRO invoice / cost-proposal PDF -> filled Word form.
 * Built-in readers: SGS, TKL, IIVS, Hilltop. To add a CRO, write a parseX()
 * for its standard invoice layout, add a detect rule and a LAB_NAMES entry.
 * Everything runs locally in the browser. Needs JSZip (global) for .docx work;
 * pdf.js is loaded on first PDF.
 *
 * This has been a Maura Lavelle production
 */
(function (global) {
  'use strict';

  var PDFJS_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.6.82/pdf.min.mjs';
  var PDFJS_WORKER = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.6.82/pdf.worker.min.mjs';
  var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var XML_NS = 'http://www.w3.org/XML/1998/namespace';

  var MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
    'August', 'September', 'October', 'November', 'December'];
  var MON3 = { JAN: 0, FEB: 1, MAR: 2, APR: 3, MAY: 4, JUN: 5, JUL: 6, AUG: 7, SEP: 8, OCT: 9, NOV: 10, DEC: 11 };

  // ─── dates ──────────────────────────────────────────────────────────────────
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function fmtLong(d) { return MONTHS[d.getMonth()] + ' ' + pad2(d.getDate()) + ', ' + d.getFullYear(); } // September 04, 2026
  function fmtMonthYear(d) { return MONTHS[d.getMonth()] + ' ' + d.getFullYear(); }
  function addDays(d, n) { var x = new Date(d.getTime()); x.setDate(x.getDate() + n); return x; }
  function addMonths(d, n) { return new Date(d.getFullYear(), d.getMonth() + n, 1); }
  function monthIndex(name) {
    var k = String(name).slice(0, 3).toUpperCase();
    return Object.prototype.hasOwnProperty.call(MON3, k) ? MON3[k] : -1;
  }

  // Accepts 22-OCT-26, 6/3/2025, August 17, 2026, Aug 08, 2025, 1 September 2026
  function parseDate(s) {
    if (!s) return null;
    s = String(s).trim();
    var m, mo, y;
    if ((m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2}|\d{4})$/))) {
      mo = monthIndex(m[2]); y = +m[3]; if (y < 100) y += 2000;
      return mo < 0 ? null : new Date(y, mo, +m[1]);
    }
    if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))) return new Date(+m[3], +m[1] - 1, +m[2]);
    if ((m = s.match(/^([A-Za-z]+)\.?\s+(\d{1,2}),?\s+(\d{4})$/))) {
      mo = monthIndex(m[1]); return mo < 0 ? null : new Date(+m[3], mo, +m[2]);
    }
    if ((m = s.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/))) {
      mo = monthIndex(m[2]); return mo < 0 ? null : new Date(+m[3], mo, +m[1]);
    }
    return null;
  }

  // ─── PDF -> lines ───────────────────────────────────────────────────────────
  var pdfjsPromise = null;
  function loadPdfJs() {
    if (!pdfjsPromise) {
      pdfjsPromise = import(PDFJS_URL).then(function (lib) {
        lib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER;
        return lib;
      });
    }
    return pdfjsPromise;
  }

  // Rebuild visual lines (like pdfplumber's extract_text) by grouping text
  // items on the same baseline and ordering them left to right.
  async function pdfToLines(arrayBuffer) {
    var pdfjs = await loadPdfJs();
    var pdf = await pdfjs.getDocument({ data: new Uint8Array(arrayBuffer) }).promise;
    var lines = [];
    for (var p = 1; p <= pdf.numPages; p++) {
      var page = await pdf.getPage(p);
      var content = await page.getTextContent();
      var rows = [];
      content.items.forEach(function (it) {
        if (!it.str || !it.str.trim()) return;
        var x = it.transform[4], y = it.transform[5];
        var fs = Math.hypot(it.transform[0], it.transform[1]) || 10;
        var row = null;
        for (var i = 0; i < rows.length; i++) {
          if (Math.abs(rows[i].y - y) <= Math.max(2, fs * 0.3)) { row = rows[i]; break; }
        }
        if (!row) { row = { y: y, items: [] }; rows.push(row); }
        row.items.push({ x: x, w: it.width, fs: fs, s: it.str });
      });
      rows.sort(function (a, b) { return b.y - a.y; });
      rows.forEach(function (r) {
        r.items.sort(function (a, b) { return a.x - b.x; });
        var out = '', lastEnd = null;
        r.items.forEach(function (i) {
          if (lastEnd !== null && i.x - lastEnd > i.fs * 0.15 && !/\s$/.test(out) && !/^\s/.test(i.s)) out += ' ';
          out += i.s;
          lastEnd = i.x + i.w;
        });
        out = out.replace(/\s+/g, ' ').trim();
        if (out) lines.push(out);
      });
    }
    return lines;
  }

  // ─── parsing helpers ────────────────────────────────────────────────────────
  function find(lines, re, grp) {
    for (var i = 0; i < lines.length; i++) {
      var m = lines[i].match(re);
      if (m) return (m[grp || 1] || '').trim();
    }
    return '';
  }
  function money(v) { return v ? '$' + v : ''; }
  function normCas(s) { return s.replace(/CAS\s*#\s*/gi, 'CAS # '); }
  function assayAbbr(study) {
    var re = /\(([^)]+)\)/g, m;
    while ((m = re.exec(study))) { if (!/^OECD/i.test(m[1].trim())) return m[1].trim(); }
    return '';
  }
  function escRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
  function safeName(s) { return s.replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim(); }

  // ─── lab detection ──────────────────────────────────────────────────────────
  // Hilltop first: its invoices name SGS or TKL as the ship-to.
  function detectLab(text) {
    if (/Hill\s*Top Research/i.test(text)) return 'hilltop';
    if (/TKL STUDY\s*#/i.test(text) || /TKL RESEARCH/i.test(text)) return 'tkl';
    if (/SGS North America/i.test(text) || /Clinical Trial Fee/i.test(text)) return 'sgs';
    if (/Institute for In Vitro Sciences|\bIIVS\b/i.test(text)) return 'iivs';
    return null;
  }

  var LAB_NAMES = { sgs: 'SGS', tkl: 'TKL', iivs: 'IIVS', hilltop: 'Hilltop' };

  // Each parser returns:
  //   fields: [{label, value}]  label = paragraph start in the template
  //   chems: lines to prefill in the chemical box (empty = leave box empty)
  //   chemHint: text found on the invoice, offered but not prefilled
  //   extraParas: paragraphs appended at the end of the form
  //   filename, kind, missing: [labels with no value]
  //
  // S = the user's details (nothing organization-specific is coded here):
  //   company  Company Name          strips "Company_" from agreement numbers
  //   docType  Output Document Type  goes in file names and the cost-proposal note
  //   cro      CRO Name              label used in file names (defaults to the reader's name)
  //   study    Type of Study         project-code prefix, e.g. "ABC" finds ABC-123

  function parseSgs(lines, today, S) {
    var inv = find(lines, /Invoice Number\s*:?\s*(\d+)/i);
    var dueRaw = find(lines, /Due Date\s*:?\s*(\d{1,2}-[A-Za-z]{3}-\d{2,4})/i);
    var dateRaw = find(lines, /^Date\s*:?\s*(\d{1,2}-[A-Za-z]{3}-\d{2,4})/i);
    var panel = find(lines, /Panel\s*:\s*([A-Za-z0-9\-]+)/i);
    var total = find(lines, /Total Amount\s*USD\s*([\d,]+\.\d{2})/i) ||
      find(lines, /Total Due\s*(?:USD)?\s*\$?([\d,]+\.\d{2})/i);

    // Dates run off the due date.
    var base = parseDate(dueRaw);
    var dueText = dueRaw;
    if (!base && parseDate(dateRaw)) { base = addDays(parseDate(dateRaw), 30); dueText = fmtLong(base); }

    // Project code: "<Type of Study>-123" anywhere on the invoice when a study
    // type is set; otherwise the first test-sample code without its trailing
    // -1/-2/-3 (e.g. ABC-123-1 -> ABC-123).
    var project = S.study ? find(lines, new RegExp('\\b(' + escRe(S.study) + '[-_]\\d+)', 'i')) : '';
    var hint = [];
    var start = lines.findIndex(function (L) { return /^Test Samples\s*:/i.test(L); });
    if (start >= 0) {
      for (var i = start + 1; i < lines.length; i++) {
        var L = lines[i];
        if (/^Actual Execution|^No Sales Tax|^\d{4,}\s+\S/i.test(L)) break;
        hint.push(L);
        var pm = !project && L.match(/^([A-Za-z]+[-_]\d+)-\d+\b/);
        if (pm) project = pm[1];
      }
    }

    // Study description: the lines between "Panel:" and "Test Samples:"
    var desc = [];
    var pIdx = lines.findIndex(function (L) { return /^Panel\s*:/i.test(L); });
    var stopAt = start > pIdx ? start : pIdx + 4;
    for (var j = pIdx + 1; pIdx >= 0 && j < Math.min(stopAt, lines.length); j++) desc.push(lines[j]);

    return {
      summary: {
        invoiceNo: inv, invoiceDate: dateRaw, dueDate: dueRaw,
        orderNo: find(lines, /Order No\.?\s*:?\s*(\d+)/i), total: money(total),
        project: project, studyDesc: (panel ? 'Panel ' + panel + (desc.length ? ': ' : '') : '') + desc.join('; '),
        started: base ? String(base.getFullYear()) : '', completion: fmtMonthYear(addDays(base || today, 60))
      },
      kind: 'SGS invoice',
      fields: [
        { label: 'Invoice Number:', value: inv ? inv + (dueText ? ' (Due ' + dueText + ')' : '') : '' },
        { label: 'Project Name:', value: project },
        { label: 'Panel:', value: panel },
        { label: 'Date Signed:', value: fmtLong(today) },
        { label: 'Total Project Cost:', value: money(total) },
        { label: 'Date Started:', value: base ? String(base.getFullYear()) : String(today.getFullYear()) },
        { label: 'Date for Testing Completion/Draft Report:', value: fmtMonthYear(addDays(base || today, 60)) }
      ],
      chems: [],
      chemHint: hint.join('\n'),
      extraParas: [],
      filename: safeName((project || S.study || 'Project') + ' ' + inv + ' (' + (S.cro || 'SGS') + ')') + '.docx'
    };
  }

  function parseTkl(lines, today, S) {
    var inv = find(lines, /Invoice Number\s*:?\s*(\d+)/i);
    var invDate = parseDate(find(lines, /Invoice Date\s*:?\s*([A-Za-z]+\.?\s+\d{1,2},?\s+\d{4})/i));
    var study = find(lines, /TKL STUDY\s*#\s*(\S+)/i);
    var panel = find(lines, /PRODUCT\s*#\s*(\S+)/i);
    var total = find(lines, /Invoice Total\s*\$?\s*([\d,]+\.\d{2})/i);
    var text = lines.join('\n');
    var phase = /UPON INITIATION/i.test(text) ? 'Initiation' : (/UPON COMPLETION/i.test(text) ? 'Completion' : '');
    var base = invDate || today;

    var descr = find(lines, /DESCRIPTION\s*:\s*(.+)/i);
    return {
      summary: {
        invoiceNo: inv, invoiceDate: invDate ? fmtLong(invDate) : '',
        dueDate: invDate ? fmtLong(addDays(invDate, 30)) : '', total: money(total),
        paymentStage: phase ? phase + ' payment' : '',
        project: panel, studyDesc: [descr, study ? 'TKL study ' + study : ''].filter(Boolean).join('; '),
        started: String(base.getFullYear()), completion: fmtMonthYear(addDays(base, 60))
      },
      kind: 'TKL invoice' + (phase ? ' (' + phase.toLowerCase() + ' payment)' : ''),
      fields: [
        { label: 'Invoice Number:', value: inv ? inv + (invDate ? ' (Due ' + fmtLong(addDays(invDate, 30)) + ')' : '') : '' },
        { label: 'Study Name:', value: study },
        { label: 'Panel:', value: panel },
        { label: 'Date Signed:', value: fmtLong(today) },
        { label: 'Total Project Cost:', value: money(total) },
        { label: 'Date Started:', value: String(base.getFullYear()) },
        { label: 'Date for Testing Completion/Draft Report:', value: fmtMonthYear(addDays(base, 60)) }
      ],
      chems: [],
      chemHint: '',
      extraParas: [],
      filename: safeName((panel || S.study || 'Project') + ' ' + phase + ' ' + inv + ' (' + (S.cro || 'TKL') + ')') + '.docx'
    };
  }

  function parseHilltop(lines, today, S) {
    var inv = find(lines, /Invoice\s*#\s*:?\s*(\d+)/i);
    var ordered = find(lines, /^Date\s*:?\s*([A-Za-z]{3}\s*\d{1,2},\s*\d{4})/i);
    var total = find(lines, /^Total\s+\$?\s*([\d,]+\.\d{2})/i);
    if (!total) {
      var d = lines.join('\n').match(/\$([\d,]+\.\d{2})/g);
      if (d) total = d[d.length - 1].slice(1);
    }
    return {
      summary: { invoiceNo: inv, invoiceDate: ordered, total: money(total) },
      kind: 'Hilltop chamber invoice',
      fields: [
        { label: 'Invoice:', value: inv },
        { label: 'Date Signed:', value: fmtLong(today) },
        { label: 'Total Project Cost:', value: money(total) },
        { label: 'Date Ordered/Shipped to CRO:', value: ordered }
      ],
      chems: [],
      chemHint: '',
      extraParas: [],
      filename: safeName((S.cro || 'Hilltop') + ' Chambers ' + S.docType + ' ' + inv) + '.docx'
    };
  }

  var FEE_LINE = /^\d+\s+\d{3,6}\s+.+\s[\d,]+\.\d{2}\s+[\d,]+\.\d{2}$/;

  function parseIivsInvoice(lines, today, S) {
    var inv = find(lines, /Invoice Number\s*:\s*(\d+)/i);
    var invDate = parseDate(find(lines, /Invoice Date\s*:\s*(\d{1,2}\/\d{1,2}\/\d{4})/i));
    var total = find(lines, /Total Due\s*USD\s*\$?\s*([\d,]+\.\d{2})/i);

    // Row under "IIVS PROJECT NO. | IIVS STUDY NUMBER | IIVS AGREEMENT NO."
    var project = '', agreeNo = '';
    var hdr = lines.findIndex(function (L) { return /IIVS PROJECT NO\./i.test(L); });
    if (hdr >= 0 && lines[hdr + 1]) {
      var pm = lines[hdr + 1].match(/^(\d{4,6})\b.*\s(\S+)$/);
      if (pm) { project = pm[1]; agreeNo = pm[2]; }
    }

    // Assay rows carry a 6-digit protocol code (176000 kDPRA, 177000 h-CLAT);
    // 3-digit codes are fees (902 GLP, 904/907 draft report).
    var assays = [];
    var lastFee = -1;
    lines.forEach(function (L, i) {
      if (FEE_LINE.test(L)) lastFee = i;
      var m = L.match(/^\d+\s+\d{6}\s+(.+?)\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}$/);
      if (m) assays.push(m[1].trim());
    });
    var study = assays.join('; ');

    // Test materials: lines after the last fee row, up to the banking block.
    var chems = [];
    if (lastFee >= 0) {
      for (var i = lastFee + 1; i < lines.length; i++) {
        if (/^(All banking fees|PLEASE NOTE|DOMESTIC ACH)/i.test(lines[i])) break;
        chems.push(normCas(lines[i]));
      }
    }

    var base = invDate || today;
    var abbr = assayAbbr(study);
    return {
      summary: {
        invoiceNo: inv, invoiceDate: invDate ? fmtLong(invDate) : '',
        dueDate: invDate ? fmtLong(addDays(invDate, 30)) : '', orderNo: agreeNo, total: money(total),
        project: project, studyDesc: study,
        started: String(base.getFullYear()), completion: fmtMonthYear(base)
      },
      kind: 'IIVS invoice',
      fields: [
        { label: 'Invoice Number:', value: inv ? inv + (invDate ? ' (Due ' + fmtLong(addDays(invDate, 30)) + ')' : '') : '' },
        { label: 'IIVS Project No.:', value: project },
        { label: 'Study Type:', value: study },
        { label: 'Date Signed:', value: fmtLong(today) },
        { label: 'Total Project Cost:', value: money(total) },
        { label: 'Date Started:', value: String(base.getFullYear()) },
        { label: 'Date for Testing Completion/Draft Report:', value: fmtMonthYear(base) }
      ],
      chems: chems,
      chemHint: '',
      extraParas: [],
      filename: safeName((S.cro || 'IIVS') + ' ' + project + ' ' + (abbr || '') + ' ' + S.docType + ' Invoice ' + inv) + '.docx'
    };
  }

  function parseIivsProposal(lines, today, S) {
    var agree = find(lines, /Agreement\s*No\.\s*:\s*([\w\-\.]+)/i);
    var prefix = S.company ? new RegExp('^' + escRe(S.company) + '_', 'i') : /^[A-Za-z]+_/;
    var projectNo = agree ? agree.replace(prefix, '').split('-')[0] : '';
    var agreeDate = find(lines, /^(\d{1,2}\s+[A-Za-z]+\s+\d{4})$/) || find(lines, /(\d{1,2}\s+[A-Za-z]+\s+\d{4})/);

    var assays = [];
    lines.forEach(function (L) {
      var m = L.match(/^\d+\s+\d{6}\s+(.+?)\s+\$[\d,]+(?:\.\d{2})?$/);
      if (m) assays.push(m[1].trim());
    });
    var study = assays.join('; ');
    var total = find(lines, /TOTAL:\s*\$?\s*([\d,]+(?:\.\d{2})?)/);

    var chems = [];
    var start = lines.findIndex(function (L) { return /Test Materials\s*:/i.test(L); });
    if (start >= 0) {
      for (var i = start; i < lines.length; i++) {
        var L = i === start ? lines[i].replace(/^.*?Test Materials\s*:\s*/i, '') : lines[i];
        if (i > start && (/will conduct/i.test(L) || /prices quoted/i.test(L))) break;
        if (/CAS/i.test(L)) chems.push(normCas(L.trim()));
      }
    }

    var abbr = assayAbbr(study);
    var cro = S.cro || 'IIVS';
    var stamp = pad2(today.getMonth() + 1) + ' ' + today.getFullYear();
    return {
      summary: {
        invoiceNo: agree, invoiceDate: agreeDate, orderNo: agree, total: money(total),
        paymentStage: 'Cost proposal (not yet invoiced)',
        project: projectNo, studyDesc: study,
        started: fmtMonthYear(today), completion: fmtMonthYear(addMonths(today, 3))
      },
      kind: 'IIVS cost proposal',
      fields: [
        { label: 'Invoice Number:', value: agree ? agree + (agreeDate ? ' (Cost Proposal ' + agreeDate + ')' : '') : '' },
        { label: 'IIVS Project No.:', value: projectNo },
        { label: 'Study Type:', value: study },
        { label: 'Date Signed:', value: fmtLong(today) },
        { label: 'Total Project Cost:', value: money(total) },
        { label: 'Date Started:', value: fmtMonthYear(today) },
        { label: 'Date for Testing Completion/Draft Report:', value: fmtMonthYear(addMonths(today, 3)) }
      ],
      chems: chems,
      chemHint: '',
      extraParas: agree ? [(S.docType || 'Form') + ' based on cost proposal ' + agree + (agreeDate ? ' dated ' + agreeDate : '') + '. Invoice will be updated when received.'] : [],
      filename: safeName(S.docType
        ? 'Initial ' + S.docType + ' for ' + cro + ' Cost Proposal for ' + projectNo + ' ' + abbr + ' ' + stamp
        : cro + ' Cost Proposal ' + projectNo + ' ' + abbr + ' ' + stamp) + '.docx'
    };
  }

  // ─── general reader (any CRO) ───────────────────────────────────────────────
  var DATE_RE = '(\\d{1,2}-[A-Za-z]{3}-\\d{2,4}|\\d{1,2}/\\d{1,2}/\\d{2,4}|[A-Za-z]{3,9}\\.?\\s+\\d{1,2},?\\s+\\d{4}|\\d{1,2}\\s+[A-Za-z]{3,9}\\s+\\d{4})';
  var AMOUNT_RE = '\\$?\\s*([\\d,]+(?:\\.\\d{2})?)';

  function parseGeneric(lines, today) {
    var text = lines.join('\n');
    function findDate(label) { return find(lines, new RegExp(label + '\\s*[:#]?\\s*' + DATE_RE, 'i')); }

    var invoiceNo = find(lines, /Invoice\s*(?:Number|No\.?|#)\s*[:#]?\s*([A-Za-z0-9][A-Za-z0-9\-]{2,})/i);
    var invoiceDate = findDate('Invoice Date') || findDate('^Date');
    if (!invoiceDate) invoiceDate = find(lines.slice(0, 15), new RegExp(DATE_RE));
    var dueDate = findDate('Due Date');
    var net = text.match(/Net\s*(?:Due\s*in\s*)?(\d{1,3})\s*Days?|Net\s*(\d{1,3})\b/i);
    if (!dueDate && net && parseDate(invoiceDate)) dueDate = fmtLong(addDays(parseDate(invoiceDate), +(net[1] || net[2])));

    // PO / order / quote / agreement number: must contain a digit, and not N/A
    var orderNo = '';
    lines.some(function (L) {
      var m = L.match(/\b(?:P\.?O\.?|Purchase Order|Order|Quote|Quotation|Agreement)\s*(?:No\.?|Number|#)?\s*[:#]?\s*([A-Za-z0-9][A-Za-z0-9_.\-\/]*\d[A-Za-z0-9_.\-\/]*)/i);
      if (m && !/^n\/?a$/i.test(m[1])) { orderNo = m[1]; return true; }
      return false;
    });

    // Total: most specific label first, then a bare "Total", then the last $ amount
    var total = '';
    [/(?:Total Amount|Total Due|Invoice Total|Amount Due|Grand Total)\s*(?:USD)?\s*:?\s*/i, /^Total\b\s*:?\s*(?:USD)?\s*/i, /\bTOTAL\s*:\s*/]
      .some(function (lab) { total = find(lines, new RegExp(lab.source + AMOUNT_RE, lab.flags)); return !!total; });
    if (!total) { var d = text.match(/\$\s*([\d,]+\.\d{2})/g); if (d) total = d[d.length - 1].replace(/[$\s]/g, ''); }

    // CRO name: who gets paid. Take the first company-looking line (not an address)
    // from "payable to" / "Beneficiary Name", the lines under "Remit to", or the top of page 1.
    var companyLike = function (L) {
      return /\b(Inc|LLC|Ltd|Limited|GmbH|Corp|Corporation|Company|Laborator(?:y|ies)|Research|Institute|Sciences?)\b/i.test(L) &&
        !/^\d/.test(L) && !/\b[A-Z]{2}\s+\d{5}\b/.test(L);
    };
    var clean = function (L) { return L.replace(/^(?:BILL TO|SHIP TO|REMIT TO)\s*:\s*/i, '').trim(); };
    var cands = [find(lines, /(?:Make checks payable to|Beneficiary Name)\s*:\s*(.+)/i)];
    lines.forEach(function (L, i) {
      if (/Remit (?:Checks |Payment )?to\s*:?\s*$/i.test(L)) cands.push.apply(cands, lines.slice(i + 1, i + 4));
    });
    cands = cands.concat(lines.slice(0, 6)).map(function (L) { return clean(L || ''); });
    var cro = cands.filter(companyLike)[0] || '';

    // Line items: a description followed by an amount, skipping totals/tax/banking
    var items = [];
    lines.forEach(function (L) {
      var m = L.match(/^(.*[A-Za-z].*?)\s+\$?\s?([\d,]+\.\d{2})$/);
      if (m && !/total|tax|subtotal|net amount|amount (?:due|usd)|account|routing/i.test(m[1]) && items.length < 15) {
        // drop trailing quantity / unit / price columns: "Fee 1 Ea 2,520.00 2,520.00" -> "Fee"
        var desc = m[1].replace(/(\s+(?:\$?[\d,]+(?:\.\d+)?|Ea|EA|each|units?))+$/i, '').trim();
        items.push((desc || m[1].trim()) + ' $' + m[2]);
      }
    });

    var terms = find(lines, /(Net\s*(?:Due\s*in\s*)?\d{1,3}(?:\s*Days?)?)/i);

    return {
      cro: cro, invoiceNo: invoiceNo, invoiceDate: invoiceDate, dueDate: dueDate, orderNo: orderNo,
      total: money(total), paymentStage: '', terms: terms, lineItems: items,
      project: '', studyDesc: '', started: '', completion: ''
    };
  }

  // Fields shown on the page (and on the fallback form): [section, key, label].
  // Contractor's number and board approval are left for the user; invoices
  // don't state them reliably.
  var SUMMARY_FIELDS = [
    ['finance', 'cro', 'Contractor'],
    ['finance', 'croPhone', 'Contractor’s number'],
    ['finance', 'invoiceNo', 'Invoice number'],
    ['finance', 'invoiceDate', 'Invoice date'],
    ['finance', 'dueDate', 'Due date'],
    ['finance', 'total', 'Total project cost'],
    ['finance', 'orderNo', 'PO / order / agreement no.'],
    ['finance', 'paymentConditions', 'Payment conditions'],
    ['finance', 'boardApproval', 'Board approval needed?'],
    ['finance', 'dateSigned', 'Date signed'],
    ['science', 'testingType', 'Testing type'],
    ['science', 'project', 'Project number/name'],
    ['science', 'started', 'Date testing started'],
    ['science', 'completion', 'Estimated completion / draft report']
  ];

  // Template labels: [label, key]. A paragraph that starts with the label gets
  // the value added after it. The first group matches the built-in Study
  // Authorization Form; the rest are also accepted in people's own templates.
  var TEMPLATE_LABELS = [
    ['Contractor:', 'cro'], ['Contractor’s Number:', 'croPhone'], ['Contractor\'s Number:', 'croPhone'],
    ['Invoice Number:', 'invoiceNo'], ['Invoice Date:', 'invoiceDate'], ['Due Date:', 'dueDate'],
    ['Testing Type:', 'testingType'], ['Project Number/Name:', 'project'], ['Date Signed:', 'dateSigned'],
    ['Total Project Cost:', 'total'], ['Payment Conditions:', 'paymentConditions'], ['Board Approval Needed?', 'boardApproval'],
    ['Date Testing Started:', 'started'], ['Estimated Date for Testing Completion/Draft Report:', 'completion'],
    ['PO Number:', 'orderNo'],
    ['Company:', 'company'], ['CRO:', 'cro'], ['Total Amount:', 'total'], ['Payment Stage:', 'paymentConditions'],
    ['Type of Study:', 'studyType'], ['Project Number:', 'project'], ['Study Description:', 'testingType'],
    ['Date Started:', 'started'], ['Completion Date:', 'completion']
  ];
  // Multi-line entries: each line goes under the paragraph that starts with the label
  var TEMPLATE_LISTS = [
    ['Materials being tested:', 'materials'], ['Other Notes:', 'notes'], ['Line Items:', 'lineItems'],
    ['Test Materials:', 'materials'], ['Notes:', 'notes']
  ];

  function parseLines(lines, lab, today, settings) {
    today = today || new Date();
    var S = {};
    ['company', 'docType', 'cro', 'study'].forEach(function (k) { S[k] = ((settings || {})[k] || '').trim(); });
    var text = lines.join('\n');
    var detected = detectLab(text);
    lab = lab || detected;
    var r = null;
    if (lab === 'sgs') r = parseSgs(lines, today, S);
    else if (lab === 'tkl') r = parseTkl(lines, today, S);
    else if (lab === 'hilltop') r = parseHilltop(lines, today, S);
    // Cost proposals have "Agreement No.: XXXX_12345"; invoices have an "IIVS AGREEMENT NO." column header.
    else if (lab === 'iivs') r = /Agreement\s*No\.\s*:/i.test(text) ? parseIivsProposal(lines, today, S) : parseIivsInvoice(lines, today, S);

    var out = r || { kind: 'Invoice (general reader)', fields: [], chems: [], chemHint: '', extraParas: [], filename: '' };
    out.lab = r ? lab : null;
    out.detected = detected;

    // General summary, with the CRO reader's values on top where it has them
    var sum = parseGeneric(lines, today);
    var extra = (r && r.summary) || {};
    Object.keys(extra).forEach(function (k) { if (extra[k]) sum[k] = extra[k]; });
    if (r) sum.cro = S.cro || LAB_NAMES[lab];
    else if (S.cro) sum.cro = S.cro;
    sum.studyType = S.study;
    sum.company = S.company;
    sum.testingType = sum.studyDesc || S.study;
    sum.paymentConditions = [sum.paymentStage, sum.terms].filter(Boolean).join('; ');
    sum.dateSigned = fmtLong(today);
    sum.croPhone = ''; sum.boardApproval = '';
    out.summary = sum;
    out.summaryFilename = safeName((S.docType || 'Study Authorization Form') + ' ' + (sum.cro || '').split(/\s+/)[0] + ' ' + (sum.invoiceNo || '')) + '.docx';

    out.missing = out.fields.filter(function (f) { return !f.value; }).map(function (f) { return f.label; });
    if (text.replace(/\s/g, '').length < 100) out.error = 'Almost no text came out of this PDF. It may be a scanned image; fill the fields by hand.';
    return out;
  }

  // ─── built-in form (.docx from scratch) ─────────────────────────────────────
  function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  function xRun(text, o) {
    o = o || {};
    return '<w:r><w:rPr>' + (o.b ? '<w:b/>' : '') + (o.color ? '<w:color w:val="' + o.color + '"/>' : '') +
      '<w:sz w:val="' + (o.sz || 20) + '"/></w:rPr><w:t xml:space="preserve">' + esc(text) + '</w:t></w:r>';
  }
  function xPara(runs, o) {
    o = o || {};
    return '<w:p><w:pPr><w:spacing w:before="' + (o.before || 0) + '" w:after="' + (o.after == null ? 80 : o.after) + '"/></w:pPr>' + runs + '</w:p>';
  }
  function xTable(rows) {
    var border = '<w:top w:val="single" w:sz="4" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:color="BFBFBF"/><w:bottom w:val="single" w:sz="4" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:color="BFBFBF"/><w:insideH w:val="single" w:sz="4" w:color="BFBFBF"/><w:insideV w:val="single" w:sz="4" w:color="BFBFBF"/>';
    var cell = function (w, content, shade) {
      return '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/>' + (shade ? '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' : '') +
        '</w:tcPr>' + content + '</w:tc>';
    };
    return '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblBorders>' + border + '</w:tblBorders>' +
      '<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="6360"/></w:tblGrid>' +
      rows.map(function (row) {
        var valueParas = String(row[1] || '').split('\n').map(function (v) { return xPara(xRun(v), { after: 0 }); }).join('');
        return '<w:tr>' + cell(3000, xPara(xRun(row[0], { b: true }), { after: 0 }), true) + cell(6360, valueParas) + '</w:tr>';
      }).join('') + '</w:tbl>';
  }

  // data: {summary, lineItems:[..], materials:[..], notes:[..], title, company, submittedBy, today}
  async function buildSummaryDocx(data) {
    var s = data.summary, today = data.today || new Date();
    var head = function (t) { return xPara(xRun(t, { b: true, sz: 24, color: '1F3864' }), { before: 240, after: 80 }); };
    var rowsFor = function (section) {
      return SUMMARY_FIELDS.filter(function (f) { return f[0] === section; }).map(function (f) { return [f[2], s[f[1]]]; });
    };
    var list = function (label, items) {
      return xPara(xRun(label, { b: true }), { before: 120, after: 40 }) +
        (items.length ? items : ['']).map(function (t) { return xPara(xRun(t), { after: 20 }); }).join('');
    };
    var body =
      xPara(xRun(data.title || 'Study Authorization Form', { b: true, sz: 32, color: '1F3864' }), { after: 40 }) +
      xPara(xRun([data.company, 'Prepared ' + fmtLong(today)].filter(Boolean).join(' • '), { color: '595959' }), { after: 120 }) +
      head('Finance / Accounting') + xTable(rowsFor('finance')) +
      list('Line items', data.lineItems || []) +
      head('Science / Operations') + xTable(rowsFor('science')) +
      list('Test materials', data.materials || []) +
      list('Notes', data.notes || []) +
      xPara(xRun('Submitted by: ' + (data.submittedBy || '____________________') + '          Approved by: ____________________'), { before: 360 });
    return packDocx(body, data.title || 'Study Authorization Form', data.submittedBy || '');
  }

  // Wrap body XML in a minimal .docx package (Arial 10, US Letter)
  function packDocx(body, title, creator) {
    var documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="' + W + '"><w:body>' + body +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>' +
      '</w:body></w:document>';
    var stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="' + W + '"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>' +
      '<w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="80"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>';
    var iso = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    var coreXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:title>' + esc(title) + '</dc:title><dc:creator>' + esc(creator) + '</dc:creator>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">' + iso + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + iso + '</dcterms:modified></cp:coreProperties>';

    var zip = new global.JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>');
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>');
    zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
    zip.file('word/document.xml', documentXml);
    zip.file('word/styles.xml', stylesXml);
    zip.file('docProps/core.xml', coreXml);
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression: 'DEFLATE' });
  }

  // ─── .docx fill ─────────────────────────────────────────────────────────────
  function wEl(doc, name) { return doc.createElementNS(W, 'w:' + name); }
  function wAttr(el, name, val) { el.setAttributeNS(W, 'w:' + name, val); }

  function makeRun(doc, text, opt) {
    var r = wEl(doc, 'r');
    var rPr = wEl(doc, 'rPr');
    var f = wEl(doc, 'rFonts');
    ['ascii', 'hAnsi', 'eastAsia', 'cs'].forEach(function (k) { wAttr(f, k, 'Verdana'); });
    rPr.appendChild(f);
    if (opt.bold) { rPr.appendChild(wEl(doc, 'b')); rPr.appendChild(wEl(doc, 'bCs')); }
    var sz = wEl(doc, 'sz'); wAttr(sz, 'val', '22'); rPr.appendChild(sz);
    var szCs = wEl(doc, 'szCs'); wAttr(szCs, 'val', '22'); rPr.appendChild(szCs);
    if (opt.underline) { var u = wEl(doc, 'u'); wAttr(u, 'val', 'single'); rPr.appendChild(u); }
    r.appendChild(rPr);
    if (opt.br) {
      r.appendChild(wEl(doc, 'br'));
    } else {
      var t = wEl(doc, 't');
      t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
      t.textContent = text;
      r.appendChild(t);
    }
    return r;
  }

  function paraText(p) {
    var ts = p.getElementsByTagNameNS(W, 't'), s = '';
    for (var i = 0; i < ts.length; i++) s += ts[i].textContent;
    return s;
  }

  // A run in the paragraph's own font (copied from its last run), not bold.
  // text null = a line break.
  function styledRun(doc, p, text) {
    var r = wEl(doc, 'r'), runs = p.getElementsByTagNameNS(W, 'r'), rPr = null;
    for (var i = runs.length - 1; i >= 0 && !rPr; i--) {
      var first = runs[i].firstElementChild;
      if (first && first.localName === 'rPr') rPr = first;
    }
    if (rPr) {
      rPr = rPr.cloneNode(true);
      ['b', 'bCs'].forEach(function (n) {
        var els = rPr.getElementsByTagNameNS(W, n);
        while (els.length) els[0].parentNode.removeChild(els[0]);
      });
      r.appendChild(rPr);
    }
    if (text == null) { r.appendChild(wEl(doc, 'br')); return r; }
    var t = wEl(doc, 't');
    t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
    t.textContent = text;
    r.appendChild(t);
    return r;
  }
  // Insert text right after the run that contains `needle`
  function insertAfter(doc, p, needle, text) {
    var runs = p.getElementsByTagNameNS(W, 'r');
    for (var i = 0; i < runs.length; i++) {
      var ts = runs[i].getElementsByTagNameNS(W, 't'), s = '';
      for (var j = 0; j < ts.length; j++) s += ts[j].textContent;
      if (s.indexOf(needle) >= 0) {
        var nr = styledRun(doc, p, text);
        var rPr = runs[i].firstElementChild;
        if (rPr && rPr.localName === 'rPr') {
          var old = nr.firstElementChild;
          var copy = rPr.cloneNode(true);
          ['b', 'bCs'].forEach(function (n) { var els = copy.getElementsByTagNameNS(W, n); while (els.length) els[0].parentNode.removeChild(els[0]); });
          if (old && old.localName === 'rPr') nr.replaceChild(copy, old); else nr.insertBefore(copy, nr.firstChild);
        }
        runs[i].parentNode.insertBefore(nr, runs[i].nextSibling);
        return true;
      }
    }
    return false;
  }

  // spec: {fields:[{label,value}], lists:[{label, lines}], chems:[string], submittedBy:string|null,
  //        signDate:string, style:'inherit'|undefined, extraParas:[string]}
  // style 'inherit' keeps the template's labels and writes values in its font
  // (used for the built-in form and people's own templates).
  // Returns {blob, unmatched:[labels not found in the template]}
  async function fillTemplate(templateBuf, spec) {
    var zip = await global.JSZip.loadAsync(templateBuf);
    var xml = await zip.file('word/document.xml').async('string');
    var doc = new DOMParser().parseFromString(xml, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('Could not read word/document.xml in this template.');

    var matched = {};
    var inherit = spec.style === 'inherit';
    var paras = Array.prototype.slice.call(doc.getElementsByTagNameNS(W, 'p'));
    paras.forEach(function (p) {
      var text = paraText(p).trim();

      for (var i = 0; i < spec.fields.length; i++) {
        var f = spec.fields[i];
        if (text.indexOf(f.label) === 0) {
          if (inherit) {
            // keep the template's label as is; add the value in the label's font, not bold
            if (f.value) p.appendChild(styledRun(doc, p, ' ' + f.value));
          } else {
            Array.prototype.slice.call(p.childNodes).forEach(function (c) {
              if (c.localName !== 'pPr') p.removeChild(c);
            });
            p.appendChild(makeRun(doc, f.label, { bold: true }));
            p.appendChild(makeRun(doc, ' ' + (f.value || ''), { bold: true, underline: true }));
          }
          matched[f.label] = true;
          return;
        }
      }

      var lists = spec.lists || [];
      for (var li = 0; li < lists.length; li++) {
        if (text.indexOf(lists[li].label) === 0) {
          if (!inherit) {
            Array.prototype.slice.call(p.childNodes).forEach(function (c) {
              if (c.localName !== 'pPr') p.removeChild(c);
            });
            p.appendChild(makeRun(doc, lists[li].label, { bold: true }));
          }
          (lists[li].lines || []).forEach(function (line) {
            p.appendChild(inherit ? styledRun(doc, p, null) : makeRun(doc, '', { br: true }));
            p.appendChild(inherit ? styledRun(doc, p, line) : makeRun(doc, line, {}));
          });
          matched[lists[li].label] = true;
          return;
        }
      }

      if (inherit && /^Submitted by/.test(text)) {
        // "Submitted by: ____    Date: ____" on one line: name after the label, today after "Date:"
        if (spec.submittedBy) insertAfter(doc, p, 'Submitted by:', ' ' + spec.submittedBy);
        if (spec.signDate) insertAfter(doc, p, 'Date:', ' ' + spec.signDate);
        return;
      }

      if (/^List all\b/.test(text) && spec.chems && spec.chems.length) {
        spec.chems.forEach(function (line, j) {
          if (j > 0) p.appendChild(makeRun(doc, '', { br: true }));
          p.appendChild(makeRun(doc, line, { bold: true }));
        });
        return;
      }

      if (/^Submitted by/.test(text) && spec.submittedBy != null) {
        // Replace the name already on the line; on a blank "Submitted by:" line, add it
        var ts = p.getElementsByTagNameNS(W, 't'), done = false;
        for (var k = 0; k < ts.length; k++) {
          if (ts[k].textContent.indexOf('Maura Lavelle') >= 0) {
            ts[k].textContent = ts[k].textContent.replace('Maura Lavelle', spec.submittedBy);
            ts[k].setAttributeNS(XML_NS, 'xml:space', 'preserve');
            done = true;
            break;
          }
        }
        if (!done && spec.submittedBy && /^Submitted by:?\s*$/.test(text)) p.appendChild(makeRun(doc, ' ' + spec.submittedBy, {}));
      }
    });

    if (spec.extraParas && spec.extraParas.length) {
      var body = doc.getElementsByTagNameNS(W, 'body')[0];
      var sectPr = null;
      for (var c = body.lastChild; c; c = c.previousSibling) {
        if (c.nodeType === 1) { if (c.localName === 'sectPr') sectPr = c; break; }
      }
      spec.extraParas.forEach(function (txt) {
        [wEl(doc, 'p'), wEl(doc, 'p')].forEach(function (np, idx) {
          if (idx === 1) np.appendChild(makeRun(doc, txt, {}));
          body.insertBefore(np, sectPr);
        });
      });
    }

    zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
    var blob = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE'
    });
    var unmatched = spec.fields.concat(spec.lists || []).filter(function (f) { return !matched[f.label]; }).map(function (f) { return f.label; });
    return { blob: blob, unmatched: unmatched, matched: Object.keys(matched) };
  }

  global.InvoiceAutomator = {
    LAB_NAMES: LAB_NAMES,
    SUMMARY_FIELDS: SUMMARY_FIELDS,
    TEMPLATE_LABELS: TEMPLATE_LABELS,
    TEMPLATE_LISTS: TEMPLATE_LISTS,
    buildSummaryDocx: buildSummaryDocx,
    pdfToLines: pdfToLines,
    parseLines: parseLines,
    detectLab: detectLab,
    parseDate: parseDate,
    fillTemplate: fillTemplate
  };
})(window);
