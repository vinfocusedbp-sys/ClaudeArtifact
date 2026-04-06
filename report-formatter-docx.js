/**
 * report-formatter-docx.js
 * Focused Business Partners — Benefits Reconciliation Tool
 *
 * Converts raw Claude structured output into a formatted Word document (.docx).
 * No API calls. Works entirely client-side in the browser.
 * All technical tags are translated to plain English.
 * LOG lines are stripped entirely.
 *
 * Requires: docx library (loaded via CDN as window.docx)
 * CDN: https://cdn.jsdelivr.net/npm/docx@9.6.1/build/index.js
 */

const ReportFormatterDocx = (() => {

  // ── Plain English label map (replaces all technical tags) ──────────────────
  const LABELS = {
    'AMOUNT_MISMATCH' : 'Amount Mismatch',
    'DED_ONLY'        : 'In Deductions, Not on Invoice',
    'INV_ONLY'        : 'On Invoice, Not in Deductions',
    'NAME_MATCH'      : 'Possible Name Match Issue',
    'PLAN_NOTE'       : 'Plan or Coverage Discrepancy',
    'CLIENT'          : 'Client Mismatch',
    'TIME'            : 'Time Period Mismatch',
  };

  // ── Variance formatting ────────────────────────────────────────────────────
  const parseVariance = raw => {
    const s = String(raw).replace(/[$,]/g, '').trim();
    const neg = s.startsWith('-') || (s.startsWith('(') && s.endsWith(')'));
    const num = parseFloat(s.replace(/[()$\-]/g, '')) || 0;
    if (num < 0.01) return { text: 'In balance', color: '166534', bold: false };
    const fmt = '$' + num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    return neg
      ? { text: '\u2212' + fmt + '  under-billed', color: '92400E', bold: true }
      : { text: '+' + fmt + '  over-billed',  color: 'B91C1C', bold: true };
  };

  // ── Raw text parsing helpers ───────────────────────────────────────────────
  const stripLogs = raw => raw
    .split('\n')
    .filter(l => !/^(LOG:|##END)/i.test(l.trim()))
    .join('\n');

  const parsePipe = (line, prefix) =>
    line.replace(new RegExp('^' + prefix + '\\s*:\\s*', 'i'), '')
        .split('|').map(s => s.trim());

  const getLines = (lines, prefix) =>
    lines.filter(l => l.trim().toUpperCase().startsWith(prefix.toUpperCase() + ':'))
         .map(l => l.trim().slice(prefix.length + 1).trim());

  const get1 = (lines, prefix) => getLines(lines, prefix)[0] || '';

  // ── docx element factories ─────────────────────────────────────────────────
  // These are resolved at runtime from window.docx (browser) or require (Node)
  let D = null;

  const initDocx = () => {
    if (D) return;
    if (typeof window !== 'undefined' && window.docx) {
      D = window.docx;
    } else {
      D = require('./node_modules/docx');
    }
  };

  const BRAND_BLUE  = '0647A1';
  const LIGHT_BLUE  = 'E8EEF8';
  const BORDER_GREY = 'CBD5E1';
  const CONTENT_W   = 9360; // US Letter - 1" margins each side

  const cellBorder = () => ({
    top:    { style: D.BorderStyle.SINGLE, size: 1, color: BORDER_GREY },
    bottom: { style: D.BorderStyle.SINGLE, size: 1, color: BORDER_GREY },
    left:   { style: D.BorderStyle.SINGLE, size: 1, color: BORDER_GREY },
    right:  { style: D.BorderStyle.SINGLE, size: 1, color: BORDER_GREY },
  });

  const cellMargins = () => ({ top: 80, bottom: 80, left: 120, right: 120 });

  const run = (text, opts = {}) => new D.TextRun({
    text,
    font: 'Arial',
    size: opts.size || 20,
    bold: opts.bold || false,
    color: opts.color || '000000',
    italics: opts.italic || false,
  });

  const para = (children, opts = {}) => new D.Paragraph({
    children: Array.isArray(children) ? children : [children],
    spacing: { before: opts.before || 0, after: opts.after || 120 },
    alignment: opts.align || D.AlignmentType.LEFT,
    ...(opts.border ? { border: { bottom: { style: D.BorderStyle.SINGLE, size: 6, color: BRAND_BLUE, space: 1 } } } : {}),
  });

  const heading = (text, level = 1) => {
    const sizes   = { 1: 28, 2: 24, 3: 22 };
    const befores = { 1: 320, 2: 240, 3: 180 };
    return para(
      run(text.toUpperCase(), { bold: true, size: sizes[level], color: level === 1 ? BRAND_BLUE : '1E293B' }),
      { before: befores[level], after: 120, border: level <= 2 }
    );
  };

  const spacer = () => para(run(''), { before: 0, after: 60 });

  const metaRow = (label, value) => para([
    run(label + ': ', { bold: true, size: 20 }),
    run(value || '—', { size: 20 }),
  ], { after: 80 });

  const alertBlock = (type, message) => {
    const isClient = type === 'CLIENT';
    const icon  = isClient ? '⚠  Client Mismatch' : '⚠  Time Period Mismatch';
    const color = isClient ? 'B91C1C' : '92400E';
    // Strip the **Esteban...** confirmation lines — shown in the UI, not needed in doc
    const clean = message
      .replace(/\*\*Esteban.*?\*\*/gi, '')
      .replace(/\s{2,}/g, ' ')
      .trim();
    return [
      para(run(icon, { bold: true, size: 20, color }), { before: 120, after: 60 }),
      para(run(clean, { size: 19, color: '374151' }), { after: 120 }),
    ];
  };

  // Generic exception table
  const exceptionTable = (headers, rows, colWidths) => {
    const headerRow = new D.TableRow({
      tableHeader: true,
      children: headers.map((h, i) => new D.TableCell({
        borders: cellBorder(),
        margins: cellMargins(),
        width: { size: colWidths[i], type: D.WidthType.DXA },
        shading: { fill: LIGHT_BLUE, type: D.ShadingType.CLEAR },
        children: [para(run(h, { bold: true, size: 18, color: '1E293B' }), { after: 0 })],
      })),
    });

    const dataRows = rows.map(r => new D.TableRow({
      children: r.map((cell, i) => {
        const isVariance = headers[i] === 'Variance' || headers[i] === 'Net Variance';
        const varInfo = isVariance ? parseVariance(cell) : null;
        return new D.TableCell({
          borders: cellBorder(),
          margins: cellMargins(),
          width: { size: colWidths[i], type: D.WidthType.DXA },
          children: [para(
            varInfo
              ? run(varInfo.text, { bold: varInfo.bold, size: 18, color: varInfo.color })
              : run(String(cell || ''), { size: 18 }),
            { after: 0 }
          )],
        });
      }),
    }));

    return new D.Table({
      width: { size: CONTENT_W, type: D.WidthType.DXA },
      columnWidths: colWidths,
      rows: [headerRow, ...dataRows],
    });
  };

  // Summary stats table
  const summaryTable = (stats) => {
    return new D.Table({
      width: { size: CONTENT_W, type: D.WidthType.DXA },
      columnWidths: [4000, 5360],
      rows: stats.map(([label, value, isVariance]) => {
        const varInfo = isVariance ? parseVariance(value) : null;
        return new D.TableRow({
          children: [
            new D.TableCell({
              borders: cellBorder(),
              margins: cellMargins(),
              width: { size: 4000, type: D.WidthType.DXA },
              shading: { fill: 'F8FAFC', type: D.ShadingType.CLEAR },
              children: [para(run(label, { bold: true, size: 18 }), { after: 0 })],
            }),
            new D.TableCell({
              borders: cellBorder(),
              margins: cellMargins(),
              width: { size: 5360, type: D.WidthType.DXA },
              children: [para(
                varInfo
                  ? run(varInfo.text, { bold: varInfo.bold, size: 18, color: varInfo.color })
                  : run(String(value || '—'), { size: 18 }),
                { after: 0 }
              )],
            }),
          ],
        });
      }),
    });
  };

  // ── Report block builder ───────────────────────────────────────────────────
  const buildReportBlock = (blockLines, carrier, period, isMulti) => {
    const els = [];

    const client   = get1(blockLines, 'CLIENT');
    const dedCount = get1(blockLines, 'DED_COUNT');
    const invCount = get1(blockLines, 'INV_COUNT');
    const dedTotal = get1(blockLines, 'DED_TOTAL');
    const invTotal = get1(blockLines, 'INV_TOTAL');
    const rawVar   = get1(blockLines, 'VARIANCE').split(':')[0] || '';

    // Carrier header
    const carrierLabel = carrier.toUpperCase() + (period ? '  ·  ' + period : '');
    els.push(heading(carrierLabel, isMulti ? 2 : 1));
    if (client) els.push(metaRow('Client', client));

    // Summary
    els.push(heading('Summary', 3));
    const summaryStats = [
      ['Employees in deductions', dedCount || '—', false],
      ['Employees on invoice',    invCount || '—', false],
      ['Total — deductions',      dedTotal || '—', false],
      ['Total — invoice',         invTotal || '—', false],
      ['Net variance',            rawVar   || '—', true],
    ];
    els.push(summaryTable(summaryStats));
    els.push(spacer());

    // Amount mismatches
    const amtRows = getLines(blockLines, 'AMOUNT_MISMATCH').map(l => parsePipe(l, ''));
    if (amtRows.length) {
      els.push(heading(LABELS.AMOUNT_MISMATCH, 3));
      els.push(exceptionTable(
        ['Employee', 'Deductions Amount', 'Invoice Amount', 'Variance'],
        amtRows,
        [2800, 2000, 2000, 2560]
      ));
      els.push(spacer());
    }

    // On deductions, not on invoice
    const dedRows = getLines(blockLines, 'DED_ONLY').map(l => parsePipe(l, ''));
    if (dedRows.length) {
      els.push(heading(LABELS.DED_ONLY, 3));
      els.push(exceptionTable(
        ['Employee', 'Monthly Amount', 'Notes'],
        dedRows,
        [3000, 2000, 4360]
      ));
      els.push(spacer());
    }

    // On invoice, not on deductions
    const invRows = getLines(blockLines, 'INV_ONLY').map(l => parsePipe(l, ''));
    if (invRows.length) {
      els.push(heading(LABELS.INV_ONLY, 3));
      els.push(exceptionTable(
        ['Employee', 'Invoice Amount', 'Notes'],
        invRows,
        [3000, 2000, 4360]
      ));
      els.push(spacer());
    }

    // Name match issues
    const nameRows = getLines(blockLines, 'NAME_MATCH').map(l => parsePipe(l, ''));
    if (nameRows.length) {
      els.push(heading(LABELS.NAME_MATCH, 3));
      els.push(exceptionTable(
        ['Deductions Name', 'Invoice Name', 'Confidence', 'Notes'],
        nameRows,
        [2400, 2400, 1200, 3360]
      ));
      els.push(spacer());
    }

    // Plan & coverage notes
    const planRows = getLines(blockLines, 'PLAN_NOTE').map(l => parsePipe(l, ''));
    if (planRows.length) {
      els.push(heading(LABELS.PLAN_NOTE, 3));
      els.push(exceptionTable(
        ['Employee', 'Deductions', 'Invoice'],
        planRows,
        [2800, 3280, 3280]
      ));
      els.push(spacer());
    }

    // No exceptions message
    const hasExceptions = amtRows.length || dedRows.length || invRows.length || nameRows.length;
    if (!hasExceptions) {
      els.push(para(
        run('No exceptions found. Deductions and invoice are in agreement.', { size: 19, color: '166534' }),
        { before: 120, after: 120 }
      ));
    }

    return els;
  };

  // ── Combined totals table (multi-carrier) ──────────────────────────────────
  const buildCombinedTotals = blocks => {
    const rows = blocks.map(b => {
      const ded    = get1(b.lines, 'DED_TOTAL');
      const inv    = get1(b.lines, 'INV_TOTAL');
      const rawVar = get1(b.lines, 'VARIANCE').split(':')[0] || '';
      return [b.carrier + (b.period ? ' · ' + b.period : ''), ded || '—', inv || '—', rawVar || '—'];
    });

    return [
      heading('Combined Totals — All Carriers', 2),
      exceptionTable(
        ['Carrier', 'Deductions Total', 'Invoice Total', 'Variance'],
        rows,
        [3000, 2000, 2000, 2360]
      ),
      spacer(),
    ];
  };

  // ── Main entry point ───────────────────────────────────────────────────────
  const buildDoc = raw => {
    initDocx();

    const cleaned = stripLogs(raw);
    const lines   = cleaned.split('\n').map(l => l.trim()).filter(l => l);

    const children = [];

    // Title
    children.push(
      new D.Paragraph({
        children: [run('Benefits Reconciliation Report', { bold: true, size: 32, color: BRAND_BLUE })],
        spacing: { before: 0, after: 80 },
        border: { bottom: { style: D.BorderStyle.SINGLE, size: 10, color: BRAND_BLUE, space: 2 } },
      })
    );
    children.push(para(
      run('Focused Business Partners  ·  ' + new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' }),
        { size: 18, color: '64748B' }),
      { after: 240 }
    ));

    // Details
    const detStart = lines.findIndex(l => l.startsWith('##DETAILS'));
    const altStart = lines.findIndex(l => l.startsWith('##ALERTS'));
    const repIdxs  = lines.reduce((a, l, i) => { if (/^##REPORT[: ]/i.test(l)) a.push(i); return a; }, []);

    if (detStart !== -1) {
      const end = altStart !== -1 ? altStart : repIdxs[0] ?? lines.length;
      const dLines = lines.slice(detStart + 1, end);
      children.push(heading('Document Details', 2));
      dLines.forEach(l => {
        if (l.startsWith('REG:'))
          children.push(metaRow('Payroll Register', l.replace(/^REG:\s*/i, '')));
        else if (l.startsWith('INV:'))
          children.push(metaRow('Carrier Invoice', l.replace(/^INV:\s*/i, '')));
      });
      children.push(spacer());
    }

    // Alerts
    if (altStart !== -1) {
      const end = repIdxs[0] ?? lines.length;
      const aLines = lines.slice(altStart + 1, end);
      const hasNone = aLines.some(l => /^NONE$/i.test(l));
      const clientLines = aLines.filter(l => /^CLIENT:/i.test(l));
      const timeLines   = aLines.filter(l => /^TIME:/i.test(l));

      children.push(heading('Mismatch Alerts', 2));
      if (hasNone && !clientLines.length && !timeLines.length) {
        children.push(para(
          run('No client or time period mismatches detected.', { size: 19, color: '166534' }),
          { after: 120 }
        ));
      } else {
        clientLines.forEach(l => alertBlock('CLIENT', l.replace(/^CLIENT:\s*/i, '')).forEach(e => children.push(e)));
        timeLines.forEach(l  => alertBlock('TIME',   l.replace(/^TIME:\s*/i,   '')).forEach(e => children.push(e)));
      }
      children.push(spacer());
    }

    // Report blocks
    const blocks = repIdxs.map((start, ri) => {
      const end   = repIdxs[ri + 1] ?? lines.length;
      const hdr   = lines[start].replace(/^##REPORT[: ]*/i, '');
      const parts = hdr.split(':');
      return {
        carrier : parts[0].trim(),
        period  : parts.slice(1).join(':').trim(),
        lines   : lines.slice(start + 1, end),
      };
    });

    const isMulti = blocks.length > 1;
    blocks.forEach((b, bi) => {
      if (isMulti && bi > 0) {
        children.push(new D.Paragraph({
          children: [],
          border: { top: { style: D.BorderStyle.SINGLE, size: 6, color: BRAND_BLUE, space: 4 } },
          spacing: { before: 240, after: 0 },
        }));
      }
      buildReportBlock(b.lines, b.carrier, b.period, isMulti)
        .forEach(el => children.push(el));
    });

    // Combined totals
    if (isMulti) {
      children.push(new D.Paragraph({
        children: [],
        border: { top: { style: D.BorderStyle.SINGLE, size: 10, color: BRAND_BLUE, space: 4 } },
        spacing: { before: 240, after: 0 },
      }));
      buildCombinedTotals(blocks).forEach(el => children.push(el));
    }

    return new D.Document({
      styles: {
        default: {
          document: { run: { font: 'Arial', size: 20 } },
        },
      },
      sections: [{
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
          },
        },
        children,
      }],
    });
  };

  // ── Browser export: returns a Blob promise ─────────────────────────────────
  const toBlob = raw => {
    initDocx();
    const doc = buildDoc(raw);
    return D.Packer.toBlob(doc);
  };

  // ── Node export: returns a Buffer promise ──────────────────────────────────
  const toBuffer = raw => {
    initDocx();
    const doc = buildDoc(raw);
    return D.Packer.toBuffer(doc);
  };

  return { toBlob, toBuffer, buildDoc };

})();

if (typeof module !== 'undefined' && module.exports) {
  module.exports = ReportFormatterDocx;
}
