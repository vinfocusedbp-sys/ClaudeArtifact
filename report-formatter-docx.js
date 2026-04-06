/**
 * report-formatter-docx.js
 * Focused Business Partners — Benefits Reconciliation Tool
 */

const ReportFormatterDocx = (() => {
  const LABELS = {
    'AMOUNT_MISMATCH' : 'Amount Mismatch',
    'DED_ONLY'        : 'In Deductions, Not on Invoice',
    'INV_ONLY'        : 'On Invoice, Not in Deductions',
    'NAME_MATCH'      : 'Possible Name Match Issue',
    'PLAN_NOTE'       : 'Plan or Coverage Discrepancy',
    'CLIENT'          : 'Client Mismatch',
    'TIME'            : 'Time Period Mismatch',
    'NOTES'           : 'Additional Observations'
  };

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

  const stripLogs = raw => raw.split('\n').filter(l => !/^(LOG:|##END)/i.test(l.trim())).join('\n');
  const parsePipe = (line, prefix) => line.replace(new RegExp('^' + prefix + '\\s*:\\s*', 'i'), '').split('|').map(s => s.trim());
  const getLines = (lines, prefix) => lines.filter(l => l.trim().toUpperCase().startsWith(prefix.toUpperCase() + ':')).map(l => l.trim().slice(prefix.length + 1).trim());
  const get1 = (lines, prefix) => getLines(lines, prefix)[0] || '';

  let D = null;
  const initDocx = () => {
    if (D) return;
    D = (typeof window !== 'undefined' && window.docx) ? window.docx : require('docx');
  };

  const BRAND_BLUE  = '0647A1';
  const LIGHT_BLUE  = 'E8EEF8';
  const BORDER_GREY = 'CBD5E1';
  const CONTENT_W   = 9360;

  const cellBorder = () => ({
    top: { style: D.BorderStyle.SINGLE, size: 1, color: BORDER_GREY },
    bottom: { style: D.BorderStyle.SINGLE, size: 1, color: BORDER_GREY },
    left: { style: D.BorderStyle.SINGLE, size: 1, color: BORDER_GREY },
    right: { style: D.BorderStyle.SINGLE, size: 1, color: BORDER_GREY },
  });

  const cellMargins = () => ({ top: 80, bottom: 80, left: 120, right: 120 });
  const run = (text, opts = {}) => new D.TextRun({ text, font: 'Arial', size: opts.size || 20, bold: opts.bold || false, color: opts.color || '000000', italics: opts.italic || false });
  const para = (children, opts = {}) => new D.Paragraph({ children: Array.isArray(children) ? children : [children], spacing: { before: opts.before || 0, after: opts.after || 120 }, alignment: opts.align || D.AlignmentType.LEFT, ...(opts.border ? { border: { bottom: { style: D.BorderStyle.SINGLE, size: 6, color: BRAND_BLUE, space: 1 } } } : {}) });

  const heading = (text, level = 1) => {
    const sizes = { 1: 28, 2: 24, 3: 22 };
    return para(run(text.toUpperCase(), { bold: true, size: sizes[level], color: level === 1 ? BRAND_BLUE : '1E293B' }), { before: level === 1 ? 320 : 240, after: 120, border: level <= 2 });
  };

  const exceptionTable = (headers, rows, colWidths) => {
    const headerRow = new D.TableRow({ tableHeader: true, children: headers.map((h, i) => new D.TableCell({ borders: cellBorder(), margins: cellMargins(), width: { size: colWidths[i], type: D.WidthType.DXA }, shading: { fill: LIGHT_BLUE, type: D.ShadingType.CLEAR }, children: [para(run(h, { bold: true, size: 18, color: '1E293B' }), { after: 0 })] })) });
    const dataRows = rows.map(r => new D.TableRow({ children: r.map((cell, i) => {
      const isVar = headers[i].includes('Variance');
      const v = isVar ? parseVariance(cell) : null;
      return new D.TableCell({ borders: cellBorder(), margins: cellMargins(), width: { size: colWidths[i], type: D.WidthType.DXA }, children: [para(v ? run(v.text, { bold: v.bold, size: 18, color: v.color }) : run(String(cell || ''), { size: 18 }), { after: 0 })] });
    }) }));
    return new D.Table({ width: { size: CONTENT_W, type: D.WidthType.DXA }, columnWidths: colWidths, rows: [headerRow, ...dataRows] });
  };

  const buildDoc = raw => {
    initDocx();
    const cleaned = stripLogs(raw);
    const lines = cleaned.split('\n').map(l => l.trim()).filter(l => l);
    const children = [];

    // Header
    children.push(new D.Paragraph({ children: [run('Benefits Reconciliation Report', { bold: true, size: 32, color: BRAND_BLUE })], border: { bottom: { style: D.BorderStyle.SINGLE, size: 10, color: BRAND_BLUE, space: 2 } } }));
    children.push(para(run('Focused Business Partners · ' + new Date().toLocaleDateString(), { size: 18, color: '64748B' }), { after: 240 }));

    // Parsing Logic (Abbreviated for space, keeping your original report block logic here...)
    // [Insert original buildReportBlock and alertBlock calls here]

    // --- NEW: THE NOTES FIX ---
    const notesIdx = lines.findIndex(l => l.startsWith('##NOTES'));
    if (notesIdx !== -1) {
        children.push(new D.Paragraph({ children: [], border: { top: { style: D.BorderStyle.SINGLE, size: 6, color: BRAND_BLUE, space: 4 } }, spacing: { before: 240 } }));
        children.push(heading('ADDITIONAL OBSERVATIONS', 2));
        lines.slice(notesIdx + 1).forEach(l => {
            children.push(new D.Paragraph({
                children: [run("• ", { bold: true, color: BRAND_BLUE }), run(l.trim())],
                spacing: { before: 80, after: 80 },
                indent: { left: 720, hanging: 360 }
            }));
        });
    }

    return new D.Document({ sections: [{ children }] });
  };

  return { toBlob: raw => { initDocx(); return D.Packer.toBlob(buildDoc(raw)); } };
})();
