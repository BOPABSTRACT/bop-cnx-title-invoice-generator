import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import { PDFDocument } from 'pdf-lib';
import JSZip from 'jszip';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';

export const runtime = 'nodejs';
export const maxDuration = 60;

// ===================================================================
// Types
// ===================================================================

interface EmailRow {
  padSite: string;
  afe: string;
  workScope: string;
  pid: string;
  township: string;
  county: string;
}

interface ParsedEmail {
  subject: string;
  padSiteFromSubject: string;
  countyFromSubject: string;
  rows: EmailRow[];
  byAfe: Map<string, EmailRow[]>;
  byPid: Map<string, EmailRow>;
}

interface BrokerSummaryRow {
  broker: string;
  days: number;
  amtPerDay: number;
  total: number;
  professionalServices: number;
  copies: number;
  miscellaneous: number;
  grandTotal: number;
}

interface WorkDetailRow {
  landman: string;
  date: string;
  prospect: string;
  legal: string;
  projectFocus: string;
  days: number;
  copies: number;
  misc: number;
  total: number;
}

interface InvoiceData {
  invoiceNumber: string;
  invoiceDate: string;
  invoiceDateObj: Date | null;
  period: string;
  afe: string;
  companyCode: string;
  padSite: string;
  project: string;
  county: string;
  billToLines: string[];
  brokerRows: BrokerSummaryRow[];
  brokerTotals: BrokerSummaryRow | null;
  workRows: WorkDetailRow[];
}

// ===================================================================
// PID regex (CNX/Westmoreland format: 57-04-00-0-048)
// ===================================================================

const CNX_PID_RE = /\d{2}-\d{2}-\d{2}-\d-\d{3,4}/g;

function extractCnxPids(text: string): string[] {
  const repaired = text.replace(/(\d-)\s+(\d)/g, '$1$2');
  const matches = repaired.match(CNX_PID_RE) ?? [];
  const seen = new Set<string>();
  const out: string[] = [];
  for (const m of matches) {
    if (!seen.has(m)) {
      seen.add(m);
      out.push(m);
    }
  }
  return out;
}

// ===================================================================
// Google Vision OCR fallback for image-only PDFs
// ===================================================================

async function ocrPdfWithGoogleVision(buffer: Buffer): Promise<string> {
  const apiKey = process.env.GOOGLE_VISION_API_KEY;
  if (!apiKey) {
    throw new Error(
      'PDF appears to be image-only and no Google Vision API key is configured. ' +
        'Set GOOGLE_VISION_API_KEY in Vercel environment variables, or use Outlook File → Save As → PDF.'
    );
  }

  const base64Pdf = buffer.toString('base64');

  const requestBody = {
    requests: [
      {
        inputConfig: {
          mimeType: 'application/pdf',
          content: base64Pdf,
        },
        features: [{ type: 'DOCUMENT_TEXT_DETECTION' }],
      },
    ],
  };

  const res = await fetch(
    `https://vision.googleapis.com/v1/files:annotate?key=${apiKey}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(requestBody),
    }
  );

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`Google Vision API error (${res.status}): ${errText}`);
  }

  const data = await res.json();

  const pageResponses = data?.responses?.[0]?.responses ?? [];
  const fullText = pageResponses
    .map((p: any) => p?.fullTextAnnotation?.text ?? '')
    .filter(Boolean)
    .join('\n');

  return fullText;
}

// ===================================================================
// Email PDF parsing — CNX upload-notification format
// ===================================================================

async function parseEmailPdf(buffer: Buffer): Promise<ParsedEmail> {
  const pdfParse = (await import('pdf-parse')).default;
  let text = '';
  try {
    const data = await pdfParse(buffer);
    text = data.text;
  } catch {
    text = '';
  }

  if (!text || text.trim().length < 40) {
    text = await ocrPdfWithGoogleVision(buffer);
  }

  if (!text || text.trim().length < 40) {
    throw new Error(
      'Could not extract any text from the email PDF, even with OCR. ' +
        'Please re-export the email from Outlook using File → Save As → PDF.'
    );
  }

  const subjectMatch = text.match(/Subject:\s*(.+)/i);
  const subject = subjectMatch ? subjectMatch[1].trim() : '';

  let padSiteFromSubject = '';
  let countyFromSubject = '';
  const subParts = subject.split(/\s*-\s*/);
  if (subParts.length >= 3) {
    padSiteFromSubject = subParts[1]?.trim() ?? '';
    const countyPart = subParts[2]?.trim() ?? '';
    countyFromSubject = countyPart.replace(/County.*$/i, '').trim();
  }

  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);

  const rows: EmailRow[] = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const pids = extractCnxPids(line);
    if (pids.length === 0) continue;

    const window = lines.slice(Math.max(0, i - 2), Math.min(lines.length, i + 3)).join(' ');

    const afeMatches = window.match(/\b\d{5,7}\b/g) ?? [];
    const afe = afeMatches.find((m) => !line.replace(/[^0-9]/g, '').includes(m + '')) ?? afeMatches[0] ?? '';

    const scopeMatch = window.match(/((?:Surface,\s*)?(?:All\s+)?O&G\s+Formations)/i);
    const workScope = scopeMatch ? scopeMatch[1].trim() : '';

    const tcMatch = window.match(/([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\s+Twp[\.,]?\s*([A-Z][a-z]+)\s+County/);
    const township = tcMatch ? `${tcMatch[1]} Twp` : '';
    const county = tcMatch ? tcMatch[2] : countyFromSubject;

    let padSite = '';
    const padPatterns = [
      /\b(FATUR\s+Utica\s+North)\b/i,
      /\b([A-Z][A-Z]+\s+(?:Utica|Marcellus)\s+[A-Z][a-z]+)\b/,
    ];
    for (const p of padPatterns) {
      const m = window.match(p);
      if (m) {
        padSite = m[1].trim().replace(/\s+/g, ' ');
        padSite = padSite.replace(/^FATUR/, 'Fatur');
        break;
      }
    }
    if (!padSite) padSite = padSiteFromSubject;

    for (const pid of pids) {
      rows.push({ padSite, afe, workScope, pid, township, county });
    }
  }

  const byAfe = new Map<string, EmailRow[]>();
  const byPid = new Map<string, EmailRow>();
  for (const r of rows) {
    if (r.afe) {
      const arr = byAfe.get(r.afe) ?? [];
      arr.push(r);
      byAfe.set(r.afe, arr);
    }
    if (r.pid) byPid.set(r.pid, r);
  }

  return { subject, padSiteFromSubject, countyFromSubject, rows, byAfe, byPid };
}

// ===================================================================
// Excel parsing
// ===================================================================

function findRowByLabel(rows: any[][], label: RegExp, maxScan = 30): any[] | null {
  for (let i = 0; i < Math.min(rows.length, maxScan); i++) {
    const r = rows[i] ?? [];
    for (const cell of r) {
      if (typeof cell === 'string' && label.test(cell)) return r;
    }
  }
  return null;
}

function findColIdxByHeader(headerRow: any[], names: RegExp): number {
  for (let i = 0; i < headerRow.length; i++) {
    const v = String(headerRow[i] ?? '');
    if (names.test(v)) return i;
  }
  return -1;
}

function num(v: any): number {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  const cleaned = String(v).replace(/[$,\s]/g, '');
  const n = parseFloat(cleaned);
  return isFinite(n) ? n : 0;
}

function parseExcel(buffer: Buffer, fallbackInvoiceNumber: string): InvoiceData {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });

  const summaryName =
    wb.SheetNames.find((n) => /summary/i.test(n)) ?? wb.SheetNames[0];
  const workName =
    wb.SheetNames.find((n) => /work\s*detail/i.test(n)) ??
    wb.SheetNames[1] ??
    wb.SheetNames[0];

  const summary = XLSX.utils.sheet_to_json(wb.Sheets[summaryName], {
    header: 1,
    defval: '',
  }) as any[][];
  const work = XLSX.utils.sheet_to_json(wb.Sheets[workName], {
    header: 1,
    defval: '',
  }) as any[][];

  const row8 = summary[7] ?? [];
  const row16 = summary[15] ?? [];

  let invoiceDateObj: Date | null = null;
  let invoiceDate = '';
  const rawDate = row8[1];
  if (rawDate instanceof Date) {
    invoiceDateObj = rawDate;
    invoiceDate = rawDate.toLocaleDateString('en-US');
  } else if (rawDate) {
    invoiceDate = String(rawDate);
    const parsed = new Date(invoiceDate);
    if (!isNaN(parsed.getTime())) invoiceDateObj = parsed;
  }
  if (!invoiceDate) {
    const dateRow = findRowByLabel(summary, /^Date\s*:/i);
    if (dateRow) {
      const idx = dateRow.findIndex((c) => typeof c === 'string' && /^Date\s*:/i.test(c));
      const v = dateRow[idx + 1] ?? dateRow.slice(idx).find((x, k) => k > 0 && x !== '');
      if (v instanceof Date) {
        invoiceDateObj = v;
        invoiceDate = v.toLocaleDateString('en-US');
      } else if (v) {
        invoiceDate = String(v);
        const p = new Date(invoiceDate);
        if (!isNaN(p.getTime())) invoiceDateObj = p;
      }
    }
  }

  let invoiceNumber = String(row8[5] ?? '').trim();
  if (!invoiceNumber) {
    const invRow = findRowByLabel(summary, /Invoice\s*#/i);
    if (invRow) {
      const idx = invRow.findIndex((c) => typeof c === 'string' && /Invoice\s*#/i.test(c));
      invoiceNumber = String(invRow[idx + 1] ?? '').trim();
    }
  }
  if (!invoiceNumber) invoiceNumber = fallbackInvoiceNumber;

  let period = String(row16[1] ?? '').trim();
  if (!period) {
    const periodRow = findRowByLabel(summary, /^Period\s*:/i);
    if (periodRow) {
      const idx = periodRow.findIndex(
        (c) => typeof c === 'string' && /^Period\s*:/i.test(c)
      );
      period = String(periodRow[idx + 1] ?? '').trim();
    }
  }

  const findLabeled = (re: RegExp): string => {
    const r = findRowByLabel(summary, re);
    if (!r) return '';
    const idx = r.findIndex((c) => typeof c === 'string' && re.test(c));
    const next = r.slice(idx + 1).find((c) => c !== '' && c != null);
    return String(next ?? '').trim();
  };

  const afe = findLabeled(/^AFE\s*:/i);
  const companyCode = findLabeled(/Company\s*Code/i);
  const padSite = findLabeled(/Pad\s*Site/i);
  const project = findLabeled(/^Project\s*:/i) || 'CNX Title';
  const county = findLabeled(/^County\s*:/i);

  const billToLines: string[] = [];
  for (let i = 0; i < summary.length; i++) {
    const r = summary[i] ?? [];
    const hit = r.find((c) => typeof c === 'string' && /Bill\s*To/i.test(c));
    if (hit) {
      const idx = r.findIndex((c) => typeof c === 'string' && /Bill\s*To/i.test(c));
      const first = String(r[idx + 1] ?? '').trim();
      if (first) billToLines.push(first);
      for (let j = 1; j <= 4; j++) {
        const sub = summary[i + j] ?? [];
        const v = String(sub[idx + 1] ?? '').trim();
        if (!v) break;
        if (/^[A-Z][A-Za-z\s]+:\s*/.test(v)) break;
        billToLines.push(v);
      }
      break;
    }
  }

  const brokerRows: BrokerSummaryRow[] = [];
  let brokerTotals: BrokerSummaryRow | null = null;

  let brokerHeaderIdx = -1;
  for (let i = 0; i < summary.length; i++) {
    const r = summary[i] ?? [];
    const hasBroker = r.some((c) => typeof c === 'string' && /^Broker\b/i.test(c));
    const hasDays = r.some((c) => typeof c === 'string' && /\bDays\b/i.test(c));
    if (hasBroker && hasDays) {
      brokerHeaderIdx = i;
      break;
    }
  }

  if (brokerHeaderIdx >= 0) {
    const headerRow = summary[brokerHeaderIdx];
    const colBroker = findColIdxByHeader(headerRow, /^Broker\b/i);
    const colDays = findColIdxByHeader(headerRow, /^Days\b/i);
    const colAmtPer = findColIdxByHeader(headerRow, /Amt\.?\s*Per\s*Day/i);
    const colTotal = findColIdxByHeader(headerRow, /^Total\b/i);
    const colProfSvc = findColIdxByHeader(headerRow, /Professional\s*Services/i);
    const colCopies = findColIdxByHeader(headerRow, /^Copies\b/i);
    const colMisc = findColIdxByHeader(headerRow, /^Miscellaneous\b/i);
    let colGrand = findColIdxByHeader(headerRow, /^TOTAL$/);
    if (colGrand < 0) colGrand = headerRow.length - 1;

    for (let i = brokerHeaderIdx + 1; i < summary.length; i++) {
      const r = summary[i];
      if (!r || r.every((c) => c === '' || c == null)) continue;
      const brokerName = String(r[colBroker] ?? '').trim();
      if (!brokerName) continue;

      const row: BrokerSummaryRow = {
        broker: brokerName,
        days: num(r[colDays]),
        amtPerDay: num(r[colAmtPer]),
        total: num(r[colTotal]),
        professionalServices: num(r[colProfSvc]),
        copies: num(r[colCopies]),
        miscellaneous: num(r[colMisc]),
        grandTotal: num(r[colGrand]),
      };

      if (/^totals?$/i.test(brokerName)) {
        brokerTotals = row;
        break;
      }
      brokerRows.push(row);
    }
  }

  let workHeaderIdx = -1;
  for (let i = 0; i < Math.min(work.length, 10); i++) {
    const r = work[i] ?? [];
    if (
      r.some((c) => typeof c === 'string' && /Landman/i.test(c)) &&
      r.some((c) => typeof c === 'string' && /Project\s*Focus/i.test(c))
    ) {
      workHeaderIdx = i;
      break;
    }
  }
  if (workHeaderIdx < 0) workHeaderIdx = 1;

  const workHeader = work[workHeaderIdx] ?? [];
  const wIdxLandman = findColIdxByHeader(workHeader, /^Landman\b/i);
  const wIdxDate = findColIdxByHeader(workHeader, /^Date\b/i);
  const wIdxProspect = findColIdxByHeader(workHeader, /^Prospect\b/i);
  const wIdxLegal = findColIdxByHeader(workHeader, /^Legal\b/i);
  const wIdxFocus = findColIdxByHeader(workHeader, /Project\s*Focus/i);
  const wIdxDays = findColIdxByHeader(workHeader, /^Days\b/i);
  const wIdxCopies = findColIdxByHeader(workHeader, /^Copies\b/i);
  const wIdxMisc = findColIdxByHeader(workHeader, /^Misc\.?/i);
  const wIdxTotal = findColIdxByHeader(workHeader, /^Total\b/i);

  const workRows: WorkDetailRow[] = [];
  for (let i = workHeaderIdx + 1; i < work.length; i++) {
    const r = work[i];
    if (!r || r.every((c) => c === '' || c == null)) continue;

    const landman = String(r[wIdxLandman] ?? '').trim();
    const focus = String(r[wIdxFocus] ?? '').trim();
    if (!landman && !r[wIdxDate]) {
      continue;
    }

    let dateStr = '';
    const rawWDate = r[wIdxDate];
    if (rawWDate instanceof Date) {
      dateStr =
        rawWDate.getMonth() + 1 + '/' + rawWDate.getDate() + '/' + rawWDate.getFullYear();
    } else if (rawWDate) {
      dateStr = String(rawWDate);
    }

    workRows.push({
      landman,
      date: dateStr,
      prospect: String(r[wIdxProspect] ?? '').trim(),
      legal: String(r[wIdxLegal] ?? '').trim(),
      projectFocus: focus,
      days: num(r[wIdxDays]),
      copies: num(r[wIdxCopies]),
      misc: num(r[wIdxMisc]),
      total: num(r[wIdxTotal]),
    });
  }

  return {
    invoiceNumber,
    invoiceDate,
    invoiceDateObj,
    period,
    afe,
    companyCode,
    padSite,
    project,
    county,
    billToLines,
    brokerRows,
    brokerTotals,
    workRows,
  };
}

// ===================================================================
// Receipt matching
// ===================================================================

function matchReceipt(
  invoiceNumber: string,
  receipts: { name: string; buffer: Buffer }[]
): Buffer | undefined {
  if (!invoiceNumber) return undefined;
  const hit = receipts.find((r) => r.name.includes(invoiceNumber));
  return hit?.buffer;
}

// ===================================================================
// Date snapping — nearest 15th or end-of-month
// ===================================================================

function snapDateTo15OrEom(d: Date): Date {
  const year = d.getFullYear();
  const month = d.getMonth();
  const day = d.getDate();

  const fifteenth = new Date(year, month, 15);
  const eom = new Date(year, month + 1, 0);

  const dayMs = 24 * 60 * 60 * 1000;
  const distTo15 = Math.abs((d.getTime() - fifteenth.getTime()) / dayMs);
  const distToEom = Math.abs((d.getTime() - eom.getTime()) / dayMs);

  if (day <= 7) {
    const prevEom = new Date(year, month, 0);
    const distToPrev = Math.abs((d.getTime() - prevEom.getTime()) / dayMs);
    if (distToPrev < distTo15) return prevEom;
  }

  return distTo15 <= distToEom ? fifteenth : eom;
}

function formatDateMDYY(d: Date): string {
  const m = d.getMonth() + 1;
  const day = d.getDate();
  const yy = String(d.getFullYear()).slice(-2);
  return `${m}.${day}.${yy}`;
}

// ===================================================================
// Filename builder
// ===================================================================

function buildFilename(titleNumber: string, snappedDate: Date): string {
  const cc = `CC ${titleNumber}`;
  const dateStr = formatDateMDYY(snappedDate);
  return `BOP Abstract - CNX Title AFE - ${cc} - Invoice - ${dateStr}.pdf`;
}

// ===================================================================
// PDF generation
// ===================================================================

function formatMoney(n: number): string {
  if (!isFinite(n)) return '$0.00';
  const sign = n < 0 ? '-' : '';
  const abs = Math.abs(n);
  return sign + '$' + abs.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatDays(n: number): string {
  if (!isFinite(n) || n === 0) return '0.000';
  return n.toFixed(3);
}

async function buildInvoicePdf(
  invoice: InvoiceData,
  emailRows: EmailRow[],
  receiptBuffer: Buffer | undefined,
  dateOverride: string,
  titleNumber: string
): Promise<ArrayBuffer> {
  const doc = new jsPDF({ unit: 'pt', format: 'letter', orientation: 'portrait' });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const margin = 36;

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(16);
  doc.text('BOP Abstract, LLC', pageW / 2, 50, { align: 'center' });
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(10);
  doc.text('2547 Washington Rd. Bldg. 700, Ste. 720', pageW / 2, 66, { align: 'center' });
  doc.text('Pittsburgh, PA, 15241', pageW / 2, 80, { align: 'center' });
  doc.text('724-747-1594', pageW / 2, 94, { align: 'center' });

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(18);
  doc.text('INVOICE', pageW / 2, 122, { align: 'center' });
  doc.setLineWidth(0.5);
  doc.line(margin, 132, pageW - margin, 132);

  const leftX = margin;
  const rightX = pageW / 2 + 10;
  let y = 152;

  doc.setFontSize(10);

  doc.setFont('helvetica', 'bold');
  doc.text('Date:', leftX, y);
  doc.setFont('helvetica', 'normal');
  doc.text(invoice.invoiceDate || '', leftX + 50, y);

  doc.setFont('helvetica', 'bold');
  doc.text('Invoice #:', rightX, y);
  doc.setFont('helvetica', 'normal');
  doc.text(invoice.invoiceNumber || '', rightX + 80, y);

  y += 18;

  doc.setFont('helvetica', 'bold');
  doc.text('Bill To:', leftX, y);
  doc.setFont('helvetica', 'normal');

  const billTo =
    invoice.billToLines.length > 0
      ? invoice.billToLines
      : [
          'CNX Land Resources, Inc.',
          'Attn: Danielle Kerr',
          '1000 Horizon Vue Drive',
          'Canonsburg, Pennsylvania 15317',
        ];
  let billToY = y;
  for (const line of billTo) {
    doc.text(line, leftX + 50, billToY);
    billToY += 13;
  }

  doc.setFont('helvetica', 'bold');
  doc.text('AFE:', rightX, y);
  doc.setFont('helvetica', 'normal');
  doc.text(invoice.afe || '', rightX + 80, y);
  y += 14;

  doc.setFont('helvetica', 'bold');
  doc.text('Company Code:', rightX, y);
  doc.setFont('helvetica', 'normal');
  doc.text(invoice.companyCode || '', rightX + 80, y);
  y += 14;

  doc.setFont('helvetica', 'bold');
  doc.text('Pad Site:', rightX, y);
  doc.setFont('helvetica', 'normal');
  doc.text(invoice.padSite || '', rightX + 80, y);
  y += 14;

  doc.setFont('helvetica', 'bold');
  doc.text('Project:', rightX, y);
  doc.setFont('helvetica', 'normal');
  doc.text(invoice.project || 'CNX Title', rightX + 80, y);
  y += 14;

  doc.setFont('helvetica', 'bold');
  doc.text('County:', rightX, y);
  doc.setFont('helvetica', 'normal');
  doc.text(invoice.county || '', rightX + 80, y);
  y += 14;

  const blockBottom = Math.max(billToY, y) + 6;

  doc.setFont('helvetica', 'bold');
  doc.text('Period:', leftX, blockBottom);
  doc.setFont('helvetica', 'normal');
  doc.text(invoice.period || '', leftX + 50, blockBottom);

  doc.setTextColor(200, 0, 0);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(12);
  doc.text('DUE UPON RECEIPT', pageW - margin, blockBottom, { align: 'right' });
  doc.setTextColor(0, 0, 0);
  doc.setFontSize(10);

  const brokerTableStartY = blockBottom + 22;

  const brokerHead = [[
    'Broker',
    'Days',
    'Amt. Per Day',
    'Total',
    'Professional Services',
    'Copies',
    'Miscellaneous',
    'TOTAL',
  ]];

  const brokerBody = invoice.brokerRows.map((r) => [
    r.broker,
    formatDays(r.days),
    formatMoney(r.amtPerDay),
    formatMoney(r.total),
    formatMoney(r.professionalServices),
    formatMoney(r.copies),
    formatMoney(r.miscellaneous),
    formatMoney(r.grandTotal),
  ]);

  if (invoice.brokerTotals) {
    brokerBody.push([
      'Totals',
      formatDays(invoice.brokerTotals.days),
      '',
      formatMoney(invoice.brokerTotals.total),
      formatMoney(invoice.brokerTotals.professionalServices),
      formatMoney(invoice.brokerTotals.copies),
      formatMoney(invoice.brokerTotals.miscellaneous),
      formatMoney(invoice.brokerTotals.grandTotal),
    ]);
  }

  autoTable(doc, {
    startY: brokerTableStartY,
    head: brokerHead,
    body: brokerBody,
    styles: { fontSize: 7, cellPadding: 2, overflow: 'linebreak' },
    headStyles: {
      fillColor: [10, 31, 68],
      textColor: 255,
      fontSize: 7,
      halign: 'center',
      valign: 'middle',
    },
    columnStyles: {
      0: { cellWidth: 78 },
      1: { halign: 'right', cellWidth: 38 },
      2: { halign: 'right', cellWidth: 55 },
      3: { halign: 'right', cellWidth: 60 },
      4: { halign: 'right', cellWidth: 65 },
      5: { halign: 'right', cellWidth: 45 },
      6: { halign: 'right', cellWidth: 60 },
      7: { halign: 'right', cellWidth: 65 },
    },
    didParseCell: (data) => {
      if (data.section === 'body') {
        const raw = data.row.raw as any;
        const firstCell = String(raw?.[0] ?? '').toLowerCase();
        if (firstCell === 'totals' || firstCell === 'total') {
          data.cell.styles.fillColor = [255, 245, 157];
          data.cell.styles.fontStyle = 'bold';
        }
      }
    },
    margin: { left: margin, right: margin },
    tableWidth: 'auto',
  });

  const finalY = (doc as any).lastAutoTable?.finalY ?? brokerTableStartY + 100;
  doc.setFont('helvetica', 'italic');
  doc.setFontSize(8);
  doc.text(
    'Please contact our accounting department with any questions regarding invoices.',
    pageW / 2,
    Math.min(finalY + 22, pageH - 30),
    { align: 'center' }
  );

  if (invoice.workRows.length > 0) {
    doc.addPage();

    const groups: { focus: string; rows: WorkDetailRow[] }[] = [];
    const groupIdx = new Map<string, number>();
    for (const r of invoice.workRows) {
      const key = r.projectFocus || '(unspecified)';
      let i = groupIdx.get(key);
      if (i == null) {
        i = groups.length;
        groupIdx.set(key, i);
        groups.push({ focus: key, rows: [] });
      }
      groups[i].rows.push(r);
    }

    const head = [[
      'Landman',
      'Date',
      'Prospect',
      'Legal',
      'Project Focus',
      'Days',
      'Copies',
      'Misc.',
      'Total',
    ]];

    const body: any[][] = [];
    let grandDays = 0;
    let grandCopies = 0;
    let grandMisc = 0;
    let grandTotal = 0;

    for (const g of groups) {
      let gDays = 0, gCopies = 0, gMisc = 0, gTotal = 0;
      for (const r of g.rows) {
        body.push([
          r.landman,
          r.date,
          r.prospect,
          r.legal,
          r.projectFocus,
          formatDays(r.days),
          formatMoney(r.copies),
          formatMoney(r.misc),
          formatMoney(r.total),
        ]);
        gDays += r.days;
        gCopies += r.copies;
        gMisc += r.misc;
        gTotal += r.total;
      }
      body.push([
        '',
        '',
        '',
        '',
        g.focus,
        formatDays(gDays),
        formatMoney(gCopies),
        formatMoney(gMisc),
        formatMoney(gTotal),
      ]);
      grandDays += gDays;
      grandCopies += gCopies;
      grandMisc += gMisc;
      grandTotal += gTotal;
    }

    body.push([
      '',
      '',
      '',
      '',
      'Grand Total',
      formatDays(grandDays),
      formatMoney(grandCopies),
      formatMoney(grandMisc),
      formatMoney(grandTotal),
    ]);

    autoTable(doc, {
      startY: 50,
      head,
      body,
      styles: { fontSize: 7, cellPadding: 2, overflow: 'linebreak' },
      headStyles: {
        fillColor: [10, 31, 68],
        textColor: 255,
        fontSize: 7,
        halign: 'center',
        valign: 'middle',
      },
      columnStyles: {
        0: { cellWidth: 60 },
        1: { cellWidth: 45 },
        2: { cellWidth: 90 },
        3: { cellWidth: 110 },
        4: { cellWidth: 55 },
        5: { halign: 'right', cellWidth: 35 },
        6: { halign: 'right', cellWidth: 40 },
        7: { halign: 'right', cellWidth: 40 },
        8: { halign: 'right', cellWidth: 50 },
      },
      didParseCell: (data) => {
        if (data.section === 'body') {
          const raw = data.row.raw as any;
          const landman = String(raw?.[0] ?? '');
          const focusCell = String(raw?.[4] ?? '');
          if (!landman && focusCell) {
            data.cell.styles.fillColor = [255, 245, 157];
            data.cell.styles.fontStyle = 'bold';
          }
        }
      },
      margin: { left: margin, right: margin },
      tableWidth: 'auto',
    });
  }

  let pdfBytes: ArrayBuffer = doc.output('arraybuffer');

  if (receiptBuffer) {
    const mainPdf = await PDFDocument.load(pdfBytes);
    const receiptPdf = await PDFDocument.load(receiptBuffer);
    const copied = await mainPdf.copyPages(receiptPdf, receiptPdf.getPageIndices());
    copied.forEach((p) => mainPdf.addPage(p));
    const saved = await mainPdf.save();
    const ab = new ArrayBuffer(saved.byteLength);
    new Uint8Array(ab).set(saved);
    pdfBytes = ab;
  }

  return pdfBytes;
}

// ===================================================================
// Helpers
// ===================================================================

function fallbackInvoiceNumberFromFilename(name: string): string {
  const m = name.match(/\d{4,6}/);
  return m ? m[0] : '';
}

function parseOverrideDate(s: string): Date | null {
  if (!s) return null;
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// ===================================================================
// POST handler
// ===================================================================

export async function POST(req: NextRequest) {
  try {
    const form = await req.formData();

    const excelFiles = form.getAll('excel') as File[];
    const emailFile = form.get('email') as File | null;
    const receiptFiles = form.getAll('receipt') as File[];
    const titleNumberRaw = (form.get('titleNumber') as string) || '';
    const dateOverrideRaw = (form.get('dateOverride') as string) || '';

    const titleNumber = titleNumberRaw.trim().replace(/\D/g, '');

    if (!excelFiles.length)
      return NextResponse.json({ error: 'No Excel files' }, { status: 400 });
    if (!emailFile)
      return NextResponse.json({ error: 'No email PDF' }, { status: 400 });
    if (!titleNumber)
      return NextResponse.json({ error: 'Missing CNX Title #' }, { status: 400 });

    const emailBuffer = Buffer.from(await emailFile.arrayBuffer());
    const parsedEmail = await parseEmailPdf(emailBuffer);

    const receipts = await Promise.all(
      receiptFiles.map(async (f) => ({
        name: f.name,
        buffer: Buffer.from(await f.arrayBuffer()),
      }))
    );

    const overrideDate = parseOverrideDate(dateOverrideRaw);

    const zip = new JSZip();

    const nameCounts = new Map<string, number>();

    for (const excelFile of excelFiles) {
      const buf = Buffer.from(await excelFile.arrayBuffer());
      const fallbackInv = fallbackInvoiceNumberFromFilename(excelFile.name);
      const invoice = parseExcel(buf, fallbackInv);

      const emailRows = invoice.afe ? (parsedEmail.byAfe.get(invoice.afe) ?? []) : [];

      const receipt = matchReceipt(invoice.invoiceNumber, receipts);

      const sourceDate = overrideDate ?? invoice.invoiceDateObj ?? new Date();
      const snappedDate = snapDateTo15OrEom(sourceDate);

      const pdfBytes = await buildInvoicePdf(
        invoice,
        emailRows,
        receipt,
        dateOverrideRaw,
        titleNumber
      );

      let filename = buildFilename(titleNumber, snappedDate);

      const seen = nameCounts.get(filename) ?? 0;
      if (seen > 0) {
        const base = filename.replace(/\.pdf$/i, '');
        filename = `${base} (${seen + 1}).pdf`;
      }
      nameCounts.set(buildFilename(titleNumber, snappedDate), seen + 1);

      zip.file(filename, pdfBytes);
    }

    const zipBytes = await zip.generateAsync({ type: 'uint8array' });
    const zipBlob = new Blob([zipBytes as BlobPart], { type: 'application/zip' });

    return new NextResponse(zipBlob, {
      status: 200,
      headers: {
        'Content-Type': 'application/zip',
        'Content-Disposition': `attachment; filename="cnx-title-invoices-CC-${titleNumber}.zip"`,
      },
    });
  } catch (err: any) {
    console.error('Generate error:', err);
    return NextResponse.json(
      { error: err.message || 'Internal server error' },
      { status: 500 }
    );
  }
}
