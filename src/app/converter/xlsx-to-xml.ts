import type { Workbook } from 'exceljs';

/**
 * Pure, framework-free port of `docs/script.py`.
 *
 * The contract is byte-for-byte parity with `docs/STT BPI 31082026-02092026.xml`,
 * enforced by the golden test in `xlsx-to-xml.spec.ts`. Nothing in this file may
 * import Angular, touch the DOM, read the filesystem, or depend on the machine's
 * timezone/locale.
 */

const FAKTUR_SHEET = 'Faktur';
const DETAIL_SHEET = 'DetailFaktur';

/** First `Faktur` data row (rows 1-3 are the TIN line, a blank, and the header). */
const FAKTUR_FIRST_DATA_ROW = 4;
/** First `DetailFaktur` data row (row 1 is the header). */
const DETAIL_FIRST_DATA_ROW = 2;

/**
 * Convert one ExcelJS cell value to its output string, matching `script.py`'s
 * `cell()`:
 *
 *   - `None`/empty          -> ""
 *   - float equal to its int -> "5"      (integer, no decimal point)
 *   - any other float        -> "5.50"   (two decimals)
 *   - anything else          -> str(v)   (verbatim)
 *
 * ExcelJS has no int/float distinction, so a whole number renders bare and a
 * fractional number renders with two decimals. Text cells (which is how the
 * sample stores every monetary column) pass straight through, so the literal
 * `"0.00"` in `Total Diskon` stays `"0.00"`.
 */
export function cell(value: unknown): string {
  if (value === null || value === undefined || value === '') return '';

  if (typeof value === 'number') {
    if (!Number.isFinite(value)) return String(value);
    return Number.isInteger(value) ? String(value) : value.toFixed(2);
  }

  if (typeof value === 'boolean') return value ? 'True' : 'False';

  if (value instanceof Date) return formatInvoiceDate(value);

  if (typeof value === 'object') {
    const v = value as Record<string, unknown>;
    if ('result' in v) return cell(v['result']); // formula cell
    if ('error' in v) return String(v['error']); // error cell (#N/A, ...)
    if (Array.isArray(v['richText'])) {
      return (v['richText'] as { text?: unknown }[])
        .map((run) => (run.text === undefined ? '' : String(run.text)))
        .join('');
    }
    if ('text' in v) return String(v['text']); // hyperlink cell
  }

  return String(value);
}

/**
 * Port of `xml.sax.saxutils.escape`: escape `&`, `<`, `>` for element text
 * content, `&` first. Quotes are intentionally left alone.
 */
export function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Format the `Tanggal Faktur` value as `YYYY-MM-DD`.
 *
 * ExcelJS builds date cells on a UTC epoch base, so the Excel wall-clock date
 * lands in the Date's UTC fields — reading them with `getUTC*` keeps the output
 * independent of the machine's timezone. A string value falls back to its first
 * ten characters, mirroring `script.py`'s `cell(v)[:10]`.
 */
export function formatInvoiceDate(value: unknown): string {
  if (value instanceof Date) {
    const pad = (n: number) => String(n).padStart(2, '0');
    return (
      `${value.getUTCFullYear()}-` +
      `${pad(value.getUTCMonth() + 1)}-` +
      `${pad(value.getUTCDate())}`
    );
  }
  if (typeof value === 'string') return value.slice(0, 10);
  return cell(value).slice(0, 10);
}

/** Convert a loaded workbook to the finished XML string. */
export function convertWorkbook(wb: Workbook): string {
  const faktur = wb.getWorksheet(FAKTUR_SHEET);
  if (!faktur) throw new Error(`Sheet "${FAKTUR_SHEET}" not found`);
  const detail = wb.getWorksheet(DETAIL_SHEET);
  if (!detail) throw new Error(`Sheet "${DETAIL_SHEET}" not found`);

  // TIN: NPWP Penjual, Faktur cell C1.
  const tin = cell(faktur.getRow(1).getCell(3).value);

  // Faktur data rows; skip (do not stop on) rows with an empty Baris.
  const fakturRows = [];
  for (let r = FAKTUR_FIRST_DATA_ROW; r <= faktur.rowCount; r++) {
    const row = faktur.getRow(r);
    if (row.getCell(1).value == null) continue;
    fakturRows.push(row);
  }

  // DetailFaktur rows grouped by Baris.
  const detailByBaris = new Map<unknown, ReturnType<typeof detail.getRow>[]>();
  for (let r = DETAIL_FIRST_DATA_ROW; r <= detail.rowCount; r++) {
    const row = detail.getRow(r);
    const baris = row.getCell(1).value;
    if (baris == null) continue;
    const key = normaliseKey(baris);
    const group = detailByBaris.get(key);
    if (group) group.push(row);
    else detailByBaris.set(key, [row]);
  }

  const lines: string[] = [];
  lines.push('<?xml version="1.0" encoding="utf-8"?>');
  lines.push(
    '<TaxInvoiceBulk xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="TaxInvoice.xsd">',
  );
  lines.push(`  <TIN>${escapeXml(tin)}</TIN>`);
  lines.push('  <ListOfTaxInvoice>');

  for (const row of fakturRows) {
    const f = (col: number) => row.getCell(col).value;
    const baris = normaliseKey(f(1));

    lines.push('    <TaxInvoice>');
    lines.push(`      <TaxInvoiceDate>${escapeXml(formatInvoiceDate(f(2)))}</TaxInvoiceDate>`);
    lines.push(`      <TaxInvoiceOpt>${escapeXml(cell(f(3)))}</TaxInvoiceOpt>`);
    lines.push(`      <TrxCode>${escapeXml(cell(f(4)))}</TrxCode>`);
    lines.push(`      <AddInfo>${escapeXml(cell(f(5)))}</AddInfo>`);
    lines.push(`      <CustomDoc>${escapeXml(cell(f(6)))}</CustomDoc>`);
    lines.push(`      <RefDesc>${escapeXml(cell(f(7)))}</RefDesc>`);
    lines.push(`      <FacilityStamp>${escapeXml(cell(f(8)))}</FacilityStamp>`);
    lines.push(`      <SellerIDTKU>${escapeXml(cell(f(9)))}</SellerIDTKU>`);
    lines.push(`      <BuyerTin>${escapeXml(cell(f(10)))}</BuyerTin>`);
    lines.push(`      <BuyerDocument>${escapeXml(cell(f(11)))}</BuyerDocument>`);
    lines.push(`      <BuyerCountry>${escapeXml(cell(f(12)))}</BuyerCountry>`);
    lines.push(`      <BuyerDocumentNumber>${escapeXml(cell(f(13)))}</BuyerDocumentNumber>`);
    lines.push(`      <BuyerName>${escapeXml(cell(f(14)))}</BuyerName>`);
    lines.push(`      <BuyerAdress>${escapeXml(cell(f(15)))}</BuyerAdress>`);
    lines.push(`      <BuyerEmail>${escapeXml(cell(f(16)))}</BuyerEmail>`);
    lines.push(`      <BuyerIDTKU>${escapeXml(cell(f(17)))}</BuyerIDTKU>`);
    lines.push('      <ListOfGoodService>');

    for (const drow of detailByBaris.get(baris) ?? []) {
      const d = (col: number) => drow.getCell(col).value;
      lines.push('        <GoodService>');
      lines.push(`          <Opt>${escapeXml(cell(d(2)))}</Opt>`);
      lines.push(`          <Code>${escapeXml(cell(d(3)))}</Code>`);
      lines.push(`          <Name>${escapeXml(cell(d(4)))}</Name>`);
      lines.push(`          <Unit>${escapeXml(cell(d(5)))}</Unit>`);
      lines.push(`          <Price>${escapeXml(cell(d(6)))}</Price>`);
      lines.push(`          <Qty>${escapeXml(cell(d(7)))}</Qty>`);
      lines.push(`          <TotalDiscount>${escapeXml(cell(d(8)))}</TotalDiscount>`);
      lines.push(`          <TaxBase>${escapeXml(cell(d(9)))}</TaxBase>`);
      lines.push(`          <OtherTaxBase>${escapeXml(cell(d(10)))}</OtherTaxBase>`);
      lines.push(`          <VATRate>${escapeXml(cell(d(11)))}</VATRate>`);
      lines.push(`          <VAT>${escapeXml(cell(d(12)))}</VAT>`);
      lines.push(`          <STLGRate>${escapeXml(cell(d(13)))}</STLGRate>`);
      lines.push(`          <STLG>${escapeXml(cell(d(14)))}</STLG>`);
      lines.push('        </GoodService>');
    }

    lines.push('      </ListOfGoodService>');
    lines.push('    </TaxInvoice>');
  }

  lines.push('  </ListOfTaxInvoice>');
  lines.push('</TaxInvoiceBulk>');

  return lines.join('\n') + '\n';
}

/** Count of `<TaxInvoice>` blocks a workbook would produce, for the UI. */
export function countInvoices(wb: Workbook): number {
  const faktur = wb.getWorksheet(FAKTUR_SHEET);
  if (!faktur) return 0;
  let n = 0;
  for (let r = FAKTUR_FIRST_DATA_ROW; r <= faktur.rowCount; r++) {
    if (faktur.getRow(r).getCell(1).value != null) n++;
  }
  return n;
}

/**
 * Group key for `Baris`. ExcelJS may hand back the join value as a number on one
 * sheet and (defensively) a string on the other; compare them as strings so the
 * grouping still lines up.
 */
function normaliseKey(value: unknown): string {
  return typeof value === 'number' ? String(value) : String(value ?? '');
}
