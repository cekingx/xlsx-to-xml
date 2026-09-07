// @vitest-environment node
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import ExcelJS from 'exceljs';
import { beforeEach, describe, expect, it } from 'vitest';

import { cell, convertWorkbook, countInvoices, escapeXml, formatInvoiceDate } from './xlsx-to-xml';

const docsPath = (name: string) => resolve(process.cwd(), 'docs', name);
const SAMPLE_XLSX = docsPath('STT BPI 31082026-02092026.xlsx');
const SAMPLE_XML = docsPath('STT BPI 31082026-02092026.xml');

const readArrayBuffer = (path: string): ArrayBuffer => new Uint8Array(readFileSync(path)).buffer;

async function loadSample(): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(readArrayBuffer(SAMPLE_XLSX));
  return wb;
}

function firstDiff(actual: string, expected: string): string {
  const a = actual.split('\n');
  const e = expected.split('\n');
  for (let i = 0; i < Math.max(a.length, e.length); i++) {
    if (a[i] !== e[i]) {
      return `line ${i + 1}\n  actual:   ${JSON.stringify(a[i])}\n  expected: ${JSON.stringify(e[i])}`;
    }
  }
  return 'no line difference (length only)';
}

describe('convertWorkbook — golden parity', () => {
  it('reproduces docs/STT BPI 31082026-02092026.xml byte-for-byte', async () => {
    const actual = convertWorkbook(await loadSample());
    const expected = readFileSync(SAMPLE_XML, 'utf8');
    if (actual !== expected) throw new Error(firstDiff(actual, expected));
    expect(actual).toBe(expected);
  });

  it('produces 39 invoices for the sample', async () => {
    expect(countInvoices(await loadSample())).toBe(39);
  });

  for (const tz of ['Asia/Jakarta', 'America/New_York', 'UTC']) {
    it(`is identical under TZ=${tz}`, async () => {
      const original = process.env['TZ'];
      process.env['TZ'] = tz;
      try {
        const actual = convertWorkbook(await loadSample());
        expect(actual).toBe(readFileSync(SAMPLE_XML, 'utf8'));
      } finally {
        process.env['TZ'] = original;
      }
    });
  }
});

describe('cell()', () => {
  it('maps empty values to an empty string', () => {
    expect(cell(null)).toBe('');
    expect(cell(undefined)).toBe('');
    expect(cell('')).toBe('');
  });

  it('renders whole numbers without a decimal point', () => {
    expect(cell(12)).toBe('12');
    expect(cell(0)).toBe('0');
    expect(cell(2750)).toBe('2750');
  });

  it('renders fractional numbers with two decimals', () => {
    expect(cell(15990.99)).toBe('15990.99');
    expect(cell(5.5)).toBe('5.50');
    expect(cell(5.1)).toBe('5.10');
  });

  it('passes text through verbatim, including "0.00"', () => {
    expect(cell('0.00')).toBe('0.00');
    expect(cell('351801.80')).toBe('351801.80');
    expect(cell('National ID')).toBe('National ID');
  });

  it('unwraps formula, rich-text and hyperlink cell shapes', () => {
    expect(cell({ formula: 'A1+B1', result: 7 })).toBe('7');
    expect(cell({ richText: [{ text: 'AB' }, { text: 'CD' }] })).toBe('ABCD');
    expect(cell({ text: 'label', hyperlink: 'https://x' })).toBe('label');
  });
});

describe('escapeXml()', () => {
  it('escapes &, <, > with & first', () => {
    expect(escapeXml('a & <b> c')).toBe('a &amp; &lt;b&gt; c');
  });

  it('leaves quotes untouched', () => {
    expect(escapeXml(`"x" 'y'`)).toBe(`"x" 'y'`);
  });
});

describe('formatInvoiceDate()', () => {
  it('formats a Date as zero-padded YYYY-MM-DD from UTC fields', () => {
    expect(formatInvoiceDate(new Date('2026-09-01T08:00:36.000Z'))).toBe('2026-09-01');
    expect(formatInvoiceDate(new Date('2026-01-05T23:59:59.000Z'))).toBe('2026-01-05');
  });

  it('takes the first ten characters of a string value', () => {
    expect(formatInvoiceDate('2026-08-31 08:00:36')).toBe('2026-08-31');
  });
});

describe('convertWorkbook — synthetic workbooks', () => {
  const FAKTUR_HEADERS = [
    'Baris',
    'Tanggal Faktur',
    'Jenis Faktur',
    'Kode Transaksi',
    'Keterangan Tambahan',
    'Dokumen Pendukung',
    'Referensi',
    'Cap Fasilitas',
    'ID TKU Penjual',
    'NPWP/NIK Pembeli',
    'Jenis ID Pembeli',
    'Negara Pembeli',
    'Nomor Dokumen Pembeli',
    'Nama Pembeli',
    'Alamat Pembeli',
    'Email Pembeli',
    'ID TKU Pembeli',
  ];
  const DETAIL_HEADERS = [
    'Baris',
    'Barang/Jasa',
    'Kode Barang Jasa',
    'Nama Barang/Jasa',
    'Nama Satuan Ukur',
    'Harga Satuan',
    'Jumlah Barang Jasa',
    'Total Diskon',
    'DPP',
    'DPP Nilai Lain',
    'Tarif PPN',
    'PPN',
    'Tarif PPnBM',
    'PPnBM',
  ];

  let wb: ExcelJS.Workbook;
  let faktur: ExcelJS.Worksheet;
  let detail: ExcelJS.Worksheet;

  const fakturRow = (baris: number | null) => [
    baris,
    new Date('2026-09-01T08:00:36.000Z'),
    'Normal',
    '04',
    null,
    null,
    'ref',
    null,
    'SELLER',
    '0000000000000000',
    'National ID',
    'IDN',
    'NIK',
    'BUYER',
    'ADDR',
    null,
    '000000',
  ];
  const detailRow = (baris: number | null) => [
    baris,
    'A',
    '271100',
    'ITEM',
    'UM.0022',
    '100.00',
    1,
    '0.00',
    '100.00',
    '90.00',
    12,
    '10.00',
    0,
    0,
  ];

  beforeEach(() => {
    wb = new ExcelJS.Workbook();
    faktur = wb.addWorksheet('Faktur');
    faktur.getRow(1).getCell(3).value = 'TIN123';
    faktur.getRow(3).values = FAKTUR_HEADERS;
    detail = wb.addWorksheet('DetailFaktur');
    detail.getRow(1).values = DETAIL_HEADERS;
  });

  it('skips a blank-Baris row in the middle of DetailFaktur without stopping', () => {
    faktur.getRow(4).values = fakturRow(1);
    detail.getRow(2).values = detailRow(1);
    detail.getRow(3).values = detailRow(null); // blank Baris — skipped
    detail.getRow(4).values = detailRow(1);

    const xml = convertWorkbook(wb);
    expect(xml.match(/<GoodService>/g)).toHaveLength(2);
  });

  it('emits an empty ListOfGoodService for an invoice with no details', () => {
    faktur.getRow(4).values = fakturRow(7);

    expect(convertWorkbook(wb)).toContain('      <ListOfGoodService>\n      </ListOfGoodService>');
  });

  it('treats a workbook with zero data rows as valid', () => {
    const xml = convertWorkbook(wb);
    expect(xml).toContain('<ListOfTaxInvoice>\n  </ListOfTaxInvoice>');
    expect(countInvoices(wb)).toBe(0);
  });

  it('throws when the Faktur sheet is missing', () => {
    const bare = new ExcelJS.Workbook();
    bare.addWorksheet('DetailFaktur');
    expect(() => convertWorkbook(bare)).toThrow('Sheet "Faktur" not found');
  });

  it('throws when the DetailFaktur sheet is missing', () => {
    const bare = new ExcelJS.Workbook();
    bare.addWorksheet('Faktur');
    expect(() => convertWorkbook(bare)).toThrow('Sheet "DetailFaktur" not found');
  });
});
