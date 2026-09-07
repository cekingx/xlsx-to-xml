import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs/dist/exceljs.min.js';

import { ConverterError } from './converter-error';
import { convertWorkbook, countInvoices } from './xlsx-to-xml';

const OOXML_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

export interface ConversionResult {
  xml: string;
  filename: string;
  invoiceCount: number;
  byteSize: number;
}

/**
 * Bridges the browser file world and the pure converter core: reads bytes,
 * drives ExcelJS, calls the core, derives the output filename, and translates
 * every failure into a {@link ConverterError}.
 */
@Injectable({ providedIn: 'root' })
export class ConverterService {
  async convert(file: File): Promise<ConversionResult> {
    if (!/\.xlsx$/i.test(file.name) && file.type !== OOXML_MIME) {
      throw new ConverterError('NOT_XLSX', 'Please choose an .xlsx file.');
    }

    let workbook: ExcelJS.Workbook;
    try {
      workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await file.arrayBuffer());
    } catch {
      throw new ConverterError(
        'PARSE_FAILED',
        "Could not read this file as an Excel workbook. Make sure it's a valid .xlsx.",
      );
    }

    try {
      const xml = convertWorkbook(workbook);
      return {
        xml,
        filename: toXmlFilename(file.name),
        invoiceCount: countInvoices(workbook),
        byteSize: new Blob([xml]).size,
      };
    } catch (err) {
      if (err instanceof Error && err.message.startsWith('Sheet "')) {
        throw new ConverterError(
          'SHEET_MISSING',
          `${err.message} — is this the DJP bulk template?`,
        );
      }
      throw new ConverterError('UNEXPECTED', 'Something went wrong during conversion.');
    }
  }
}

/** `STT BPI ....xlsx` -> `STT BPI ....xml`; `report.v2.xlsx` -> `report.v2.xml`. */
function toXmlFilename(name: string): string {
  return name.replace(/\.[^.]+$/, '') + '.xml';
}
