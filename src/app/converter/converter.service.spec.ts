import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { TestBed } from '@angular/core/testing';
import { beforeEach, describe, expect, it } from 'vitest';

import { ConverterError } from './converter-error';
import { ConverterService } from './converter.service';

const docsPath = (name: string) => resolve(process.cwd(), 'docs', name);

function fileFrom(diskName: string, asName = diskName): File {
  const bytes = new Uint8Array(readFileSync(docsPath(diskName))).buffer;
  return new File([bytes], asName, {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

describe('ConverterService', () => {
  let service: ConverterService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(ConverterService);
  });

  it('converts the sample workbook and reports its metadata', async () => {
    const result = await service.convert(fileFrom('STT BPI 31082026-02092026.xlsx'));

    expect(result.filename).toBe('STT BPI 31082026-02092026.xml');
    expect(result.invoiceCount).toBe(39);
    expect(result.xml).toBe(readFileSync(docsPath('STT BPI 31082026-02092026.xml'), 'utf8'));
    expect(result.byteSize).toBe(87151);
  });

  it('derives the output name by swapping the final extension', async () => {
    const result = await service.convert(
      fileFrom('STT BPI 31082026-02092026.xlsx', 'report.v2.xlsx'),
    );
    expect(result.filename).toBe('report.v2.xml');
  });

  it('rejects a non-xlsx file with code NOT_XLSX', async () => {
    const txt = new File(['nope'], 'notes.txt', { type: 'text/plain' });
    await expect(service.convert(txt)).rejects.toMatchObject({
      code: 'NOT_XLSX',
    });
  });

  it('rejects an unreadable .xlsx with code PARSE_FAILED', async () => {
    const junk = new File([new Uint8Array([1, 2, 3, 4])], 'broken.xlsx', {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const err = await service.convert(junk).catch((e) => e);
    expect(err).toBeInstanceOf(ConverterError);
    expect(err.code).toBe('PARSE_FAILED');
  });
});
