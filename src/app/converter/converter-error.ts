export type ConverterErrorCode = 'NOT_XLSX' | 'PARSE_FAILED' | 'SHEET_MISSING' | 'UNEXPECTED';

/**
 * The single error type that reaches the UI. The pure core throws plain `Error`s
 * with stable messages; `ConverterService` catches them and re-throws as this.
 */
export class ConverterError extends Error {
  constructor(
    readonly code: ConverterErrorCode,
    message: string,
  ) {
    super(message);
    this.name = 'ConverterError';
  }
}
