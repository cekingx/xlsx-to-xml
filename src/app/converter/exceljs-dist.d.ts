/**
 * The app imports the pre-built browser bundle so Angular's esbuild pipeline
 * never tries to resolve ExcelJS's Node entry (which pulls `stream`/`fs`). That
 * path has no bundled types, so alias it to the package's own declarations.
 */
declare module 'exceljs/dist/exceljs.min.js' {
  import ExcelJS = require('exceljs');
  export = ExcelJS;
}
