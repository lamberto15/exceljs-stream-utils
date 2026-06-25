import ExcelJS from '@protobi/exceljs';

import { DateTime } from 'luxon';
import { Readable } from 'node:stream';
import type { Writable } from 'node:stream';
import type { ReadableStream as NodeWebReadableStream } from 'node:stream/web';

export type ExcelInput = string | Readable | NodeWebReadableStream<Uint8Array>;
export type ExcelOutput = string | Writable;
export type RowObject = Record<string, unknown>;

type MaybeAsyncIterable<T> = Iterable<T> | AsyncIterable<T>;

/**
 * Options for reading worksheet rows as plain JavaScript objects.
 *
 * The `/v2` reader adds duplicate-header handling so imports fail loudly by
 * default instead of silently overwriting earlier columns.
 */
export interface XlsxReadOptions {
  /** Read only the matching worksheet name. Defaults to the first emitted sheet. */
  sheetName?: string;
  /** 1-based row number used as the header row. Defaults to `1`. */
  headerRowNumber?: number;
  /** Trim surrounding whitespace from header text before using it as an object key. */
  trimHeaders?: boolean;
  /** Trim all string cell values. */
  trimTextValues?: boolean;
  /** Trim only these columns when `trimTextValues` is `false`. */
  trimTextColumns?: ReadonlyArray<string>;
  /** Convert these columns from delimited strings into arrays. */
  arrayColumns?: ReadonlyArray<string>;
  /** Delimiter used when splitting `arrayColumns`. Defaults to `,`. */
  arrayDelimiter?: string;
  /** Trim whitespace from each split array item. */
  trimArrayItems?: boolean;
  /** Drop empty items after splitting array columns. */
  removeEmptyArrayItems?: boolean;
  /** Skip rows where every cell is empty. */
  skipEmptyRows?: boolean;
  /** Transform each header before using it as an object key. */
  normalizeHeader?: (header: string, index: number) => string;
  /** Parse date-like numeric Excel cells into `Date` objects. */
  parseDates?: boolean;
  /** Interpret numeric dates using Excel's 1904 date system. */
  date1904?: boolean;
  /** Reinterpret parsed dates in the provided IANA time zone. */
  timeZone?: string;
  /** Fail on duplicate headers or suffix them as `_2`, `_3`, and so on. */
  duplicateHeaders?: 'error' | 'suffix';
}

/**
 * Column metadata used when writing rows to a worksheet.
 */
export interface XlsxWriteColumn<T extends RowObject = RowObject> {
  /** Header text written to the worksheet. */
  header: string;
  /** Object key to read from each row. */
  key: keyof T & string;
  /** Optional worksheet column width. */
  width?: number;
}

/**
 * Options for writing objects to a worksheet.
 */
export interface XlsxWriteOptions<T extends RowObject = RowObject> {
  /** Worksheet name. Defaults to `Sheet1`. */
  sheetName?: string;
  /** Explicit column order and labels. If omitted, columns are inferred from the first row. */
  columns?: ReadonlyArray<XlsxWriteColumn<T>>;
  /** Enable ExcelJS style support while streaming writes. */
  useStyles?: boolean;
  /** Enable ExcelJS shared string support while streaming writes. */
  useSharedStrings?: boolean;
  /** Convert written `Date` values into the provided IANA time zone wall-clock. */
  timeZone?: string;
  /** Limit time zone conversion to these date columns. Empty means all `Date` fields. */
  dateColumns?: ReadonlyArray<keyof T & string>;
}

/**
 * Options for processing large worksheets in bounded batches.
 */
export interface ProcessLargeOptions extends XlsxReadOptions {
  /** Number of rows to collect before processing a batch. */
  batchSize?: number;
  /** Maximum number of row handlers running at once inside a batch. */
  concurrency?: number;
}

function isAsyncIterable<T>(value: unknown): value is AsyncIterable<T> {
  return (
    typeof value === 'object' &&
    value !== null &&
    Symbol.asyncIterator in value &&
    typeof (value as AsyncIterable<T>)[Symbol.asyncIterator] === 'function'
  );
}

function isWritableOutput(output: ExcelOutput): output is Writable {
  return typeof output !== 'string';
}

async function waitForWritableDrain(output: Writable): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    const onDrain = () => {
      cleanup();
      resolve();
    };

    const onError = (error: Error) => {
      cleanup();
      reject(error);
    };

    const cleanup = () => {
      output.off('drain', onDrain);
      output.off('error', onError);
    };

    output.on('drain', onDrain);
    output.on('error', onError);
  });
}

async function* toAsyncIterable<T>(
  rows: MaybeAsyncIterable<T>,
): AsyncIterable<T> {
  if (isAsyncIterable<T>(rows)) {
    yield* rows;
    return;
  }

  for (const row of rows) {
    yield row;
  }
}

function isEmptyRow(values: unknown[]): boolean {
  return values.every((value) => value == null || value === '');
}

function getRowValues(row: ExcelJS.Row): unknown[] {
  const values = row.values;
  if (!Array.isArray(values)) {
    return [];
  }

  return values.slice(1);
}

function getWorksheetName(
  worksheet: ExcelJS.stream.xlsx.WorksheetReader,
): string | undefined {
  const candidate = (worksheet as { name?: unknown }).name;
  return typeof candidate === 'string' ? candidate : undefined;
}

function excelSerialToDate(serial: number, date1904: boolean): Date {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  const msPerDay = 24 * 60 * 60 * 1000;
  return new Date(epoch + serial * msPerDay);
}

function reinterpretDateToTimeZone(date: Date, timeZone: string): Date {
  const reinterpreted = DateTime.fromObject(
    {
      year: date.getUTCFullYear(),
      month: date.getUTCMonth() + 1,
      day: date.getUTCDate(),
      hour: date.getUTCHours(),
      minute: date.getUTCMinutes(),
      second: date.getUTCSeconds(),
      millisecond: date.getUTCMilliseconds(),
    },
    { zone: timeZone },
  );

  if (!reinterpreted.isValid) {
    throw new Error(`Invalid time zone: ${timeZone}`);
  }

  return reinterpreted.toUTC().toJSDate();
}

function utcDateToExcelLocalDate(dateUtc: Date, timeZone: string): Date {
  const userDate = DateTime.fromJSDate(dateUtc, { zone: 'utc' }).setZone(
    timeZone,
  );

  if (!userDate.isValid) {
    throw new Error(`Invalid time zone: ${timeZone}`);
  }

  return new Date(
    userDate.year,
    userDate.month - 1,
    userDate.day,
    userDate.hour,
    userDate.minute,
    userDate.second,
    userDate.millisecond,
  );
}

function maybeTrimTextValue(value: unknown, trimTextValues: boolean): unknown {
  if (!trimTextValues || typeof value !== 'string') {
    return value;
  }

  return value.trim();
}

function shouldTrimColumn(
  trimTextValues: boolean,
  trimTextColumns: ReadonlySet<string>,
  columnName: string,
): boolean {
  if (trimTextValues) {
    return true;
  }

  return trimTextColumns.has(columnName);
}

function maybeConvertToArray(
  value: unknown,
  shouldConvert: boolean,
  delimiter: string,
  trimItems: boolean,
  removeEmpty: boolean,
): unknown {
  if (!shouldConvert) {
    return value;
  }

  if (value == null || value === '') {
    return [];
  }

  if (Array.isArray(value)) {
    return value;
  }

  if (typeof value !== 'string') {
    return [value];
  }

  let items = value.split(delimiter);

  if (trimItems) {
    items = items.map((item) => item.trim());
  }

  if (removeEmpty) {
    items = items.filter((item) => item.length > 0);
  }

  return items;
}

function shouldConvertDateColumn(
  columnName: string,
  dateColumnsSet: ReadonlySet<string>,
): boolean {
  if (dateColumnsSet.size === 0) {
    return true;
  }

  return dateColumnsSet.has(columnName);
}

function mapRowForExcelWrite<T extends RowObject>(
  row: T,
  timeZone?: string,
  dateColumnsSet: ReadonlySet<string> = new Set<string>(),
): RowObject {
  if (!timeZone) {
    return row;
  }

  const converted: RowObject = { ...row };

  for (const [columnName, value] of Object.entries(row)) {
    if (
      value instanceof Date &&
      shouldConvertDateColumn(columnName, dateColumnsSet)
    ) {
      converted[columnName] = utcDateToExcelLocalDate(value, timeZone);
    }
  }

  return converted;
}

function inferColumnsFromRow<T extends RowObject>(
  row: T,
): ReadonlyArray<XlsxWriteColumn<T>> {
  return (Object.keys(row) as Array<keyof T & string>).map((key) => ({
    header: key,
    key,
  }));
}

function toWorkbookInput(input: ExcelInput): string | Readable {
  if (typeof input === 'string' || input instanceof Readable) {
    return input;
  }

  return Readable.fromWeb(
    input as unknown as Parameters<typeof Readable.fromWeb>[0],
  );
}

function headerValueToString(value: unknown): string {
  if (value == null) {
    return '';
  }

  if (typeof value === 'string') {
    return value;
  }

  if (
    typeof value === 'number' ||
    typeof value === 'boolean' ||
    typeof value === 'bigint'
  ) {
    return `${value}`;
  }

  if (value instanceof Date) {
    return value.toISOString();
  }

  if (typeof value === 'object') {
    const withText = value as { text?: unknown };
    if (typeof withText.text === 'string') {
      return withText.text;
    }

    const withRichText = value as {
      richText?: Array<{ text?: unknown }>;
    };
    if (Array.isArray(withRichText.richText)) {
      return withRichText.richText
        .map((part) => (typeof part.text === 'string' ? part.text : ''))
        .join('');
    }

    const withResult = value as { result?: unknown };
    if (withResult.result !== undefined) {
      return headerValueToString(withResult.result);
    }
  }

  return '';
}

function normalizeCellRawValue(cell: ExcelJS.Cell): unknown {
  const value = cell.value;

  if (value == null) {
    return null;
  }

  if (
    typeof value === 'object' &&
    value !== null &&
    'result' in value
  ) {
    return (value as ExcelJS.CellFormulaValue).result ?? null;
  }

  return value;
}

function normalizeCellValue(
  cell: ExcelJS.Cell,
  parseDates: boolean,
  date1904: boolean,
  timeZone?: string,
): unknown {
  const value = normalizeCellRawValue(cell);

  if (value == null) {
    return null;
  }

  if (value instanceof Date) {
    return timeZone ? reinterpretDateToTimeZone(value, timeZone) : value;
  }

  if (
    parseDates &&
    typeof value === 'number' &&
    typeof cell.numFmt === 'string' &&
    /[dmyhs]/i.test(cell.numFmt)
  ) {
    const date = excelSerialToDate(value, date1904);
    return timeZone ? reinterpretDateToTimeZone(date, timeZone) : date;
  }

  return value;
}

function buildHeaders(
  values: unknown[],
  trimHeaders: boolean,
  normalizeHeader: (header: string, index: number) => string,
  duplicateHeaders: 'error' | 'suffix',
): string[] {
  const usedHeaders = new Map<string, number>();

  return Array.from({ length: values.length }, (_, index) => {
    const value = values[index];
    let header = headerValueToString(value);

    if (trimHeaders) {
      header = header.trim();
    }

    header = normalizeHeader(header, index) || `column_${index + 1}`;
    const seen = usedHeaders.get(header) ?? 0;

    if (seen === 0) {
      usedHeaders.set(header, 1);
      return header;
    }

    if (duplicateHeaders === 'error') {
      throw new Error(
        `Duplicate header "${header}" at column ${index + 1}.`,
      );
    }

    const uniqueHeader = `${header}_${seen + 1}`;
    usedHeaders.set(header, seen + 1);
    usedHeaders.set(uniqueHeader, 1);
    return uniqueHeader;
  });
}

async function waitForBatchSlot(
  inFlight: Set<Promise<void>>,
  errors: Error[],
): Promise<void> {
  while (inFlight.size > 0) {
    await Promise.race(inFlight);

    if (errors.length > 0) {
      await Promise.allSettled(inFlight);
      throw errors[0];
    }

    return;
  }
}

/**
 * Stream worksheet rows as plain objects using the selected header row for keys.
 *
 * Accepts a file path, a Node readable stream, or a Web readable stream.
 * In `/v2`, duplicate headers are handled explicitly instead of silently
 * overwriting earlier columns.
 */
export async function* readXlsxRows<T extends RowObject = RowObject>(
  input: ExcelInput,
  opts: XlsxReadOptions = {},
): AsyncGenerator<T> {
  const {
    sheetName,
    headerRowNumber = 1,
    trimHeaders = true,
    trimTextValues = false,
    trimTextColumns = [],
    arrayColumns = [],
    arrayDelimiter = ',',
    trimArrayItems = true,
    removeEmptyArrayItems = true,
    skipEmptyRows = true,
    parseDates = true,
    date1904 = false,
    timeZone,
    duplicateHeaders = 'error',
  } = opts;
  const trimTextColumnsSet = new Set(trimTextColumns);
  const arrayColumnsSet = new Set(arrayColumns);
  const normalizeHeader: (header: string, index: number) => string =
    opts.normalizeHeader ?? ((header) => header);

  const workbook = new ExcelJS.stream.xlsx.WorkbookReader(
    toWorkbookInput(input),
    {
      worksheets: 'emit',
      sharedStrings: 'cache',
      styles: parseDates ? 'cache' : 'ignore',
      hyperlinks: 'ignore',
      entries: 'emit',
    },
  );

  for await (const worksheet of workbook) {
    if (sheetName) {
      const currentSheetName = getWorksheetName(worksheet);
      if (currentSheetName !== sheetName) {
        continue;
      }
    }

    let headers: string[] | null = null;

    for await (const row of worksheet) {
      const values = getRowValues(row);

      if (skipEmptyRows && isEmptyRow(values)) {
        continue;
      }

      if (row.number === headerRowNumber) {
        headers = buildHeaders(
          values,
          trimHeaders,
          normalizeHeader,
          duplicateHeaders,
        );
        continue;
      }

      if (!headers) {
        continue;
      }

      const obj: RowObject = {};
      for (let i = 0; i < headers.length; i += 1) {
        const columnName = headers[i];
        const cell = row.getCell(i + 1);
        const normalizedValue = normalizeCellValue(
          cell,
          parseDates,
          date1904,
          timeZone,
        );
        const trimmedValue = maybeTrimTextValue(
          normalizedValue,
          shouldTrimColumn(trimTextValues, trimTextColumnsSet, columnName),
        );
        obj[columnName] = maybeConvertToArray(
          trimmedValue,
          arrayColumnsSet.has(columnName),
          arrayDelimiter,
          trimArrayItems,
          removeEmptyArrayItems,
        );
      }

      yield obj as T;
    }
  }
}

/**
 * Write iterable or async-iterable row objects to an `.xlsx` file or writable stream.
 *
 * When `columns` is omitted, the first row is used to infer column order and headers.
 * In `/v2`, the options object is optional for the common case.
 */
export async function writeXlsxRows<T extends RowObject>(
  output: ExcelOutput,
  rows: MaybeAsyncIterable<T>,
  opts: XlsxWriteOptions<T> = {},
): Promise<void> {
  const {
    sheetName = 'Sheet1',
    columns,
    useStyles = false,
    useSharedStrings = false,
    timeZone,
    dateColumns = [],
  } = opts;
  const dateColumnsSet = new Set<string>(dateColumns);
  const writableOutput = isWritableOutput(output) ? output : null;
  const rowIterator = toAsyncIterable(rows)[Symbol.asyncIterator]();
  const firstRowResult = await rowIterator.next();

  const resolvedColumns =
    columns && columns.length > 0
      ? columns
      : firstRowResult.done
        ? null
        : inferColumnsFromRow(firstRowResult.value);

  if (!resolvedColumns || resolvedColumns.length === 0) {
    throw new Error(
      'Unable to determine columns. Provide `columns` or include at least one row with keys.',
    );
  }

  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(
    typeof output === 'string'
      ? { filename: output, useStyles, useSharedStrings }
      : { stream: output, useStyles, useSharedStrings },
  );

  const sheet = workbook.addWorksheet(sheetName);
  sheet.columns = resolvedColumns.map((column) => ({
    header: column.header,
    key: column.key,
    width: column.width,
  }));

  if (!firstRowResult.done) {
    sheet
      .addRow(
        mapRowForExcelWrite(firstRowResult.value, timeZone, dateColumnsSet),
      )
      .commit();

    if (writableOutput?.writableNeedDrain) {
      await waitForWritableDrain(writableOutput);
    }
  }

  while (true) {
    const rowResult = await rowIterator.next();
    if (rowResult.done) {
      break;
    }

    sheet
      .addRow(mapRowForExcelWrite(rowResult.value, timeZone, dateColumnsSet))
      .commit();

    if (writableOutput?.writableNeedDrain) {
      await waitForWritableDrain(writableOutput);
    }
  }

  sheet.commit();
  await workbook.commit();
}

/**
 * Read a large worksheet and process rows in batches with bounded concurrency.
 *
 * `/v2` drains in-flight handlers before surfacing the first batch error so
 * callers do not end up with stray unhandled rejections.
 */
export async function processXlsxRows<T extends RowObject = RowObject>(
  input: ExcelInput,
  handler: (row: T) => Promise<void> | void,
  opts: ProcessLargeOptions = {},
): Promise<void> {
  const { batchSize = 2000, concurrency = 8, ...readOpts } = opts;

  if (batchSize <= 0) {
    throw new Error('batchSize must be > 0');
  }

  if (concurrency <= 0) {
    throw new Error('concurrency must be > 0');
  }

  let batch: T[] = [];

  const runBatch = async (items: T[]): Promise<void> => {
    const inFlight = new Set<Promise<void>>();
    const errors: Error[] = [];

    for (const item of items) {
      const promise = Promise.resolve()
        .then(() => handler(item))
        .catch((error: unknown) => {
          errors.push(
            error instanceof Error ? error : new Error(String(error)),
          );
        });

      inFlight.add(promise);
      void promise.finally(() => {
        inFlight.delete(promise);
      });

      if (inFlight.size >= concurrency) {
        await waitForBatchSlot(inFlight, errors);
      }
    }

    await Promise.allSettled(inFlight);

    if (errors.length > 0) {
      throw errors[0];
    }
  };

  for await (const row of readXlsxRows<T>(input, readOpts)) {
    batch.push(row);

    if (batch.length >= batchSize) {
      await runBatch(batch);
      batch = [];
    }
  }

  if (batch.length > 0) {
    await runBatch(batch);
  }
}
