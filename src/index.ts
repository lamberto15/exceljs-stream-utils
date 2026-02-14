import ExcelJS from 'exceljs';
import { Readable } from 'node:stream';
import type { Writable } from 'node:stream';
import type { ReadableStream as NodeWebReadableStream } from 'node:stream/web';

export type ExcelInput = string | Readable | NodeWebReadableStream<Uint8Array>;
export type ExcelOutput = string | Writable;
export type RowObject = Record<string, unknown>;

type MaybeAsyncIterable<T> = Iterable<T> | AsyncIterable<T>;

type ReaderCacheMode = 'cache' | 'emit' | 'ignore';
type ReaderStylesMode = 'cache' | 'ignore';

interface DateTimeParts {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
  millisecond: number;
}

export interface XlsxReadOptions {
  sheetName?: string;
  headerRowNumber?: number;
  trimHeaders?: boolean;
  trimTextValues?: boolean;
  trimTextColumns?: ReadonlyArray<string>;
  arrayColumns?: ReadonlyArray<string>;
  arrayDelimiter?: string;
  trimArrayItems?: boolean;
  removeEmptyArrayItems?: boolean;
  skipEmptyRows?: boolean;
  normalizeHeader?: (header: string, index: number) => string;
  parseDates?: boolean;
  date1904?: boolean;
  timeZone?: string;
  sharedStringsMode?: ReaderCacheMode;
  stylesMode?: ReaderStylesMode;
}

export interface XlsxWriteColumn<T extends RowObject = RowObject> {
  header: string;
  key: keyof T & string;
  width?: number;
}

export interface XlsxWriteOptions<T extends RowObject = RowObject> {
  sheetName?: string;
  columns?: ReadonlyArray<XlsxWriteColumn<T>>;
  useStyles?: boolean;
  useSharedStrings?: boolean;
  timeZone?: string;
  dateColumns?: ReadonlyArray<keyof T & string>;
}

export interface ProcessLargeOptions extends XlsxReadOptions {
  batchSize?: number;
  concurrency?: number;
}

const formatterCache = new Map<string, Intl.DateTimeFormat>();

function isAsyncIterable<T>(value: unknown): value is AsyncIterable<T> {
  return (
    typeof value === 'object' &&
    value !== null &&
    Symbol.asyncIterator in value &&
    typeof (value as AsyncIterable<T>)[Symbol.asyncIterator] === 'function'
  );
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

function assertValidDate(date: Date, label: string): void {
  if (Number.isNaN(date.getTime())) {
    throw new Error(`Invalid ${label}`);
  }
}

function assertValidTimeZone(timeZone: string): void {
  try {
    new Intl.DateTimeFormat('en-US', { timeZone }).format(new Date());
  } catch {
    throw new Error(`Invalid time zone: ${timeZone}`);
  }
}

function getTimeZoneFormatter(timeZone: string): Intl.DateTimeFormat {
  const cached = formatterCache.get(timeZone);
  if (cached) {
    return cached;
  }

  const formatter = new Intl.DateTimeFormat('en-CA', {
    timeZone,
    hour12: false,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  });

  formatterCache.set(timeZone, formatter);
  return formatter;
}

function getTimeZoneParts(date: Date, timeZone: string): DateTimeParts {
  const formatter = getTimeZoneFormatter(timeZone);
  const parts = formatter.formatToParts(date);
  const values = new Map<string, string>();

  for (const part of parts) {
    if (part.type !== 'literal') {
      values.set(part.type, part.value);
    }
  }

  const year = Number(values.get('year'));
  const month = Number(values.get('month'));
  const day = Number(values.get('day'));
  const hour = Number(values.get('hour'));
  const minute = Number(values.get('minute'));
  const second = Number(values.get('second'));

  if (
    [year, month, day, hour, minute, second].some((value) =>
      Number.isNaN(value),
    )
  ) {
    throw new Error(`Unable to parse date parts for time zone: ${timeZone}`);
  }

  return {
    year,
    month,
    day,
    hour,
    minute,
    second,
    millisecond: date.getUTCMilliseconds(),
  };
}

function toUtcMillis(parts: DateTimeParts): number {
  return Date.UTC(
    parts.year,
    parts.month - 1,
    parts.day,
    parts.hour,
    parts.minute,
    parts.second,
    parts.millisecond,
  );
}

function getTimeZoneOffsetMs(date: Date, timeZone: string): number {
  const zoned = getTimeZoneParts(date, timeZone);
  return toUtcMillis(zoned) - date.getTime();
}

function localDateTimeInZoneToUtc(
  parts: DateTimeParts,
  timeZone: string,
): Date {
  const localAsUtc = toUtcMillis(parts);
  let guess = localAsUtc;

  for (let attempt = 0; attempt < 4; attempt += 1) {
    const offset = getTimeZoneOffsetMs(new Date(guess), timeZone);
    const next = localAsUtc - offset;
    if (next === guess) {
      break;
    }

    guess = next;
  }

  return new Date(guess);
}

function replaceDateTimeZoneByUserTimeZone(date: Date, timeZone: string): Date {
  assertValidDate(date, 'date');
  assertValidTimeZone(timeZone);

  return localDateTimeInZoneToUtc(
    {
      year: date.getUTCFullYear(),
      month: date.getUTCMonth() + 1,
      day: date.getUTCDate(),
      hour: date.getUTCHours(),
      minute: date.getUTCMinutes(),
      second: date.getUTCSeconds(),
      millisecond: date.getUTCMilliseconds(),
    },
    timeZone,
  );
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
  return replaceDateTimeZoneByUserTimeZone(date, timeZone);
}

function utcDateToExcelLocalDate(dateUtc: Date, timeZone: string): Date {
  assertValidDate(dateUtc, 'date');
  assertValidTimeZone(timeZone);

  const userDate = getTimeZoneParts(dateUtc, timeZone);
  return new Date(
    userDate.year,
    userDate.month - 1,
    userDate.day,
    userDate.hour,
    userDate.minute,
    userDate.second,
    dateUtc.getUTCMilliseconds(),
  );
}

function normalizeCellValue(
  cell: ExcelJS.Cell,
  parseDates: boolean,
  date1904: boolean,
  timeZone?: string,
): unknown {
  const value = cell.value;

  if (value == null) {
    return null;
  }

  if (value instanceof Date) {
    return timeZone ? reinterpretDateToTimeZone(value, timeZone) : value;
  }

  if (typeof value === 'object' && 'result' in value) {
    return (value as ExcelJS.CellFormulaValue).result ?? null;
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

export async function* xlsxToObjects<T extends RowObject = RowObject>(
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
    sharedStringsMode = 'cache',
    stylesMode = parseDates ? 'cache' : 'ignore',
  } = opts;
  const trimTextColumnsSet = new Set(trimTextColumns);
  const arrayColumnsSet = new Set(arrayColumns);
  const normalizeHeader: (header: string, index: number) => string =
    opts.normalizeHeader ?? ((header) => header);

  if (timeZone) {
    assertValidTimeZone(timeZone);
  }

  const workbook = new ExcelJS.stream.xlsx.WorkbookReader(
    toWorkbookInput(input),
    {
      worksheets: 'emit',
      sharedStrings: sharedStringsMode,
      styles: stylesMode,
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
        headers = values.map((value, index) => {
          let header = headerValueToString(value);

          if (trimHeaders) {
            header = header.trim();
          }

          header = normalizeHeader(header, index);
          return header || `column_${index + 1}`;
        });
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

export async function objectsToXlsx<T extends RowObject>(
  output: ExcelOutput,
  rows: MaybeAsyncIterable<T>,
  opts: XlsxWriteOptions<T>,
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

  if (timeZone) {
    assertValidTimeZone(timeZone);
  }

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
  }

  while (true) {
    const rowResult = await rowIterator.next();
    if (rowResult.done) {
      break;
    }

    sheet
      .addRow(mapRowForExcelWrite(rowResult.value, timeZone, dateColumnsSet))
      .commit();
  }

  sheet.commit();
  await workbook.commit();
}

export async function processXlsxLarge<T extends RowObject = RowObject>(
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

    for (const item of items) {
      const promise = Promise.resolve().then(() => handler(item));
      inFlight.add(promise);
      void promise.finally(() => {
        inFlight.delete(promise);
      });

      if (inFlight.size >= concurrency) {
        await Promise.race(inFlight);
      }
    }

    await Promise.all(inFlight);
  };

  for await (const row of xlsxToObjects<T>(input, readOpts)) {
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

export function xlsxFileToObjects<T extends RowObject = RowObject>(
  filePath: string,
  opts: XlsxReadOptions = {},
): AsyncGenerator<T> {
  return xlsxToObjects<T>(filePath, opts);
}

export function xlsxReadableToObjects<T extends RowObject = RowObject>(
  stream: Readable,
  opts: XlsxReadOptions = {},
): AsyncGenerator<T> {
  return xlsxToObjects<T>(stream, opts);
}

export function xlsxWebReadableToObjects<T extends RowObject = RowObject>(
  stream: NodeWebReadableStream<Uint8Array>,
  opts: XlsxReadOptions = {},
): AsyncGenerator<T> {
  return xlsxToObjects<T>(stream, opts);
}

export function objectsToXlsxFile<T extends RowObject>(
  filePath: string,
  rows: MaybeAsyncIterable<T>,
  opts: XlsxWriteOptions<T>,
): Promise<void> {
  return objectsToXlsx<T>(filePath, rows, opts);
}

export function objectsToXlsxWritable<T extends RowObject>(
  writable: Writable,
  rows: MaybeAsyncIterable<T>,
  opts: XlsxWriteOptions<T>,
): Promise<void> {
  return objectsToXlsx<T>(writable, rows, opts);
}
