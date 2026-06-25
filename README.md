# exceljs-stream-utils

Streaming Excel helpers built on top of `@protobi/exceljs`, a maintained ExcelJS-based streaming implementation.

- Read `.xlsx` as async objects
- Write iterable or async-iterable rows to `.xlsx`
- Process large files in bounded batches
- Built-in timezone-aware date handling for read/write flows

## Requirements

- Node.js `>= 18`

## Install

```bash
npm install exceljs-stream-utils
```

`@protobi/exceljs` is installed automatically as a dependency of this package.

This package builds on top of:

- [`@protobi/exceljs`](https://www.npmjs.com/package/@protobi/exceljs) for Excel workbook streaming support
- [`luxon`](https://www.npmjs.com/package/luxon) for time zone aware date conversion

## API

The package now has two entrypoints:

- `exceljs-stream-utils`: existing behavior, kept stable for compatibility
- `exceljs-stream-utils/v2`: improved behavior for new code

### Read

- `readXlsxRows(input, options)`

### Write

- `writeXlsxRows(output, rows, options)`

### Processing

- `processXlsxRows(input, handler, options)`

Compatibility aliases are still exported:

- `xlsxStreamToObjects`
- `objectsToXlsxStream`
- `processXlsxLarge`

## V2

Import from `exceljs-stream-utils/v2` to opt into the safer API without affecting existing users.

```ts
import {
  readXlsxRows,
  writeXlsxRows,
  processXlsxRows,
} from 'exceljs-stream-utils/v2';
```

`v2` currently improves a few edge cases:

- `writeXlsxRows(output, rows)` now works without requiring an options object
- Formula result cells are normalized through the same date handling path as other cells
- Duplicate headers are detected by default instead of silently overwriting earlier columns
- `processXlsxRows` drains in-flight work before surfacing batch errors

Automated Bun tests currently cover:

- file path input and output
- Node readable and writable stream transport
- timezone date round-trips
- array parsing, trimming, and header normalization
- root API compatibility and legacy aliases
- a realistic 10,000-row write-and-process scenario

## Parameters

Unless noted otherwise, the parameter tables below describe the `/v2` entrypoint. The root entrypoint keeps the older compatibility behavior.

### `readXlsxRows(input, options?)`

`input` types:

- `string` (file path)
- Node `Readable`
- Web `ReadableStream<Uint8Array>`

`options` (`XlsxReadOptions`, all optional):

| Name                    | Type                            | Default                           | Description                                                     |
| ----------------------- | ------------------------------- | --------------------------------- | --------------------------------------------------------------- |
| `sheetName`             | `string`                        | first matching sheet              | Read only this sheet name                                       |
| `headerRowNumber`       | `number`                        | `1`                               | Row number used as header row                                   |
| `trimHeaders`           | `boolean`                       | `true`                            | Trim header cell text                                           |
| `trimTextValues`        | `boolean`                       | `false`                           | Trim all string cell values                                     |
| `trimTextColumns`       | `string[]`                      | `[]`                              | Trim only these columns (used when `trimTextValues` is `false`) |
| `arrayColumns`          | `string[]`                      | `[]`                              | Convert these string columns to arrays                          |
| `arrayDelimiter`        | `string`                        | `","`                             | Delimiter for `arrayColumns`                                    |
| `trimArrayItems`        | `boolean`                       | `true`                            | Trim each array item after split                                |
| `removeEmptyArrayItems` | `boolean`                       | `true`                            | Remove empty items from split arrays                            |
| `skipEmptyRows`         | `boolean`                       | `true`                            | Skip rows where all cells are empty                             |
| `normalizeHeader`       | `(header, index) => string`     | identity                          | Custom header normalizer                                        |
| `duplicateHeaders`      | `'error' \| 'suffix'`           | `'error'`                         | `/v2` only. Error on duplicate headers or suffix them as `_2`, `_3`, etc. |
| `parseDates`            | `boolean`                       | `true`                            | Convert date-like numeric cells to `Date`                       |
| `date1904`              | `boolean`                       | `false`                           | Use Excel 1904 epoch mode                                       |
| `timeZone`              | `string`                        | `undefined`                       | Reinterpret read date wall-clock in this IANA time zone         |

### `writeXlsxRows(output, rows, options?)`

`output` types:

- `string` (output file path)
- Node `Writable`

`rows` types:

- `Iterable<Record<string, unknown>>`
- `AsyncIterable<Record<string, unknown>>`

`options` (`XlsxWriteOptions`, all optional):

| Name               | Type                                                | Default                 | Description                                                                       |
| ------------------ | --------------------------------------------------- | ----------------------- | --------------------------------------------------------------------------------- |
| `sheetName`        | `string`                                            | `'Sheet1'`              | Worksheet name                                                                    |
| `columns`          | `{ header: string; key: string; width?: number }[]` | inferred from first row | Explicit column order and labels                                                  |
| `useStyles`        | `boolean`                                           | `false`                 | Enable ExcelJS style writing                                                      |
| `useSharedStrings` | `boolean`                                           | `false`                 | Enable ExcelJS shared strings writing                                             |
| `timeZone`         | `string`                                            | `undefined`             | Convert `Date` values to target time zone wall-clock                              |
| `dateColumns`      | `string[]`                                          | `[]`                    | Restrict timezone conversion to these date columns; empty means all `Date` fields |

### `processXlsxRows(input, handler, options?)`

Parameters:

- `input` (required): same accepted types as `readXlsxRows`
- `handler` (required): `(row) => Promise<void> | void`
- `options` (optional): all `XlsxReadOptions` plus:

| Name          | Type     | Default | Description                            |
| ------------- | -------- | ------- | -------------------------------------- |
| `batchSize`   | `number` | `2000`  | Rows per processing batch              |
| `concurrency` | `number` | `8`     | Max concurrent handlers within a batch |

## Quick start

```ts
import { readXlsxRows, writeXlsxRows } from 'exceljs-stream-utils';

for await (const row of readXlsxRows('/tmp/input.xlsx')) {
  console.log(row);
}

await writeXlsxRows('/tmp/output.xlsx', [{ id: 1, name: 'John' }], {
  sheetName: 'Sheet1',
});
```

## How to use

### 1) Read from file

```ts
import { readXlsxRows } from 'exceljs-stream-utils';

for await (const row of readXlsxRows('/tmp/input.xlsx', {
  headerRowNumber: 1,
  trimHeaders: true,
  trimTextValues: true,
  parseDates: true,
  timeZone: 'Asia/Manila',
  arrayColumns: ['tags'],
  arrayDelimiter: ',',
})) {
  console.log(row);
}
```

### 2) Write to file

```ts
import { writeXlsxRows } from 'exceljs-stream-utils';

await writeXlsxRows(
  '/tmp/out.xlsx',
  [
    { accountNumber: '1234', amount: 100, dueDate: new Date() },
    { accountNumber: '5678', amount: 200, dueDate: new Date() },
  ],
  {
    sheetName: 'Debtors',
    columns: [
      { header: 'Account #', key: 'accountNumber' },
      { header: 'Amount', key: 'amount' },
      { header: 'Due Date', key: 'dueDate' },
    ],
    timeZone: 'Asia/Manila',
    dateColumns: ['dueDate'],
  },
);
```

### 3) Write from async iterator

```ts
import { writeXlsxRows } from 'exceljs-stream-utils';

async function* debtorRows() {
  yield { accountNumber: '1234', amount: 100, dueDate: new Date() };
  yield { accountNumber: '5678', amount: 200, dueDate: new Date() };
}

await writeXlsxRows('/tmp/streamed.xlsx', debtorRows(), {
  sheetName: 'Debtors',
  columns: [
    { header: 'Account #', key: 'accountNumber' },
    { header: 'Amount', key: 'amount' },
    { header: 'Due Date', key: 'dueDate' },
  ],
});
```

### 4) Stream directly to S3

```ts
import { PassThrough } from 'node:stream';
import { Upload } from '@aws-sdk/lib-storage';
import { writeXlsxRows } from 'exceljs-stream-utils';

const pass = new PassThrough();

const upload = new Upload({
  client: s3Client,
  params: {
    Bucket: 'my-bucket',
    Key: 'exports/debtors.xlsx',
    Body: pass,
  },
});

const uploadPromise = upload.done();

await writeXlsxRows(pass, debtorRows(), {
  sheetName: 'Debtors',
  timeZone: 'Asia/Manila',
  dateColumns: ['dueDate'],
});

await uploadPromise;
```

### 5) Process very large files in batches

```ts
import { processXlsxRows } from 'exceljs-stream-utils';

await processXlsxRows(
  '/tmp/input.xlsx',
  async (row) => {
    // handle one row
  },
  {
    batchSize: 2000,
    concurrency: 8,
    trimTextValues: true,
    parseDates: true,
  },
);
```

### 6) Export a realistic large dataset

```ts
import { writeXlsxRows } from 'exceljs-stream-utils/v2';

async function* debtorExportRows() {
  for (let id = 1; id <= 10000; id += 1) {
    yield {
      debtorId: `D-${id.toString().padStart(5, '0')}`,
      email: `user${id}@example.com`,
      balanceCents: id * 125,
      dueDate: new Date('2024-01-01T00:00:00.000Z'),
      tags: id % 2 === 0 ? 'priority,renewal' : 'standard',
    };
  }
}

await writeXlsxRows('/tmp/debtors.xlsx', debtorExportRows(), {
  sheetName: 'Debtors',
  columns: [
    { header: 'Debtor ID', key: 'debtorId' },
    { header: 'Email', key: 'email' },
    { header: 'Balance', key: 'balanceCents' },
    { header: 'Due Date', key: 'dueDate' },
    { header: 'Tags', key: 'tags' },
  ],
  timeZone: 'Asia/Manila',
  dateColumns: ['dueDate'],
});
```

## Notes

- `writeXlsxRows` supports both `Iterable` and `AsyncIterable` rows.
- `columns` is optional; if omitted, columns are inferred from the first row.
- Set `timeZone` in read/write options to enable built-in timezone date conversion.
- `writeXlsxRows` accepts either a file path or a Node `Writable` and handles writable backpressure.

## Type Behavior

- String cells are returned as strings.
- Boolean cells are returned as booleans.
- Integer and decimal numeric cells are returned as numbers.
- Empty cells are returned as `null`.
- Date cells are returned as `Date` objects when date parsing is enabled.

Excel date values are commonly stored as numeric serials plus cell formatting metadata. Because of that, automatic date parsing depends on the workbook indicating that a numeric cell is date-like, usually through its Excel number format. This is a normal Excel parsing constraint, not a Node-specific quirk.

If you want raw numeric serials instead of parsed `Date` values, set:

```ts
{ parseDates: false }
```

## Build and publish

```bash
bun run build
```

Release flow follows the local `npm-publish` skill process.
