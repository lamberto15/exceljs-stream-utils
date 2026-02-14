# exceljs-stream-utils

Streaming Excel helpers built on top of `exceljs`.

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

`exceljs` is installed automatically as a dependency of this package.

## API

### Read

- `xlsxStreamToObjects(input, options)`

### Write

- `objectsToXlsxStream(output, rows, options)`

### Processing

- `processXlsxLarge(input, handler, options)`

## Parameters

### `xlsxStreamToObjects(input, options?)`

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
| `parseDates`            | `boolean`                       | `true`                            | Convert date-like numeric cells to `Date`                       |
| `date1904`              | `boolean`                       | `false`                           | Use Excel 1904 epoch mode                                       |
| `timeZone`              | `string`                        | `undefined`                       | Reinterpret read date wall-clock in this IANA time zone         |

### `objectsToXlsxStream(output, rows, options?)`

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

### `processXlsxLarge(input, handler, options?)`

Parameters:

- `input` (required): same accepted types as `xlsxStreamToObjects`
- `handler` (required): `(row) => Promise<void> | void`
- `options` (optional): all `XlsxReadOptions` plus:

| Name          | Type     | Default | Description                            |
| ------------- | -------- | ------- | -------------------------------------- |
| `batchSize`   | `number` | `2000`  | Rows per processing batch              |
| `concurrency` | `number` | `8`     | Max concurrent handlers within a batch |

## Quick start

```ts
import { xlsxStreamToObjects, objectsToXlsxStream } from 'exceljs-stream-utils';

for await (const row of xlsxStreamToObjects('/tmp/input.xlsx')) {
  console.log(row);
}

await objectsToXlsxStream('/tmp/output.xlsx', [{ id: 1, name: 'John' }], {
  sheetName: 'Sheet1',
});
```

## How to use

### 1) Read from file

```ts
import { xlsxStreamToObjects } from 'exceljs-stream-utils';

for await (const row of xlsxStreamToObjects('/tmp/input.xlsx', {
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
import { objectsToXlsxStream } from 'exceljs-stream-utils';

await objectsToXlsxStream(
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
import { objectsToXlsxStream } from 'exceljs-stream-utils';

async function* debtorRows() {
  yield { accountNumber: '1234', amount: 100, dueDate: new Date() };
  yield { accountNumber: '5678', amount: 200, dueDate: new Date() };
}

await objectsToXlsxStream('/tmp/streamed.xlsx', debtorRows(), {
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
import { objectsToXlsxStream } from 'exceljs-stream-utils';

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

await objectsToXlsxStream(pass, debtorRows(), {
  sheetName: 'Debtors',
  timeZone: 'Asia/Manila',
  dateColumns: ['dueDate'],
});

await uploadPromise;
```

### 5) Process very large files in batches

```ts
import { processXlsxLarge } from 'exceljs-stream-utils';

await processXlsxLarge(
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

## Notes

- `objectsToXlsxStream` supports both `Iterable` and `AsyncIterable` rows.
- `columns` is optional; if omitted, columns are inferred from the first row.
- Set `timeZone` in read/write options to enable built-in timezone date conversion.
- `objectsToXlsxStream` accepts either a file path or a Node `Writable` and handles writable backpressure.

## Build and publish

```bash
bun run build
```

Release flow follows the local `npm-publish` skill process.
