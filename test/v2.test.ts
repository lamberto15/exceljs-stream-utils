import { afterEach, describe, expect, test } from 'bun:test';
import ExcelJS from '@protobi/exceljs';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { PassThrough, Readable } from 'node:stream';

import {
  objectsToXlsxStream,
  processXlsxLarge,
  readXlsxRows as readRootXlsxRows,
  writeXlsxRows as writeRootXlsxRows,
  xlsxStreamToObjects,
} from '../src/index';
import {
  processXlsxRows,
  readXlsxRows,
  writeXlsxRows,
} from '../src/v2';

const tempDirs: string[] = [];

afterEach(async () => {
  await Promise.all(
    tempDirs.splice(0).map((dir) => rm(dir, { recursive: true, force: true })),
  );
});

async function createTempDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'exceljs-stream-utils-'));
  tempDirs.push(dir);
  return dir;
}

async function createWorkbookFile(
  filename: string,
  configure: (worksheet: ExcelJS.Worksheet) => void,
): Promise<string> {
  const dir = await createTempDir();
  const filePath = join(dir, filename);
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');

  configure(worksheet);

  await workbook.xlsx.writeFile(filePath);
  return filePath;
}

async function collectRows<T>(
  iterable: AsyncIterable<T>,
): Promise<T[]> {
  const rows: T[] = [];
  for await (const row of iterable) {
    rows.push(row);
  }
  return rows;
}

async function passThroughToBuffer(pass: PassThrough): Promise<Buffer> {
  const chunks: Buffer[] = [];

  return new Promise<Buffer>((resolve, reject) => {
    pass.on('data', (chunk) => {
      chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    });
    pass.on('end', () => resolve(Buffer.concat(chunks)));
    pass.on('error', reject);
  });
}

describe('root API compatibility', () => {
  test('keeps duplicate-header overwrite behavior', async () => {
    const filePath = await createWorkbookFile('root-duplicate-headers.xlsx', (ws) => {
      ws.addRow(['name', 'name']);
      ws.addRow(['Alice', 'Bob']);
    });

    const rows = await collectRows(
      readRootXlsxRows<{ name: string }>(filePath),
    );

    expect(rows).toEqual([{ name: 'Bob' }]);
  });

  test('supports legacy aliases', async () => {
    const dir = await createTempDir();
    const filePath = join(dir, 'legacy-aliases.xlsx');

    await objectsToXlsxStream(
      filePath,
      [{ id: 1, name: 'Alice' }],
      { sheetName: 'Sheet1' },
    );

    const rows = await collectRows(
      xlsxStreamToObjects<{ id: number; name: string }>(filePath),
    );

    expect(rows).toEqual([{ id: 1, name: 'Alice' }]);
  });

  test('supports current root named exports', async () => {
    const dir = await createTempDir();
    const filePath = join(dir, 'root-named-exports.xlsx');
    const handledIds: number[] = [];

    await writeRootXlsxRows(
      filePath,
      [{ id: 1 }, { id: 2 }, { id: 3 }],
      { sheetName: 'Sheet1' },
    );

    await processXlsxLarge<{ id: number }>(
      filePath,
      (row) => {
        handledIds.push(row.id);
      },
      { batchSize: 2, concurrency: 2 },
    );

    expect(handledIds).toEqual([1, 2, 3]);
  });
});

describe('/v2 readXlsxRows', () => {
  test('reads common cell types as expected', async () => {
    const filePath = await createWorkbookFile('common-types.xlsx', (ws) => {
      ws.addRow([
        'name',
        'active',
        'count',
        'ratio',
        'notes',
        'optional',
        'createdAt',
      ]);

      ws.getCell('A2').value = 'Alice';
      ws.getCell('B2').value = true;
      ws.getCell('C2').value = 42;
      ws.getCell('D2').value = 123.45;
      ws.getCell('E2').value = 'hello';
      ws.getCell('F2').value = null;
      ws.getCell('G2').value = new Date('2024-01-02T03:04:05.000Z');
      ws.getCell('G2').numFmt = 'mm/dd/yyyy hh:mm:ss';
    });

    const rows = await collectRows(
      readXlsxRows<{
        name: string;
        active: boolean;
        count: number;
        ratio: number;
        notes: string;
        optional: null;
        createdAt: Date;
      }>(filePath),
    );

    expect(rows).toHaveLength(1);
    expect(rows[0].name).toBe('Alice');
    expect(rows[0].active).toBe(true);
    expect(rows[0].count).toBe(42);
    expect(rows[0].ratio).toBe(123.45);
    expect(rows[0].notes).toBe('hello');
    expect(rows[0].optional).toBeNull();
    expect(rows[0].createdAt).toBeInstanceOf(Date);
    expect(rows[0].createdAt.toISOString()).toBe('2024-01-02T03:04:05.000Z');
  });

  test('throws on duplicate headers by default', async () => {
    const filePath = await createWorkbookFile('duplicate-headers.xlsx', (ws) => {
      ws.addRow(['name', 'name']);
      ws.addRow(['Alice', 'Bob']);
    });

    await expect(
      collectRows(readXlsxRows(filePath)),
    ).rejects.toThrow('Duplicate header "name"');
  });

  test('can suffix duplicate headers when configured', async () => {
    const filePath = await createWorkbookFile('duplicate-headers-suffix.xlsx', (ws) => {
      ws.addRow(['name', 'name']);
      ws.addRow(['Alice', 'Bob']);
    });

    const rows = await collectRows(
      readXlsxRows(filePath, { duplicateHeaders: 'suffix' }),
    );

    expect(rows).toEqual([{ name: 'Alice', name_2: 'Bob' }]);
  });

  test('trims, normalizes, parses arrays, and fills blank headers', async () => {
    const filePath = await createWorkbookFile('read-options.xlsx', (ws) => {
      ws.addRow([' Account ', '', 'Tags', 'Note']);
      ws.addRow(['  A-1  ', 'x', ' red, blue, , green ', '  keep me  ']);
      ws.addRow([null, null, null, null]);
    });

    const rows = await collectRows(
      readXlsxRows<{
        account: string;
        column_2: string;
        tags: string[];
        note: string;
      }>(filePath, {
        trimTextColumns: ['account', 'note'],
        arrayColumns: ['tags'],
        normalizeHeader: (header) =>
          header.toLowerCase().replace(/\s+/g, '_'),
      }),
    );

    expect(rows).toEqual([
      {
        account: 'A-1',
        column_2: 'x',
        tags: ['red', 'blue', 'green'],
        note: 'keep me',
      },
    ]);
  });

  test('normalizes formula date results to Date values', async () => {
    const filePath = await createWorkbookFile('formula-date.xlsx', (ws) => {
      ws.getCell('A1').value = 'dueDate';
      ws.getCell('A2').value = { formula: 'DATE(2024,1,2)', result: 45293 };
      ws.getCell('A2').numFmt = 'mm/dd/yyyy';
    });

    const rows = await collectRows(
      readXlsxRows<{ dueDate: Date }>(filePath),
    );

    expect(rows).toHaveLength(1);
    expect(rows[0].dueDate).toBeInstanceOf(Date);
    expect(rows[0].dueDate.toISOString()).toBe('2024-01-02T00:00:00.000Z');
  });

  test('can disable date parsing for numeric Excel dates', async () => {
    const filePath = await createWorkbookFile('raw-date-serial.xlsx', (ws) => {
      ws.getCell('A1').value = 'dueDate';
      ws.getCell('A2').value = 45293;
      ws.getCell('A2').numFmt = 'mm/dd/yyyy';
    });

    const rows = await collectRows(
      readXlsxRows<{ dueDate: number }>(filePath, { parseDates: false }),
    );

    expect(rows).toEqual([{ dueDate: 45293 }]);
  });

  test('round-trips through a Node readable stream', async () => {
    const sourcePath = await createWorkbookFile('stream-input-source.xlsx', (ws) => {
      ws.addRow(['id', 'name']);
      ws.addRow([1, 'Alice']);
      ws.addRow([2, 'Bob']);
    });
    const sourceBuffer = await readFile(sourcePath);

    const rows = await collectRows(
      readXlsxRows<{ id: number; name: string }>(Readable.from(sourceBuffer)),
    );

    expect(rows).toEqual([
      { id: 1, name: 'Alice' },
      { id: 2, name: 'Bob' },
    ]);
  });
});

describe('/v2 writeXlsxRows', () => {
  test('round-trips common JavaScript values as expected', async () => {
    const dir = await createTempDir();
    const filePath = join(dir, 'write-common-types.xlsx');

    await writeXlsxRows(filePath, [
      {
        name: 'Alice',
        active: true,
        count: 42,
        ratio: 123.45,
        notes: 'hello',
        optional: null,
      },
    ]);

    const rows = await collectRows(
      readXlsxRows<{
        name: string;
        active: boolean;
        count: number;
        ratio: number;
        notes: string;
        optional: null;
      }>(filePath),
    );

    expect(rows).toEqual([
      {
        name: 'Alice',
        active: true,
        count: 42,
        ratio: 123.45,
        notes: 'hello',
        optional: null,
      },
    ]);
  });

  test('accepts omitted options', async () => {
    const dir = await createTempDir();
    const filePath = join(dir, 'write-no-opts.xlsx');

    await writeXlsxRows(filePath, [{ id: 1, name: 'Alice' }]);

    const rows = await collectRows(
      readXlsxRows<{ id: number; name: string }>(filePath),
    );

    expect(rows).toEqual([{ id: 1, name: 'Alice' }]);
  });

  test('supports Node writable stream output', async () => {
    const pass = new PassThrough();
    const bufferPromise = passThroughToBuffer(pass);

    await writeXlsxRows(pass, [
      { id: 1, name: 'Alice' },
      { id: 2, name: 'Bob' },
    ]);

    const buffer = await bufferPromise;
    const dir = await createTempDir();
    const filePath = join(dir, 'stream-output.xlsx');
    await writeFile(filePath, buffer);

    const rows = await collectRows(
      readXlsxRows<{ id: number; name: string }>(filePath),
    );

    expect(rows).toEqual([
      { id: 1, name: 'Alice' },
      { id: 2, name: 'Bob' },
    ]);
  });

  test('round-trips time-zoned date values', async () => {
    const dir = await createTempDir();
    const filePath = join(dir, 'timezone-roundtrip.xlsx');
    const dueDate = new Date('2024-01-02T16:30:45.123Z');

    await writeXlsxRows(
      filePath,
      [{ id: 1, dueDate }],
      {
        timeZone: 'Asia/Manila',
        dateColumns: ['dueDate'],
      },
    );

    const rows = await collectRows(
      readXlsxRows<{ id: number; dueDate: Date }>(filePath, {
        timeZone: 'Asia/Manila',
      }),
    );

    expect(rows).toHaveLength(1);
    expect(rows[0].dueDate).toBeInstanceOf(Date);
    expect(rows[0].dueDate.toISOString()).toBe(dueDate.toISOString());
  });

  test('throws clearly for empty rows without explicit columns', async () => {
    const dir = await createTempDir();
    const filePath = join(dir, 'empty-write.xlsx');

    await expect(
      writeXlsxRows(filePath, []),
    ).rejects.toThrow('Unable to determine columns');
  });
});

describe('/v2 processXlsxRows', () => {
  test('waits for in-flight handlers to settle before rejecting', async () => {
    const filePath = await createWorkbookFile('process-errors.xlsx', (ws) => {
      ws.addRow(['id']);
      ws.addRow([1]);
      ws.addRow([2]);
    });

    let secondHandlerSettled = false;

    await expect(
      processXlsxRows(
        filePath,
        async (row: { id: number }) => {
          if (row.id === 1) {
            await Bun.sleep(10);
            throw new Error('first failure');
          }

          if (row.id === 2) {
            await Bun.sleep(40);
            secondHandlerSettled = true;
          }
        },
        { batchSize: 2, concurrency: 2 },
      ),
    ).rejects.toThrow('first failure');

    expect(secondHandlerSettled).toBe(true);
  });

  test('respects the concurrency ceiling', async () => {
    const filePath = await createWorkbookFile('process-concurrency.xlsx', (ws) => {
      ws.addRow(['id']);
      for (let id = 1; id <= 24; id += 1) {
        ws.addRow([id]);
      }
    });

    let inFlight = 0;
    let maxInFlight = 0;

    await processXlsxRows<{ id: number }>(
      filePath,
      async () => {
        inFlight += 1;
        maxInFlight = Math.max(maxInFlight, inFlight);
        await Bun.sleep(5);
        inFlight -= 1;
      },
      { batchSize: 12, concurrency: 3 },
    );

    expect(maxInFlight).toBeLessThanOrEqual(3);
  });

  test('processes a realistic 10000-row file end to end', async () => {
    const dir = await createTempDir();
    const filePath = join(dir, 'large-dataset.xlsx');
    const rowCount = 10000;

    async function* largeRowSource() {
      for (let id = 1; id <= rowCount; id += 1) {
        yield {
          debtorId: `D-${id.toString().padStart(5, '0')}`,
          email: `user${id}@example.com`,
          balanceCents: id * 125,
          dueDate: new Date('2024-01-01T00:00:00.000Z'),
          tags: id % 2 === 0 ? 'priority,renewal' : 'standard',
        };
      }
    }

    await writeXlsxRows(filePath, largeRowSource(), {
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

    let processed = 0;
    let totalBalance = 0;

    await processXlsxRows<{
      'Debtor ID': string;
      Email: string;
      Balance: number;
      'Due Date': Date;
      Tags: string;
    }>(
      filePath,
      async (row) => {
        processed += 1;
        totalBalance += row.Balance;
      },
      {
        batchSize: 1000,
        concurrency: 16,
        timeZone: 'Asia/Manila',
      },
    );

    expect(processed).toBe(rowCount);
    expect(totalBalance).toBe((125 * rowCount * (rowCount + 1)) / 2);
  });
});
