import { describe, it, expect, beforeEach } from 'vitest';
import { WorkbookRegistry } from '../registry';
import { ToolExecutor } from '../executor';
import type { ExcelRunner } from '../registry';
import type { ToolCall } from '../../types';
import {
  READ_RANGE,
  LIST_SHEETS,
  GET_SHEET_CONTEXT,
  GET_SELECTION,
  LIST_WORKBOOKS,
  GET_ACTIVE_WORKBOOK,
} from '../tools/specs';
import {
  handleReadRange,
  handleListSheets,
  handleGetSheetContext,
  handleGetSelection,
} from '../tools/range';
import { handleListWorkbooks, handleGetActiveWorkbook } from '../tools/workbook_tools';
import { COPY_RANGE_FORMAT, FORMAT_RANGE, handleCopyRangeFormat, handleFormatRange } from '../tools/write';

// ── Helpers ────────────────────────────────────────────────────────────────

const SCOPE = { workbookId: 'wb1' };
const HOST_ID = 'wb1';

function makeRunner(ctx: object): ExcelRunner {
  return async fn => fn(ctx as Excel.RequestContext);
}

function makeCall(name: string, args: Record<string, unknown> = {}): ToolCall {
  return { id: `call_${name}`, name, arguments: args, workbookId: HOST_ID, mutating: false };
}

// Pre-populated registry used by most tests
function makeRegistry(): WorkbookRegistry {
  const r = new WorkbookRegistry();
  // Inject fixture state without calling Excel.run
  (r as unknown as { handles: Map<string, unknown>; activeId: string; hostId: string }).handles.set(HOST_ID, {
    workbookId: HOST_ID,
    name: 'TestBook.xlsx',
    isActive: true,
    isHost: true,
    sheets: [{ name: 'Sheet1', position: 0, visible: true }],
    lastRefreshed: new Date().toISOString(),
    capability: 'host-only',
  });
  (r as unknown as { activeId: string; hostId: string }).activeId = HOST_ID;
  (r as unknown as { hostId: string }).hostId = HOST_ID;
  return r;
}

// ── ToolExecutor validation ────────────────────────────────────────────────

describe('ToolExecutor — validation', () => {
  let executor: ToolExecutor;

  beforeEach(() => {
    const registry = makeRegistry();
    executor = new ToolExecutor(registry, makeRunner({}));
    executor.register(READ_RANGE, handleReadRange);
    executor.register(LIST_SHEETS, handleListSheets);
  });

  it('unknown tool → ValidationError', async () => {
    const r = await executor.execute(makeCall('does_not_exist'), SCOPE);
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe('ValidationError');
    expect(r.error?.message).toMatch(/Unknown tool/);
  });

  it('missing required arg → ValidationError', async () => {
    const r = await executor.execute(makeCall('read_range', { workbook_id: HOST_ID, sheet: 'Sheet1' }), SCOPE);
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe('ValidationError');
    expect(r.error?.message).toMatch(/"address"/);
  });

  it('unknown workbook_id → WorkbookNotFound', async () => {
    const r = await executor.execute(makeCall('list_sheets', { workbook_id: 'unknown-id' }), SCOPE);
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe('WorkbookNotFound');
  });

  it('getToolSpecs returns registered specs', () => {
    const specs = executor.getToolSpecs();
    expect(specs.map(s => s.name)).toContain('read_range');
    expect(specs.map(s => s.name)).toContain('list_sheets');
  });
});

// ── read_range ─────────────────────────────────────────────────────────────

describe('read_range handler', () => {
  it('returns values, rowCount, colCount from mock context', async () => {
    const mockRange = {
      values: [['Name', 'Age'], ['Alice', 30]],
      rowCount: 2,
      columnCount: 2,
      address: 'Sheet1!A1:B2',
      load: () => {},
    };
    const ctx = {
      workbook: { worksheets: { getItem: () => ({ getRange: () => mockRange }) } },
      sync: async () => {},
    };

    const registry = makeRegistry();
    const executor = new ToolExecutor(registry, makeRunner(ctx));
    executor.register(READ_RANGE, handleReadRange);

    const r = await executor.execute(
      makeCall('read_range', { workbook_id: HOST_ID, sheet: 'Sheet1', address: 'A1:B2' }),
      SCOPE
    );

    expect(r.ok).toBe(true);
    expect(r.data).toMatchObject({
      address: 'Sheet1!A1:B2',
      rowCount: 2,
      colCount: 2,
      values: [['Name', 'Age'], ['Alice', 30]],
    });
  });

  it('rejects ranges over 10k cells with ValidationError', async () => {
    const mockRange = {
      rowCount: 200,
      columnCount: 100, // 20k cells
      address: 'Sheet1!A1:CV200',
      load: () => {},
      values: [],
    };
    const ctx = {
      workbook: { worksheets: { getItem: () => ({ getRange: () => mockRange }) } },
      sync: async () => {},
    };

    const registry = makeRegistry();
    const executor = new ToolExecutor(registry, makeRunner(ctx));
    executor.register(READ_RANGE, handleReadRange);

    const r = await executor.execute(
      makeCall('read_range', { workbook_id: HOST_ID, sheet: 'Sheet1', address: 'A1:CV200' }),
      SCOPE
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe('ValidationError');
    expect(r.error?.message).toMatch(/too large/);
  });

  it('includes formulas when requested', async () => {
    const mockRange = {
      values: [[100]],
      formulas: [['=SUM(A1:A10)']],
      rowCount: 1,
      columnCount: 1,
      address: 'Sheet1!B1',
      load: () => {},
    };
    const ctx = {
      workbook: { worksheets: { getItem: () => ({ getRange: () => mockRange }) } },
      sync: async () => {},
    };

    const registry = makeRegistry();
    const executor = new ToolExecutor(registry, makeRunner(ctx));
    executor.register(READ_RANGE, handleReadRange);

    const r = await executor.execute(
      makeCall('read_range', { workbook_id: HOST_ID, sheet: 'Sheet1', address: 'B1', include: ['values', 'formulas'] }),
      SCOPE
    );
    expect(r.ok).toBe(true);
    expect((r.data as { formulas: unknown[][] }).formulas).toEqual([['=SUM(A1:A10)']]);
  });
});

// ── list_sheets ────────────────────────────────────────────────────────────

describe('copy_range_format handler', () => {
  beforeEach(() => {
    (globalThis as unknown as Record<string, unknown>).Excel = {
      RangeCopyType: { formats: 'Formats' },
    };
  });

  it('copies formats and column width from source to target', async () => {
    let copiedFrom: unknown;
    let copiedType: unknown;
    const source = {
      address: 'Sheet1!F6:F16',
      rowCount: 11,
      columnCount: 1,
      format: { columnWidth: 12, load: () => {} },
      load: () => {},
    };
    const target = {
      address: 'Sheet1!G6:G16',
      rowCount: 11,
      columnCount: 1,
      format: { columnWidth: 8, load: () => {} },
      load: () => {},
      copyFrom: (src: unknown, type: unknown) => {
        copiedFrom = src;
        copiedType = type;
      },
    };
    const ws = {
      getRange: (address: string) => address === 'F6:F16' ? source : target,
    };
    const ctx = {
      workbook: { worksheets: { getItem: () => ws } },
      sync: async () => {},
    };

    const registry = makeRegistry();
    const executor = new ToolExecutor(registry, makeRunner(ctx));
    executor.register(COPY_RANGE_FORMAT, handleCopyRangeFormat);

    const r = await executor.execute(
      makeCall('copy_range_format', {
        workbook_id: HOST_ID,
        sheet: 'Sheet1',
        source_address: 'F6:F16',
        target_address: 'G6:G16',
      }),
      SCOPE
    );

    expect(r.ok).toBe(true);
    expect(copiedFrom).toBe(source);
    expect(copiedType).toBe('Formats');
    expect(target.format.columnWidth).toBe(12);
    expect(r.data).toMatchObject({
      sheet: 'Sheet1',
      source: 'Sheet1!F6:F16',
      target: 'Sheet1!G6:G16',
      copied: 'formats',
      copiedColumnWidth: true,
    });
  });

  it('rejects mismatched source and target shapes', async () => {
    const source = {
      address: 'Sheet1!F6:F16',
      rowCount: 11,
      columnCount: 1,
      format: { columnWidth: 12, load: () => {} },
      load: () => {},
    };
    const target = {
      address: 'Sheet1!G6:G20',
      rowCount: 15,
      columnCount: 1,
      format: { columnWidth: 8, load: () => {} },
      load: () => {},
      copyFrom: () => {},
    };
    const ws = {
      getRange: (address: string) => address === 'F6:F16' ? source : target,
    };
    const ctx = {
      workbook: { worksheets: { getItem: () => ws } },
      sync: async () => {},
    };

    const registry = makeRegistry();
    const executor = new ToolExecutor(registry, makeRunner(ctx));
    executor.register(COPY_RANGE_FORMAT, handleCopyRangeFormat);

    const r = await executor.execute(
      makeCall('copy_range_format', {
        workbook_id: HOST_ID,
        sheet: 'Sheet1',
        source_address: 'F6:F16',
        target_address: 'G6:G20',
      }),
      SCOPE
    );

    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe('ValidationError');
    expect(r.error?.message).toMatch(/Dimension mismatch/);
  });
});

describe('format_range handler', () => {
  it('applies direct worksheet formatting to a range', async () => {
    let autofitColumnsCalled = false;
    let autofitRowsCalled = false;
    const borders: Record<string, { style?: unknown; color?: unknown; weight?: unknown }> = {};
    const range = {
      address: 'Sheet1!A2:E13',
      rowCount: 12,
      columnCount: 5,
      numberFormat: [] as string[][],
      format: {
        font: {} as { bold?: boolean; italic?: boolean; color?: string; size?: number },
        fill: {} as { color?: string },
        horizontalAlignment: undefined as unknown,
        verticalAlignment: undefined as unknown,
        wrapText: false,
        columnWidth: 0,
        rowHeight: 0,
        borders: {
          getItem: (index: string) => {
            borders[index] ??= {};
            return borders[index];
          },
        },
        autofitColumns: () => { autofitColumnsCalled = true; },
        autofitRows: () => { autofitRowsCalled = true; },
      },
      load: () => {},
    };
    const ctx = {
      workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } },
      sync: async () => {},
    };

    const registry = makeRegistry();
    const executor = new ToolExecutor(registry, makeRunner(ctx));
    executor.register(FORMAT_RANGE, handleFormatRange);

    const r = await executor.execute(
      makeCall('format_range', {
        workbook_id: HOST_ID,
        sheet: 'Sheet1',
        address: 'A2:E13',
        number_format: '#,##0',
        bold: true,
        fill_color: '#D9EAF7',
        horizontal_alignment: 'center',
        border_style: 'continuous',
        border_color: '#B7C9D6',
        border_weight: 'thin',
        autofit_columns: true,
        autofit_rows: true,
      }),
      SCOPE
    );

    expect(r.ok).toBe(true);
    expect(range.numberFormat).toHaveLength(12);
    expect(range.numberFormat[0]).toEqual(['#,##0', '#,##0', '#,##0', '#,##0', '#,##0']);
    expect(range.format.font.bold).toBe(true);
    expect(range.format.fill.color).toBe('#D9EAF7');
    expect(range.format.horizontalAlignment).toBe('Center');
    expect(borders.EdgeTop).toMatchObject({ style: 'Continuous', color: '#B7C9D6', weight: 'Thin' });
    expect(borders.InsideHorizontal).toMatchObject({ style: 'Continuous', color: '#B7C9D6', weight: 'Thin' });
    expect(autofitColumnsCalled).toBe(true);
    expect(autofitRowsCalled).toBe(true);
    expect(r.data).toMatchObject({
      sheet: 'Sheet1',
      address: 'Sheet1!A2:E13',
      formatted: expect.arrayContaining(['numberFormat', 'bold', 'fillColor', 'borders', 'autofitColumns', 'autofitRows']),
    });
  });

  it('rejects unsupported formatting enum values', async () => {
    const registry = makeRegistry();
    const executor = new ToolExecutor(registry, makeRunner({}));
    executor.register(FORMAT_RANGE, handleFormatRange);

    const r = await executor.execute(
      makeCall('format_range', {
        workbook_id: HOST_ID,
        sheet: 'Sheet1',
        address: 'A1',
        horizontal_alignment: 'sideways',
      }),
      SCOPE
    );

    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe('ValidationError');
    expect(r.error?.message).toMatch(/horizontal_alignment/);
  });
});

describe('list_sheets handler', () => {
  it('returns sheet list with name, position, visible', async () => {
    const mockSheets = {
      items: [
        { name: 'Sheet1', position: 0, visibility: 'Visible' },
        { name: 'Sheet2', position: 1, visibility: 'Hidden' },
      ],
      load: () => {},
    };
    const ctx = {
      workbook: { worksheets: mockSheets },
      sync: async () => {},
    };
    // Stub Excel.SheetVisibility for node test environment
    (globalThis as unknown as Record<string, unknown>).Excel = {
      SheetVisibility: { visible: 'Visible' },
    };

    const registry = makeRegistry();
    const executor = new ToolExecutor(registry, makeRunner(ctx));
    executor.register(LIST_SHEETS, handleListSheets);

    const r = await executor.execute(makeCall('list_sheets', { workbook_id: HOST_ID }), SCOPE);
    expect(r.ok).toBe(true);
    const data = r.data as Array<{ name: string; position: number; visible: boolean }>;
    expect(data).toHaveLength(2);
    expect(data[0]).toMatchObject({ name: 'Sheet1', position: 0, visible: true });
    expect(data[1]).toMatchObject({ name: 'Sheet2', position: 1, visible: false });
  });
});

// ── get_sheet_context ──────────────────────────────────────────────────────

describe('get_sheet_context handler', () => {
  it('returns usedRange, headers, sampleValues', async () => {
    const mockUsed = {
      address: 'Sheet1!A1:C4',
      rowCount: 4,
      columnCount: 3,
      rowIndex: 0,
      columnIndex: 0,
      isNullObject: false,
      load: () => {},
    };
    const mockSample = {
      values: [['Name', 'Age', 'City'], ['Alice', 30, 'NY'], ['Bob', 25, 'LA']],
      load: () => {},
    };
    const ctx = {
      workbook: {
        worksheets: {
          getItem: () => ({
            getUsedRangeOrNullObject: () => mockUsed,
            getRangeByIndexes: () => mockSample,
          }),
        },
      },
      sync: async () => {},
    };

    const registry = makeRegistry();
    const executor = new ToolExecutor(registry, makeRunner(ctx));
    executor.register(GET_SHEET_CONTEXT, handleGetSheetContext);

    const r = await executor.execute(makeCall('get_sheet_context', { workbook_id: HOST_ID, sheet: 'Sheet1' }), SCOPE);
    expect(r.ok).toBe(true);
    const data = r.data as { usedRange: object; headers: unknown[]; sampleValues: unknown[][] };
    expect(data.usedRange).toMatchObject({ address: 'Sheet1!A1:C4', rowCount: 4, colCount: 3 });
    expect(data.headers).toEqual(['Name', 'Age', 'City']);
    expect(data.sampleValues).toHaveLength(2);
  });

  it('returns null usedRange for empty sheet', async () => {
    const mockUsed = { isNullObject: true, load: () => {} };
    const ctx = {
      workbook: {
        worksheets: {
          getItem: () => ({ getUsedRangeOrNullObject: () => mockUsed }),
        },
      },
      sync: async () => {},
    };

    const registry = makeRegistry();
    const executor = new ToolExecutor(registry, makeRunner(ctx));
    executor.register(GET_SHEET_CONTEXT, handleGetSheetContext);

    const r = await executor.execute(makeCall('get_sheet_context', { workbook_id: HOST_ID, sheet: 'Empty' }), SCOPE);
    expect(r.ok).toBe(true);
    expect((r.data as { usedRange: unknown }).usedRange).toBeNull();
  });
});

// ── get_selection ──────────────────────────────────────────────────────────

describe('get_selection handler', () => {
  it('returns sheet, address, values', async () => {
    const mockSel = {
      address: 'Sheet1!B2:C3',
      worksheet: { name: 'Sheet1' },
      values: [[1, 2], [3, 4]],
      rowCount: 2,
      columnCount: 2,
      load: () => {},
    };
    const ctx = {
      workbook: { getSelectedRange: () => mockSel },
      sync: async () => {},
    };

    const registry = makeRegistry();
    const executor = new ToolExecutor(registry, makeRunner(ctx));
    executor.register(GET_SELECTION, handleGetSelection);

    const r = await executor.execute(makeCall('get_selection', { workbook_id: HOST_ID }), SCOPE);
    expect(r.ok).toBe(true);
    expect(r.data).toMatchObject({ sheet: 'Sheet1', address: 'Sheet1!B2:C3', rowCount: 2, colCount: 2 });
  });
});

// ── list_workbooks / get_active_workbook ───────────────────────────────────

describe('registry-backed tools', () => {
  it('list_workbooks returns manifest', async () => {
    const registry = makeRegistry();
    const executor = new ToolExecutor(registry, makeRunner({}));
    executor.register(LIST_WORKBOOKS, handleListWorkbooks);

    const r = await executor.execute(makeCall('list_workbooks'), SCOPE);
    expect(r.ok).toBe(true);
    const data = r.data as { active: string; workbooks: unknown[] };
    expect(data.active).toBe(HOST_ID);
    expect(data.workbooks).toHaveLength(1);
  });

  it('get_active_workbook returns workbook_id and name', async () => {
    const registry = makeRegistry();
    const executor = new ToolExecutor(registry, makeRunner({}));
    executor.register(GET_ACTIVE_WORKBOOK, handleGetActiveWorkbook);

    const r = await executor.execute(makeCall('get_active_workbook'), SCOPE);
    expect(r.ok).toBe(true);
    expect(r.data).toMatchObject({ workbook_id: HOST_ID, name: 'TestBook.xlsx' });
  });
});

// ── Office API error mapping ───────────────────────────────────────────────

describe('ToolExecutor — Office error mapping', () => {
  it('maps thrown error to OfficeApiError', async () => {
    const ctx = {
      workbook: {
        worksheets: {
          getItem: () => { throw new Error('ItemNotFound: Sheet "Missing" not found'); },
        },
      },
      sync: async () => {},
    };
    const registry = makeRegistry();
    const executor = new ToolExecutor(registry, makeRunner(ctx));
    executor.register(READ_RANGE, handleReadRange);

    const r = await executor.execute(
      makeCall('read_range', { workbook_id: HOST_ID, sheet: 'Missing', address: 'A1' }),
      SCOPE
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe('OfficeApiError');
    expect(r.error?.message).toMatch(/ItemNotFound/);
  });
});
