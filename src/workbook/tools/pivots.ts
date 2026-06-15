import type { ToolSpec } from '../../types';
import type { ToolHandler } from '../executor';
import { ToolValidationError } from '../executor';
import { ToolUnsupportedError } from '../unsupported-error';
import { optionalBooleanArg, optionalStringArg, stringArg } from './args';

function assertPivotApi(ctx: Excel.RequestContext): void {
  if (!('pivotTables' in ctx.workbook)) {
    throw new ToolUnsupportedError('Pivot tables require ExcelApi 1.8+');
  }
}

// ── Aggregation function mapping ───────────────────────────────────────────

function getAggregation(fn?: string): Excel.AggregationFunction {
  switch ((fn ?? '').toLowerCase()) {
    case 'count':   return Excel.AggregationFunction.count;
    case 'average': return Excel.AggregationFunction.average;
    case 'max':     return Excel.AggregationFunction.max;
    case 'min':     return Excel.AggregationFunction.min;
    case 'product': return Excel.AggregationFunction.product;
    default:        return Excel.AggregationFunction.sum;
  }
}

// ── Specs ──────────────────────────────────────────────────────────────────

export const LIST_PIVOTS: ToolSpec = {
  name: 'list_pivots',
  description: 'List all pivot tables in a worksheet, or across the whole workbook if sheet is omitted.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name (optional — omit to list all pivot tables in workbook)' },
    },
    required: ['workbook_id'],
  },
  mutating: false,
};

export const GET_PIVOT: ToolSpec = {
  name: 'get_pivot',
  description: 'Get the field layout of a pivot table: which fields are in rows, columns, data, and filter areas.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      name: { type: 'string', description: 'Pivot table name' },
    },
    required: ['workbook_id', 'name'],
  },
  mutating: false,
};

export const CREATE_PIVOT: ToolSpec = {
  name: 'create_pivot',
  description: 'Create a new pivot table from a source data range. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id:  { type: 'string', description: 'Workbook ID' },
      sheet:        { type: 'string', description: 'Source data worksheet name' },
      source_range: { type: 'string', description: 'A1 range of source data including headers, e.g. "A1:D100"' },
      destination:  { type: 'string', description: 'Top-left cell address where the pivot will be placed, e.g. "F1"' },
      dest_sheet:   { type: 'string', description: 'Destination worksheet name (defaults to same as source sheet)' },
      name:         { type: 'string', description: 'Pivot table name (optional — Excel auto-assigns if omitted)' },
    },
    required: ['workbook_id', 'sheet', 'source_range', 'destination'],
  },
  mutating: true,
};

export const ADD_PIVOT_FIELD: ToolSpec = {
  name: 'add_pivot_field',
  description: 'Add a field to a pivot table area (row, column, data, or filter). Use get_pivot first to see available fields.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      name:        { type: 'string', description: 'Pivot table name' },
      field:       { type: 'string', description: 'Field name to add (must match a column header from the source range)' },
      area:        { type: 'string', description: 'Target area: "row", "column", "data", or "filter"' },
      function:    { type: 'string', description: 'Aggregation for data area: sum (default), count, average, max, min, product' },
    },
    required: ['workbook_id', 'name', 'field', 'area'],
  },
  mutating: true,
};

export const REFRESH_PIVOT: ToolSpec = {
  name: 'refresh_pivot',
  description: 'Refresh a pivot table to reflect changes in its source data. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      name: { type: 'string', description: 'Pivot table name' },
    },
    required: ['workbook_id', 'name'],
  },
  mutating: true,
};

export const REMOVE_PIVOT_FIELD: ToolSpec = {
  name: 'remove_pivot_field',
  description: 'Remove a field from a pivot table area. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      name: { type: 'string', description: 'Pivot table name' },
      field: { type: 'string', description: 'Field name to remove' },
      area: { type: 'string', description: 'Area: row, column, data, or filter' },
    },
    required: ['workbook_id', 'name', 'field', 'area'],
  },
  mutating: true,
};

export const SET_PIVOT_STYLE: ToolSpec = {
  name: 'set_pivot_style',
  description: 'Set the pivot table style. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      name: { type: 'string', description: 'Pivot table name' },
      style: { type: 'string', description: 'Pivot style name, e.g. PivotStyleMedium9' },
    },
    required: ['workbook_id', 'name', 'style'],
  },
  mutating: true,
};

export const SET_PIVOT_LAYOUT: ToolSpec = {
  name: 'set_pivot_layout',
  description: 'Set pivot table layout and grand total options. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      name: { type: 'string', description: 'Pivot table name' },
      layout_type: { type: 'string', description: 'Compact, Outline, or Tabular' },
      show_row_grand_totals: { type: 'boolean', description: 'Show row grand totals' },
      show_column_grand_totals: { type: 'boolean', description: 'Show column grand totals' },
      show_banded_rows: { type: 'boolean', description: 'Show banded rows' },
      show_banded_columns: { type: 'boolean', description: 'Show banded columns' },
    },
    required: ['workbook_id', 'name'],
  },
  mutating: true,
};

export const REFRESH_ALL_PIVOTS: ToolSpec = {
  name: 'refresh_all_pivots',
  description: 'Refresh every pivot table in the workbook. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
    },
    required: ['workbook_id'],
  },
  mutating: true,
};

export const DELETE_PIVOT: ToolSpec = {
  name: 'delete_pivot',
  description: 'Delete a pivot table. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      name: { type: 'string', description: 'Pivot table name' },
    },
    required: ['workbook_id', 'name'],
  },
  mutating: true,
};

export const PIVOT_SPECS: ToolSpec[] = [
  LIST_PIVOTS,
  GET_PIVOT,
  CREATE_PIVOT,
  ADD_PIVOT_FIELD,
  REFRESH_PIVOT,
  REMOVE_PIVOT_FIELD,
  SET_PIVOT_STYLE,
  SET_PIVOT_LAYOUT,
  REFRESH_ALL_PIVOTS,
  DELETE_PIVOT,
];

// ── Handlers ───────────────────────────────────────────────────────────────

export const handleListPivots: ToolHandler = async (args, ctx) => {
  assertPivotApi(ctx);
  if (args.sheet) {
    const sheet = ctx.workbook.worksheets.getItem(args.sheet as string);
    sheet.pivotTables.load('items/name');
    await ctx.sync();
    return sheet.pivotTables.items.map(p => ({ name: p.name }));
  } else {
    ctx.workbook.pivotTables.load('items/name,items/worksheet/name');
    await ctx.sync();
    return ctx.workbook.pivotTables.items.map(p => ({
      name: p.name,
      sheet: (p as unknown as { worksheet?: { name?: string } }).worksheet?.name,
    }));
  }
};

export const handleGetPivot: ToolHandler = async (args, ctx) => {
  assertPivotApi(ctx);
  const pivot = ctx.workbook.pivotTables.getItem(args.name as string);
  pivot.rowHierarchies.load('items/name');
  pivot.columnHierarchies.load('items/name');
  pivot.dataHierarchies.load('items/name,items/summarizeBy');
  pivot.filterHierarchies.load('items/name');
  pivot.hierarchies.load('items/name');
  await ctx.sync();
  return {
    name: args.name,
    availableFields: pivot.hierarchies.items.map(h => h.name),
    rows:    pivot.rowHierarchies.items.map(h => h.name),
    columns: pivot.columnHierarchies.items.map(h => h.name),
    data:    pivot.dataHierarchies.items.map(h => ({ name: h.name, summarizeBy: h.summarizeBy })),
    filters: pivot.filterHierarchies.items.map(h => h.name),
  };
};

export const handleCreatePivot: ToolHandler = async (args, ctx) => {
  assertPivotApi(ctx);
  const srcSheet  = ctx.workbook.worksheets.getItem(args.sheet as string);
  const destSheet = args.dest_sheet
    ? ctx.workbook.worksheets.getItem(args.dest_sheet as string)
    : srcSheet;

  const sourceRange = srcSheet.getRange(args.source_range as string);
  const destRange   = destSheet.getRange(args.destination as string);
  const pivotName   = (args.name as string | undefined) ?? '';

  const pivot = destSheet.pivotTables.add(pivotName, sourceRange, destRange);
  pivot.load('name');
  await ctx.sync();
  return { name: pivot.name, created: true };
};

export const handleAddPivotField: ToolHandler = async (args, ctx) => {
  assertPivotApi(ctx);
  const area  = args.area as string;
  const field = args.field as string;

  if (!['row', 'column', 'data', 'filter'].includes(area)) {
    throw new ToolValidationError(`Invalid area "${area}". Must be: row, column, data, filter`);
  }

  const pivot = ctx.workbook.pivotTables.getItem(args.name as string);
  pivot.hierarchies.load('items/name');
  await ctx.sync();

  const hierarchy = pivot.hierarchies.items.find(h => h.name === field);
  if (!hierarchy) {
    const available = pivot.hierarchies.items.map(h => h.name).join(', ');
    throw new ToolValidationError(`Field "${field}" not found. Available: ${available}`);
  }

  switch (area) {
    case 'row':    pivot.rowHierarchies.add(hierarchy); break;
    case 'column': pivot.columnHierarchies.add(hierarchy); break;
    case 'filter': pivot.filterHierarchies.add(hierarchy); break;
    case 'data': {
      const dh = pivot.dataHierarchies.add(hierarchy);
      dh.summarizeBy = getAggregation(args.function as string | undefined);
      break;
    }
  }
  await ctx.sync();
  return { name: args.name, field, area };
};

export const handleRefreshPivot: ToolHandler = async (args, ctx) => {
  assertPivotApi(ctx);
  const pivot = ctx.workbook.pivotTables.getItem(args.name as string);
  pivot.refresh();
  await ctx.sync();
  return { name: args.name, refreshed: true };
};

type PivotCollectionLike = {
  load: (properties: string) => void;
  items: Array<{ name: string; delete?: () => void }>;
  remove?: (hierarchy: unknown) => void;
};

function pivotByName(args: Record<string, unknown>, ctx: Excel.RequestContext): Excel.PivotTable {
  assertPivotApi(ctx);
  return ctx.workbook.pivotTables.getItem(stringArg(args, 'name'));
}

function pivotAreaCollection(pivot: Excel.PivotTable, area: string): PivotCollectionLike {
  const pivotAny = pivot as unknown as {
    rowHierarchies: PivotCollectionLike;
    columnHierarchies: PivotCollectionLike;
    dataHierarchies: PivotCollectionLike;
    filterHierarchies: PivotCollectionLike;
  };
  switch (area) {
    case 'row': return pivotAny.rowHierarchies;
    case 'column': return pivotAny.columnHierarchies;
    case 'data': return pivotAny.dataHierarchies;
    case 'filter': return pivotAny.filterHierarchies;
    default:
      throw new ToolValidationError(`Invalid area "${area}". Must be: row, column, data, filter`);
  }
}

export const handleRemovePivotField: ToolHandler = async (args, ctx) => {
  const pivot = pivotByName(args, ctx);
  const field = stringArg(args, 'field');
  const area = stringArg(args, 'area').toLowerCase();
  const collection = pivotAreaCollection(pivot, area);
  collection.load('items/name');
  await ctx.sync();
  const item = collection.items.find(h => h.name === field);
  if (!item) {
    throw new ToolValidationError(`Field "${field}" is not in the ${area} area.`);
  }
  if (collection.remove) {
    collection.remove(item);
  } else if (item.delete) {
    item.delete();
  } else {
    throw new ToolUnsupportedError(`Removing fields from the ${area} area is not supported by this Excel host.`);
  }
  await ctx.sync();
  return { name: args.name, field, area, removed: true };
};

export const handleSetPivotStyle: ToolHandler = async (args, ctx) => {
  const pivot = pivotByName(args, ctx) as unknown as { style?: string };
  pivot.style = stringArg(args, 'style');
  await ctx.sync();
  return { name: args.name, style: args.style };
};

export const handleSetPivotLayout: ToolHandler = async (args, ctx) => {
  const pivot = pivotByName(args, ctx) as unknown as {
    layout?: { layoutType?: string };
    rowGrandTotals?: boolean;
    columnGrandTotals?: boolean;
    showBandedRows?: boolean;
    showBandedColumns?: boolean;
  };
  const applied: string[] = [];
  const layoutType = optionalStringArg(args, 'layout_type');
  const rowGrandTotals = optionalBooleanArg(args, 'show_row_grand_totals');
  const columnGrandTotals = optionalBooleanArg(args, 'show_column_grand_totals');
  const bandedRows = optionalBooleanArg(args, 'show_banded_rows');
  const bandedColumns = optionalBooleanArg(args, 'show_banded_columns');

  if (layoutType !== undefined) {
    pivot.layout ??= {};
    pivot.layout.layoutType = layoutType;
    applied.push('layoutType');
  }
  if (rowGrandTotals !== undefined) {
    pivot.rowGrandTotals = rowGrandTotals;
    applied.push('rowGrandTotals');
  }
  if (columnGrandTotals !== undefined) {
    pivot.columnGrandTotals = columnGrandTotals;
    applied.push('columnGrandTotals');
  }
  if (bandedRows !== undefined) {
    pivot.showBandedRows = bandedRows;
    applied.push('showBandedRows');
  }
  if (bandedColumns !== undefined) {
    pivot.showBandedColumns = bandedColumns;
    applied.push('showBandedColumns');
  }

  await ctx.sync();
  return { name: args.name, applied };
};

export const handleRefreshAllPivots: ToolHandler = async (_args, ctx) => {
  assertPivotApi(ctx);
  ctx.workbook.pivotTables.refreshAll();
  await ctx.sync();
  return { refreshedAll: true };
};

export const handleDeletePivot: ToolHandler = async (args, ctx) => {
  const name = stringArg(args, 'name');
  const pivot = pivotByName(args, ctx) as unknown as { delete: () => void };
  pivot.delete();
  await ctx.sync();
  return { name, deleted: true };
};
