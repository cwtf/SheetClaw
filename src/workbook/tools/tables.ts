import type { ToolSpec } from '../../types';
import type { ToolHandler } from '../executor';
import { enumArg, matrixArg, optionalBooleanArg, optionalNumberArg, optionalStringArg, stringArg } from './args';

export const LIST_TABLES: ToolSpec = {
  name: 'list_tables',
  description: 'List Excel tables in a workbook or worksheet, including name, sheet, range, style, header, and totals settings.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Optional worksheet name' },
    },
    required: ['workbook_id'],
  },
  mutating: false,
};

export const CREATE_TABLE: ToolSpec = {
  name: 'create_table',
  description: 'Create an Excel table from an existing range. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range address for the table' },
      has_headers: { type: 'boolean', description: 'Whether the first row contains headers. Default: true' },
      name: { type: 'string', description: 'Optional table name' },
      style: { type: 'string', description: 'Optional Excel table style, e.g. "TableStyleMedium2"' },
    },
    required: ['workbook_id', 'sheet', 'address'],
  },
  mutating: true,
};

export const RESIZE_TABLE: ToolSpec = {
  name: 'resize_table',
  description: 'Resize an existing Excel table to a new range. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      table: { type: 'string', description: 'Table name' },
      address: { type: 'string', description: 'New A1 range address' },
    },
    required: ['workbook_id', 'table', 'address'],
  },
  mutating: true,
};

export const ADD_TABLE_ROWS: ToolSpec = {
  name: 'add_table_rows',
  description: 'Add rows to an Excel table. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      table: { type: 'string', description: 'Table name' },
      values: { type: 'array', items: { type: 'array', items: { description: 'Cell value' } }, description: 'Row-major values to add' },
      index: { type: 'number', description: 'Optional zero-based insertion index. Omit to append.' },
      always_insert: { type: 'boolean', description: 'Whether to always insert rows. Default: true' },
    },
    required: ['workbook_id', 'table', 'values'],
  },
  mutating: true,
};

export const ADD_TABLE_COLUMNS: ToolSpec = {
  name: 'add_table_columns',
  description: 'Add a column to an Excel table. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      table: { type: 'string', description: 'Table name' },
      name: { type: 'string', description: 'New column name' },
      values: { type: 'array', items: { type: 'array', items: { description: 'Cell value' } }, description: 'Optional one-column values, including header if needed' },
      index: { type: 'number', description: 'Optional zero-based insertion index. Omit to append.' },
    },
    required: ['workbook_id', 'table', 'name'],
  },
  mutating: true,
};

export const SET_TABLE_STYLE: ToolSpec = {
  name: 'set_table_style',
  description: 'Set table style and header/totals/banded display options. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      table: { type: 'string', description: 'Table name' },
      style: { type: 'string', description: 'Excel table style, e.g. "TableStyleMedium2"' },
      show_headers: { type: 'boolean', description: 'Show header row' },
      show_totals: { type: 'boolean', description: 'Show totals row' },
      show_banded_rows: { type: 'boolean', description: 'Show banded rows' },
      show_banded_columns: { type: 'boolean', description: 'Show banded columns' },
      show_filter_button: { type: 'boolean', description: 'Show filter buttons' },
    },
    required: ['workbook_id', 'table'],
  },
  mutating: true,
};

export const SET_TABLE_TOTALS: ToolSpec = {
  name: 'set_table_totals',
  description: 'Set totals row labels or functions for table columns. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      table: { type: 'string', description: 'Table name' },
      column: { type: 'string', description: 'Column name' },
      function: { type: 'string', description: 'Totals function: none, sum, average, count, count_numbers, min, max, standard_deviation, variance, custom' },
      label: { type: 'string', description: 'Optional totals row label for text columns' },
      formula: { type: 'string', description: 'Optional custom totals formula' },
    },
    required: ['workbook_id', 'table', 'column'],
  },
  mutating: true,
};

export const DELETE_TABLE: ToolSpec = {
  name: 'delete_table',
  description: 'Delete an Excel table. Data remains by default; set convert_to_range false to remove the table object only if host supports delete. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      table: { type: 'string', description: 'Table name' },
      convert_to_range: { type: 'boolean', description: 'Convert to ordinary range before deleting table object. Default: true' },
    },
    required: ['workbook_id', 'table'],
  },
  mutating: true,
};

export const TABLE_SPECS: ToolSpec[] = [
  LIST_TABLES,
  CREATE_TABLE,
  RESIZE_TABLE,
  ADD_TABLE_ROWS,
  ADD_TABLE_COLUMNS,
  SET_TABLE_STYLE,
  SET_TABLE_TOTALS,
  DELETE_TABLE,
];

const TOTAL_FUNCTIONS = ['none', 'sum', 'average', 'count', 'count_numbers', 'min', 'max', 'standard_deviation', 'variance', 'custom'] as const;

const TOTAL_FORMULAS: Record<string, number> = {
  average: 101,
  count: 103,
  count_numbers: 102,
  max: 104,
  min: 105,
  standard_deviation: 107,
  sum: 109,
  variance: 110,
};

function getTable(ctx: Excel.RequestContext, name: string): Excel.Table {
  return ctx.workbook.tables.getItem(name);
}

export const handleListTables: ToolHandler = async (args, ctx) => {
  const sheetName = optionalStringArg(args, 'sheet');
  const collection = sheetName
    ? ctx.workbook.worksheets.getItem(sheetName).tables
    : ctx.workbook.tables;
  collection.load('items/name,items/style,items/showHeaders,items/showTotals,items/showBandedRows,items/showBandedColumns,items/showFilterButton,items/worksheet/name');
  await ctx.sync();

  const rows = collection.items.map(table => {
    const range = table.getRange();
    range.load('address');
    return { table, range };
  });
  await ctx.sync();

  return rows.map(({ table, range }) => ({
    name: table.name,
    sheet: (table as Excel.Table & { worksheet?: Excel.Worksheet }).worksheet?.name,
    address: range.address,
    style: table.style,
    showHeaders: table.showHeaders,
    showTotals: table.showTotals,
    showBandedRows: table.showBandedRows,
    showBandedColumns: table.showBandedColumns,
    showFilterButton: table.showFilterButton,
  }));
};

export const handleCreateTable: ToolHandler = async (args, ctx) => {
  const sheet = ctx.workbook.worksheets.getItem(stringArg(args, 'sheet'));
  const table = sheet.tables.add(stringArg(args, 'address'), optionalBooleanArg(args, 'has_headers') ?? true);
  const name = optionalStringArg(args, 'name');
  const style = optionalStringArg(args, 'style');
  if (name) table.name = name;
  if (style) table.style = style;
  table.load('name');
  const range = table.getRange();
  range.load('address');
  await ctx.sync();
  return { name: table.name, address: range.address, created: true };
};

export const handleResizeTable: ToolHandler = async (args, ctx) => {
  const table = getTable(ctx, stringArg(args, 'table'));
  table.resize(stringArg(args, 'address'));
  await ctx.sync();
  return { table: args.table, resizedTo: args.address };
};

export const handleAddTableRows: ToolHandler = async (args, ctx) => {
  const table = getTable(ctx, stringArg(args, 'table'));
  const index = optionalNumberArg(args, 'index');
  const values = matrixArg(args, 'values');
  table.rows.add(index, values, optionalBooleanArg(args, 'always_insert') ?? true);
  await ctx.sync();
  return { table: args.table, rowsAdded: values.length };
};

export const handleAddTableColumns: ToolHandler = async (args, ctx) => {
  const table = getTable(ctx, stringArg(args, 'table'));
  const index = optionalNumberArg(args, 'index');
  const name = stringArg(args, 'name');
  const rawValues = args.values === undefined ? undefined : matrixArg(args, 'values');
  const values = rawValues ?? [[name]];
  const column = table.columns.add(index, values, name);
  column.load('name');
  await ctx.sync();
  return { table: args.table, column: column.name };
};

export const handleSetTableStyle: ToolHandler = async (args, ctx) => {
  const table = getTable(ctx, stringArg(args, 'table'));
  const applied: string[] = [];
  const style = optionalStringArg(args, 'style');
  if (style) { table.style = style; applied.push('style'); }
  const showHeaders = optionalBooleanArg(args, 'show_headers');
  if (showHeaders !== undefined) { table.showHeaders = showHeaders; applied.push('showHeaders'); }
  const showTotals = optionalBooleanArg(args, 'show_totals');
  if (showTotals !== undefined) { table.showTotals = showTotals; applied.push('showTotals'); }
  const showBandedRows = optionalBooleanArg(args, 'show_banded_rows');
  if (showBandedRows !== undefined) { table.showBandedRows = showBandedRows; applied.push('showBandedRows'); }
  const showBandedColumns = optionalBooleanArg(args, 'show_banded_columns');
  if (showBandedColumns !== undefined) { table.showBandedColumns = showBandedColumns; applied.push('showBandedColumns'); }
  const showFilterButton = optionalBooleanArg(args, 'show_filter_button');
  if (showFilterButton !== undefined) { table.showFilterButton = showFilterButton; applied.push('showFilterButton'); }
  await ctx.sync();
  return { table: args.table, applied };
};

export const handleSetTableTotals: ToolHandler = async (args, ctx) => {
  const table = getTable(ctx, stringArg(args, 'table'));
  table.showTotals = true;
  const columnName = stringArg(args, 'column');
  const column = table.columns.getItem(columnName);
  const totalsFunction = enumArg(args, 'function', TOTAL_FUNCTIONS, 'none');
  const label = optionalStringArg(args, 'label');
  const formula = optionalStringArg(args, 'formula');
  const totalCell = column.getTotalRowRange();
  if (label) {
    totalCell.values = [[label]];
  } else if (formula) {
    totalCell.formulas = [[formula]];
  } else if (totalsFunction !== 'none' && totalsFunction !== 'custom') {
    totalCell.formulas = [[`=SUBTOTAL(${TOTAL_FORMULAS[totalsFunction]},[${columnName}])`]];
  }
  await ctx.sync();
  return { table: args.table, column: args.column, totalsFunction, label, formula };
};

export const handleDeleteTable: ToolHandler = async (args, ctx) => {
  const table = getTable(ctx, stringArg(args, 'table'));
  if (optionalBooleanArg(args, 'convert_to_range') ?? true) {
    table.convertToRange();
  } else {
    table.delete();
  }
  await ctx.sync();
  return { table: args.table, deleted: true };
};
