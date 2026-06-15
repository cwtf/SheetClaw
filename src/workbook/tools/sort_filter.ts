import type { ToolSpec } from '../../types';
import type { ToolHandler } from '../executor';
import { ToolValidationError } from '../executor';
import { enumArg, optionalBooleanArg, optionalStringArg, stringArg } from './args';

export const SORT_RANGE: ToolSpec = {
  name: 'sort_range',
  description: 'Sort a worksheet range by one or more columns. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range address to sort' },
      fields: { type: 'array', items: { type: 'object', properties: {} }, description: 'Sort fields: [{key: 0-based column index, ascending: true, dataOption?: normal|text_as_number}]' },
      has_headers: { type: 'boolean', description: 'Whether the first row is headers. Default: true' },
      match_case: { type: 'boolean', description: 'Case-sensitive sort. Default: false' },
      orientation: { type: 'string', description: 'rows or columns. Default: rows' },
    },
    required: ['workbook_id', 'sheet', 'address', 'fields'],
  },
  mutating: true,
};

export const SORT_TABLE: ToolSpec = {
  name: 'sort_table',
  description: 'Sort an Excel table by one or more columns. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      table: { type: 'string', description: 'Table name' },
      fields: { type: 'array', items: { type: 'object', properties: {} }, description: 'Sort fields: [{key: column name or 0-based index, ascending: true}]' },
      match_case: { type: 'boolean', description: 'Case-sensitive sort. Default: false' },
    },
    required: ['workbook_id', 'table', 'fields'],
  },
  mutating: true,
};

export const APPLY_FILTER: ToolSpec = {
  name: 'apply_filter',
  description: 'Apply a filter to a range or table column. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name for range filtering' },
      address: { type: 'string', description: 'Range to filter when table is omitted' },
      table: { type: 'string', description: 'Optional table name' },
      column: { description: 'Column name for table filters or zero-based column index for range filters' },
      criterion1: { type: 'string', description: 'Primary criterion, e.g. ">100", "North", or "*" wildcard' },
      criterion2: { type: 'string', description: 'Secondary criterion' },
      operator: { type: 'string', description: 'and, or, top_items, bottom_items, top_percent, bottom_percent, values, dynamic' },
      values: { type: 'array', items: { type: 'string' }, description: 'Values filter list' },
      filter_on: { type: 'string', description: 'custom, values, top_items, bottom_items, top_percent, bottom_percent, dynamic, cell_color, font_color, icon' },
      dynamic_criteria: { type: 'string', description: 'Dynamic filter criteria such as today, this_month, above_average' },
    },
    required: ['workbook_id', 'column'],
  },
  mutating: true,
};

export const CLEAR_FILTERS: ToolSpec = {
  name: 'clear_filters',
  description: 'Clear filters from a worksheet range or table. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name for range filters' },
      address: { type: 'string', description: 'Optional filtered range address' },
      table: { type: 'string', description: 'Optional table name' },
    },
    required: ['workbook_id'],
  },
  mutating: true,
};

export const REAPPLY_FILTERS: ToolSpec = {
  name: 'reapply_filters',
  description: 'Reapply filters on a worksheet range or table. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name for range filters' },
      table: { type: 'string', description: 'Optional table name' },
    },
    required: ['workbook_id'],
  },
  mutating: true,
};

export const SORT_FILTER_SPECS: ToolSpec[] = [
  SORT_RANGE,
  SORT_TABLE,
  APPLY_FILTER,
  CLEAR_FILTERS,
  REAPPLY_FILTERS,
];

const FILTER_ON = ['custom', 'values', 'top_items', 'bottom_items', 'top_percent', 'bottom_percent', 'dynamic', 'cell_color', 'font_color', 'icon'] as const;
const FILTER_OPERATORS = ['and', 'or', 'top_items', 'bottom_items', 'top_percent', 'bottom_percent', 'values', 'dynamic'] as const;
const ORIENTATIONS = ['rows', 'columns'] as const;

function parseSortFields(value: unknown): Excel.SortField[] {
  if (!Array.isArray(value) || !value.length) {
    throw new ToolValidationError('"fields" must be a non-empty array');
  }
  return value.map((field, index) => {
    if (!field || typeof field !== 'object') {
      throw new ToolValidationError(`"fields[${index}]" must be an object`);
    }
    const row = field as Record<string, unknown>;
    const key = row.key;
    if (typeof key !== 'number' && typeof key !== 'string') {
      throw new ToolValidationError(`"fields[${index}].key" must be a number or string`);
    }
    return {
      key,
      ascending: row.ascending === undefined ? true : Boolean(row.ascending),
      dataOption: row.dataOption === 'text_as_number' ? 'TextAsNumber' : 'Normal',
    } as Excel.SortField;
  });
}

function filterCriteria(args: Record<string, unknown>): Excel.FilterCriteria {
  const values = args.values;
  const filterOn = enumArg(args, 'filter_on', FILTER_ON, values ? 'values' : 'custom');
  const operator = enumArg(args, 'operator', FILTER_OPERATORS, 'and');
  const criteria: Excel.FilterCriteria = {
    filterOn: filterOn.replace(/_([a-z])/g, (_, char: string) => char.toUpperCase()) as Excel.FilterOn,
  };
  const criterion1 = optionalStringArg(args, 'criterion1');
  const criterion2 = optionalStringArg(args, 'criterion2');
  if (criterion1) criteria.criterion1 = criterion1;
  if (criterion2) criteria.criterion2 = criterion2;
  if (operator) criteria.operator = operator === 'or' ? 'Or' : operator === 'and' ? 'And' : undefined;
  if (Array.isArray(values)) criteria.values = values as string[];
  const dynamicCriteria = optionalStringArg(args, 'dynamic_criteria');
  if (dynamicCriteria) {
    criteria.dynamicCriteria = dynamicCriteria.replace(/_([a-z])/g, (_, char: string) => char.toUpperCase()) as Excel.DynamicFilterCriteria;
  }
  return criteria;
}

function columnKey(value: unknown): number | string {
  if (typeof value === 'number' || typeof value === 'string') return value;
  throw new ToolValidationError('"column" must be a table column name or zero-based range column index');
}

export const handleSortRange: ToolHandler = async (args, ctx) => {
  const range = ctx.workbook.worksheets.getItem(stringArg(args, 'sheet')).getRange(stringArg(args, 'address'));
  range.sort.apply(
    parseSortFields(args.fields),
    optionalBooleanArg(args, 'match_case') ?? false,
    optionalBooleanArg(args, 'has_headers') ?? true,
    enumArg(args, 'orientation', ORIENTATIONS, 'rows') === 'columns' ? 'Columns' : 'Rows'
  );
  await ctx.sync();
  return { sheet: args.sheet, address: args.address, sorted: true };
};

export const handleSortTable: ToolHandler = async (args, ctx) => {
  const table = ctx.workbook.tables.getItem(stringArg(args, 'table'));
  table.sort.apply(parseSortFields(args.fields), optionalBooleanArg(args, 'match_case') ?? false);
  await ctx.sync();
  return { table: args.table, sorted: true };
};

export const handleApplyFilter: ToolHandler = async (args, ctx) => {
  const tableName = optionalStringArg(args, 'table');
  const criteria = filterCriteria(args);
  if (tableName) {
    const table = ctx.workbook.tables.getItem(tableName);
    table.columns.getItem(columnKey(args.column) as string).filter.apply(criteria);
    await ctx.sync();
    return { table: tableName, column: args.column, filtered: true };
  }

  const sheet = stringArg(args, 'sheet');
  const address = stringArg(args, 'address');
  const column = columnKey(args.column);
  if (typeof column !== 'number') throw new ToolValidationError('Range filters require a zero-based numeric "column"');
  ctx.workbook.worksheets.getItem(sheet).autoFilter.apply(address, column, criteria);
  await ctx.sync();
  return { sheet, address, column, filtered: true };
};

export const handleClearFilters: ToolHandler = async (args, ctx) => {
  const tableName = optionalStringArg(args, 'table');
  if (tableName) {
    ctx.workbook.tables.getItem(tableName).autoFilter.clearCriteria();
    await ctx.sync();
    return { table: tableName, cleared: true };
  }
  const sheet = stringArg(args, 'sheet');
  const ws = ctx.workbook.worksheets.getItem(sheet);
  ws.autoFilter.clearCriteria();
  await ctx.sync();
  return { sheet, cleared: true };
};

export const handleReapplyFilters: ToolHandler = async (args, ctx) => {
  const tableName = optionalStringArg(args, 'table');
  if (tableName) {
    ctx.workbook.tables.getItem(tableName).autoFilter.reapply();
    await ctx.sync();
    return { table: tableName, reapplied: true };
  }
  const sheet = stringArg(args, 'sheet');
  ctx.workbook.worksheets.getItem(sheet).autoFilter.reapply();
  await ctx.sync();
  return { sheet, reapplied: true };
};
