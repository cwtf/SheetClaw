import type { ToolSpec } from '../../types';
import type { ToolHandler } from '../executor';
import { enumArg, optionalBooleanArg, optionalStringArg, stringArg } from './args';

export const INSERT_CELLS: ToolSpec = {
  name: 'insert_cells',
  description: 'Insert cells at a range, shifting existing cells down or right. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range address' },
      shift: { type: 'string', description: 'down or right. Default: down' },
    },
    required: ['workbook_id', 'sheet', 'address'],
  },
  mutating: true,
};

export const DELETE_CELLS: ToolSpec = {
  name: 'delete_cells',
  description: 'Delete cells at a range, shifting remaining cells up or left. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range address' },
      shift: { type: 'string', description: 'up or left. Default: up' },
    },
    required: ['workbook_id', 'sheet', 'address'],
  },
  mutating: true,
};

export const INSERT_ROWS: ToolSpec = {
  name: 'insert_rows',
  description: 'Insert whole worksheet rows at the rows intersecting a range. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range whose rows should be inserted' },
    },
    required: ['workbook_id', 'sheet', 'address'],
  },
  mutating: true,
};

export const DELETE_ROWS: ToolSpec = {
  name: 'delete_rows',
  description: 'Delete whole worksheet rows intersecting a range. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range whose rows should be deleted' },
    },
    required: ['workbook_id', 'sheet', 'address'],
  },
  mutating: true,
};

export const INSERT_COLUMNS: ToolSpec = {
  name: 'insert_columns',
  description: 'Insert whole worksheet columns at the columns intersecting a range. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range whose columns should be inserted' },
    },
    required: ['workbook_id', 'sheet', 'address'],
  },
  mutating: true,
};

export const DELETE_COLUMNS: ToolSpec = {
  name: 'delete_columns',
  description: 'Delete whole worksheet columns intersecting a range. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range whose columns should be deleted' },
    },
    required: ['workbook_id', 'sheet', 'address'],
  },
  mutating: true,
};

export const MERGE_RANGE: ToolSpec = {
  name: 'merge_range',
  description: 'Merge cells in a range. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range address' },
      across: { type: 'boolean', description: 'Merge cells across each row instead of all together. Default: false' },
    },
    required: ['workbook_id', 'sheet', 'address'],
  },
  mutating: true,
};

export const UNMERGE_RANGE: ToolSpec = {
  name: 'unmerge_range',
  description: 'Unmerge cells in a range. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range address' },
    },
    required: ['workbook_id', 'sheet', 'address'],
  },
  mutating: true,
};

export const FIND_REPLACE: ToolSpec = {
  name: 'find_replace',
  description: 'Find and replace text in a worksheet range or whole sheet. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'Optional A1 range address. Omit to search the sheet.' },
      find: { type: 'string', description: 'Text to find' },
      replace: { type: 'string', description: 'Replacement text' },
      complete_match: { type: 'boolean', description: 'Match entire cell contents. Default: false' },
      match_case: { type: 'boolean', description: 'Case-sensitive match. Default: false' },
    },
    required: ['workbook_id', 'sheet', 'find', 'replace'],
  },
  mutating: true,
};

export const GET_SPECIAL_CELLS: ToolSpec = {
  name: 'get_special_cells',
  description: 'Find special cells in a range, such as blanks, formulas, constants, visible cells, data validations, or conditional formats.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range address' },
      cell_type: { type: 'string', description: 'blanks, constants, formulas, visible, data_validations, conditional_formats, same_data_validation, or same_conditional_format' },
      value_type: { type: 'string', description: 'Optional value type: all, errors, logical, numbers, text, numbers_text, etc.' },
    },
    required: ['workbook_id', 'sheet', 'address', 'cell_type'],
  },
  mutating: false,
};

export const RANGE_OP_SPECS: ToolSpec[] = [
  INSERT_CELLS,
  DELETE_CELLS,
  INSERT_ROWS,
  DELETE_ROWS,
  INSERT_COLUMNS,
  DELETE_COLUMNS,
  MERGE_RANGE,
  UNMERGE_RANGE,
  FIND_REPLACE,
  GET_SPECIAL_CELLS,
];

const INSERT_SHIFTS = ['down', 'right'] as const;
const DELETE_SHIFTS = ['up', 'left'] as const;
const SPECIAL_TYPES = ['blanks', 'constants', 'formulas', 'visible', 'data_validations', 'conditional_formats', 'same_data_validation', 'same_conditional_format'] as const;

function range(args: Record<string, unknown>, ctx: Excel.RequestContext): Excel.Range {
  return ctx.workbook.worksheets.getItem(stringArg(args, 'sheet')).getRange(stringArg(args, 'address'));
}

function pascal(value: string): string {
  return value.replace(/(^|_)([a-z])/g, (_match, _sep: string, char: string) => char.toUpperCase());
}

export const handleInsertCells: ToolHandler = async (args, ctx) => {
  range(args, ctx).insert(enumArg(args, 'shift', INSERT_SHIFTS, 'down') === 'right' ? 'Right' : 'Down');
  await ctx.sync();
  return { sheet: args.sheet, address: args.address, inserted: true };
};

export const handleDeleteCells: ToolHandler = async (args, ctx) => {
  range(args, ctx).delete(enumArg(args, 'shift', DELETE_SHIFTS, 'up') === 'left' ? 'Left' : 'Up');
  await ctx.sync();
  return { sheet: args.sheet, address: args.address, deleted: true };
};

export const handleInsertRows: ToolHandler = async (args, ctx) => {
  range(args, ctx).getEntireRow().insert('Down');
  await ctx.sync();
  return { sheet: args.sheet, address: args.address, rowsInserted: true };
};

export const handleDeleteRows: ToolHandler = async (args, ctx) => {
  range(args, ctx).getEntireRow().delete('Up');
  await ctx.sync();
  return { sheet: args.sheet, address: args.address, rowsDeleted: true };
};

export const handleInsertColumns: ToolHandler = async (args, ctx) => {
  range(args, ctx).getEntireColumn().insert('Right');
  await ctx.sync();
  return { sheet: args.sheet, address: args.address, columnsInserted: true };
};

export const handleDeleteColumns: ToolHandler = async (args, ctx) => {
  range(args, ctx).getEntireColumn().delete('Left');
  await ctx.sync();
  return { sheet: args.sheet, address: args.address, columnsDeleted: true };
};

export const handleMergeRange: ToolHandler = async (args, ctx) => {
  range(args, ctx).merge(optionalBooleanArg(args, 'across') ?? false);
  await ctx.sync();
  return { sheet: args.sheet, address: args.address, merged: true };
};

export const handleUnmergeRange: ToolHandler = async (args, ctx) => {
  range(args, ctx).unmerge();
  await ctx.sync();
  return { sheet: args.sheet, address: args.address, unmerged: true };
};

export const handleFindReplace: ToolHandler = async (args, ctx) => {
  const target = optionalStringArg(args, 'address')
    ? range(args, ctx)
    : ctx.workbook.worksheets.getItem(stringArg(args, 'sheet'));
  const result = target.replaceAll(stringArg(args, 'find'), stringArg(args, 'replace'), {
    completeMatch: optionalBooleanArg(args, 'complete_match') ?? false,
    matchCase: optionalBooleanArg(args, 'match_case') ?? false,
  });
  await ctx.sync();
  return { sheet: args.sheet, address: optionalStringArg(args, 'address'), replaced: result.value };
};

export const handleGetSpecialCells: ToolHandler = async (args, ctx) => {
  const cellType = pascal(enumArg(args, 'cell_type', SPECIAL_TYPES)) as Excel.SpecialCellType;
  const valueType = optionalStringArg(args, 'value_type');
  const areas = range(args, ctx).getSpecialCellsOrNullObject(cellType, valueType ? pascal(valueType) as Excel.SpecialCellValueType : undefined);
  areas.load('isNullObject,address,areaCount');
  await ctx.sync();
  return areas.isNullObject ? { address: null, areaCount: 0 } : { address: areas.address, areaCount: areas.areaCount };
};
