import type { ToolSpec } from '../../types';
import type { ToolHandler } from '../executor';
import { enumArg, numberArg, optionalNumberArg, optionalStringArg, stringArg } from './args';

export const ADD_SHEET: ToolSpec = {
  name: 'add_sheet',
  description: 'Add a new worksheet. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      name: { type: 'string', description: 'Optional worksheet name' },
      activate: { type: 'boolean', description: 'Activate the new worksheet. Default: true' },
    },
    required: ['workbook_id'],
  },
  mutating: true,
};

export const RENAME_SHEET: ToolSpec = {
  name: 'rename_sheet',
  description: 'Rename a worksheet. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Current worksheet name' },
      new_name: { type: 'string', description: 'New worksheet name' },
    },
    required: ['workbook_id', 'sheet', 'new_name'],
  },
  mutating: true,
};

export const DELETE_SHEET: ToolSpec = {
  name: 'delete_sheet',
  description: 'Delete a worksheet. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
    },
    required: ['workbook_id', 'sheet'],
  },
  mutating: true,
};

export const COPY_SHEET: ToolSpec = {
  name: 'copy_sheet',
  description: 'Copy a worksheet. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet to copy' },
      new_name: { type: 'string', description: 'Optional copied worksheet name' },
      position: { type: 'string', description: 'none, before, after, beginning, or end. Default: end' },
      relative_to: { type: 'string', description: 'Worksheet used for before/after positioning' },
    },
    required: ['workbook_id', 'sheet'],
  },
  mutating: true,
};

export const MOVE_SHEET: ToolSpec = {
  name: 'move_sheet',
  description: 'Move a worksheet to a zero-based workbook position. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      position: { type: 'number', description: 'Zero-based worksheet position' },
    },
    required: ['workbook_id', 'sheet', 'position'],
  },
  mutating: true,
};

export const HIDE_SHEET: ToolSpec = {
  name: 'hide_sheet',
  description: 'Set worksheet visibility. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      visibility: { type: 'string', description: 'visible, hidden, or very_hidden. Default: hidden' },
    },
    required: ['workbook_id', 'sheet'],
  },
  mutating: true,
};

export const ACTIVATE_SHEET: ToolSpec = {
  name: 'activate_sheet',
  description: 'Activate a worksheet. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
    },
    required: ['workbook_id', 'sheet'],
  },
  mutating: true,
};

export const FREEZE_PANES: ToolSpec = {
  name: 'freeze_panes',
  description: 'Freeze rows, columns, a range, or unfreeze panes on a worksheet. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      mode: { type: 'string', description: 'rows, columns, at, or none' },
      count: { type: 'number', description: 'Number of rows or columns for rows/columns mode' },
      address: { type: 'string', description: 'Range address for at mode' },
    },
    required: ['workbook_id', 'sheet', 'mode'],
  },
  mutating: true,
};

export const SHEET_SPECS: ToolSpec[] = [
  ADD_SHEET,
  RENAME_SHEET,
  DELETE_SHEET,
  COPY_SHEET,
  MOVE_SHEET,
  HIDE_SHEET,
  ACTIVATE_SHEET,
  FREEZE_PANES,
];

const POSITIONS = ['none', 'before', 'after', 'beginning', 'end'] as const;
const VISIBILITIES = ['visible', 'hidden', 'very_hidden'] as const;
const FREEZE_MODES = ['rows', 'columns', 'at', 'none'] as const;

function visibilityValue(value: (typeof VISIBILITIES)[number]): Excel.SheetVisibility {
  if (value === 'visible') return 'Visible' as Excel.SheetVisibility;
  if (value === 'very_hidden') return 'VeryHidden' as Excel.SheetVisibility;
  return 'Hidden' as Excel.SheetVisibility;
}

function positionValue(value: (typeof POSITIONS)[number]): Excel.WorksheetPositionType {
  if (value === 'before') return 'Before' as Excel.WorksheetPositionType;
  if (value === 'after') return 'After' as Excel.WorksheetPositionType;
  if (value === 'beginning') return 'Beginning' as Excel.WorksheetPositionType;
  if (value === 'end') return 'End' as Excel.WorksheetPositionType;
  return 'None' as Excel.WorksheetPositionType;
}

export const handleAddSheet: ToolHandler = async (args, ctx) => {
  const sheet = ctx.workbook.worksheets.add(optionalStringArg(args, 'name'));
  if (args.activate !== false) sheet.activate();
  sheet.load('name,position,visibility');
  await ctx.sync();
  return { name: sheet.name, position: sheet.position, visible: sheet.visibility === 'Visible' };
};

export const handleRenameSheet: ToolHandler = async (args, ctx) => {
  const ws = ctx.workbook.worksheets.getItem(stringArg(args, 'sheet'));
  ws.name = stringArg(args, 'new_name');
  await ctx.sync();
  return { sheet: args.sheet, renamedTo: args.new_name };
};

export const handleDeleteSheet: ToolHandler = async (args, ctx) => {
  ctx.workbook.worksheets.getItem(stringArg(args, 'sheet')).delete();
  await ctx.sync();
  return { sheet: args.sheet, deleted: true };
};

export const handleCopySheet: ToolHandler = async (args, ctx) => {
  const source = ctx.workbook.worksheets.getItem(stringArg(args, 'sheet'));
  const position = positionValue(enumArg(args, 'position', POSITIONS, 'end'));
  const relativeName = optionalStringArg(args, 'relative_to');
  const relative = relativeName ? ctx.workbook.worksheets.getItem(relativeName) : undefined;
  const copied = source.copy(position, relative);
  const newName = optionalStringArg(args, 'new_name');
  if (newName) copied.name = newName;
  copied.load('name,position');
  await ctx.sync();
  return { source: args.sheet, name: copied.name, position: copied.position };
};

export const handleMoveSheet: ToolHandler = async (args, ctx) => {
  const ws = ctx.workbook.worksheets.getItem(stringArg(args, 'sheet'));
  ws.position = numberArg(args, 'position');
  await ctx.sync();
  return { sheet: args.sheet, position: args.position };
};

export const handleHideSheet: ToolHandler = async (args, ctx) => {
  const visibility = visibilityValue(enumArg(args, 'visibility', VISIBILITIES, 'hidden'));
  ctx.workbook.worksheets.getItem(stringArg(args, 'sheet')).visibility = visibility;
  await ctx.sync();
  return { sheet: args.sheet, visibility };
};

export const handleActivateSheet: ToolHandler = async (args, ctx) => {
  ctx.workbook.worksheets.getItem(stringArg(args, 'sheet')).activate();
  await ctx.sync();
  return { sheet: args.sheet, activated: true };
};

export const handleFreezePanes: ToolHandler = async (args, ctx) => {
  const ws = ctx.workbook.worksheets.getItem(stringArg(args, 'sheet'));
  const mode = enumArg(args, 'mode', FREEZE_MODES);
  if (mode === 'none') ws.freezePanes.unfreeze();
  if (mode === 'rows') ws.freezePanes.freezeRows(optionalNumberArg(args, 'count') ?? 1);
  if (mode === 'columns') ws.freezePanes.freezeColumns(optionalNumberArg(args, 'count') ?? 1);
  if (mode === 'at') ws.freezePanes.freezeAt(stringArg(args, 'address'));
  await ctx.sync();
  return { sheet: args.sheet, mode };
};
