import type { ToolSpec } from '../../types';
import type { ToolHandler } from '../executor';
import { optionalBooleanArg, optionalStringArg, stringArg } from './args';

export const CREATE_NAMED_RANGE: ToolSpec = {
  name: 'create_named_range',
  description: 'Create a workbook- or worksheet-scoped named range/formula. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      name: { type: 'string', description: 'Name to create' },
      reference: { type: 'string', description: 'A formula or range reference, e.g. "=Sheet1!$A$1:$D$10"' },
      comment: { type: 'string', description: 'Optional name comment' },
      sheet: { type: 'string', description: 'Optional worksheet name for sheet-scoped names' },
      visible: { type: 'boolean', description: 'Whether the name is visible. Default: true' },
    },
    required: ['workbook_id', 'name', 'reference'],
  },
  mutating: true,
};

export const UPDATE_NAMED_RANGE: ToolSpec = {
  name: 'update_named_range',
  description: 'Update an existing named range/formula reference, comment, or visibility. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      name: { type: 'string', description: 'Existing name' },
      reference: { type: 'string', description: 'New formula or range reference' },
      comment: { type: 'string', description: 'New comment' },
      visible: { type: 'boolean', description: 'Whether the name is visible' },
      sheet: { type: 'string', description: 'Optional worksheet scope' },
    },
    required: ['workbook_id', 'name'],
  },
  mutating: true,
};

export const DELETE_NAMED_RANGE: ToolSpec = {
  name: 'delete_named_range',
  description: 'Delete a workbook- or worksheet-scoped named range/formula. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      name: { type: 'string', description: 'Name to delete' },
      sheet: { type: 'string', description: 'Optional worksheet scope' },
    },
    required: ['workbook_id', 'name'],
  },
  mutating: true,
};

export const NAME_SPECS: ToolSpec[] = [
  CREATE_NAMED_RANGE,
  UPDATE_NAMED_RANGE,
  DELETE_NAMED_RANGE,
];

function namesCollection(args: Record<string, unknown>, ctx: Excel.RequestContext): Excel.NamedItemCollection {
  const sheet = optionalStringArg(args, 'sheet');
  return sheet ? ctx.workbook.worksheets.getItem(sheet).names : ctx.workbook.names;
}

export const handleCreateNamedRange: ToolHandler = async (args, ctx) => {
  const item = namesCollection(args, ctx).add(stringArg(args, 'name'), stringArg(args, 'reference'), optionalStringArg(args, 'comment'));
  const visible = optionalBooleanArg(args, 'visible');
  if (visible !== undefined) item.visible = visible;
  item.load('name,formula,scope,comment,visible');
  await ctx.sync();
  return item.toJSON();
};

export const handleUpdateNamedRange: ToolHandler = async (args, ctx) => {
  const item = namesCollection(args, ctx).getItem(stringArg(args, 'name'));
  const reference = optionalStringArg(args, 'reference');
  if (reference) item.formula = reference;
  const comment = optionalStringArg(args, 'comment');
  if (comment !== undefined) item.comment = comment;
  const visible = optionalBooleanArg(args, 'visible');
  if (visible !== undefined) item.visible = visible;
  item.load('name,formula,scope,comment,visible');
  await ctx.sync();
  return item.toJSON();
};

export const handleDeleteNamedRange: ToolHandler = async (args, ctx) => {
  namesCollection(args, ctx).getItem(stringArg(args, 'name')).delete();
  await ctx.sync();
  return { name: args.name, deleted: true };
};
