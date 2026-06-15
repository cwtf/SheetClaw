import type { ToolSpec } from '../../types';
import type { ToolHandler } from '../executor';
import { optionalBooleanArg, optionalNumberArg, optionalStringArg, stringArg } from './args';

export const LIST_COMMENTS: ToolSpec = {
  name: 'list_comments',
  description: 'List threaded comments in a workbook or worksheet.',
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

export const ADD_COMMENT: ToolSpec = {
  name: 'add_comment',
  description: 'Add a threaded comment to a single cell. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'Single-cell A1 address' },
      content: { type: 'string', description: 'Comment text' },
    },
    required: ['workbook_id', 'sheet', 'address', 'content'],
  },
  mutating: true,
};

export const REPLY_TO_COMMENT: ToolSpec = {
  name: 'reply_to_comment',
  description: 'Reply to a threaded comment by ID or by cell. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      comment_id: { type: 'string', description: 'Comment ID' },
      sheet: { type: 'string', description: 'Worksheet name if replying by cell' },
      address: { type: 'string', description: 'Single-cell A1 address if replying by cell' },
      content: { type: 'string', description: 'Reply text' },
    },
    required: ['workbook_id', 'content'],
  },
  mutating: true,
};

export const DELETE_COMMENT: ToolSpec = {
  name: 'delete_comment',
  description: 'Delete a threaded comment by ID or by cell. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      comment_id: { type: 'string', description: 'Comment ID' },
      sheet: { type: 'string', description: 'Worksheet name if deleting by cell' },
      address: { type: 'string', description: 'Single-cell A1 address if deleting by cell' },
    },
    required: ['workbook_id'],
  },
  mutating: true,
};

export const ADD_NOTE: ToolSpec = {
  name: 'add_note',
  description: 'Add or update an Excel note on a single cell. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'Single-cell A1 address' },
      content: { type: 'string', description: 'Note text' },
      visible: { type: 'boolean', description: 'Show the note' },
      width: { type: 'number', description: 'Note width' },
      height: { type: 'number', description: 'Note height' },
    },
    required: ['workbook_id', 'sheet', 'address', 'content'],
  },
  mutating: true,
};

export const DELETE_NOTE: ToolSpec = {
  name: 'delete_note',
  description: 'Delete an Excel note from a single cell. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'Single-cell A1 address' },
    },
    required: ['workbook_id', 'sheet', 'address'],
  },
  mutating: true,
};

export const COMMENT_SPECS: ToolSpec[] = [
  LIST_COMMENTS,
  ADD_COMMENT,
  REPLY_TO_COMMENT,
  DELETE_COMMENT,
  ADD_NOTE,
  DELETE_NOTE,
];

function fullAddress(args: Record<string, unknown>): string {
  const address = stringArg(args, 'address');
  return address.includes('!') ? address : `${stringArg(args, 'sheet')}!${address}`;
}

function comments(args: Record<string, unknown>, ctx: Excel.RequestContext): Excel.CommentCollection {
  const sheet = optionalStringArg(args, 'sheet');
  return sheet ? ctx.workbook.worksheets.getItem(sheet).comments : ctx.workbook.comments;
}

function commentByIdOrCell(args: Record<string, unknown>, ctx: Excel.RequestContext): Excel.Comment {
  const id = optionalStringArg(args, 'comment_id');
  if (id) return ctx.workbook.comments.getItem(id);
  return ctx.workbook.comments.getItemByCell(fullAddress(args));
}

export const handleListComments: ToolHandler = async (args, ctx) => {
  const collection = comments(args, ctx);
  collection.load('items/id,items/content,items/authorName,items/creationDate,items/resolved');
  await ctx.sync();
  return collection.items.map(comment => ({
    id: comment.id,
    content: comment.content,
    authorName: comment.authorName,
    creationDate: comment.creationDate,
    resolved: comment.resolved,
  }));
};

export const handleAddComment: ToolHandler = async (args, ctx) => {
  const comment = ctx.workbook.comments.add(fullAddress(args), stringArg(args, 'content'), 'Plain');
  comment.load('id,content');
  await ctx.sync();
  return { id: comment.id, content: comment.content };
};

export const handleReplyToComment: ToolHandler = async (args, ctx) => {
  const reply = commentByIdOrCell(args, ctx).replies.add(stringArg(args, 'content'), 'Plain');
  reply.load('id,content');
  await ctx.sync();
  return { id: reply.id, content: reply.content };
};

export const handleDeleteComment: ToolHandler = async (args, ctx) => {
  commentByIdOrCell(args, ctx).delete();
  await ctx.sync();
  return { comment: optionalStringArg(args, 'comment_id') ?? fullAddress(args), deleted: true };
};

export const handleAddNote: ToolHandler = async (args, ctx) => {
  const note = ctx.workbook.worksheets.getItem(stringArg(args, 'sheet')).notes.add(stringArg(args, 'address'), stringArg(args, 'content'));
  const visible = optionalBooleanArg(args, 'visible');
  if (visible !== undefined) note.visible = visible;
  const width = optionalNumberArg(args, 'width');
  if (width !== undefined) note.width = width;
  const height = optionalNumberArg(args, 'height');
  if (height !== undefined) note.height = height;
  note.load('content,authorName,visible,width,height');
  await ctx.sync();
  return note.toJSON();
};

export const handleDeleteNote: ToolHandler = async (args, ctx) => {
  ctx.workbook.worksheets.getItem(stringArg(args, 'sheet')).notes.getItem(stringArg(args, 'address')).delete();
  await ctx.sync();
  return { sheet: args.sheet, address: args.address, deleted: true };
};
