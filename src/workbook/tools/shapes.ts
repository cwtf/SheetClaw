import type { ToolSpec } from '../../types';
import type { ToolHandler } from '../executor';
import { optionalNumberArg, optionalStringArg, stringArg } from './args';

const shapePositionProps = {
  left: { type: 'number', description: 'Left position in points from the worksheet edge' },
  top: { type: 'number', description: 'Top position in points from the worksheet edge' },
  width: { type: 'number', description: 'Shape width in points' },
  height: { type: 'number', description: 'Shape height in points' },
  name: { type: 'string', description: 'Optional shape name' },
} as const;

export const LIST_SHAPES: ToolSpec = {
  name: 'list_shapes',
  description: 'List shapes, images, and text boxes on a worksheet.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
    },
    required: ['workbook_id', 'sheet'],
  },
  mutating: false,
};

export const ADD_IMAGE: ToolSpec = {
  name: 'add_image',
  description: 'Add a PNG or JPEG image to a worksheet from a base64 string. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      base64_image: { type: 'string', description: 'Base64 image payload without a data URL prefix' },
      ...shapePositionProps,
    },
    required: ['workbook_id', 'sheet', 'base64_image'],
  },
  mutating: true,
};

export const ADD_TEXTBOX: ToolSpec = {
  name: 'add_textbox',
  description: 'Add a text box to a worksheet. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      text: { type: 'string', description: 'Text box contents' },
      ...shapePositionProps,
    },
    required: ['workbook_id', 'sheet', 'text'],
  },
  mutating: true,
};

export const ADD_SHAPE: ToolSpec = {
  name: 'add_shape',
  description: 'Add a geometric shape such as Rectangle, Ellipse, RightArrow, or Callout1. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      shape_type: { type: 'string', description: 'Excel geometric shape type, defaults to Rectangle' },
      fill_color: { type: 'string', description: 'Optional fill color such as #E6F2FF' },
      line_color: { type: 'string', description: 'Optional outline color such as #1F4E79' },
      ...shapePositionProps,
    },
    required: ['workbook_id', 'sheet'],
  },
  mutating: true,
};

export const MOVE_SHAPE: ToolSpec = {
  name: 'move_shape',
  description: 'Move an existing worksheet shape by name or ID. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      shape: { type: 'string', description: 'Shape name or ID' },
      left: shapePositionProps.left,
      top: shapePositionProps.top,
    },
    required: ['workbook_id', 'sheet', 'shape'],
  },
  mutating: true,
};

export const RESIZE_SHAPE: ToolSpec = {
  name: 'resize_shape',
  description: 'Resize an existing worksheet shape by name or ID. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      shape: { type: 'string', description: 'Shape name or ID' },
      width: shapePositionProps.width,
      height: shapePositionProps.height,
    },
    required: ['workbook_id', 'sheet', 'shape'],
  },
  mutating: true,
};

export const DELETE_SHAPE: ToolSpec = {
  name: 'delete_shape',
  description: 'Delete a worksheet shape by name or ID. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      shape: { type: 'string', description: 'Shape name or ID' },
    },
    required: ['workbook_id', 'sheet', 'shape'],
  },
  mutating: true,
};

export const SHAPE_SPECS: ToolSpec[] = [
  LIST_SHAPES,
  ADD_IMAGE,
  ADD_TEXTBOX,
  ADD_SHAPE,
  MOVE_SHAPE,
  RESIZE_SHAPE,
  DELETE_SHAPE,
];

function worksheet(args: Record<string, unknown>, ctx: Excel.RequestContext): Excel.Worksheet {
  return ctx.workbook.worksheets.getItem(stringArg(args, 'sheet'));
}

function shapeFor(args: Record<string, unknown>, ctx: Excel.RequestContext): Excel.Shape {
  return worksheet(args, ctx).shapes.getItem(stringArg(args, 'shape'));
}

function applyShapeOptions(shape: Excel.Shape, args: Record<string, unknown>): void {
  const left = optionalNumberArg(args, 'left');
  const top = optionalNumberArg(args, 'top');
  const width = optionalNumberArg(args, 'width');
  const height = optionalNumberArg(args, 'height');
  const name = optionalStringArg(args, 'name');
  if (left !== undefined) shape.left = left;
  if (top !== undefined) shape.top = top;
  if (width !== undefined) shape.width = width;
  if (height !== undefined) shape.height = height;
  if (name !== undefined) shape.name = name;
}

async function loadShape(shape: Excel.Shape, ctx: Excel.RequestContext): Promise<Record<string, unknown>> {
  shape.load('id,name,type,left,top,width,height,visible');
  await ctx.sync();
  return shape.toJSON() as Record<string, unknown>;
}

export const handleListShapes: ToolHandler = async (args, ctx) => {
  const shapes = worksheet(args, ctx).shapes;
  shapes.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height,items/visible');
  await ctx.sync();
  return shapes.items.map(shape => shape.toJSON());
};

export const handleAddImage: ToolHandler = async (args, ctx) => {
  const shape = worksheet(args, ctx).shapes.addImage(stringArg(args, 'base64_image'));
  applyShapeOptions(shape, args);
  return { created: true, ...(await loadShape(shape, ctx)) };
};

export const handleAddTextbox: ToolHandler = async (args, ctx) => {
  const shape = worksheet(args, ctx).shapes.addTextBox(stringArg(args, 'text'));
  applyShapeOptions(shape, args);
  return { created: true, ...(await loadShape(shape, ctx)) };
};

export const handleAddShape: ToolHandler = async (args, ctx) => {
  const shapeType = (optionalStringArg(args, 'shape_type') ?? 'Rectangle') as Excel.GeometricShapeType;
  const shape = worksheet(args, ctx).shapes.addGeometricShape(shapeType);
  applyShapeOptions(shape, args);
  const fillColor = optionalStringArg(args, 'fill_color');
  const lineColor = optionalStringArg(args, 'line_color');
  if (fillColor !== undefined) shape.fill.setSolidColor(fillColor);
  if (lineColor !== undefined) shape.lineFormat.color = lineColor;
  return { created: true, ...(await loadShape(shape, ctx)) };
};

export const handleMoveShape: ToolHandler = async (args, ctx) => {
  const shape = shapeFor(args, ctx);
  applyShapeOptions(shape, args);
  return { moved: true, ...(await loadShape(shape, ctx)) };
};

export const handleResizeShape: ToolHandler = async (args, ctx) => {
  const shape = shapeFor(args, ctx);
  applyShapeOptions(shape, args);
  return { resized: true, ...(await loadShape(shape, ctx)) };
};

export const handleDeleteShape: ToolHandler = async (args, ctx) => {
  const name = stringArg(args, 'shape');
  shapeFor(args, ctx).delete();
  await ctx.sync();
  return { deleted: true, shape: name };
};
