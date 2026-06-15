import type { ToolSpec } from '../../types';
import type { ToolHandler } from '../executor';
import { enumArg, optionalNumberArg, optionalStringArg, stringArg } from './args';

export const ADD_CONDITIONAL_FORMAT: ToolSpec = {
  name: 'add_conditional_format',
  description: 'Add conditional formatting to a range. Supports cell_value, contains_text, top_bottom, color_scale, and data_bar. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range address' },
      type: { type: 'string', description: 'cell_value, contains_text, top_bottom, color_scale, or data_bar' },
      operator: { type: 'string', description: 'For cell_value: greater_than, less_than, between, equal_to, not_equal_to, greater_than_or_equal_to, less_than_or_equal_to' },
      formula1: { type: 'string', description: 'Primary value/formula for cell value rules' },
      formula2: { type: 'string', description: 'Secondary value/formula for between rules' },
      text: { type: 'string', description: 'Text for contains_text rules' },
      rank: { type: 'number', description: 'Rank for top/bottom rules. Default: 10' },
      fill_color: { type: 'string', description: 'Fill color for simple formats' },
      font_color: { type: 'string', description: 'Font color for simple formats' },
      lowest_color: { type: 'string', description: 'Color scale low color' },
      midpoint_color: { type: 'string', description: 'Color scale midpoint color' },
      highest_color: { type: 'string', description: 'Color scale high color' },
      bar_color: { type: 'string', description: 'Data bar color' },
    },
    required: ['workbook_id', 'sheet', 'address', 'type'],
  },
  mutating: true,
};

export const LIST_CONDITIONAL_FORMATS: ToolSpec = {
  name: 'list_conditional_formats',
  description: 'List conditional formats active on a range.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range address' },
    },
    required: ['workbook_id', 'sheet', 'address'],
  },
  mutating: false,
};

export const CLEAR_CONDITIONAL_FORMATS: ToolSpec = {
  name: 'clear_conditional_formats',
  description: 'Clear conditional formatting from a range. Requires user confirmation.',
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

export const CONDITIONAL_FORMAT_SPECS: ToolSpec[] = [
  ADD_CONDITIONAL_FORMAT,
  LIST_CONDITIONAL_FORMATS,
  CLEAR_CONDITIONAL_FORMATS,
];

const FORMAT_TYPES = ['cell_value', 'contains_text', 'top_bottom', 'color_scale', 'data_bar'] as const;
const OPERATORS = ['greater_than', 'less_than', 'between', 'equal_to', 'not_equal_to', 'greater_than_or_equal_to', 'less_than_or_equal_to'] as const;

function pascal(value: string): string {
  return value.replace(/(^|_)([a-z])/g, (_match, _sep: string, char: string) => char.toUpperCase());
}

function formatRange(args: Record<string, unknown>, ctx: Excel.RequestContext): Excel.Range {
  return ctx.workbook.worksheets.getItem(stringArg(args, 'sheet')).getRange(stringArg(args, 'address'));
}

function applyBasicFormat(cf: unknown, args: Record<string, unknown>): void {
  const fillColor = optionalStringArg(args, 'fill_color');
  const fontColor = optionalStringArg(args, 'font_color');
  const format = (cf as { format?: { fill?: { color?: string }; font?: { color?: string } } }).format;
  if (fillColor && format?.fill) format.fill.color = fillColor;
  if (fontColor && format?.font) format.font.color = fontColor;
}

export const handleAddConditionalFormat: ToolHandler = async (args, ctx) => {
  const type = enumArg(args, 'type', FORMAT_TYPES);
  const range = formatRange(args, ctx);
  const addType =
    type === 'cell_value' ? 'CellValue' :
    type === 'contains_text' ? 'ContainsText' :
    type === 'top_bottom' ? 'TopBottom' :
    type === 'color_scale' ? 'ColorScale' :
    'DataBar';
  const cf = range.conditionalFormats.add(addType as Excel.ConditionalFormatType);
  const anyCf = cf as unknown as Record<string, unknown>;

  if (type === 'cell_value') {
    const cellValue = anyCf.cellValue as { rule?: unknown; format?: unknown };
    cellValue.rule = {
      formula1: stringArg(args, 'formula1'),
      formula2: optionalStringArg(args, 'formula2'),
      operator: pascal(enumArg(args, 'operator', OPERATORS, 'greater_than')),
    };
    applyBasicFormat(cellValue, args);
  }
  if (type === 'contains_text') {
    const containsText = anyCf.textComparison as { rule?: unknown; format?: unknown };
    containsText.rule = { operator: 'Contains', text: stringArg(args, 'text') };
    applyBasicFormat(containsText, args);
  }
  if (type === 'top_bottom') {
    const topBottom = anyCf.topBottom as { rule?: unknown; format?: unknown };
    topBottom.rule = { rank: optionalNumberArg(args, 'rank') ?? 10, type: 'TopItems' };
    applyBasicFormat(topBottom, args);
  }
  if (type === 'color_scale') {
    const colorScale = anyCf.colorScale as { criteria?: { minimum?: unknown; midpoint?: unknown; maximum?: unknown } };
    colorScale.criteria = {
      minimum: { type: 'LowestValue', color: optionalStringArg(args, 'lowest_color') ?? '#F8696B' },
      midpoint: { type: 'Percentile', formula: '50', color: optionalStringArg(args, 'midpoint_color') ?? '#FFEB84' },
      maximum: { type: 'HighestValue', color: optionalStringArg(args, 'highest_color') ?? '#63BE7B' },
    };
  }
  if (type === 'data_bar') {
    const dataBar = anyCf.dataBar as { barColor?: string };
    dataBar.barColor = optionalStringArg(args, 'bar_color') ?? '#638EC6';
  }

  cf.load('id,type');
  await ctx.sync();
  return { sheet: args.sheet, address: args.address, id: cf.id, type: cf.type };
};

export const handleListConditionalFormats: ToolHandler = async (args, ctx) => {
  const formats = formatRange(args, ctx).conditionalFormats;
  formats.load('items/id,items/type,items/priority');
  await ctx.sync();
  return formats.items.map(item => ({ id: item.id, type: item.type, priority: item.priority }));
};

export const handleClearConditionalFormats: ToolHandler = async (args, ctx) => {
  formatRange(args, ctx).conditionalFormats.clearAll();
  await ctx.sync();
  return { sheet: args.sheet, address: args.address, cleared: true };
};
