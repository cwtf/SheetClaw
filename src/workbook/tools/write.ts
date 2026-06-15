import type { ToolSpec } from '../../types';
import type { ToolHandler } from '../executor';
import { ToolValidationError } from '../executor';

// ── Specs ──────────────────────────────────────────────────────────────────

export const WRITE_RANGE: ToolSpec = {
  name: 'write_range',
  description: 'Write values or formulas into a cell range. The 2D values array dimensions must exactly match the address. Requires user confirmation before applying.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range address matching the values array dimensions' },
      values: { type: 'array', items: { type: 'array', items: { description: 'Cell value: string, number, boolean, null, or formula string starting with =' } }, description: 'Row-major 2D array of cell values' },
      as_text: { type: 'boolean', description: 'Force all values to literal text, skipping formula detection. Default: false' },
    },
    required: ['workbook_id', 'sheet', 'address', 'values'],
  },
  mutating: true,
};

export const CLEAR_RANGE: ToolSpec = {
  name: 'clear_range',
  description: 'Clear contents and/or formats of a range. Requires user confirmation before applying.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range address' },
      apply_to: { type: 'string', description: '"contents" (default), "formats", or "all"' },
    },
    required: ['workbook_id', 'sheet', 'address'],
  },
  mutating: true,
};

export const COPY_RANGE_FORMAT: ToolSpec = {
  name: 'copy_range_format',
  description: 'Copy cell formatting from one range to another range on the same worksheet. Use this to match an existing table/column style, including number format, fill, font, borders, alignment, and optionally column width. Requires user confirmation before applying.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      source_address: { type: 'string', description: 'A1 range address to copy formatting from, e.g. "F6:F16"' },
      target_address: { type: 'string', description: 'A1 range address to apply formatting to, e.g. "G6:G16". Use the same shape as source_address when possible.' },
      copy_column_width: { type: 'boolean', description: 'Also copy the source column width to the target column(s). Default: true' },
    },
    required: ['workbook_id', 'sheet', 'source_address', 'target_address'],
  },
  mutating: true,
};

export const FORMAT_RANGE: ToolSpec = {
  name: 'format_range',
  description: 'Apply cell formatting directly to a worksheet range, including number format, font styling, fill, alignment, borders, row/column sizing, and autofit. Use this when the user asks to format a table, header row, numbers, borders, or worksheet layout. Requires user confirmation before applying.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range address to format, e.g. "A1:E13"' },
      number_format: { type: 'string', description: 'Excel number format code to apply, e.g. "#,##0", "$#,##0.00", "0.0%", or "yyyy-mm-dd"' },
      bold: { type: 'boolean', description: 'Set font bold on or off' },
      italic: { type: 'boolean', description: 'Set font italic on or off' },
      font_color: { type: 'string', description: 'Font color as a hex value such as "#1F2937"' },
      fill_color: { type: 'string', description: 'Cell fill color as a hex value such as "#D9EAF7"' },
      font_size: { type: 'number', description: 'Font size in points' },
      horizontal_alignment: { type: 'string', enum: ['general', 'left', 'center', 'right', 'fill', 'justify', 'center_across_selection', 'distributed'], description: 'Horizontal alignment' },
      vertical_alignment: { type: 'string', enum: ['top', 'middle', 'center', 'bottom', 'justify', 'distributed'], description: 'Vertical alignment' },
      wrap_text: { type: 'boolean', description: 'Enable or disable wrapped text' },
      border_style: { type: 'string', enum: ['none', 'continuous', 'dash', 'dashed', 'dash_dot', 'dash_dot_dot', 'dot', 'dotted', 'double'], description: 'Border line style for all outside and inside borders' },
      border_color: { type: 'string', description: 'Border color as a hex value such as "#B7C9D6"' },
      border_weight: { type: 'string', enum: ['hairline', 'thin', 'medium', 'thick'], description: 'Border weight' },
      column_width: { type: 'number', description: 'Set column width for the range columns' },
      row_height: { type: 'number', description: 'Set row height for the range rows' },
      autofit_columns: { type: 'boolean', description: 'Autofit columns after applying other formatting' },
      autofit_rows: { type: 'boolean', description: 'Autofit rows after applying other formatting' },
    },
    required: ['workbook_id', 'sheet', 'address'],
  },
  mutating: true,
};

export const WRITE_SPECS: ToolSpec[] = [WRITE_RANGE, CLEAR_RANGE, COPY_RANGE_FORMAT, FORMAT_RANGE];

// ── Handlers ───────────────────────────────────────────────────────────────
// These run after the snapshot is captured and the user has confirmed.
// They just perform the write — no snapshot or confirmation logic here.

type StringMap<T extends string> = Record<string, T>;

const HORIZONTAL_ALIGNMENTS: StringMap<Excel.HorizontalAlignment | 'General' | 'Left' | 'Center' | 'Right' | 'Fill' | 'Justify' | 'CenterAcrossSelection' | 'Distributed'> = {
  general: 'General',
  left: 'Left',
  center: 'Center',
  right: 'Right',
  fill: 'Fill',
  justify: 'Justify',
  center_across_selection: 'CenterAcrossSelection',
  distributed: 'Distributed',
};

const VERTICAL_ALIGNMENTS: StringMap<Excel.VerticalAlignment | 'Top' | 'Center' | 'Bottom' | 'Justify' | 'Distributed'> = {
  top: 'Top',
  middle: 'Center',
  center: 'Center',
  bottom: 'Bottom',
  justify: 'Justify',
  distributed: 'Distributed',
};

const BORDER_STYLES: StringMap<Excel.BorderLineStyle | 'None' | 'Continuous' | 'Dash' | 'DashDot' | 'DashDotDot' | 'Dot' | 'Double'> = {
  none: 'None',
  continuous: 'Continuous',
  dash: 'Dash',
  dashed: 'Dash',
  dash_dot: 'DashDot',
  dash_dot_dot: 'DashDotDot',
  dot: 'Dot',
  dotted: 'Dot',
  double: 'Double',
};

const BORDER_WEIGHTS: StringMap<Excel.BorderWeight | 'Hairline' | 'Thin' | 'Medium' | 'Thick'> = {
  hairline: 'Hairline',
  thin: 'Thin',
  medium: 'Medium',
  thick: 'Thick',
};

const BORDER_INDEXES: Array<Excel.BorderIndex | 'EdgeTop' | 'EdgeBottom' | 'EdgeLeft' | 'EdgeRight' | 'InsideVertical' | 'InsideHorizontal'> = [
  'EdgeTop',
  'EdgeBottom',
  'EdgeLeft',
  'EdgeRight',
  'InsideVertical',
  'InsideHorizontal',
];

function optionalString(args: Record<string, unknown>, key: string): string | undefined {
  const value = args[key];
  if (value === undefined) return undefined;
  if (typeof value !== 'string') throw new ToolValidationError(`"${key}" must be a string`);
  return value;
}

function optionalNumber(args: Record<string, unknown>, key: string): number | undefined {
  const value = args[key];
  if (value === undefined) return undefined;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new ToolValidationError(`"${key}" must be a finite number`);
  }
  return value;
}

function optionalBoolean(args: Record<string, unknown>, key: string): boolean | undefined {
  const value = args[key];
  if (value === undefined) return undefined;
  if (typeof value !== 'boolean') throw new ToolValidationError(`"${key}" must be a boolean`);
  return value;
}

function enumValue<T extends string>(map: StringMap<T>, raw: string | undefined, key: string): T | undefined {
  if (raw === undefined) return undefined;
  const normalized = raw.toLowerCase().replace(/[\s-]+/g, '_');
  const value = map[normalized];
  if (!value) {
    throw new ToolValidationError(`Unsupported "${key}" value "${raw}"`);
  }
  return value;
}

function makeMatrix(value: string, rows: number, cols: number): string[][] {
  return Array.from({ length: rows }, () => Array.from({ length: cols }, () => value));
}

export const handleWriteRange: ToolHandler = async (args, ctx) => {
  const sheet = args.sheet as string;
  const address = args.address as string;
  const values = args.values as (string | number | boolean | null)[][];
  const asText = (args.as_text as boolean | undefined) ?? false;

  if (!Array.isArray(values) || !values.every(r => Array.isArray(r))) {
    throw new ToolValidationError('"values" must be a 2D array');
  }

  const range = ctx.workbook.worksheets.getItem(sheet).getRange(address);
  range.load('rowCount,columnCount');
  await ctx.sync();

  if (range.rowCount !== values.length || range.columnCount !== (values[0]?.length ?? 0)) {
    throw new ToolValidationError(
      `Dimension mismatch: range ${address} is ${range.rowCount}×${range.columnCount} ` +
      `but values are ${values.length}×${values[0]?.length ?? 0}. Adjust address or values.`
    );
  }

  if (asText) {
    range.values = values as (string | number | boolean)[][];
  } else {
    // formulas handles both formula strings (=SUM…) and plain values
    range.formulas = values.map(row =>
      row.map(v => (v === null || v === undefined ? '' : v))
    ) as (string | number | boolean)[][];
  }
  await ctx.sync();

  return { address, written: { rows: range.rowCount, cols: range.columnCount } };
};

export const handleClearRange: ToolHandler = async (args, ctx) => {
  const sheet = args.sheet as string;
  const address = args.address as string;
  const applyTo = (args.apply_to as string | undefined) ?? 'contents';

  const clearType =
    applyTo === 'formats' ? Excel.ClearApplyTo.formats :
    applyTo === 'all' ? Excel.ClearApplyTo.all :
    Excel.ClearApplyTo.contents;

  const range = ctx.workbook.worksheets.getItem(sheet).getRange(address);
  range.clear(clearType);
  await ctx.sync();

  return { address, cleared: applyTo };
};

export const handleCopyRangeFormat: ToolHandler = async (args, ctx) => {
  const sheet = args.sheet as string;
  const sourceAddress = args.source_address as string;
  const targetAddress = args.target_address as string;
  const copyColumnWidth = (args.copy_column_width as boolean | undefined) ?? true;

  const ws = ctx.workbook.worksheets.getItem(sheet);
  const source = ws.getRange(sourceAddress);
  const target = ws.getRange(targetAddress);
  source.load('address,rowCount,columnCount');
  target.load('address,rowCount,columnCount');
  if (copyColumnWidth) source.format.load('columnWidth');
  await ctx.sync();

  if (source.rowCount !== target.rowCount || source.columnCount !== target.columnCount) {
    throw new ToolValidationError(
      `Dimension mismatch: source ${sourceAddress} is ${source.rowCount}x${source.columnCount} ` +
      `but target ${targetAddress} is ${target.rowCount}x${target.columnCount}. Use matching range shapes.`
    );
  }

  target.copyFrom(source, Excel.RangeCopyType.formats);
  if (copyColumnWidth) {
    target.format.columnWidth = source.format.columnWidth;
  }
  await ctx.sync();

  return {
    sheet,
    source: source.address,
    target: target.address,
    copied: 'formats',
    copiedColumnWidth: copyColumnWidth,
  };
};

export const handleFormatRange: ToolHandler = async (args, ctx) => {
  const sheet = args.sheet as string;
  const address = args.address as string;
  const numberFormat = optionalString(args, 'number_format');
  const bold = optionalBoolean(args, 'bold');
  const italic = optionalBoolean(args, 'italic');
  const fontColor = optionalString(args, 'font_color');
  const fillColor = optionalString(args, 'fill_color');
  const fontSize = optionalNumber(args, 'font_size');
  const horizontalAlignment = enumValue(HORIZONTAL_ALIGNMENTS, optionalString(args, 'horizontal_alignment'), 'horizontal_alignment');
  const verticalAlignment = enumValue(VERTICAL_ALIGNMENTS, optionalString(args, 'vertical_alignment'), 'vertical_alignment');
  const wrapText = optionalBoolean(args, 'wrap_text');
  const borderStyle = enumValue(BORDER_STYLES, optionalString(args, 'border_style'), 'border_style');
  const borderColor = optionalString(args, 'border_color');
  const borderWeight = enumValue(BORDER_WEIGHTS, optionalString(args, 'border_weight'), 'border_weight');
  const columnWidth = optionalNumber(args, 'column_width');
  const rowHeight = optionalNumber(args, 'row_height');
  const autofitColumns = optionalBoolean(args, 'autofit_columns') ?? false;
  const autofitRows = optionalBoolean(args, 'autofit_rows') ?? false;

  const hasFormatting =
    numberFormat !== undefined ||
    bold !== undefined ||
    italic !== undefined ||
    fontColor !== undefined ||
    fillColor !== undefined ||
    fontSize !== undefined ||
    horizontalAlignment !== undefined ||
    verticalAlignment !== undefined ||
    wrapText !== undefined ||
    borderStyle !== undefined ||
    borderColor !== undefined ||
    borderWeight !== undefined ||
    columnWidth !== undefined ||
    rowHeight !== undefined ||
    autofitColumns ||
    autofitRows;

  if (!hasFormatting) {
    throw new ToolValidationError('Provide at least one formatting option');
  }

  const range = ctx.workbook.worksheets.getItem(sheet).getRange(address);
  range.load('address,rowCount,columnCount');
  await ctx.sync();

  const applied: string[] = [];

  if (numberFormat !== undefined) {
    range.numberFormat = makeMatrix(numberFormat, range.rowCount, range.columnCount);
    applied.push('numberFormat');
  }
  if (bold !== undefined) {
    range.format.font.bold = bold;
    applied.push('bold');
  }
  if (italic !== undefined) {
    range.format.font.italic = italic;
    applied.push('italic');
  }
  if (fontColor !== undefined) {
    range.format.font.color = fontColor;
    applied.push('fontColor');
  }
  if (fillColor !== undefined) {
    range.format.fill.color = fillColor;
    applied.push('fillColor');
  }
  if (fontSize !== undefined) {
    range.format.font.size = fontSize;
    applied.push('fontSize');
  }
  if (horizontalAlignment !== undefined) {
    range.format.horizontalAlignment = horizontalAlignment;
    applied.push('horizontalAlignment');
  }
  if (verticalAlignment !== undefined) {
    range.format.verticalAlignment = verticalAlignment;
    applied.push('verticalAlignment');
  }
  if (wrapText !== undefined) {
    range.format.wrapText = wrapText;
    applied.push('wrapText');
  }
  if (columnWidth !== undefined) {
    range.format.columnWidth = columnWidth;
    applied.push('columnWidth');
  }
  if (rowHeight !== undefined) {
    range.format.rowHeight = rowHeight;
    applied.push('rowHeight');
  }
  if (borderStyle !== undefined || borderColor !== undefined || borderWeight !== undefined) {
    for (const index of BORDER_INDEXES) {
      const border = range.format.borders.getItem(index as Excel.BorderIndex);
      border.style = borderStyle ?? 'Continuous';
      if (borderColor !== undefined) border.color = borderColor;
      if (borderWeight !== undefined) border.weight = borderWeight;
    }
    applied.push('borders');
  }
  if (autofitColumns) {
    range.format.autofitColumns();
    applied.push('autofitColumns');
  }
  if (autofitRows) {
    range.format.autofitRows();
    applied.push('autofitRows');
  }

  await ctx.sync();

  return {
    sheet,
    address: range.address,
    formatted: applied,
  };
};
