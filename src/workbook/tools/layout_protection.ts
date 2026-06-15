import type { ToolSpec } from '../../types';
import type { ToolHandler } from '../executor';
import { optionalBooleanArg, optionalNumberArg, optionalStringArg, stringArg } from './args';

export const PROTECT_SHEET: ToolSpec = {
  name: 'protect_sheet',
  description: 'Protect a worksheet, optionally allowing selected actions. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      password: { type: 'string', description: 'Optional protection password' },
      allow_sort: { type: 'boolean', description: 'Allow sorting while protected' },
      allow_filter: { type: 'boolean', description: 'Allow AutoFilter while protected' },
      allow_format_cells: { type: 'boolean', description: 'Allow formatting cells while protected' },
      allow_insert_rows: { type: 'boolean', description: 'Allow inserting rows while protected' },
      allow_delete_rows: { type: 'boolean', description: 'Allow deleting rows while protected' },
      allow_pivots: { type: 'boolean', description: 'Allow PivotTable use while protected' },
    },
    required: ['workbook_id', 'sheet'],
  },
  mutating: true,
};

export const UNPROTECT_SHEET: ToolSpec = {
  name: 'unprotect_sheet',
  description: 'Remove worksheet protection. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      password: { type: 'string', description: 'Optional protection password' },
    },
    required: ['workbook_id', 'sheet'],
  },
  mutating: true,
};

export const PROTECT_WORKBOOK: ToolSpec = {
  name: 'protect_workbook',
  description: 'Protect workbook structure. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      password: { type: 'string', description: 'Optional protection password' },
    },
    required: ['workbook_id'],
  },
  mutating: true,
};

export const UNPROTECT_WORKBOOK: ToolSpec = {
  name: 'unprotect_workbook',
  description: 'Remove workbook structure protection. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      password: { type: 'string', description: 'Optional protection password' },
    },
    required: ['workbook_id'],
  },
  mutating: true,
};

export const SET_PRINT_AREA: ToolSpec = {
  name: 'set_print_area',
  description: 'Set the worksheet print area. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'Range address to print, e.g. A1:H40' },
    },
    required: ['workbook_id', 'sheet', 'address'],
  },
  mutating: true,
};

export const CLEAR_PRINT_AREA: ToolSpec = {
  name: 'clear_print_area',
  description: 'Clear the worksheet print area. Requires user confirmation.',
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

export const SET_PAGE_LAYOUT: ToolSpec = {
  name: 'set_page_layout',
  description: 'Set page orientation, paper size, print gridlines/headings, margins, or fit-to-page options. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      orientation: { type: 'string', description: 'Portrait or Landscape' },
      paper_size: { type: 'string', description: 'Excel paper size such as A4, Letter, Legal' },
      print_gridlines: { type: 'boolean', description: 'Whether to print gridlines' },
      print_headings: { type: 'boolean', description: 'Whether to print row/column headings' },
      center_horizontally: { type: 'boolean', description: 'Center printed output horizontally' },
      center_vertically: { type: 'boolean', description: 'Center printed output vertically' },
      margin_unit: { type: 'string', description: 'Points, Inches, or Centimeters; defaults to Inches' },
      top_margin: { type: 'number', description: 'Top margin in the chosen unit' },
      bottom_margin: { type: 'number', description: 'Bottom margin in the chosen unit' },
      left_margin: { type: 'number', description: 'Left margin in the chosen unit' },
      right_margin: { type: 'number', description: 'Right margin in the chosen unit' },
      fit_to_width: { type: 'number', description: 'Pages wide for fit-to-page printing' },
      fit_to_height: { type: 'number', description: 'Pages tall for fit-to-page printing' },
      scale: { type: 'number', description: 'Print scale percentage from 10 to 400' },
    },
    required: ['workbook_id', 'sheet'],
  },
  mutating: true,
};

export const ADD_PAGE_BREAK: ToolSpec = {
  name: 'add_page_break',
  description: 'Add a horizontal or vertical manual page break before a cell. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'Cell address immediately after the break, e.g. A25 or H1' },
      direction: { type: 'string', description: 'horizontal or vertical' },
    },
    required: ['workbook_id', 'sheet', 'address', 'direction'],
  },
  mutating: true,
};

export const LAYOUT_PROTECTION_SPECS: ToolSpec[] = [
  PROTECT_SHEET,
  UNPROTECT_SHEET,
  PROTECT_WORKBOOK,
  UNPROTECT_WORKBOOK,
  SET_PRINT_AREA,
  CLEAR_PRINT_AREA,
  SET_PAGE_LAYOUT,
  ADD_PAGE_BREAK,
];

function worksheet(args: Record<string, unknown>, ctx: Excel.RequestContext): Excel.Worksheet {
  return ctx.workbook.worksheets.getItem(stringArg(args, 'sheet'));
}

function protectionOptions(args: Record<string, unknown>): Excel.WorksheetProtectionOptions {
  return {
    allowSort: optionalBooleanArg(args, 'allow_sort'),
    allowAutoFilter: optionalBooleanArg(args, 'allow_filter'),
    allowFormatCells: optionalBooleanArg(args, 'allow_format_cells'),
    allowInsertRows: optionalBooleanArg(args, 'allow_insert_rows'),
    allowDeleteRows: optionalBooleanArg(args, 'allow_delete_rows'),
    allowPivotTables: optionalBooleanArg(args, 'allow_pivots'),
  };
}

export const handleProtectSheet: ToolHandler = async (args, ctx) => {
  const options = protectionOptions(args);
  worksheet(args, ctx).protection.protect(options, optionalStringArg(args, 'password'));
  await ctx.sync();
  return { sheet: args.sheet, protected: true, options };
};

export const handleUnprotectSheet: ToolHandler = async (args, ctx) => {
  worksheet(args, ctx).protection.unprotect(optionalStringArg(args, 'password'));
  await ctx.sync();
  return { sheet: args.sheet, protected: false };
};

export const handleProtectWorkbook: ToolHandler = async (args, ctx) => {
  ctx.workbook.protection.protect(optionalStringArg(args, 'password'));
  await ctx.sync();
  return { workbookProtected: true };
};

export const handleUnprotectWorkbook: ToolHandler = async (args, ctx) => {
  ctx.workbook.protection.unprotect(optionalStringArg(args, 'password'));
  await ctx.sync();
  return { workbookProtected: false };
};

export const handleSetPrintArea: ToolHandler = async (args, ctx) => {
  worksheet(args, ctx).pageLayout.setPrintArea(stringArg(args, 'address'));
  await ctx.sync();
  return { sheet: args.sheet, printArea: args.address };
};

export const handleClearPrintArea: ToolHandler = async (args, ctx) => {
  worksheet(args, ctx).pageLayout.setPrintArea('');
  await ctx.sync();
  return { sheet: args.sheet, printArea: null };
};

export const handleSetPageLayout: ToolHandler = async (args, ctx) => {
  const layout = worksheet(args, ctx).pageLayout;
  const applied: string[] = [];
  const orientation = optionalStringArg(args, 'orientation');
  const paperSize = optionalStringArg(args, 'paper_size');
  const printGridlines = optionalBooleanArg(args, 'print_gridlines');
  const printHeadings = optionalBooleanArg(args, 'print_headings');
  const centerHorizontally = optionalBooleanArg(args, 'center_horizontally');
  const centerVertically = optionalBooleanArg(args, 'center_vertically');

  if (orientation !== undefined) {
    layout.orientation = orientation as Excel.PageOrientation;
    applied.push('orientation');
  }
  if (paperSize !== undefined) {
    layout.paperSize = paperSize as Excel.PaperType;
    applied.push('paperSize');
  }
  if (printGridlines !== undefined) {
    layout.printGridlines = printGridlines;
    applied.push('printGridlines');
  }
  if (printHeadings !== undefined) {
    layout.printHeadings = printHeadings;
    applied.push('printHeadings');
  }
  if (centerHorizontally !== undefined) {
    layout.centerHorizontally = centerHorizontally;
    applied.push('centerHorizontally');
  }
  if (centerVertically !== undefined) {
    layout.centerVertically = centerVertically;
    applied.push('centerVertically');
  }

  const margins: Excel.PageLayoutMarginOptions = {};
  const top = optionalNumberArg(args, 'top_margin');
  const bottom = optionalNumberArg(args, 'bottom_margin');
  const left = optionalNumberArg(args, 'left_margin');
  const right = optionalNumberArg(args, 'right_margin');
  if (top !== undefined) margins.top = top;
  if (bottom !== undefined) margins.bottom = bottom;
  if (left !== undefined) margins.left = left;
  if (right !== undefined) margins.right = right;
  if (Object.keys(margins).length > 0) {
    layout.setPrintMargins((optionalStringArg(args, 'margin_unit') ?? 'Inches') as Excel.PrintMarginUnit, margins);
    applied.push('margins');
  }

  const fitToWidth = optionalNumberArg(args, 'fit_to_width');
  const fitToHeight = optionalNumberArg(args, 'fit_to_height');
  const scale = optionalNumberArg(args, 'scale');
  if (fitToWidth !== undefined || fitToHeight !== undefined || scale !== undefined) {
    layout.zoom = {
      horizontalFitToPages: fitToWidth,
      verticalFitToPages: fitToHeight,
      scale,
    };
    applied.push('zoom');
  }

  await ctx.sync();
  return { sheet: args.sheet, applied };
};

export const handleAddPageBreak: ToolHandler = async (args, ctx) => {
  const sheet = worksheet(args, ctx);
  const direction = stringArg(args, 'direction').toLowerCase();
  const address = stringArg(args, 'address');
  const pageBreak = direction === 'vertical'
    ? sheet.verticalPageBreaks.add(address)
    : sheet.horizontalPageBreaks.add(address);
  pageBreak.load('rowIndex,columnIndex');
  await ctx.sync();
  return { sheet: args.sheet, direction, address, rowIndex: pageBreak.rowIndex, columnIndex: pageBreak.columnIndex };
};
