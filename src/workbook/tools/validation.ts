import type { ToolSpec } from '../../types';
import type { ToolHandler } from '../executor';
import { enumArg, optionalBooleanArg, optionalStringArg, stringArg } from './args';

export const SET_DATA_VALIDATION: ToolSpec = {
  name: 'set_data_validation',
  description: 'Set data validation on a range, such as dropdown lists, numeric constraints, dates, text length, or custom formulas. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      address: { type: 'string', description: 'A1 range address' },
      type: { type: 'string', description: 'list, whole_number, decimal, date, time, text_length, or custom' },
      source: { type: 'string', description: 'List source such as "Yes,No" or "=Sheet1!$A$1:$A$5"' },
      formula1: { type: 'string', description: 'Primary formula or comparison value' },
      formula2: { type: 'string', description: 'Secondary formula or comparison value for between/not_between' },
      operator: { type: 'string', description: 'between, not_between, equal_to, not_equal_to, greater_than, less_than, greater_than_or_equal_to, less_than_or_equal_to' },
      ignore_blanks: { type: 'boolean', description: 'Ignore blank values. Default: true' },
      show_dropdown: { type: 'boolean', description: 'Show dropdown for list validation. Default: true' },
      prompt_title: { type: 'string', description: 'Input prompt title' },
      prompt_message: { type: 'string', description: 'Input prompt message' },
      error_title: { type: 'string', description: 'Error alert title' },
      error_message: { type: 'string', description: 'Error alert message' },
      error_style: { type: 'string', description: 'stop, warning, or information. Default: stop' },
    },
    required: ['workbook_id', 'sheet', 'address', 'type'],
  },
  mutating: true,
};

export const CLEAR_DATA_VALIDATION: ToolSpec = {
  name: 'clear_data_validation',
  description: 'Clear data validation from a range. Requires user confirmation.',
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

export const GET_DATA_VALIDATION: ToolSpec = {
  name: 'get_data_validation',
  description: 'Read data validation settings from a range.',
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

export const VALIDATION_SPECS: ToolSpec[] = [
  SET_DATA_VALIDATION,
  CLEAR_DATA_VALIDATION,
  GET_DATA_VALIDATION,
];

const VALIDATION_TYPES = ['list', 'whole_number', 'decimal', 'date', 'time', 'text_length', 'custom'] as const;
const OPERATORS = ['between', 'not_between', 'equal_to', 'not_equal_to', 'greater_than', 'less_than', 'greater_than_or_equal_to', 'less_than_or_equal_to'] as const;
const ERROR_STYLES = ['stop', 'warning', 'information'] as const;

function pascal(value: string): string {
  return value.replace(/(^|_)([a-z])/g, (_match, _sep: string, char: string) => char.toUpperCase());
}

function validationRule(args: Record<string, unknown>): Excel.DataValidationRule {
  const type = enumArg(args, 'type', VALIDATION_TYPES);
  const operator = pascal(enumArg(args, 'operator', OPERATORS, 'between')) as Excel.DataValidationOperator;
  const formula1 = optionalStringArg(args, 'formula1');
  const formula2 = optionalStringArg(args, 'formula2');
  if (type === 'list') {
    return { list: { source: stringArg(args, 'source'), inCellDropDown: optionalBooleanArg(args, 'show_dropdown') ?? true } };
  }
  if (type === 'custom') return { custom: { formula: stringArg(args, 'formula1') } };
  const rule = { formula1: formula1 ?? '', formula2, operator };
  if (type === 'whole_number') return { wholeNumber: rule };
  if (type === 'text_length') return { textLength: rule };
  return { [type]: rule } as Excel.DataValidationRule;
}

export const handleSetDataValidation: ToolHandler = async (args, ctx) => {
  const range = ctx.workbook.worksheets.getItem(stringArg(args, 'sheet')).getRange(stringArg(args, 'address'));
  const validation = range.dataValidation;
  validation.rule = validationRule(args);
  validation.ignoreBlanks = optionalBooleanArg(args, 'ignore_blanks') ?? true;
  const promptTitle = optionalStringArg(args, 'prompt_title');
  const promptMessage = optionalStringArg(args, 'prompt_message');
  if (promptTitle || promptMessage) {
    validation.prompt = { showPrompt: true, title: promptTitle ?? '', message: promptMessage ?? '' };
  }
  const errorTitle = optionalStringArg(args, 'error_title');
  const errorMessage = optionalStringArg(args, 'error_message');
  if (errorTitle || errorMessage) {
    validation.errorAlert = {
      showAlert: true,
      title: errorTitle ?? '',
      message: errorMessage ?? '',
      style: pascal(enumArg(args, 'error_style', ERROR_STYLES, 'stop')) as Excel.DataValidationAlertStyle,
    };
  }
  await ctx.sync();
  return { sheet: args.sheet, address: args.address, validationType: args.type };
};

export const handleClearDataValidation: ToolHandler = async (args, ctx) => {
  ctx.workbook.worksheets.getItem(stringArg(args, 'sheet')).getRange(stringArg(args, 'address')).dataValidation.clear();
  await ctx.sync();
  return { sheet: args.sheet, address: args.address, cleared: true };
};

export const handleGetDataValidation: ToolHandler = async (args, ctx) => {
  const validation = ctx.workbook.worksheets.getItem(stringArg(args, 'sheet')).getRange(stringArg(args, 'address')).dataValidation;
  validation.load('type,valid,rule,ignoreBlanks,prompt,errorAlert');
  await ctx.sync();
  return validation.toJSON();
};
