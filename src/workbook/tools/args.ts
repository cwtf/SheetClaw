import { ToolValidationError } from '../executor';

export function stringArg(args: Record<string, unknown>, key: string): string {
  const value = args[key];
  if (typeof value !== 'string' || !value.trim()) {
    throw new ToolValidationError(`"${key}" must be a non-empty string`);
  }
  return value;
}

export function optionalStringArg(args: Record<string, unknown>, key: string): string | undefined {
  const value = args[key];
  if (value === undefined || value === null || value === '') return undefined;
  if (typeof value !== 'string') throw new ToolValidationError(`"${key}" must be a string`);
  return value;
}

export function numberArg(args: Record<string, unknown>, key: string): number {
  const value = args[key];
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new ToolValidationError(`"${key}" must be a finite number`);
  }
  return value;
}

export function optionalNumberArg(args: Record<string, unknown>, key: string): number | undefined {
  const value = args[key];
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new ToolValidationError(`"${key}" must be a finite number`);
  }
  return value;
}

export function optionalBooleanArg(args: Record<string, unknown>, key: string): boolean | undefined {
  const value = args[key];
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'boolean') throw new ToolValidationError(`"${key}" must be a boolean`);
  return value;
}

export function stringArrayArg(args: Record<string, unknown>, key: string): string[] {
  const value = args[key];
  if (!Array.isArray(value) || value.some(item => typeof item !== 'string')) {
    throw new ToolValidationError(`"${key}" must be an array of strings`);
  }
  return value as string[];
}

export function optionalStringArrayArg(args: Record<string, unknown>, key: string): string[] | undefined {
  const value = args[key];
  if (value === undefined || value === null) return undefined;
  if (!Array.isArray(value) || value.some(item => typeof item !== 'string')) {
    throw new ToolValidationError(`"${key}" must be an array of strings`);
  }
  return value as string[];
}

export function matrixArg(args: Record<string, unknown>, key: string): Array<Array<boolean | string | number>> {
  const value = args[key];
  if (!Array.isArray(value) || value.some(row => !Array.isArray(row))) {
    throw new ToolValidationError(`"${key}" must be a 2D array`);
  }
  return value as Array<Array<boolean | string | number>>;
}

export function enumArg<T extends string>(
  args: Record<string, unknown>,
  key: string,
  allowed: readonly T[],
  fallback?: T
): T {
  const value = args[key];
  if (value === undefined || value === null || value === '') {
    if (fallback !== undefined) return fallback;
    throw new ToolValidationError(`"${key}" is required`);
  }
  if (typeof value !== 'string') throw new ToolValidationError(`"${key}" must be a string`);
  const normalized = value.toLowerCase().replace(/[\s-]+/g, '_') as T;
  if (!allowed.includes(normalized)) {
    throw new ToolValidationError(`Invalid "${key}" value "${value}". Supported: ${allowed.join(', ')}`);
  }
  return normalized;
}
