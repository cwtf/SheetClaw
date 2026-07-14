import type { WorkbookSelection } from '../types';
import type { ExcelRunner } from '../workbook/registry';

export async function getCurrentWorkbookSelection(
  runner: ExcelRunner = fn => Excel.run(fn),
): Promise<WorkbookSelection> {
  return runner(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load('address,worksheet/name');
    await ctx.sync();

    return {
      sheet: range.worksheet.name,
      address: localAddress(range.address),
    };
  });
}

export function localAddress(address: string): string {
  const separator = address.lastIndexOf('!');
  const local = separator >= 0 ? address.slice(separator + 1) : address;
  return local.replace(/\$/g, '');
}