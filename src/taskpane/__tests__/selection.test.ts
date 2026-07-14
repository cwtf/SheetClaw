import { describe, expect, it, vi } from 'vitest';
import type { ExcelRunner } from '../../workbook/registry';
import { getCurrentWorkbookSelection, localAddress } from '../selection';

describe('workbook selection', () => {
  it('reads and normalizes the selected range', async () => {
    const range = {
      address: "'Revenue 2026'!$C$7:$D$9",
      worksheet: { name: 'Revenue 2026' },
      load: vi.fn(),
    };
    const ctx = {
      workbook: { getSelectedRange: () => range },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const runner: ExcelRunner = fn => fn(ctx as unknown as Excel.RequestContext);

    await expect(getCurrentWorkbookSelection(runner)).resolves.toEqual({
      sheet: 'Revenue 2026',
      address: 'C7:D9',
    });
    expect(range.load).toHaveBeenCalledWith('address,worksheet/name');
    expect(ctx.sync).toHaveBeenCalledOnce();
  });

  it('keeps already-local addresses unchanged', () => {
    expect(localAddress('A1:B2')).toBe('A1:B2');
  });
});