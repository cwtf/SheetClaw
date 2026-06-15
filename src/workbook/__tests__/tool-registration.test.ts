import { describe, expect, it } from 'vitest';
import { createWorkbookLayer } from '../index';

describe('workbook tool registration', () => {
  it('exposes the expanded Excel tool surface to the agent', () => {
    const { executor } = createWorkbookLayer();
    const names = executor.getToolSpecs().map(spec => spec.name);

    expect(names).toEqual(expect.arrayContaining([
      'format_range',
      'create_table',
      'sort_range',
      'apply_filter',
      'add_sheet',
      'set_data_validation',
      'add_conditional_format',
      'insert_cells',
      'create_named_range',
      'add_comment',
      'add_image',
      'protect_sheet',
      'set_print_area',
      'format_chart',
      'set_chart_axes',
      'set_pivot_style',
      'delete_pivot',
    ]));
  });
});
