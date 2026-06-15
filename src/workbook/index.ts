export { WorkbookRegistry, WorkbookNotFoundError } from './registry';
export type { ExcelRunner } from './registry';
export { ToolExecutor, ToolValidationError, ToolUnsupportedError } from './executor';
export type { ToolHandler } from './executor';
export { SnapshotManager } from './snapshot';
export { PHASE4_READ_SPECS } from './tools/specs';
export { WRITE_SPECS } from './tools/write';
export { CHART_SPECS } from './tools/charts';
export { PIVOT_SPECS } from './tools/pivots';
export { TABLE_SPECS } from './tools/tables';
export { SORT_FILTER_SPECS } from './tools/sort_filter';
export { SHEET_SPECS } from './tools/sheets';
export { VALIDATION_SPECS } from './tools/validation';
export { CONDITIONAL_FORMAT_SPECS } from './tools/conditional_formats';
export { RANGE_OP_SPECS } from './tools/range_ops';
export { NAME_SPECS } from './tools/names';
export { COMMENT_SPECS } from './tools/comments';
export { SHAPE_SPECS } from './tools/shapes';
export { LAYOUT_PROTECTION_SPECS } from './tools/layout_protection';
export { computeRangeDiff, cellAddress, parseRangeTopLeft } from './a1notation';

import type { ToolSpec } from '../types';
import { ToolExecutor } from './executor';
import type { ToolHandler } from './executor';
import { WorkbookRegistry } from './registry';
import { SnapshotManager } from './snapshot';
import {
  CHART_SPECS,
  handleAddTrendline,
  handleCreateChart,
  handleDeleteChart,
  handleFormatChart,
  handleListCharts,
  handleModifyChart,
  handleSetChartAxes,
  handleSetChartData,
  handleSetChartLabels,
} from './tools/charts';
import {
  COMMENT_SPECS,
  handleAddComment,
  handleAddNote,
  handleDeleteComment,
  handleDeleteNote,
  handleListComments,
  handleReplyToComment,
} from './tools/comments';
import {
  CONDITIONAL_FORMAT_SPECS,
  handleAddConditionalFormat,
  handleClearConditionalFormats,
  handleListConditionalFormats,
} from './tools/conditional_formats';
import {
  LAYOUT_PROTECTION_SPECS,
  handleAddPageBreak,
  handleClearPrintArea,
  handleProtectSheet,
  handleProtectWorkbook,
  handleSetPageLayout,
  handleSetPrintArea,
  handleUnprotectSheet,
  handleUnprotectWorkbook,
} from './tools/layout_protection';
import {
  NAME_SPECS,
  handleCreateNamedRange,
  handleDeleteNamedRange,
  handleUpdateNamedRange,
} from './tools/names';
import {
  PIVOT_SPECS,
  handleAddPivotField,
  handleCreatePivot,
  handleDeletePivot,
  handleGetPivot,
  handleListPivots,
  handleRefreshAllPivots,
  handleRefreshPivot,
  handleRemovePivotField,
  handleSetPivotLayout,
  handleSetPivotStyle,
} from './tools/pivots';
import {
  handleGetSelection,
  handleGetSheetContext,
  handleListSheets,
  handleReadRange,
} from './tools/range';
import {
  RANGE_OP_SPECS,
  handleDeleteCells,
  handleDeleteColumns,
  handleDeleteRows,
  handleFindReplace,
  handleGetSpecialCells,
  handleInsertCells,
  handleInsertColumns,
  handleInsertRows,
  handleMergeRange,
  handleUnmergeRange,
} from './tools/range_ops';
import {
  SHAPE_SPECS,
  handleAddImage,
  handleAddShape,
  handleAddTextbox,
  handleDeleteShape,
  handleListShapes,
  handleMoveShape,
  handleResizeShape,
} from './tools/shapes';
import {
  SHEET_SPECS,
  handleActivateSheet,
  handleAddSheet,
  handleCopySheet,
  handleDeleteSheet,
  handleFreezePanes,
  handleHideSheet,
  handleMoveSheet,
  handleRenameSheet,
} from './tools/sheets';
import {
  SORT_FILTER_SPECS,
  handleApplyFilter,
  handleClearFilters,
  handleReapplyFilters,
  handleSortRange,
  handleSortTable,
} from './tools/sort_filter';
import { PHASE4_READ_SPECS } from './tools/specs';
import {
  TABLE_SPECS,
  handleAddTableColumns,
  handleAddTableRows,
  handleCreateTable,
  handleDeleteTable,
  handleListTables,
  handleResizeTable,
  handleSetTableStyle,
  handleSetTableTotals,
} from './tools/tables';
import {
  VALIDATION_SPECS,
  handleClearDataValidation,
  handleGetDataValidation,
  handleSetDataValidation,
} from './tools/validation';
import {
  WRITE_SPECS,
  handleClearRange,
  handleCopyRangeFormat,
  handleFormatRange,
  handleWriteRange,
} from './tools/write';
import {
  handleGetActiveWorkbook,
  handleGetNamedRanges,
  handleListWorkbooks,
  handleSetScopeWorkbook,
} from './tools/workbook_tools';

export interface WorkbookLayer {
  registry: WorkbookRegistry;
  executor: ToolExecutor;
  snapshots: SnapshotManager;
}

export function createWorkbookLayer(): WorkbookLayer {
  const registry = new WorkbookRegistry();
  const snapshots = new SnapshotManager();
  const executor = new ToolExecutor(registry);

  const registrations: [ToolSpec, ToolHandler][] = [
    [PHASE4_READ_SPECS[0], handleReadRange],
    [PHASE4_READ_SPECS[1], handleListSheets],
    [PHASE4_READ_SPECS[2], handleGetSheetContext],
    [PHASE4_READ_SPECS[3], handleGetSelection],
    [PHASE4_READ_SPECS[4], handleListWorkbooks],
    [PHASE4_READ_SPECS[5], handleGetActiveWorkbook],
    [PHASE4_READ_SPECS[6], handleSetScopeWorkbook],
    [PHASE4_READ_SPECS[7], handleGetNamedRanges],
    [WRITE_SPECS[0], handleWriteRange],
    [WRITE_SPECS[1], handleClearRange],
    [WRITE_SPECS[2], handleCopyRangeFormat],
    [WRITE_SPECS[3], handleFormatRange],
    [CHART_SPECS[0], handleListCharts],
    [CHART_SPECS[1], handleCreateChart],
    [CHART_SPECS[2], handleModifyChart],
    [CHART_SPECS[3], handleDeleteChart],
    [CHART_SPECS[4], handleSetChartData],
    [CHART_SPECS[5], handleFormatChart],
    [CHART_SPECS[6], handleSetChartAxes],
    [CHART_SPECS[7], handleSetChartLabels],
    [CHART_SPECS[8], handleAddTrendline],
    [PIVOT_SPECS[0], handleListPivots],
    [PIVOT_SPECS[1], handleGetPivot],
    [PIVOT_SPECS[2], handleCreatePivot],
    [PIVOT_SPECS[3], handleAddPivotField],
    [PIVOT_SPECS[4], handleRefreshPivot],
    [PIVOT_SPECS[5], handleRemovePivotField],
    [PIVOT_SPECS[6], handleSetPivotStyle],
    [PIVOT_SPECS[7], handleSetPivotLayout],
    [PIVOT_SPECS[8], handleRefreshAllPivots],
    [PIVOT_SPECS[9], handleDeletePivot],
    [TABLE_SPECS[0], handleListTables],
    [TABLE_SPECS[1], handleCreateTable],
    [TABLE_SPECS[2], handleResizeTable],
    [TABLE_SPECS[3], handleAddTableRows],
    [TABLE_SPECS[4], handleAddTableColumns],
    [TABLE_SPECS[5], handleSetTableStyle],
    [TABLE_SPECS[6], handleSetTableTotals],
    [TABLE_SPECS[7], handleDeleteTable],
    [SORT_FILTER_SPECS[0], handleSortRange],
    [SORT_FILTER_SPECS[1], handleSortTable],
    [SORT_FILTER_SPECS[2], handleApplyFilter],
    [SORT_FILTER_SPECS[3], handleClearFilters],
    [SORT_FILTER_SPECS[4], handleReapplyFilters],
    [SHEET_SPECS[0], handleAddSheet],
    [SHEET_SPECS[1], handleRenameSheet],
    [SHEET_SPECS[2], handleDeleteSheet],
    [SHEET_SPECS[3], handleCopySheet],
    [SHEET_SPECS[4], handleMoveSheet],
    [SHEET_SPECS[5], handleHideSheet],
    [SHEET_SPECS[6], handleActivateSheet],
    [SHEET_SPECS[7], handleFreezePanes],
    [VALIDATION_SPECS[0], handleSetDataValidation],
    [VALIDATION_SPECS[1], handleClearDataValidation],
    [VALIDATION_SPECS[2], handleGetDataValidation],
    [CONDITIONAL_FORMAT_SPECS[0], handleAddConditionalFormat],
    [CONDITIONAL_FORMAT_SPECS[1], handleListConditionalFormats],
    [CONDITIONAL_FORMAT_SPECS[2], handleClearConditionalFormats],
    [RANGE_OP_SPECS[0], handleInsertCells],
    [RANGE_OP_SPECS[1], handleDeleteCells],
    [RANGE_OP_SPECS[2], handleInsertRows],
    [RANGE_OP_SPECS[3], handleDeleteRows],
    [RANGE_OP_SPECS[4], handleInsertColumns],
    [RANGE_OP_SPECS[5], handleDeleteColumns],
    [RANGE_OP_SPECS[6], handleMergeRange],
    [RANGE_OP_SPECS[7], handleUnmergeRange],
    [RANGE_OP_SPECS[8], handleFindReplace],
    [RANGE_OP_SPECS[9], handleGetSpecialCells],
    [NAME_SPECS[0], handleCreateNamedRange],
    [NAME_SPECS[1], handleUpdateNamedRange],
    [NAME_SPECS[2], handleDeleteNamedRange],
    [COMMENT_SPECS[0], handleListComments],
    [COMMENT_SPECS[1], handleAddComment],
    [COMMENT_SPECS[2], handleReplyToComment],
    [COMMENT_SPECS[3], handleDeleteComment],
    [COMMENT_SPECS[4], handleAddNote],
    [COMMENT_SPECS[5], handleDeleteNote],
    [SHAPE_SPECS[0], handleListShapes],
    [SHAPE_SPECS[1], handleAddImage],
    [SHAPE_SPECS[2], handleAddTextbox],
    [SHAPE_SPECS[3], handleAddShape],
    [SHAPE_SPECS[4], handleMoveShape],
    [SHAPE_SPECS[5], handleResizeShape],
    [SHAPE_SPECS[6], handleDeleteShape],
    [LAYOUT_PROTECTION_SPECS[0], handleProtectSheet],
    [LAYOUT_PROTECTION_SPECS[1], handleUnprotectSheet],
    [LAYOUT_PROTECTION_SPECS[2], handleProtectWorkbook],
    [LAYOUT_PROTECTION_SPECS[3], handleUnprotectWorkbook],
    [LAYOUT_PROTECTION_SPECS[4], handleSetPrintArea],
    [LAYOUT_PROTECTION_SPECS[5], handleClearPrintArea],
    [LAYOUT_PROTECTION_SPECS[6], handleSetPageLayout],
    [LAYOUT_PROTECTION_SPECS[7], handleAddPageBreak],
  ];

  for (const [spec, handler] of registrations) {
    executor.register(spec, handler);
  }

  return { registry, executor, snapshots };
}
