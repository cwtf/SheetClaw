export { WorkbookRegistry, WorkbookNotFoundError } from './registry';
export type { ExcelRunner } from './registry';
export { ToolExecutor, ToolValidationError, ToolUnsupportedError } from './executor';
export type { ToolHandler } from './executor';
export { SnapshotManager } from './snapshot';
export { PHASE4_READ_SPECS } from './tools/specs';
export { WRITE_SPECS } from './tools/write';
export { CHART_SPECS } from './tools/charts';
export { PIVOT_SPECS } from './tools/pivots';
export { computeRangeDiff, cellAddress, parseRangeTopLeft } from './a1notation';

import { WorkbookRegistry } from './registry';
import { ToolExecutor } from './executor';
import { SnapshotManager } from './snapshot';
import type { ToolSpec } from '../types';
import type { ToolHandler } from './executor';
import { PHASE4_READ_SPECS } from './tools/specs';
import { WRITE_SPECS } from './tools/write';
import {
  handleReadRange,
  handleListSheets,
  handleGetSheetContext,
  handleGetSelection,
} from './tools/range';
import {
  handleListWorkbooks,
  handleGetActiveWorkbook,
  handleSetScopeWorkbook,
  handleGetNamedRanges,
} from './tools/workbook_tools';
import { handleWriteRange, handleClearRange, handleCopyRangeFormat, handleFormatRange } from './tools/write';
import {
  CHART_SPECS,
  handleListCharts, handleCreateChart, handleModifyChart, handleDeleteChart, handleSetChartData,
} from './tools/charts';
import {
  PIVOT_SPECS,
  handleListPivots, handleGetPivot, handleCreatePivot, handleAddPivotField, handleRefreshPivot,
} from './tools/pivots';

// ── Factory ────────────────────────────────────────────────────────────────

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
    // Phase 4 — read-only
    [PHASE4_READ_SPECS[0], handleReadRange],
    [PHASE4_READ_SPECS[1], handleListSheets],
    [PHASE4_READ_SPECS[2], handleGetSheetContext],
    [PHASE4_READ_SPECS[3], handleGetSelection],
    [PHASE4_READ_SPECS[4], handleListWorkbooks],
    [PHASE4_READ_SPECS[5], handleGetActiveWorkbook],
    [PHASE4_READ_SPECS[6], handleSetScopeWorkbook],
    [PHASE4_READ_SPECS[7], handleGetNamedRanges],
    // Phase 5 — write
    [WRITE_SPECS[0], handleWriteRange],
    [WRITE_SPECS[1], handleClearRange],
    [WRITE_SPECS[2], handleCopyRangeFormat],
    [WRITE_SPECS[3], handleFormatRange],
    // Phase 7 — charts
    [CHART_SPECS[0], handleListCharts],
    [CHART_SPECS[1], handleCreateChart],
    [CHART_SPECS[2], handleModifyChart],
    [CHART_SPECS[3], handleDeleteChart],
    [CHART_SPECS[4], handleSetChartData],
    // Phase 7 — pivots
    [PIVOT_SPECS[0], handleListPivots],
    [PIVOT_SPECS[1], handleGetPivot],
    [PIVOT_SPECS[2], handleCreatePivot],
    [PIVOT_SPECS[3], handleAddPivotField],
    [PIVOT_SPECS[4], handleRefreshPivot],
  ];

  for (const [spec, handler] of registrations) {
    executor.register(spec, handler);
  }

  return { registry, executor, snapshots };
}
