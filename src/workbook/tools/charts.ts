import type { ToolSpec } from '../../types';
import type { ToolHandler } from '../executor';
import { ToolUnsupportedError } from '../unsupported-error';
import { optionalBooleanArg, optionalNumberArg, optionalStringArg, stringArg } from './args';

// ── Chart type allow-list ──────────────────────────────────────────────────

function getChartType(name: string): Excel.ChartType {
  switch (name.toLowerCase()) {
    case 'column':   return Excel.ChartType.columnClustered;
    case 'bar':      return Excel.ChartType.barClustered;
    case 'line':     return Excel.ChartType.line;
    case 'pie':      return Excel.ChartType.pie;
    case 'area':     return Excel.ChartType.area;
    case 'scatter':  return Excel.ChartType.xyscatter;
    case 'doughnut': return Excel.ChartType.doughnut;
    case 'radar':    return Excel.ChartType.radar;
    default: throw new ToolUnsupportedError(
      `Unsupported chart_type "${name}". Supported: column, bar, line, pie, area, scatter, doughnut, radar`
    );
  }
}

function getSeriesBy(val?: string): Excel.ChartSeriesBy {
  if (val === 'rows')    return Excel.ChartSeriesBy.rows;
  if (val === 'columns') return Excel.ChartSeriesBy.columns;
  return Excel.ChartSeriesBy.auto;
}

// ── Specs ──────────────────────────────────────────────────────────────────

export const LIST_CHARTS: ToolSpec = {
  name: 'list_charts',
  description: 'List all charts embedded in a worksheet, with their names and types.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
    },
    required: ['workbook_id', 'sheet'],
  },
  mutating: false,
};

export const CREATE_CHART: ToolSpec = {
  name: 'create_chart',
  description: 'Create a new chart on a worksheet from a data range. Requires user confirmation. Supported chart_type values: column, bar, line, pie, area, scatter, doughnut, radar.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id:  { type: 'string', description: 'Workbook ID' },
      sheet:        { type: 'string', description: 'Worksheet name' },
      chart_type:   { type: 'string', description: 'Chart type: column, bar, line, pie, area, scatter, doughnut, radar' },
      data_range:   { type: 'string', description: 'A1 range address of the source data, e.g. "A1:C10"' },
      title:        { type: 'string', description: 'Optional chart title' },
      series_by:    { type: 'string', description: 'How to interpret data: "auto" (default), "rows", or "columns"' },
    },
    required: ['workbook_id', 'sheet', 'chart_type', 'data_range'],
  },
  mutating: true,
};

export const MODIFY_CHART: ToolSpec = {
  name: 'modify_chart',
  description: 'Modify properties of an existing chart: title, type, or data range. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet:       { type: 'string', description: 'Worksheet name' },
      chart_name:  { type: 'string', description: 'Name of the chart to modify' },
      title:       { type: 'string', description: 'New chart title' },
      chart_type:  { type: 'string', description: 'New chart type (column, bar, line, pie, area, scatter, doughnut, radar)' },
      data_range:  { type: 'string', description: 'New source data range' },
      series_by:   { type: 'string', description: '"auto", "rows", or "columns"' },
    },
    required: ['workbook_id', 'sheet', 'chart_name'],
  },
  mutating: true,
};

export const DELETE_CHART: ToolSpec = {
  name: 'delete_chart',
  description: 'Delete a chart from a worksheet. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet:       { type: 'string', description: 'Worksheet name' },
      chart_name:  { type: 'string', description: 'Name of the chart to delete' },
    },
    required: ['workbook_id', 'sheet', 'chart_name'],
  },
  mutating: true,
};

export const SET_CHART_DATA: ToolSpec = {
  name: 'set_chart_data',
  description: 'Change the source data range of an existing chart. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet:       { type: 'string', description: 'Worksheet name' },
      chart_name:  { type: 'string', description: 'Name of the chart' },
      data_range:  { type: 'string', description: 'New source data range, e.g. "A1:C10"' },
      series_by:   { type: 'string', description: '"auto", "rows", or "columns"' },
    },
    required: ['workbook_id', 'sheet', 'chart_name', 'data_range'],
  },
  mutating: true,
};

export const FORMAT_CHART: ToolSpec = {
  name: 'format_chart',
  description: 'Format chart title, legend, style, and size/position. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      chart_name: { type: 'string', description: 'Name of the chart' },
      title: { type: 'string', description: 'Chart title text' },
      show_title: { type: 'boolean', description: 'Whether the chart title is visible' },
      show_legend: { type: 'boolean', description: 'Whether the legend is visible' },
      legend_position: { type: 'string', description: 'Top, Bottom, Left, Right, or Corner' },
      style: { type: 'number', description: 'Excel chart style number' },
      left: { type: 'number', description: 'Left position in points' },
      top: { type: 'number', description: 'Top position in points' },
      width: { type: 'number', description: 'Chart width in points' },
      height: { type: 'number', description: 'Chart height in points' },
    },
    required: ['workbook_id', 'sheet', 'chart_name'],
  },
  mutating: true,
};

export const SET_CHART_AXES: ToolSpec = {
  name: 'set_chart_axes',
  description: 'Set chart axis titles and visibility. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      chart_name: { type: 'string', description: 'Name of the chart' },
      category_axis_title: { type: 'string', description: 'Category axis title' },
      value_axis_title: { type: 'string', description: 'Value axis title' },
      category_axis_visible: { type: 'boolean', description: 'Whether the category axis is visible' },
      value_axis_visible: { type: 'boolean', description: 'Whether the value axis is visible' },
      value_axis_number_format: { type: 'string', description: 'Number format code for the value axis' },
    },
    required: ['workbook_id', 'sheet', 'chart_name'],
  },
  mutating: true,
};

export const SET_CHART_LABELS: ToolSpec = {
  name: 'set_chart_labels',
  description: 'Configure chart data labels. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      chart_name: { type: 'string', description: 'Name of the chart' },
      show_value: { type: 'boolean', description: 'Show values in data labels' },
      show_category: { type: 'boolean', description: 'Show category names in data labels' },
      show_series_name: { type: 'boolean', description: 'Show series names in data labels' },
      show_percentage: { type: 'boolean', description: 'Show percentages in data labels' },
      position: { type: 'string', description: 'Data label position such as OutsideEnd, InsideEnd, BestFit, or Center' },
    },
    required: ['workbook_id', 'sheet', 'chart_name'],
  },
  mutating: true,
};

export const ADD_TRENDLINE: ToolSpec = {
  name: 'add_trendline',
  description: 'Add a trendline to a chart series. Requires user confirmation.',
  parameters: {
    type: 'object',
    properties: {
      workbook_id: { type: 'string', description: 'Workbook ID' },
      sheet: { type: 'string', description: 'Worksheet name' },
      chart_name: { type: 'string', description: 'Name of the chart' },
      series_index: { type: 'number', description: 'Zero-based series index; defaults to 0' },
      trendline_type: { type: 'string', description: 'Linear, Exponential, Logarithmic, MovingAverage, Polynomial, or Power' },
      name: { type: 'string', description: 'Optional trendline name' },
      show_equation: { type: 'boolean', description: 'Show trendline equation' },
      show_r_squared: { type: 'boolean', description: 'Show R-squared value' },
    },
    required: ['workbook_id', 'sheet', 'chart_name'],
  },
  mutating: true,
};

export const CHART_SPECS: ToolSpec[] = [
  LIST_CHARTS,
  CREATE_CHART,
  MODIFY_CHART,
  DELETE_CHART,
  SET_CHART_DATA,
  FORMAT_CHART,
  SET_CHART_AXES,
  SET_CHART_LABELS,
  ADD_TRENDLINE,
];

// ── Handlers ───────────────────────────────────────────────────────────────

export const handleListCharts: ToolHandler = async (args, ctx) => {
  const sheet = ctx.workbook.worksheets.getItem(args.sheet as string);
  const charts = sheet.charts;
  charts.load('items/name,items/chartType');
  await ctx.sync();
  return charts.items.map(c => ({ name: c.name, chartType: c.chartType }));
};

export const handleCreateChart: ToolHandler = async (args, ctx) => {
  const sheet = ctx.workbook.worksheets.getItem(args.sheet as string);
  const range = sheet.getRange(args.data_range as string);
  const chartType = getChartType(args.chart_type as string);
  const seriesBy = getSeriesBy(args.series_by as string | undefined);

  const chart = sheet.charts.add(chartType, range, seriesBy);
  if (args.title) {
    chart.title.text = args.title as string;
    chart.title.visible = true;
  }
  chart.load('name');
  await ctx.sync();
  return { name: chart.name, created: true };
};

export const handleModifyChart: ToolHandler = async (args, ctx) => {
  const sheet = ctx.workbook.worksheets.getItem(args.sheet as string);
  const chart = sheet.charts.getItem(args.chart_name as string);
  const applied: string[] = [];

  if (args.title !== undefined) {
    chart.title.text = args.title as string;
    chart.title.visible = true;
    applied.push('title');
  }
  if (args.chart_type !== undefined) {
    chart.chartType = getChartType(args.chart_type as string);
    applied.push('chartType');
  }
  if (args.data_range !== undefined) {
    const range = sheet.getRange(args.data_range as string);
    chart.setData(range, getSeriesBy(args.series_by as string | undefined));
    applied.push('dataRange');
  }
  await ctx.sync();
  return { name: args.chart_name, applied };
};

export const handleDeleteChart: ToolHandler = async (args, ctx) => {
  const sheet = ctx.workbook.worksheets.getItem(args.sheet as string);
  const chart = sheet.charts.getItem(args.chart_name as string);
  chart.delete();
  await ctx.sync();
  return { deleted: true, name: args.chart_name };
};

export const handleSetChartData: ToolHandler = async (args, ctx) => {
  const sheet = ctx.workbook.worksheets.getItem(args.sheet as string);
  const chart = sheet.charts.getItem(args.chart_name as string);
  const range = sheet.getRange(args.data_range as string);
  chart.setData(range, getSeriesBy(args.series_by as string | undefined));
  await ctx.sync();
  return { name: args.chart_name, dataRange: args.data_range };
};

function getChart(args: Record<string, unknown>, ctx: Excel.RequestContext): Excel.Chart {
  const sheet = ctx.workbook.worksheets.getItem(stringArg(args, 'sheet'));
  return sheet.charts.getItem(stringArg(args, 'chart_name'));
}

export const handleFormatChart: ToolHandler = async (args, ctx) => {
  const chart = getChart(args, ctx);
  const applied: string[] = [];
  const chartBox = chart as unknown as {
    style?: number;
    left?: number;
    top?: number;
    width?: number;
    height?: number;
  };
  const title = optionalStringArg(args, 'title');
  const showTitle = optionalBooleanArg(args, 'show_title');
  const showLegend = optionalBooleanArg(args, 'show_legend');
  const legendPosition = optionalStringArg(args, 'legend_position');
  const style = optionalNumberArg(args, 'style');
  const left = optionalNumberArg(args, 'left');
  const top = optionalNumberArg(args, 'top');
  const width = optionalNumberArg(args, 'width');
  const height = optionalNumberArg(args, 'height');

  if (title !== undefined) {
    chart.title.text = title;
    chart.title.visible = true;
    applied.push('title');
  }
  if (showTitle !== undefined) {
    chart.title.visible = showTitle;
    applied.push('showTitle');
  }
  if (showLegend !== undefined) {
    chart.legend.visible = showLegend;
    applied.push('showLegend');
  }
  if (legendPosition !== undefined) {
    chart.legend.position = legendPosition as Excel.ChartLegendPosition;
    applied.push('legendPosition');
  }
  if (style !== undefined) {
    chartBox.style = style;
    applied.push('style');
  }
  if (left !== undefined) {
    chartBox.left = left;
    applied.push('left');
  }
  if (top !== undefined) {
    chartBox.top = top;
    applied.push('top');
  }
  if (width !== undefined) {
    chartBox.width = width;
    applied.push('width');
  }
  if (height !== undefined) {
    chartBox.height = height;
    applied.push('height');
  }

  await ctx.sync();
  return { name: args.chart_name, applied };
};

export const handleSetChartAxes: ToolHandler = async (args, ctx) => {
  const chart = getChart(args, ctx);
  const applied: string[] = [];
  const categoryTitle = optionalStringArg(args, 'category_axis_title');
  const valueTitle = optionalStringArg(args, 'value_axis_title');
  const categoryVisible = optionalBooleanArg(args, 'category_axis_visible');
  const valueVisible = optionalBooleanArg(args, 'value_axis_visible');
  const valueNumberFormat = optionalStringArg(args, 'value_axis_number_format');

  if (categoryTitle !== undefined) {
    chart.axes.categoryAxis.title.text = categoryTitle;
    chart.axes.categoryAxis.title.visible = true;
    applied.push('categoryAxisTitle');
  }
  if (valueTitle !== undefined) {
    chart.axes.valueAxis.title.text = valueTitle;
    chart.axes.valueAxis.title.visible = true;
    applied.push('valueAxisTitle');
  }
  if (categoryVisible !== undefined) {
    chart.axes.categoryAxis.visible = categoryVisible;
    applied.push('categoryAxisVisible');
  }
  if (valueVisible !== undefined) {
    chart.axes.valueAxis.visible = valueVisible;
    applied.push('valueAxisVisible');
  }
  if (valueNumberFormat !== undefined) {
    chart.axes.valueAxis.numberFormat = valueNumberFormat;
    applied.push('valueAxisNumberFormat');
  }

  await ctx.sync();
  return { name: args.chart_name, applied };
};

export const handleSetChartLabels: ToolHandler = async (args, ctx) => {
  const chart = getChart(args, ctx);
  const labels = chart.dataLabels;
  const applied: string[] = [];
  const showValue = optionalBooleanArg(args, 'show_value');
  const showCategory = optionalBooleanArg(args, 'show_category');
  const showSeriesName = optionalBooleanArg(args, 'show_series_name');
  const showPercentage = optionalBooleanArg(args, 'show_percentage');
  const position = optionalStringArg(args, 'position');

  if (showValue !== undefined) {
    labels.showValue = showValue;
    applied.push('showValue');
  }
  if (showCategory !== undefined) {
    labels.showCategoryName = showCategory;
    applied.push('showCategoryName');
  }
  if (showSeriesName !== undefined) {
    labels.showSeriesName = showSeriesName;
    applied.push('showSeriesName');
  }
  if (showPercentage !== undefined) {
    labels.showPercentage = showPercentage;
    applied.push('showPercentage');
  }
  if (position !== undefined) {
    labels.position = position as Excel.ChartDataLabelPosition;
    applied.push('position');
  }

  await ctx.sync();
  return { name: args.chart_name, applied };
};

export const handleAddTrendline: ToolHandler = async (args, ctx) => {
  const chart = getChart(args, ctx);
  const seriesIndex = optionalNumberArg(args, 'series_index') ?? 0;
  const type = (optionalStringArg(args, 'trendline_type') ?? 'Linear') as Excel.ChartTrendlineType;
  const trendline = chart.series.getItemAt(seriesIndex).trendlines.add(type);
  const name = optionalStringArg(args, 'name');
  const showEquation = optionalBooleanArg(args, 'show_equation');
  const showRSquared = optionalBooleanArg(args, 'show_r_squared');
  if (name !== undefined) trendline.name = name;
  if (showEquation !== undefined) trendline.showEquation = showEquation;
  if (showRSquared !== undefined) trendline.showRSquared = showRSquared;
  trendline.load('name,type,showEquation,showRSquared');
  await ctx.sync();
  return { chart: args.chart_name, seriesIndex, trendline: trendline.toJSON() };
};
