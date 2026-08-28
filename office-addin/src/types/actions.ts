/**
 * Shared types for Office Action protocol.
 * These map to the commands defined in the backend prompt_builder.py
 */

// ─── Base ────────────────────────────────────────────────────────
export interface BaseAction {
  command: string;
}

// ─── Word Actions ────────────────────────────────────────────────
export interface InsertTextAction extends BaseAction {
  command: 'insertText';
  text: string;
}

export interface InsertHtmlAction extends BaseAction {
  command: 'insertHtml';
  html: string;
}

export interface ReplaceSelectionAction extends BaseAction {
  command: 'replaceSelection';
  text: string;
}

export interface ApplyStyleAction extends BaseAction {
  command: 'applyStyle';
  styleName: string;
  targetText?: string;
}

export interface InsertPageBreakAction extends BaseAction {
  command: 'insertPageBreak';
}

export interface InsertTOCAction extends BaseAction {
  command: 'insertTOC';
  levels?: number;
}

export interface FindAndReplaceAction extends BaseAction {
  command: 'findAndReplace';
  find: string;
  replace: string;
  matchCase?: boolean;
}

export interface SetDocumentPropertyAction extends BaseAction {
  command: 'setDocumentProperty';
  property: string;
  value: string;
}

export type WordAction =
  | InsertTextAction
  | InsertHtmlAction
  | ReplaceSelectionAction
  | ApplyStyleAction
  | InsertPageBreakAction
  | InsertTOCAction
  | FindAndReplaceAction
  | SetDocumentPropertyAction;

// ─── Excel Actions ───────────────────────────────────────────────
export interface InsertDataAction extends BaseAction {
  command: 'insertData';
  data: (string | number | boolean)[][];
  rangeStart?: string;
}

export interface AddChartAction extends BaseAction {
  command: 'addChart';
  chartType?: string;
  dataRange: string;
  title?: string;
  xAxisLabel?: string;
  yAxisLabel?: string;
  sheet?: string;
  placement?: { top: number; left: number; width: number; height: number };
}

export interface InsertFormulaAction extends BaseAction {
  command: 'insertFormula';
  cell: string;
  formula: string;
}

export interface FormatRangeAction extends BaseAction {
  command: 'formatRange';
  range: string;
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  backgroundColor?: string;
  fontColor?: string;
  horizontalAlign?: string;
  numberFormat?: string;
}

export interface AddConditionalFormatAction extends BaseAction {
  command: 'addConditionalFormat';
  range: string;
  type: 'colorScale' | 'dataBar' | 'iconSet';
  minColor?: string;
  midColor?: string;
  maxColor?: string;
}

export interface CreatePivotTableAction extends BaseAction {
  command: 'createPivotTable';
  sourceRange: string;
  destinationSheet?: string;
  rows?: string[];
  values?: { field: string; aggregation: string }[];
}

export interface AddSheetAction extends BaseAction {
  command: 'addSheet';
  sheetName: string;
  position?: number;
}

export interface SetColumnWidthAction extends BaseAction {
  command: 'setColumnWidth';
  sheet?: string;
  columns: Record<string, number>;
}

export interface SetRowHeightAction extends BaseAction {
  command: 'setRowHeight';
  sheet?: string;
  rows: Record<string, number>;
}

export interface MergeRangeAction extends BaseAction {
  command: 'mergeRange';
  range: string;
  sheet?: string;
  value?: string | number;
  across?: boolean;
}

export interface UnmergeRangeAction extends BaseAction {
  command: 'unmergeRange';
  range: string;
}

export interface AutoFitColumnsAction extends BaseAction {
  command: 'autoFitColumns';
  sheet?: string;
  range?: string;
}

export interface AddDataValidationAction extends BaseAction {
  command: 'addDataValidation';
  range: string;
  type: 'list' | 'whole' | 'decimal' | 'date';
  listItems?: string[];
  min?: number;
  max?: number;
  minDate?: string;
  maxDate?: string;
  errorTitle?: string;
  errorMessage?: string;
}

export interface ClearRangeAction extends BaseAction {
  command: 'clearRange';
  range: string;
  clearType?: 'contents' | 'formats' | 'all';
}

export type ExcelAction =
  | InsertDataAction
  | AddChartAction
  | InsertFormulaAction
  | FormatRangeAction
  | AddConditionalFormatAction
  | CreatePivotTableAction
  | AddSheetAction
  | SetColumnWidthAction
  | SetRowHeightAction
  | MergeRangeAction
  | UnmergeRangeAction
  | AutoFitColumnsAction
  | AddDataValidationAction
  | ClearRangeAction;

// ─── PowerPoint Actions ──────────────────────────────────────────
export interface AddSlideAction extends BaseAction {
  command: 'addSlide';
  title?: string;
  bulletPoints?: string[];
}

export interface AddSlideWithChartAction extends BaseAction {
  command: 'addSlideWithChart';
  title?: string;
  chartType?: string;
  chartData?: {
    labels: string[];
    series: { name: string; values: number[] }[];
  };
  xAxisLabel?: string;
  yAxisLabel?: string;
}

export interface AddSlideWithTableAction extends BaseAction {
  command: 'addSlideWithTable';
  title?: string;
  tableData?: {
    headers: string[];
    rows: string[][];
  };
  headerBackgroundColor?: string;
  headerFontColor?: string;
}

export interface AddSpeakerNotesAction extends BaseAction {
  command: 'addSpeakerNotes';
  slideIndex?: number;
  notes: string;
}

export interface FormatSlideAction extends BaseAction {
  command: 'formatSlide';
  slideIndex?: number;
  titleFontSize?: number;
  titleFontColor?: string;
  bodyFontSize?: number;
  backgroundColor?: string;
}

export type PowerPointAction =
  | AddSlideAction
  | AddSlideWithChartAction
  | AddSlideWithTableAction
  | AddSpeakerNotesAction
  | FormatSlideAction;

// ─── Union ───────────────────────────────────────────────────────
export type OfficeAction = WordAction | ExcelAction | PowerPointAction | BaseAction;
