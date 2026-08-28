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

export interface InsertParagraphAction extends BaseAction {
  command: 'insertParagraph';
  text: string;
  styleName?: string;
  bold?: boolean;
  fontSize?: number;
  fontColor?: string;
  alignment?: string;
}

export interface InsertHeadingAction extends BaseAction {
  command: 'insertHeading';
  text: string;
  level?: number;
}

export interface InsertBulletListAction extends BaseAction {
  command: 'insertBulletList';
  items: string[];
}

export interface InsertTableAction extends BaseAction {
  command: 'insertTable';
  tableData: string[][];
  headerRow?: boolean;
  headerColor?: string;
}

export interface InsertCommentAction extends BaseAction {
  command: 'insertComment';
  commentText: string;
  targetText?: string;
}

export interface FormatSelectionAction extends BaseAction {
  command: 'formatSelection';
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontColor?: string;
  highlightColor?: string;
  styleName?: string;
}

export interface InsertHorizontalRuleAction extends BaseAction {
  command: 'insertHorizontalRule';
}

export interface ClearDocumentAction extends BaseAction {
  command: 'clearDocument';
}

export interface InsertAtEndAction extends BaseAction {
  command: 'insertAtEnd';
  text?: string;
  html?: string;
}

export interface AppendSectionAction extends BaseAction {
  command: 'appendSection';
  heading?: string;
  level?: number;
  content?: string;
  contentType?: 'text' | 'html';
}

export interface InsertSectionBreakAction extends BaseAction {
  command: 'insertSectionBreak';
  breakType?: 'nextPage' | 'continuous' | 'evenPage';
}

export interface TrackChangesAction extends BaseAction {
  command: 'trackChanges';
  enable: boolean;
}

export interface AcceptAllChangesAction extends BaseAction {
  command: 'acceptAllChanges';
}

export interface DeleteContentAction extends BaseAction {
  command: 'deleteContent';
}

export interface DeleteSelectionAction extends BaseAction {
  command: 'deleteSelection';
}

export interface DeleteTextAction extends BaseAction {
  command: 'deleteText';
  text: string;
}

export type WordAction =
  | InsertTextAction
  | InsertHtmlAction
  | ReplaceSelectionAction
  | ApplyStyleAction
  | InsertPageBreakAction
  | InsertSectionBreakAction
  | InsertTOCAction
  | FindAndReplaceAction
  | SetDocumentPropertyAction
  | InsertParagraphAction
  | InsertHeadingAction
  | InsertBulletListAction
  | InsertTableAction
  | InsertCommentAction
  | FormatSelectionAction
  | InsertHorizontalRuleAction
  | ClearDocumentAction
  | DeleteContentAction
  | DeleteSelectionAction
  | DeleteTextAction
  | InsertAtEndAction
  | AppendSectionAction
  | TrackChangesAction
  | AcceptAllChangesAction;

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

export interface InsertFormulaRangeAction extends BaseAction {
  command: 'insertFormulaRange';
  startCell: string;
  formula: string;
  count: number;
}

export interface RenameSheetAction extends BaseAction {
  command: 'renameSheet';
  oldName?: string;
  newName: string;
}

export interface ProtectSheetAction extends BaseAction {
  command: 'protectSheet';
  password?: string;
  allowSelect?: boolean;
}

export interface FreezePanesAction extends BaseAction {
  command: 'freezePanes';
  row?: number;
  column?: number;
}

export interface SortRangeAction extends BaseAction {
  command: 'sortRange';
  range: string;
  columnIndex?: number;
  ascending?: boolean;
}

export interface SetNumberFormatAction extends BaseAction {
  command: 'setNumberFormat';
  range: string;
  format: string;
}

export type ExcelAction =
  | InsertDataAction
  | AddChartAction
  | InsertFormulaAction
  | InsertFormulaRangeAction
  | FormatRangeAction
  | AddConditionalFormatAction
  | CreatePivotTableAction
  | AddSheetAction
  | RenameSheetAction
  | SetColumnWidthAction
  | SetRowHeightAction
  | MergeRangeAction
  | UnmergeRangeAction
  | AutoFitColumnsAction
  | AddDataValidationAction
  | ClearRangeAction
  | ProtectSheetAction
  | FreezePanesAction
  | SortRangeAction
  | SetNumberFormatAction;

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

export interface AddSlideWithImageAction extends BaseAction {
  command: 'addSlideWithImage';
  title?: string;
  imagePlaceholderText?: string;
  caption?: string;
}

export interface MoveSlideAction extends BaseAction {
  command: 'moveSlide';
  fromIndex: number;
  toIndex: number;
}

export interface SetThemeAction extends BaseAction {
  command: 'setTheme';
  primaryColor?: string;
  accentColor?: string;
  fontHeading?: string;
  fontBody?: string;
  backgroundColor?: string;
}

export interface DeleteSlideAction extends BaseAction {
  command: 'deleteSlide';
  slideIndex?: number;
}

export interface DuplicateSlideAction extends BaseAction {
  command: 'duplicateSlide';
  slideIndex?: number;
}

export interface EditShapeTextAction extends BaseAction {
  command: 'editShapeText';
  slideIndex?: number;
  shapeName?: string;
  shapeIndex?: number;
  newText: string;
  fontSize?: number;
  fontColor?: string;
}

export type PowerPointAction =
  | AddSlideAction
  | AddSlideWithChartAction
  | AddSlideWithTableAction
  | AddSpeakerNotesAction
  | FormatSlideAction
  | AddSlideWithImageAction
  | MoveSlideAction
  | SetThemeAction
  | DeleteSlideAction
  | DuplicateSlideAction
  | EditShapeTextAction;

// ─── Union ───────────────────────────────────────────────────────
export type OfficeAction = WordAction | ExcelAction | PowerPointAction | BaseAction;
