import type { ExcelAction } from '../types/actions';

/**
 * Maps a chart type string from the LLM to an Excel.ChartType enum value.
 */
function resolveChartType(input: string): Excel.ChartType {
  const t = (input || '').toLowerCase();
  if (t.includes('bar'))       return Excel.ChartType.barClustered;
  if (t.includes('column'))    return Excel.ChartType.columnClustered;
  if (t.includes('line'))      return Excel.ChartType.line;
  if (t.includes('pie'))       return Excel.ChartType.pie;
  if (t.includes('scatter'))   return Excel.ChartType.xyscatter;
  if (t.includes('area'))      return Excel.ChartType.area;
  if (t.includes('radar'))     return Excel.ChartType.radar;
  if (t.includes('doughnut'))  return Excel.ChartType.doughnut;
  return Excel.ChartType.columnClustered;
}

/**
 * Maps a horizontal alignment string to Excel.HorizontalAlignment.
 */
function resolveAlignment(input?: string): Excel.HorizontalAlignment | undefined {
  if (!input) return undefined;
  const map: Record<string, Excel.HorizontalAlignment> = {
    'left':    Excel.HorizontalAlignment.left,
    'center':  Excel.HorizontalAlignment.center,
    'right':   Excel.HorizontalAlignment.right,
    'justify': Excel.HorizontalAlignment.justify,
  };
  return map[input.toLowerCase()];
}

export async function executeExcelAction(action: ExcelAction) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      switch (action.command) {

        // ── Insert Data ──────────────────────────────────────────
        case 'insertData': {
          if (!action.data || !Array.isArray(action.data)) break;

          let targetCell = context.workbook.getSelectedRange().getCell(0, 0);
          if (action.rangeStart) {
            targetCell = sheet.getRange(action.rangeStart).getCell(0, 0);
          }

          const data2D = action.data.map((row) => Array.isArray(row) ? row : [row]);
          const rowCount = data2D.length;
          const colCount = data2D[0].length;

          const targetRange = targetCell.getBoundingRect(
            targetCell.getOffsetRange(rowCount - 1, colCount - 1)
          );
          targetRange.values = data2D;
          break;
        }

        // ── Add Chart ────────────────────────────────────────────
        case 'addChart': {
          if (!action.dataRange) break;

          const dataRange = sheet.getRange(action.dataRange);
          const cType = resolveChartType(action.chartType || '');
          const chart = sheet.charts.add(cType, dataRange, Excel.ChartSeriesBy.columns);

          if (action.title) {
            chart.title.text = action.title;
            chart.title.visible = true;
          }

          if (action.placement) {
            chart.top  = action.placement.top;
            chart.left = action.placement.left;
            chart.width  = action.placement.width;
            chart.height = action.placement.height;
          }

          if (action.xAxisLabel) {
            const xAxis = chart.axes.getItem(Excel.ChartAxisType.category);
            xAxis.title.text = action.xAxisLabel;
            xAxis.title.visible = true;
          }
          if (action.yAxisLabel) {
            const yAxis = chart.axes.getItem(Excel.ChartAxisType.value);
            yAxis.title.text = action.yAxisLabel;
            yAxis.title.visible = true;
          }
          break;
        }

        // ── Insert Formula ───────────────────────────────────────
        case 'insertFormula': {
          if (!action.cell || !action.formula) break;
          const cell = sheet.getRange(action.cell);

          // If it's a range of formulas (formulaRange)
          const a = action as any;
          if (a.formulaRange && Array.isArray(a.formulaRange)) {
            // Insert a 2D array of formulas
            const rangeRef = sheet.getRange(action.cell);
            rangeRef.formulas = a.formulaRange;
          } else {
            // Single formula — ensure it starts with =
            const formula = action.formula.startsWith('=')
              ? action.formula
              : `=${action.formula}`;
            cell.formulas = [[formula]];
          }
          break;
        }

        // ── Insert Formula Range (batch formulas into a column) ──
        case 'insertFormulaRange': {
          const a = action as any;
          if (!a.startCell || !a.formula || !a.count) break;

          const startRange = sheet.getRange(a.startCell);
          startRange.load(['rowIndex', 'columnIndex']);
          await context.sync();

          const rowStart = startRange.rowIndex;
          const col = startRange.columnIndex;
          const count = a.count as number;

          for (let i = 0; i < count; i++) {
            const cellRef = sheet.getCell(rowStart + i, col);
            const adjustedFormula = (a.formula as string)
              .replace(/\$?(\d+)/g, (_m: string, num: string) => String(parseInt(num) + i));
            cellRef.formulas = [[adjustedFormula.startsWith('=') ? adjustedFormula : `=${adjustedFormula}`]];
          }
          break;
        }

        // ── Format Range ─────────────────────────────────────────
        case 'formatRange': {
          if (!action.range) break;
          const range = sheet.getRange(action.range);

          if (action.bold !== undefined)      range.format.font.bold = action.bold;
          if (action.italic !== undefined)     range.format.font.italic = action.italic;
          if (action.fontSize !== undefined)   range.format.font.size = action.fontSize;
          if (action.fontColor)                range.format.font.color = action.fontColor;
          if (action.backgroundColor)          range.format.fill.color = action.backgroundColor;
          if (action.numberFormat)             range.numberFormat = [[action.numberFormat]];

          const align = resolveAlignment(action.horizontalAlign);
          if (align) range.format.horizontalAlignment = align;
          break;
        }

        // ── Conditional Format ───────────────────────────────────
        case 'addConditionalFormat': {
          if (!action.range) break;
          const cfRange = sheet.getRange(action.range);

          if (action.type === 'colorScale') {
            const cf = cfRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
            const colorScale = cf.colorScale;

            colorScale.criteria = {
              minimum: { color: action.minColor || '#FF0000', type: Excel.ConditionalFormatColorCriterionType.lowestValue },
              midpoint: action.midColor
                ? { color: action.midColor, type: Excel.ConditionalFormatColorCriterionType.percentile, formula: '50' }
                : undefined,
              maximum: { color: action.maxColor || '#00B050', type: Excel.ConditionalFormatColorCriterionType.highestValue },
            } as Excel.ConditionalColorScaleCriteria;
          } else if (action.type === 'dataBar') {
            cfRange.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
          } else if (action.type === 'iconSet') {
            cfRange.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
          }
          break;
        }

        // ── Create Pivot Table ───────────────────────────────────
        case 'createPivotTable': {
          if (!action.sourceRange) break;

          const destSheetName = action.destinationSheet || 'PivotTable';

          // Create destination sheet if it doesn't exist
          let destSheet: Excel.Worksheet;
          try {
            destSheet = context.workbook.worksheets.getItem(destSheetName);
          } catch {
            destSheet = context.workbook.worksheets.add(destSheetName);
          }

          const sourceRange = sheet.getRange(action.sourceRange);
          const pivotTable = context.workbook.pivotTables.add(
            `PT_${Date.now()}`,
            sourceRange,
            destSheet.getRange('A1')
          );

          // Add row hierarchies
          if (action.rows && action.rows.length > 0) {
            action.rows.forEach((fieldName) => {
              pivotTable.rowHierarchies.add(
                pivotTable.hierarchies.getItem(fieldName)
              );
            });
          }

          // Add data values
          if (action.values && action.values.length > 0) {
            action.values.forEach(({ field, aggregation }) => {
              const dataHierarchy = pivotTable.dataHierarchies.add(
                pivotTable.hierarchies.getItem(field)
              );
              const aggMap: Record<string, Excel.AggregationFunction> = {
                'sum':     Excel.AggregationFunction.sum,
                'count':   Excel.AggregationFunction.count,
                'average': Excel.AggregationFunction.average,
                'max':     Excel.AggregationFunction.max,
                'min':     Excel.AggregationFunction.min,
              };
              dataHierarchy.summarizeBy =
                aggMap[(aggregation || 'sum').toLowerCase()] ??
                Excel.AggregationFunction.sum;
            });
          }

          destSheet.activate();
          break;
        }

        // ── Add Sheet ────────────────────────────────────────────
        case 'addSheet': {
          if (!action.sheetName) break;
          const newSheet = context.workbook.worksheets.add(action.sheetName);
          newSheet.activate();
          break;
        }

        // ── Set Column Width ─────────────────────────────────────
        case 'setColumnWidth': {
          if (!action.columns) break;
          const targetSheet = action.sheet
            ? context.workbook.worksheets.getItem(action.sheet)
            : sheet;
          for (const [col, width] of Object.entries(action.columns)) {
            const colRange = targetSheet.getRange(`${col}:${col}`);
            colRange.format.columnWidth = width * 7;
          }
          break;
        }

        // ── Set Row Height ────────────────────────────────────────
        case 'setRowHeight': {
          if (!action.rows) break;
          const targetSheet = (action as any).sheet
            ? context.workbook.worksheets.getItem((action as any).sheet)
            : sheet;
          for (const [rowNum, height] of Object.entries(action.rows as Record<string, number>)) {
            const rowRange = targetSheet.getRange(`${rowNum}:${rowNum}`);
            rowRange.format.rowHeight = height;
          }
          break;
        }

        // ── Merge / Unmerge Range ─────────────────────────────────
        case 'mergeRange': {
          if (!action.range) break;
          const mergeSheet = (action as any).sheet
            ? context.workbook.worksheets.getItem((action as any).sheet)
            : sheet;
          const mergeRange = mergeSheet.getRange(action.range);
          mergeRange.merge((action as any).across ?? false);
          if ((action as any).value !== undefined) {
            mergeRange.values = [[(action as any).value]];
          }
          break;
        }

        case 'unmergeRange': {
          if (!action.range) break;
          sheet.getRange(action.range).unmerge();
          break;
        }

        // ── Auto-fit Columns ──────────────────────────────────────
        case 'autoFitColumns': {
          const targetSheet = (action as any).sheet
            ? context.workbook.worksheets.getItem((action as any).sheet)
            : sheet;
          if ((action as any).range) {
            targetSheet.getRange((action as any).range).format.autofitColumns();
          } else {
            // Auto-fit all used columns
            targetSheet.getUsedRange().format.autofitColumns();
          }
          break;
        }

        // ── Add Data Validation ───────────────────────────────────
        case 'addDataValidation': {
          if (!action.range) break;
          const dvRange = sheet.getRange(action.range);
          const dv = dvRange.dataValidation;

          const dvType = ((action as any).type || 'list').toLowerCase();

          if (dvType === 'list') {
            const listItems: string[] = (action as any).listItems ?? [];
            dv.rule = {
              list: {
                inCellDropDown: true,
                source: listItems.join(','),
              },
            };
          } else if (dvType === 'whole' || dvType === 'decimal') {
            dv.rule = {
              [dvType]: {
                formula1: String((action as any).min ?? 0),
                formula2: String((action as any).max ?? 100),
                operator: Excel.DataValidationOperator.between,
              },
            };
          } else if (dvType === 'date') {
            dv.rule = {
              date: {
                formula1: (action as any).minDate ?? '2000-01-01',
                formula2: (action as any).maxDate ?? '2100-12-31',
                operator: Excel.DataValidationOperator.between,
              },
            };
          }

          if ((action as any).errorMessage) {
            dv.errorAlert = {
              showAlert: true,
              style: Excel.DataValidationAlertStyle.stop,
              title: (action as any).errorTitle || 'Invalid Input',
              message: (action as any).errorMessage,
            };
          }
          break;
        }

        // ── Clear Range ───────────────────────────────────────────
        case 'clearRange': {
          if (!action.range) break;
          const clearRange = sheet.getRange(action.range);
          const clearType = ((action as any).clearType || 'contents').toLowerCase();
          if (clearType === 'all') {
            clearRange.clear(Excel.ClearApplyTo.all);
          } else if (clearType === 'formats') {
            clearRange.clear(Excel.ClearApplyTo.formats);
          } else {
            clearRange.clear(Excel.ClearApplyTo.contents);
          }
          break;
        }

        // ── Rename Sheet ──────────────────────────────────────────
        case 'renameSheet': {
          const a = action as any;
          const targetSheet = a.oldName
            ? context.workbook.worksheets.getItem(a.oldName)
            : sheet;
          targetSheet.name = a.newName;
          break;
        }

        // ── Freeze Panes ──────────────────────────────────────────
        case 'freezePanes': {
          const a = action as any;
          if (a.row !== undefined && a.column !== undefined) {
            sheet.freezePanes.freezeAt(sheet.getCell(a.row, a.column));
          } else if (a.row !== undefined) {
            sheet.freezePanes.freezeRows(a.row);
          } else if (a.column !== undefined) {
            sheet.freezePanes.freezeColumns(a.column);
          }
          break;
        }

        // ── Sort Range ────────────────────────────────────────────
        case 'sortRange': {
          const a = action as any;
          if (!a.range) break;
          const sortRange = sheet.getRange(a.range);
          sortRange.sort.apply([{
            key: a.columnIndex ?? 0,
            ascending: a.ascending !== false,
          }]);
          break;
        }

        // ── Set Number Format ────────────────────────────────────
        case 'setNumberFormat': {
          const a = action as any;
          if (!a.range || !a.format) break;
          sheet.getRange(a.range).numberFormat = [[a.format]];
          break;
        }

        // ── Protect Sheet ─────────────────────────────────────────
        case 'protectSheet': {
          const a = action as any;
          sheet.protection.protect({
            allowAutoFilter: true,
            allowSort: true,
            allowInsertRows: false,
            allowDeleteRows: false,
          }, a.password);
          break;
        }

        default:
          console.warn(`[RiskVir AI] Unknown Excel command: ${(action as any).command}`);
      }

      await context.sync();
    });
  } catch (error) {
    console.error("Failed to execute Excel action:", error);
    throw error;
  }
}
