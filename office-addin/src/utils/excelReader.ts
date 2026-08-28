/**
 * Rich Excel context collector.
 *
 * Returns a structured snapshot of the active workbook so the LLM
 * can reason precisely about what is already in the spreadsheet and
 * propose correct cell addresses, ranges, formula references, etc.
 */
export interface ExcelContext {
  selectedRange: {
    address: string;
    values: unknown[][];
    formulas: string[][];
    numberFormats: string[][];
    rowCount: number;
    columnCount: number;
  };
  worksheets: { name: string; isActive: boolean }[];
  activeSheet: string;
  tables: { name: string; range: string; headers: string[] }[];
  namedRanges: { name: string; address: string }[];
  usedRange: {
    address: string;
    rowCount: number;
    columnCount: number;
    firstRow: unknown[][];   // first row (headers)
    lastRow: unknown[][];    // last data row
  } | null;
}

export async function getSelectedExcelData(): Promise<ExcelContext> {
  return Excel.run(async (context) => {
    const wb = context.workbook;

    // ── Selection ────────────────────────────────────────────────
    const sel = wb.getSelectedRange();
    sel.load(['address', 'values', 'formulas', 'numberFormat', 'rowCount', 'columnCount']);

    // ── Worksheets ───────────────────────────────────────────────
    const sheets = wb.worksheets;
    sheets.load('items/name');
    const activeSheet = sheets.getActiveWorksheet();
    activeSheet.load('name');

    // ── Tables ───────────────────────────────────────────────────
    const tables = wb.tables;
    tables.load('items/name,items/range');

    // ── Named ranges ─────────────────────────────────────────────
    const namedItems = wb.names;
    namedItems.load('items/name,items/type');

    await context.sync();

    // ── Used range on active sheet ────────────────────────────────
    let usedRangeInfo: ExcelContext['usedRange'] = null;
    try {
      const used = activeSheet.getUsedRange();
      used.load(['address', 'rowCount', 'columnCount']);
      await context.sync();

      const usedRows = used.rowCount;
      const usedCols = used.columnCount;

      if (usedRows > 0 && usedCols > 0) {
        // Load first row (headers) and last row separately
        const headerRow = used.getRow(0);
        headerRow.load('values');

        const lastRow = used.getRow(Math.max(0, usedRows - 1));
        lastRow.load('values');

        await context.sync();

        usedRangeInfo = {
          address: used.address,
          rowCount: usedRows,
          columnCount: usedCols,
          firstRow: headerRow.values,
          lastRow: lastRow.values,
        };
      }
    } catch {
      // sheet might be empty
    }

    // ── Tables: load header ranges ────────────────────────────────
    const tableInfos: ExcelContext['tables'] = [];
    for (const table of tables.items) {
      try {
        const tRange = table.getRange();
        const headerRange = table.getHeaderRowRange();
        headerRange.load('values,address');
        tRange.load('address');
        await context.sync();
        tableInfos.push({
          name: table.name,
          range: tRange.address,
          headers: headerRange.values[0]?.map(String) ?? [],
        });
      } catch {
        // ignore inaccessible tables
      }
    }

    // ── Named ranges ─────────────────────────────────────────────
    const namedRangeInfos: ExcelContext['namedRanges'] = [];
    for (const item of namedItems.items) {
      if (item.type === Excel.NamedItemType.range) {
        try {
          const r = item.getRange();
          r.load('address');
          await context.sync();
          namedRangeInfos.push({ name: item.name, address: r.address });
        } catch {
          // ignore
        }
      }
    }

    return {
      selectedRange: {
        address: sel.address,
        values: sel.values,
        formulas: sel.formulas as string[][],
        numberFormats: sel.numberFormat as string[][],
        rowCount: sel.rowCount,
        columnCount: sel.columnCount,
      },
      worksheets: sheets.items.map((s) => ({
        name: s.name,
        isActive: s.name === activeSheet.name,
      })),
      activeSheet: activeSheet.name,
      tables: tableInfos,
      namedRanges: namedRangeInfos,
      usedRange: usedRangeInfo,
    };
  });
}
