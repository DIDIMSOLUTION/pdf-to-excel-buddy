import XLSX from "xlsx-js-style";

/**
 * Parse Microsoft Excel SpreadsheetML XML and convert to xlsx
 * Handles styles, merges, column widths, row heights
 */

const SS = "urn:schemas-microsoft-com:office:spreadsheet";

interface ParsedStyle {
  font?: { name?: string; size?: number; bold?: boolean; color?: string };
  alignment?: { horizontal?: string; vertical?: string };
  borders?: { top?: boolean; bottom?: boolean; left?: boolean; right?: boolean };
  numberFormat?: string;
}

interface ParsedCell {
  value: string;
  type: string; // String, Number, DateTime
  styleId: string;
  mergeAcross: number;
  mergeDown: number;
  colIndex?: number; // ss:Index (1-based)
}

interface ParsedRow {
  cells: ParsedCell[];
  height?: number;
  index?: number; // ss:Index (1-based)
}

interface ParsedSheet {
  name: string;
  columns: { width: number; index?: number }[];
  rows: ParsedRow[];
  defaultColWidth: number;
  defaultRowHeight: number;
}

export interface TemplateInfo {
  sheets: ParsedSheet[];
  styles: Map<string, ParsedStyle>;
  dataRowIndices: number[]; // which rows are "data rows" (0-based in rows array)
  headerRowCount: number;
  xmlRaw: string;
}

function getAttr(el: Element, ns: string, name: string): string | null {
  return el.getAttributeNS(ns, name) || el.getAttribute(`ss:${name}`) || null;
}

function parseStyles(doc: Document): Map<string, ParsedStyle> {
  const map = new Map<string, ParsedStyle>();
  const styleEls = doc.getElementsByTagName("Style");

  for (let i = 0; i < styleEls.length; i++) {
    const el = styleEls[i];
    const id = getAttr(el, SS, "ID") || "";
    const style: ParsedStyle = {};

    // Font
    const fontEl = el.getElementsByTagName("Font")[0];
    if (fontEl) {
      style.font = {
        name: getAttr(fontEl, SS, "FontName") || undefined,
        size: fontEl.getAttributeNS(SS, "Size")
          ? parseFloat(fontEl.getAttributeNS(SS, "Size")!)
          : (fontEl.getAttribute("ss:Size") ? parseFloat(fontEl.getAttribute("ss:Size")!) : undefined),
        bold: getAttr(fontEl, SS, "Bold") === "1",
        color: getAttr(fontEl, SS, "Color") || undefined,
      };
    }

    // Alignment
    const alignEl = el.getElementsByTagName("Alignment")[0];
    if (alignEl) {
      style.alignment = {
        horizontal: getAttr(alignEl, SS, "Horizontal") || undefined,
        vertical: getAttr(alignEl, SS, "Vertical") || undefined,
      };
    }

    // Borders
    const bordersEl = el.getElementsByTagName("Borders")[0];
    if (bordersEl) {
      const borderEls = bordersEl.getElementsByTagName("Border");
      const borders: ParsedStyle["borders"] = {};
      for (let b = 0; b < borderEls.length; b++) {
        const pos = getAttr(borderEls[b], SS, "Position");
        if (pos === "Top") borders.top = true;
        if (pos === "Bottom") borders.bottom = true;
        if (pos === "Left") borders.left = true;
        if (pos === "Right") borders.right = true;
      }
      style.borders = borders;
    }

    // NumberFormat
    const nfEl = el.getElementsByTagName("NumberFormat")[0];
    if (nfEl) {
      style.numberFormat = getAttr(nfEl, SS, "Format") || undefined;
    }

    map.set(id, style);
  }

  return map;
}

function parseSheet(wsEl: Element): ParsedSheet {
  const name = getAttr(wsEl, SS, "Name") || "Sheet1";
  const tableEl = wsEl.getElementsByTagName("Table")[0];

  const defaultColWidth = parseFloat(getAttr(tableEl, SS, "DefaultColumnWidth") || "75");
  const defaultRowHeight = parseFloat(getAttr(tableEl, SS, "DefaultRowHeight") || "18");

  // Columns
  const columns: { width: number; index?: number }[] = [];
  const colEls = tableEl.getElementsByTagName("Column");
  for (let i = 0; i < colEls.length; i++) {
    const el = colEls[i];
    const idx = getAttr(el, SS, "Index");
    const w = parseFloat(getAttr(el, SS, "Width") || String(defaultColWidth));
    columns.push({ width: w, index: idx ? parseInt(idx) : undefined });
  }

  // Rows
  const rows: ParsedRow[] = [];
  const rowEls = tableEl.getElementsByTagName("Row");
  for (let i = 0; i < rowEls.length; i++) {
    const rowEl = rowEls[i];
    const rowIdx = getAttr(rowEl, SS, "Index");
    const height = getAttr(rowEl, SS, "Height");
    const cells: ParsedCell[] = [];

    const cellEls = rowEl.getElementsByTagName("Cell");
    for (let c = 0; c < cellEls.length; c++) {
      const cellEl = cellEls[c];
      const dataEl = cellEl.getElementsByTagName("Data")[0];
      const value = dataEl?.textContent || "";
      const type = dataEl ? (getAttr(dataEl, SS, "Type") || "String") : "String";
      const styleId = getAttr(cellEl, SS, "StyleID") || "Default";
      const mergeAcross = parseInt(getAttr(cellEl, SS, "MergeAcross") || "0");
      const mergeDown = parseInt(getAttr(cellEl, SS, "MergeDown") || "0");
      const colIdx = getAttr(cellEl, SS, "Index");

      cells.push({
        value,
        type,
        styleId,
        mergeAcross,
        mergeDown,
        colIndex: colIdx ? parseInt(colIdx) : undefined,
      });
    }

    rows.push({
      cells,
      height: height ? parseFloat(height) : undefined,
      index: rowIdx ? parseInt(rowIdx) : undefined,
    });
  }

  return { name, columns, rows, defaultColWidth, defaultRowHeight };
}

export function parseSpreadsheetML(xmlStr: string): TemplateInfo {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlStr, "text/xml");

  const parseError = doc.querySelector("parsererror");
  if (parseError) {
    throw new Error("XML 파싱 오류: " + parseError.textContent);
  }

  const styles = parseStyles(doc);
  const wsEls = doc.getElementsByTagName("Worksheet");
  const sheets: ParsedSheet[] = [];

  for (let i = 0; i < wsEls.length; i++) {
    sheets.push(parseSheet(wsEls[i]));
  }

  if (sheets.length === 0) throw new Error("시트를 찾을 수 없습니다.");

  // Detect data rows: rows that have style s81 (bordered, no bold) cells are likely data rows
  // Heuristic: find rows where most cells have border styles and contain actual data
  const dataRowIndices: number[] = [];
  let headerRowCount = 0;
  const sheet = sheets[0];

  for (let r = 0; r < sheet.rows.length; r++) {
    const row = sheet.rows[r];
    const hasManyBorderedCells = row.cells.filter((c) => {
      const s = styles.get(c.styleId);
      return s?.borders && (s.borders.top || s.borders.bottom || s.borders.left || s.borders.right);
    }).length;

    const hasWideMerge = row.cells.some((c) => c.mergeAcross >= 5);
    const isSingleCellMerged = row.cells.length <= 2 && hasWideMerge;

    // Data row: many bordered cells, no wide merges, has actual data
    if (hasManyBorderedCells >= 5 && !isSingleCellMerged && row.cells.some((c) => c.value)) {
      dataRowIndices.push(r);
    }

    if (dataRowIndices.length === 0) {
      headerRowCount = r + 1;
    }
  }

  return { sheets, styles, dataRowIndices, headerRowCount, xmlRaw: xmlStr };
}

function toXlsxStyle(style: ParsedStyle | undefined): any {
  if (!style) return {};

  const s: any = {};

  if (style.font) {
    s.font = {
      name: style.font.name || "맑은 고딕",
      sz: style.font.size || 12,
      bold: style.font.bold || false,
    };
    if (style.font.color) s.font.color = { rgb: style.font.color.replace("#", "") };
  }

  if (style.alignment) {
    s.alignment = {
      horizontal: style.alignment.horizontal?.toLowerCase() || "left",
      vertical: style.alignment.vertical?.toLowerCase() || "center",
      wrapText: true,
    };
  }

  if (style.borders) {
    const thin = { style: "thin", color: { rgb: "000000" } };
    s.border = {};
    if (style.borders.top) s.border.top = thin;
    if (style.borders.bottom) s.border.bottom = thin;
    if (style.borders.left) s.border.left = thin;
    if (style.borders.right) s.border.right = thin;
  }

  if (style.numberFormat) {
    s.numFmt = style.numberFormat;
  }

  return s;
}

/**
 * Resolve actual column positions for cells considering ss:Index gaps
 */
function resolveCellPositions(cells: ParsedCell[]): { cell: ParsedCell; col: number }[] {
  const result: { cell: ParsedCell; col: number }[] = [];
  let currentCol = 0;

  for (const cell of cells) {
    if (cell.colIndex !== undefined) {
      currentCol = cell.colIndex - 1; // 1-based → 0-based
    }
    result.push({ cell, col: currentCol });
    currentCol += 1 + cell.mergeAcross;
  }

  return result;
}

/**
 * Resolve actual row positions considering ss:Index gaps
 */
function resolveRowPositions(rows: ParsedRow[]): { row: ParsedRow; rowIdx: number }[] {
  const result: { row: ParsedRow; rowIdx: number }[] = [];
  let currentRow = 0;

  for (const row of rows) {
    if (row.index !== undefined) {
      currentRow = row.index - 1; // 1-based → 0-based
    }
    result.push({ row, rowIdx: currentRow });
    currentRow++;
  }

  return result;
}

export function buildXlsxFromTemplate(
  template: TemplateInfo,
  jsonData: any[]
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  for (const sheet of template.sheets) {
    const ws: XLSX.WorkSheet = {};
    const merges: XLSX.Range[] = [];

    // Resolve column widths
    const totalCols = 10; // from the template
    const colWidthMap = new Map<number, number>();
    let colPos = 0;
    for (const col of sheet.columns) {
      if (col.index !== undefined) colPos = col.index - 1;
      colWidthMap.set(colPos, col.width);
      colPos++;
    }

    // Find the data row template (first data row)
    const dataRowTemplate = template.dataRowIndices.length > 0
      ? sheet.rows[template.dataRowIndices[0]]
      : null;

    // Find subtotal row (row after first data row, typically has wide merge)
    const subtotalRowIdx = template.dataRowIndices.length > 0
      ? template.dataRowIndices[0] + 1
      : -1;
    const subtotalRow = subtotalRowIdx >= 0 && subtotalRowIdx < sheet.rows.length
      ? sheet.rows[subtotalRowIdx]
      : null;

    // Build rows: header rows + JSON data rows + footer
    let outRow = 0;
    const rowHeights: { r: number; hpt: number }[] = [];

    // Write header rows (before data)
    const resolvedRows = resolveRowPositions(sheet.rows);

    for (let i = 0; i < template.headerRowCount && i < resolvedRows.length; i++) {
      const { row, rowIdx } = resolvedRows[i];
      // Use the actual row index for gaps
      outRow = rowIdx;
      writeRow(ws, merges, outRow, row, template.styles);
      if (row.height) rowHeights.push({ r: outRow, hpt: row.height });
      outRow++;
    }

    // Write data rows from JSON
    if (dataRowTemplate && jsonData.length > 0) {
      for (let d = 0; d < jsonData.length; d++) {
        const entry = jsonData[d];
        const cells = resolveCellPositions(dataRowTemplate.cells);

        // Map JSON values to cells
        const keys = Array.isArray(entry) ? entry : Object.values(entry);

        for (const { cell, col } of cells) {
          const ref = XLSX.utils.encode_cell({ r: outRow, c: col });
          const jsonVal = keys[col];
          const val = jsonVal !== undefined && jsonVal !== null ? String(jsonVal) : cell.value;

          const style = toXlsxStyle(template.styles.get(cell.styleId));

          // Determine cell type
          let cellType = "s";
          let cellVal: any = val;
          if (cell.type === "Number" || (!isNaN(Number(val)) && val !== "")) {
            cellType = "n";
            cellVal = Number(val);
          }

          ws[ref] = { t: cellType, v: cellVal, s: style };

          if (cell.mergeAcross > 0 || cell.mergeDown > 0) {
            merges.push({
              s: { r: outRow, c: col },
              e: { r: outRow + cell.mergeDown, c: col + cell.mergeAcross },
            });
          }
        }

        if (dataRowTemplate.height) rowHeights.push({ r: outRow, hpt: dataRowTemplate.height });
        outRow++;
      }
    }

    // Write remaining rows (footer/summary rows after data)
    const footerStart = template.dataRowIndices.length > 0
      ? template.dataRowIndices[template.dataRowIndices.length - 1] + 1
      : template.headerRowCount;

    // Skip subtotal rows that were interspersed with data
    for (let i = footerStart; i < sheet.rows.length; i++) {
      // Skip rows that look like subtotals between data rows
      const isSubtotalBetweenData = template.dataRowIndices.includes(i - 1) &&
        sheet.rows[i]?.cells.some((c) => c.mergeAcross >= 5);

      if (isSubtotalBetweenData && i < sheet.rows.length - 2) continue;

      const row = sheet.rows[i];
      writeRow(ws, merges, outRow, row, template.styles);
      if (row.height) rowHeights.push({ r: outRow, hpt: row.height });
      outRow++;
    }

    // Set sheet properties
    const maxCol = Math.max(totalCols - 1, 9);
    ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: outRow - 1, c: maxCol } });

    ws["!merges"] = merges;

    ws["!cols"] = Array.from({ length: maxCol + 1 }, (_, i) => ({
      wch: Math.round((colWidthMap.get(i) || sheet.defaultColWidth) / 7),
    }));

    ws["!rows"] = [];
    for (const { r, hpt } of rowHeights) {
      if (!ws["!rows"]![r]) ws["!rows"]![r] = {};
      ws["!rows"]![r].hpt = hpt;
    }

    XLSX.utils.book_append_sheet(wb, ws, sheet.name.slice(0, 31));
  }

  return wb;
}

function writeRow(
  ws: XLSX.WorkSheet,
  merges: XLSX.Range[],
  outRow: number,
  row: ParsedRow,
  styles: Map<string, ParsedStyle>
) {
  const cells = resolveCellPositions(row.cells);

  for (const { cell, col } of cells) {
    const ref = XLSX.utils.encode_cell({ r: outRow, c: col });
    const style = toXlsxStyle(styles.get(cell.styleId));

    let cellType = "s";
    let cellVal: any = cell.value;

    if (cell.type === "Number" && cell.value !== "") {
      cellType = "n";
      cellVal = Number(cell.value);
    } else if (cell.type === "DateTime" && cell.value) {
      // Convert DateTime to date string
      cellType = "s";
      cellVal = cell.value.split("T")[0];
    }

    ws[ref] = { t: cellType, v: cellVal, s: style };

    if (cell.mergeAcross > 0 || cell.mergeDown > 0) {
      merges.push({
        s: { r: outRow, c: col },
        e: { r: outRow + cell.mergeDown, c: col + cell.mergeAcross },
      });
    }
  }
}

/**
 * Extracts column headers from the template for JSON field reference
 */
export function getColumnHeaders(template: TemplateInfo): string[] {
  const sheet = template.sheets[0];
  if (!sheet) return [];

  // Find the header row (usually last row before data, with column labels)
  for (let i = template.headerRowCount - 1; i >= 0; i--) {
    const row = sheet.rows[i];
    const cells = resolveCellPositions(row.cells);
    const hasMultipleLabels = cells.filter((c) => c.cell.value && c.cell.mergeAcross === 0).length >= 3;
    if (hasMultipleLabels) {
      const headers: string[] = [];
      for (const { cell, col } of cells) {
        headers[col] = cell.value;
      }
      return headers;
    }
  }
  return [];
}
