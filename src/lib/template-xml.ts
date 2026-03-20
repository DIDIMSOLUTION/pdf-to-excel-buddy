import XLSX from "xlsx-js-style";

const SS = "urn:schemas-microsoft-com:office:spreadsheet";

// ─── Interfaces ───────────────────────────────────────────

interface ParsedStyle {
  font?: { name?: string; size?: number; bold?: boolean; color?: string };
  alignment?: { horizontal?: string; vertical?: string };
  borders?: { top?: boolean; bottom?: boolean; left?: boolean; right?: boolean };
  numberFormat?: string;
}

interface ParsedCell {
  value: string;
  type: string;
  styleId: string;
  mergeAcross: number;
  mergeDown: number;
  colIndex?: number;
}

interface ParsedRow {
  cells: ParsedCell[];
  height?: number;
  index?: number;
}

// ─── Default Styles XML (from sample template) ───────────

export const DEFAULT_STYLES_XML = `
<Style ss:ID="Default" ss:Name="Normal">
 <Alignment ss:Vertical="Center"/>
 <Borders/>
 <Font ss:FontName="맑은 고딕" x:CharSet="129" ss:Size="12" ss:Color="#000000"/>
 <Interior/><NumberFormat/><Protection/>
</Style>
<Style ss:ID="m34466542392">
 <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
 <Borders>
  <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
 </Borders><NumberFormat/>
</Style>
<Style ss:ID="s72">
 <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
 <Font ss:FontName="맑은 고딕" x:CharSet="129" ss:Size="20" ss:Color="#000000" ss:Bold="1"/>
 <NumberFormat/>
</Style>
<Style ss:ID="s73">
 <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
 <Font ss:FontName="맑은 고딕" x:CharSet="129" ss:Size="12" ss:Color="#000000" ss:Bold="1"/>
 <NumberFormat/>
</Style>
<Style ss:ID="s74">
 <Font ss:FontName="맑은 고딕" x:CharSet="129" ss:Size="12" ss:Color="#000000" ss:Bold="1"/>
 <NumberFormat/>
</Style>
<Style ss:ID="s75">
 <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>
 <Font ss:FontName="맑은 고딕" x:CharSet="129" ss:Size="12" ss:Color="#000000" ss:Bold="1"/>
 <NumberFormat/>
</Style>
<Style ss:ID="s80">
 <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
 <Borders>
  <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
 </Borders><NumberFormat/>
</Style>
<Style ss:ID="s81">
 <Borders>
  <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
 </Borders><NumberFormat/>
</Style>
<Style ss:ID="s82">
 <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
 <Borders>
  <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
 </Borders>
 <NumberFormat ss:Format="yyyy\\-mm\\-dd;@"/>
</Style>
<Style ss:ID="s94">
 <Alignment ss:Vertical="Center"/>
 <Borders>
  <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
 </Borders>
</Style>
<Style ss:ID="s95">
 <Alignment ss:Vertical="Center"/>
 <Borders>
  <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
 </Borders>
</Style>
<Style ss:ID="s96">
 <Alignment ss:Vertical="Center"/>
 <Borders>
  <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
 </Borders>
 <Font ss:FontName="맑은 고딕" x:CharSet="129" ss:Size="12" ss:Color="#000000" ss:Bold="1"/>
</Style>
<Style ss:ID="s107">
 <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>
 <Borders>
  <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
 </Borders>
 <Font ss:FontName="맑은 고딕" x:CharSet="129" ss:Size="12" ss:Color="#000000" ss:Bold="1"/>
</Style>
<Style ss:ID="s109">
 <Borders>
  <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
  <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
 </Borders>
</Style>
<Style ss:ID="s111">
 <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>
 <Borders>
  <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
 </Borders>
 <Font ss:FontName="맑은 고딕" x:CharSet="129" ss:Size="12" ss:Color="#000000" ss:Bold="1"/>
</Style>`.trim();

// ─── Default section XML ──────────────────────────────────

export const DEFAULT_HEADER_XML = `<Column ss:AutoFitWidth="0" ss:Width="41"/>
<Column ss:AutoFitWidth="0" ss:Width="107"/>
<Column ss:Index="5" ss:AutoFitWidth="0" ss:Width="43"/>
<Column ss:AutoFitWidth="0" ss:Width="59"/>
<Column ss:AutoFitWidth="0" ss:Width="86"/>
<Column ss:Index="9" ss:AutoFitWidth="0" ss:Width="99"/>
<Column ss:AutoFitWidth="0" ss:Width="63"/>
<Row>
 <Cell ss:MergeAcross="9" ss:MergeDown="2" ss:StyleID="s72"><Data ss:Type="String">납품서</Data></Cell>
</Row>
<Row ss:Index="4" ss:AutoFitHeight="0" ss:Height="27">
 <Cell ss:MergeAcross="1" ss:StyleID="s73"><Data ss:Type="String">경찰서 : {{경찰서}}</Data></Cell>
 <Cell ss:MergeAcross="1" ss:StyleID="s73"><Data ss:Type="String">품목 : {{품목}}</Data></Cell>
 <Cell ss:StyleID="s74"></Cell>
 <Cell ss:StyleID="s74"></Cell>
 <Cell ss:StyleID="s74"></Cell>
 <Cell ss:StyleID="s74"></Cell>
 <Cell ss:MergeAcross="1" ss:StyleID="s75"><Data ss:Type="String">{{연도차수}}</Data></Cell>
</Row>
<Row ss:AutoFitHeight="0" ss:Height="27">
 <Cell ss:StyleID="s80"><Data ss:Type="String">No</Data></Cell>
 <Cell ss:StyleID="s80"><Data ss:Type="String">부서</Data></Cell>
 <Cell ss:StyleID="s80"><Data ss:Type="String">이름</Data></Cell>
 <Cell ss:StyleID="s80"><Data ss:Type="String">계급</Data></Cell>
 <Cell ss:StyleID="s80"><Data ss:Type="String">성별</Data></Cell>
 <Cell ss:StyleID="s80"><Data ss:Type="String">수량</Data></Cell>
 <Cell ss:StyleID="s80"><Data ss:Type="String">구분</Data></Cell>
 <Cell ss:StyleID="s80"><Data ss:Type="String">스펙</Data></Cell>
 <Cell ss:StyleID="s80"><Data ss:Type="String">출고일</Data></Cell>
 <Cell ss:StyleID="s80"><Data ss:Type="String">비고</Data></Cell>
</Row>
<Row ss:AutoFitHeight="0" ss:Height="4">
 <Cell ss:MergeAcross="9" ss:StyleID="m34466542392"></Cell>
</Row>`.trim();

export const DEFAULT_ROW_XML = `<Row ss:AutoFitHeight="0" ss:Height="27">
 <Cell ss:StyleID="s80"><Data ss:Type="Number">1</Data></Cell>
 <Cell ss:StyleID="s81"><Data ss:Type="String">부서</Data></Cell>
 <Cell ss:StyleID="s81"><Data ss:Type="String">이름</Data></Cell>
 <Cell ss:StyleID="s81"><Data ss:Type="String">계급</Data></Cell>
 <Cell ss:StyleID="s81"><Data ss:Type="String">남</Data></Cell>
 <Cell ss:StyleID="s81"><Data ss:Type="Number">1</Data></Cell>
 <Cell ss:StyleID="s81"><Data ss:Type="String">구분</Data></Cell>
 <Cell ss:StyleID="s81"><Data ss:Type="String">스펙</Data></Cell>
 <Cell ss:StyleID="s82"><Data ss:Type="DateTime">2025-11-28T00:00:00.000</Data></Cell>
 <Cell ss:StyleID="s81"><Data ss:Type="String">비고</Data></Cell>
</Row>`.trim();

export const DEFAULT_SUMMARY_XML = `<Row ss:AutoFitHeight="0" ss:Height="27">
 <Cell ss:MergeAcross="5" ss:StyleID="s107"><Data ss:Type="String">(부서계) : 1</Data></Cell>
 <Cell ss:StyleID="s95"></Cell>
 <Cell ss:StyleID="s95"></Cell>
 <Cell ss:StyleID="s96"></Cell>
 <Cell ss:StyleID="s94"></Cell>
</Row>`.trim();

export const DEFAULT_FOOTER_XML = `<Row ss:AutoFitHeight="0" ss:Height="12">
 <Cell ss:StyleID="s109"></Cell>
 <Cell ss:StyleID="s109"></Cell>
 <Cell ss:StyleID="s109"></Cell>
 <Cell ss:StyleID="s109"></Cell>
 <Cell ss:StyleID="s109"></Cell>
 <Cell ss:StyleID="s109"></Cell>
 <Cell ss:StyleID="s109"></Cell>
 <Cell ss:StyleID="s109"></Cell>
 <Cell ss:StyleID="s109"></Cell>
 <Cell ss:StyleID="s109"></Cell>
</Row>
<Row ss:AutoFitHeight="0" ss:Height="27">
 <Cell ss:MergeAcross="5" ss:StyleID="s111"><Data ss:Type="String">경찰서계 : 2</Data></Cell>
</Row>`.trim();

export const DEFAULT_JSON = `[
  [
    {"type":"row","No":"1","부서":"청문-감찰","이름":"서문륜","계급":"경감","성별":"남","수량":"1","구분":"","스펙":"무","출고일":"","비고":""},
    {"type":"summary","label":"(부서계) : 1"}
  ],
  [
    {"type":"row","No":"2","부서":"생안-교통","이름":"황정근","계급":"경감","성별":"남","수량":"1","구분":"","스펙":"무","출고일":"","비고":""},
    {"type":"summary","label":"(부서계) : 1"}
  ]
]`;

// ─── Utility Functions ────────────────────────────────────

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
    const fontEl = el.getElementsByTagName("Font")[0];
    if (fontEl) {
      style.font = {
        name: getAttr(fontEl, SS, "FontName") || undefined,
        size: getAttr(fontEl, SS, "Size") ? parseFloat(getAttr(fontEl, SS, "Size")!) : undefined,
        bold: getAttr(fontEl, SS, "Bold") === "1",
        color: getAttr(fontEl, SS, "Color") || undefined,
      };
    }
    const alignEl = el.getElementsByTagName("Alignment")[0];
    if (alignEl) {
      style.alignment = {
        horizontal: getAttr(alignEl, SS, "Horizontal") || undefined,
        vertical: getAttr(alignEl, SS, "Vertical") || undefined,
      };
    }
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
    const nfEl = el.getElementsByTagName("NumberFormat")[0];
    if (nfEl) {
      style.numberFormat = getAttr(nfEl, SS, "Format") || undefined;
    }
    map.set(id, style);
  }
  return map;
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

function resolveCellPositions(cells: ParsedCell[]): { cell: ParsedCell; col: number }[] {
  const result: { cell: ParsedCell; col: number }[] = [];
  let currentCol = 0;
  for (const cell of cells) {
    if (cell.colIndex !== undefined) currentCol = cell.colIndex - 1;
    result.push({ cell, col: currentCol });
    currentCol += 1 + cell.mergeAcross;
  }
  return result;
}

// ─── Section Parsing ──────────────────────────────────────

function wrapInWorkbook(stylesXml: string, tableContent: string): string {
  return `<?xml version="1.0"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <Styles>${stylesXml}</Styles>
 <Worksheet ss:Name="Sheet1">
  <Table>${tableContent}</Table>
 </Worksheet>
</Workbook>`;
}

function parseRowElements(tableEl: Element): ParsedRow[] {
  const rows: ParsedRow[] = [];
  // Only get direct child Row elements of this Table
  for (let i = 0; i < tableEl.children.length; i++) {
    const rowEl = tableEl.children[i];
    if (rowEl.tagName !== "Row") continue;
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
      cells.push({ value, type, styleId, mergeAcross, mergeDown, colIndex: colIdx ? parseInt(colIdx) : undefined });
    }
    rows.push({ cells, height: height ? parseFloat(height) : undefined, index: rowIdx ? parseInt(rowIdx) : undefined });
  }
  return rows;
}

function parseColumnElements(tableEl: Element): { width: number; index?: number }[] {
  const columns: { width: number; index?: number }[] = [];
  const colEls = tableEl.getElementsByTagName("Column");
  for (let i = 0; i < colEls.length; i++) {
    const el = colEls[i];
    const idx = getAttr(el, SS, "Index");
    const w = parseFloat(getAttr(el, SS, "Width") || "75");
    columns.push({ width: w, index: idx ? parseInt(idx) : undefined });
  }
  return columns;
}

export function parseSectionXml(sectionXml: string, stylesXml: string): {
  rows: ParsedRow[];
  columns: { width: number; index?: number }[];
  styles: Map<string, ParsedStyle>;
} {
  const xml = wrapInWorkbook(stylesXml, sectionXml);
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "text/xml");
  const err = doc.querySelector("parsererror");
  if (err) throw new Error("XML 파싱 오류: " + err.textContent?.slice(0, 100));
  const styles = parseStyles(doc);
  const tableEl = doc.getElementsByTagName("Table")[0];
  const rows = parseRowElements(tableEl);
  const columns = parseColumnElements(tableEl);
  return { rows, columns, styles };
}

// ─── Extract Column Headers ──────────────────────────────

export function extractColumnHeaders(rows: ParsedRow[]): string[] {
  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i];
    const cells = resolveCellPositions(row.cells);
    const labels = cells.filter((c) => c.cell.value && c.cell.mergeAcross === 0);
    if (labels.length >= 3) {
      const headers: string[] = [];
      for (const { cell, col } of cells) {
        headers[col] = cell.value;
      }
      return headers;
    }
  }
  return [];
}

// ─── Write Row to Worksheet ──────────────────────────────

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

// ─── Build XLSX from Sections ────────────────────────────

export function buildXlsxFromSections(params: {
  stylesXml: string;
  headerXml: string;
  rowXml: string;
  summaryXml: string;
  footerXml: string;
  jsonData: any[];
}): XLSX.WorkBook {
  const { stylesXml, headerXml, rowXml, summaryXml, footerXml, jsonData } = params;

  const headerSection = parseSectionXml(headerXml, stylesXml);
  const rowSection = parseSectionXml(rowXml, stylesXml);
  const summarySection = parseSectionXml(summaryXml, stylesXml);
  const footerSection = parseSectionXml(footerXml, stylesXml);

  // Merge all styles
  const allStyles = new Map<string, ParsedStyle>([
    ...headerSection.styles,
    ...rowSection.styles,
    ...summarySection.styles,
    ...footerSection.styles,
  ]);

  // Column headers for key matching
  const columnHeaders = extractColumnHeaders(headerSection.rows);

  const rowTemplate = rowSection.rows[0];
  const summaryTemplate = summarySection.rows[0];

  if (!rowTemplate) throw new Error("Row 서식에 <Row> 요소가 없습니다.");

  const wb = XLSX.utils.book_new();
  const ws: XLSX.WorkSheet = {};
  const merges: XLSX.Range[] = [];
  const rowHeights: { r: number; hpt: number }[] = [];

  // Column widths
  const colWidthMap = new Map<number, number>();
  let colPos = 0;
  for (const col of headerSection.columns) {
    if (col.index !== undefined) colPos = col.index - 1;
    colWidthMap.set(colPos, col.width);
    colPos++;
  }

  let outRow = 0;

  // 1) Header rows (preserve ss:Index gaps)
  let currentRow = 0;
  for (const row of headerSection.rows) {
    if (row.index !== undefined) currentRow = row.index - 1;
    outRow = currentRow;
    writeRow(ws, merges, outRow, row, allStyles);
    if (row.height) rowHeights.push({ r: outRow, hpt: row.height });
    currentRow++;
  }
  outRow = currentRow;

  // 2) Data rows from JSON (type = "row" | "summary")
  for (const entry of jsonData) {
    const type = entry.type || "row";

    if (type === "row") {
      const cells = resolveCellPositions(rowTemplate.cells);
      for (const { cell, col } of cells) {
        const ref = XLSX.utils.encode_cell({ r: outRow, c: col });
        const style = toXlsxStyle(allStyles.get(cell.styleId));

        // Match JSON key to column header
        let val = "";
        if (Array.isArray(entry)) {
          val = String(entry[col] ?? "");
        } else {
          const colHeader = columnHeaders[col];
          if (colHeader && colHeader in entry) {
            val = String(entry[colHeader] ?? "");
          }
        }

        let cellType = "s";
        let cellVal: any = val;
        if (val !== "" && !isNaN(Number(val))) {
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
      if (rowTemplate.height) rowHeights.push({ r: outRow, hpt: rowTemplate.height });
    } else if (type === "summary" && summaryTemplate) {
      const cells = resolveCellPositions(summaryTemplate.cells);
      for (const { cell, col } of cells) {
        const ref = XLSX.utils.encode_cell({ r: outRow, c: col });
        const style = toXlsxStyle(allStyles.get(cell.styleId));

        // For the merged label cell, use JSON "label" or "text" field
        let val = cell.value;
        if (cell.mergeAcross > 0) {
          val = entry.label || entry.text || cell.value;
        }
        ws[ref] = { t: "s", v: val, s: style };

        if (cell.mergeAcross > 0 || cell.mergeDown > 0) {
          merges.push({
            s: { r: outRow, c: col },
            e: { r: outRow + cell.mergeDown, c: col + cell.mergeAcross },
          });
        }
      }
      if (summaryTemplate.height) rowHeights.push({ r: outRow, hpt: summaryTemplate.height });
    }
    outRow++;
  }

  // 3) Footer rows
  for (const row of footerSection.rows) {
    writeRow(ws, merges, outRow, row, allStyles);
    if (row.height) rowHeights.push({ r: outRow, hpt: row.height });
    outRow++;
  }

  // Sheet properties
  const maxCol = 9;
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(outRow - 1, 0), c: maxCol } });
  ws["!merges"] = merges;
  ws["!cols"] = Array.from({ length: maxCol + 1 }, (_, i) => ({
    wch: Math.round((colWidthMap.get(i) || 75) / 7),
  }));
  ws["!rows"] = [];
  for (const { r, hpt } of rowHeights) {
    if (!ws["!rows"]![r]) ws["!rows"]![r] = {};
    ws["!rows"]![r].hpt = hpt;
  }

  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  return wb;
}
