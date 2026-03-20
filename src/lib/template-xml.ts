import type { PageResult, CellData } from "./pdf-extract";

/**
 * XML 서식 템플릿 생성/파싱
 * PDF에서 추출한 서식 정보를 XML로 저장하고, 다시 읽어올 수 있게 함
 */

function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function unescapeXml(s: string): string {
  return s
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, ">")
    .replace(/&lt;/g, "<")
    .replace(/&amp;/g, "&");
}

export function pagesToXml(pages: PageResult[]): string {
  let xml = '<?xml version="1.0" encoding="UTF-8"?>\n<template>\n';

  for (const page of pages) {
    xml += `  <sheet name="${escapeXml(page.name)}">\n`;
    xml += `    <colWidths>${page.colWidths.join(",")}</colWidths>\n`;
    xml += `    <rowHeights>${page.rowHeights.join(",")}</rowHeights>\n`;
    xml += `    <grid rows="${page.grid.length}" cols="${page.grid[0]?.length || 0}">\n`;

    for (let r = 0; r < page.grid.length; r++) {
      xml += `      <row index="${r}">\n`;
      for (let c = 0; c < page.grid[r].length; c++) {
        const cell = page.grid[r][c];
        xml += `        <cell col="${c}"`;
        xml += ` fontSize="${cell.fontSize}"`;
        xml += ` fontName="${escapeXml(cell.fontName)}"`;
        xml += ` bold="${cell.bold}"`;
        xml += ` borderTop="${cell.borderTop}"`;
        xml += ` borderBottom="${cell.borderBottom}"`;
        xml += ` borderLeft="${cell.borderLeft}"`;
        xml += ` borderRight="${cell.borderRight}"`;
        xml += ` />\n`;
      }
      xml += `      </row>\n`;
    }

    xml += `    </grid>\n`;
    xml += `  </sheet>\n`;
  }

  xml += "</template>";
  return xml;
}

export interface TemplateSheet {
  name: string;
  colWidths: number[];
  rowHeights: number[];
  rows: number;
  cols: number;
  cells: CellData[][]; // text will be empty, only style info
}

export function parseTemplateXml(xmlStr: string): TemplateSheet[] {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlStr, "text/xml");
  const sheets: TemplateSheet[] = [];

  const sheetEls = doc.querySelectorAll("sheet");
  for (const sheetEl of sheetEls) {
    const name = sheetEl.getAttribute("name") || "Sheet";
    const colWidthsStr = sheetEl.querySelector("colWidths")?.textContent || "";
    const rowHeightsStr = sheetEl.querySelector("rowHeights")?.textContent || "";
    const gridEl = sheetEl.querySelector("grid");
    const rows = parseInt(gridEl?.getAttribute("rows") || "0");
    const cols = parseInt(gridEl?.getAttribute("cols") || "0");

    const colWidths = colWidthsStr.split(",").map(Number);
    const rowHeights = rowHeightsStr.split(",").map(Number);

    const cells: CellData[][] = [];
    const rowEls = sheetEl.querySelectorAll("row");
    for (const rowEl of rowEls) {
      const rowCells: CellData[] = [];
      const cellEls = rowEl.querySelectorAll("cell");
      for (const cellEl of cellEls) {
        rowCells.push({
          text: "",
          fontSize: parseFloat(cellEl.getAttribute("fontSize") || "10"),
          fontName: unescapeXml(cellEl.getAttribute("fontName") || ""),
          bold: cellEl.getAttribute("bold") === "true",
          borderTop: cellEl.getAttribute("borderTop") === "true",
          borderBottom: cellEl.getAttribute("borderBottom") === "true",
          borderLeft: cellEl.getAttribute("borderLeft") === "true",
          borderRight: cellEl.getAttribute("borderRight") === "true",
        });
      }
      cells.push(rowCells);
    }

    sheets.push({ name, colWidths, rowHeights, rows, cols, cells });
  }

  return sheets;
}

/**
 * JSON 데이터를 템플릿에 합쳐서 PageResult[]로 변환
 * JSON 형식: [ { "col0": "값", "col1": "값", ... }, ... ] 또는
 *            [ ["값", "값", ...], ... ]
 */
export function mergeJsonWithTemplate(
  template: TemplateSheet[],
  jsonData: any[]
): PageResult[] {
  const results: PageResult[] = [];

  for (const sheet of template) {
    const grid: CellData[][] = [];
    const maxRows = Math.max(sheet.rows, jsonData.length);

    for (let r = 0; r < maxRows; r++) {
      const row: CellData[] = [];
      const dataRow = jsonData[r];

      for (let c = 0; c < sheet.cols; c++) {
        // 스타일: 템플릿에서 가져오기 (행 범위 초과 시 마지막 행 스타일 사용)
        const styleRow = r < sheet.cells.length ? r : sheet.cells.length - 1;
        const styleCell = sheet.cells[styleRow]?.[c] || {
          text: "",
          fontSize: 10,
          fontName: "",
          bold: false,
          borderTop: false,
          borderBottom: false,
          borderLeft: false,
          borderRight: false,
        };

        // 데이터: JSON에서 가져오기
        let text = "";
        if (dataRow) {
          if (Array.isArray(dataRow)) {
            text = String(dataRow[c] ?? "");
          } else {
            // 객체인 경우 키 순서대로 또는 col0, col1... 형식
            const keys = Object.keys(dataRow);
            text = String(dataRow[keys[c]] ?? "");
          }
        }

        row.push({ ...styleCell, text });
      }
      grid.push(row);
    }

    results.push({
      name: sheet.name,
      grid,
      colWidths: sheet.colWidths,
      rowHeights: sheet.rowHeights,
    });
  }

  return results;
}

export function generateSampleJson(pages: PageResult[]): string {
  if (pages.length === 0) return "[]";
  const page = pages[0];
  const sample = page.grid.slice(0, 3).map((row) => {
    const obj: Record<string, string> = {};
    row.forEach((cell, i) => {
      obj[`col${i}`] = cell.text || `값${i + 1}`;
    });
    return obj;
  });
  return JSON.stringify(sample, null, 2);
}
