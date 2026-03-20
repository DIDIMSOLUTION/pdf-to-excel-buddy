import XLSX from "xlsx-js-style";
import type { PageResult } from "./pdf-extract";

function borderStyle(hasBorder: boolean) {
  return hasBorder ? { style: "thin", color: { rgb: "000000" } } : undefined;
}

export function buildWorkbook(pages: PageResult[]): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  for (const page of pages) {
    const ws: XLSX.WorkSheet = {};
    const numRows = page.grid.length;
    const numCols = page.grid[0]?.length || 0;

    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < numCols; c++) {
        const cell = page.grid[r][c];
        const ref = XLSX.utils.encode_cell({ r, c });

        ws[ref] = {
          t: "s",
          v: cell.text,
          s: {
            font: {
              name: "맑은 고딕",
              sz: Math.round(cell.fontSize * 0.75) || 9,
              bold: cell.bold,
            },
            alignment: {
              vertical: "center",
              wrapText: true,
            },
            border: {
              top: borderStyle(cell.borderTop),
              bottom: borderStyle(cell.borderBottom),
              left: borderStyle(cell.borderLeft),
              right: borderStyle(cell.borderRight),
            },
          },
        };
      }
    }

    ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: numRows - 1, c: numCols - 1 } });
    ws["!cols"] = page.colWidths.map((w) => ({ wch: Math.min(w, 60) }));
    ws["!rows"] = page.rowHeights.map((h) => ({ hpt: Math.min(h, 80) }));

    XLSX.utils.book_append_sheet(wb, ws, page.name.slice(0, 31));
  }

  return wb;
}

export function downloadWorkbook(wb: XLSX.WorkBook, fileName: string) {
  XLSX.writeFile(wb, `${fileName}.xlsx`);
}
