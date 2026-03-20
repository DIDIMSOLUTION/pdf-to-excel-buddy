import * as pdfjsLib from "pdfjs-dist";

export interface TextItem {
  text: string;
  x: number;
  y: number;
  width: number;
  height: number;
  fontName: string;
  fontSize: number;
}

export interface LineSegment {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

export interface CellData {
  text: string;
  fontSize: number;
  fontName: string;
  bold: boolean;
  borderTop: boolean;
  borderBottom: boolean;
  borderLeft: boolean;
  borderRight: boolean;
}

export interface PageResult {
  name: string;
  grid: CellData[][];
  colWidths: number[];
  rowHeights: number[];
}

const TOLERANCE = 3;

function clusterValues(values: number[], tolerance: number): number[] {
  if (values.length === 0) return [];
  const sorted = [...new Set(values)].sort((a, b) => a - b);
  const clusters: number[] = [sorted[0]];
  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i] - clusters[clusters.length - 1] > tolerance) {
      clusters.push(sorted[i]);
    }
  }
  return clusters;
}

function findCluster(value: number, clusters: number[], tolerance: number): number {
  let best = 0;
  let bestDist = Math.abs(value - clusters[0]);
  for (let i = 1; i < clusters.length; i++) {
    const dist = Math.abs(value - clusters[i]);
    if (dist < bestDist) {
      bestDist = dist;
      best = i;
    }
  }
  return bestDist <= tolerance * 2 ? best : -1;
}

function extractLines(ops: any[]): LineSegment[] {
  const lines: LineSegment[] = [];
  let cx = 0, cy = 0, mx = 0, my = 0;

  for (const op of ops) {
    const { fn, args } = op;
    if (fn === "moveTo") {
      mx = args[0]; my = args[1]; cx = mx; cy = my;
    } else if (fn === "lineTo") {
      lines.push({ x1: cx, y1: cy, x2: args[0], y2: args[1] });
      cx = args[0]; cy = args[1];
    } else if (fn === "rectangle") {
      const [rx, ry, rw, rh] = args;
      lines.push({ x1: rx, y1: ry, x2: rx + rw, y2: ry });
      lines.push({ x1: rx + rw, y1: ry, x2: rx + rw, y2: ry + rh });
      lines.push({ x1: rx + rw, y1: ry + rh, x2: rx, y2: ry + rh });
      lines.push({ x1: rx, y1: ry + rh, x2: rx, y2: ry });
    }
  }
  return lines;
}

function isHorizontal(line: LineSegment): boolean {
  return Math.abs(line.y1 - line.y2) < TOLERANCE;
}

function isVertical(line: LineSegment): boolean {
  return Math.abs(line.x1 - line.x2) < TOLERANCE;
}

function hasBorder(
  cellX1: number, cellY1: number, cellX2: number, cellY2: number,
  lines: LineSegment[], side: "top" | "bottom" | "left" | "right"
): boolean {
  const t = TOLERANCE * 2;
  for (const l of lines) {
    if (side === "top" && isHorizontal(l)) {
      const ly = (l.y1 + l.y2) / 2;
      if (Math.abs(ly - cellY1) < t) {
        const lx1 = Math.min(l.x1, l.x2), lx2 = Math.max(l.x1, l.x2);
        if (lx1 <= cellX1 + t && lx2 >= cellX2 - t) return true;
      }
    } else if (side === "bottom" && isHorizontal(l)) {
      const ly = (l.y1 + l.y2) / 2;
      if (Math.abs(ly - cellY2) < t) {
        const lx1 = Math.min(l.x1, l.x2), lx2 = Math.max(l.x1, l.x2);
        if (lx1 <= cellX1 + t && lx2 >= cellX2 - t) return true;
      }
    } else if (side === "left" && isVertical(l)) {
      const lx = (l.x1 + l.x2) / 2;
      if (Math.abs(lx - cellX1) < t) {
        const ly1 = Math.min(l.y1, l.y2), ly2 = Math.max(l.y1, l.y2);
        if (ly1 <= cellY1 + t && ly2 >= cellY2 - t) return true;
      }
    } else if (side === "right" && isVertical(l)) {
      const lx = (l.x1 + l.x2) / 2;
      if (Math.abs(lx - cellX2) < t) {
        const ly1 = Math.min(l.y1, l.y2), ly2 = Math.max(l.y1, l.y2);
        if (ly1 <= cellY1 + t && ly2 >= cellY2 - t) return true;
      }
    }
  }
  return false;
}

export async function extractPageData(
  pdf: pdfjsLib.PDFDocumentProxy,
  pageNum: number
): Promise<PageResult | null> {
  const page = await pdf.getPage(pageNum);
  const content = await page.getTextContent();
  const opList = await page.getOperatorList();

  const items: TextItem[] = [];
  for (const item of content.items as any[]) {
    if (!item.str || item.str.trim() === "") continue;
    const tx = item.transform;
    const fontSize = Math.abs(tx[0]) || Math.abs(tx[3]) || 12;
    items.push({
      text: item.str,
      x: tx[4],
      y: tx[5],
      width: item.width || 0,
      height: item.height || fontSize,
      fontName: item.fontName || "",
      fontSize,
    });
  }

  if (items.length === 0) return null;

  // Extract drawing operations for lines
  const ops: { fn: string; args: number[] }[] = [];
  const fnMap: Record<number, string> = {};
  for (const [key, val] of Object.entries(pdfjsLib.OPS)) {
    fnMap[val as number] = key;
  }
  for (let i = 0; i < opList.fnArray.length; i++) {
    const fnName = fnMap[opList.fnArray[i]];
    if (fnName === "moveTo" || fnName === "lineTo" || fnName === "rectangle") {
      ops.push({ fn: fnName, args: opList.argsArray[i] as number[] });
    }
  }
  const lines = extractLines(ops);

  // Cluster Y positions (rows) and X positions (columns)
  const yValues = items.map((it) => Math.round(it.y));
  const xValues = items.map((it) => Math.round(it.x));

  const rowClusters = clusterValues(yValues, TOLERANCE).sort((a, b) => b - a); // top-to-bottom
  const colClusters = clusterValues(xValues, TOLERANCE).sort((a, b) => a - b); // left-to-right

  if (rowClusters.length === 0 || colClusters.length === 0) return null;

  // Calculate row boundaries (Y: high = top in PDF coordinate)
  const rowBounds: { y1: number; y2: number }[] = [];
  for (let r = 0; r < rowClusters.length; r++) {
    const midAbove = r > 0 ? (rowClusters[r - 1] + rowClusters[r]) / 2 : rowClusters[r] + 20;
    const midBelow = r < rowClusters.length - 1 ? (rowClusters[r] + rowClusters[r + 1]) / 2 : rowClusters[r] - 20;
    rowBounds.push({ y1: midAbove, y2: midBelow }); // y1 > y2 in PDF coords
  }

  // Calculate column boundaries
  const colBounds: { x1: number; x2: number }[] = [];
  for (let c = 0; c < colClusters.length; c++) {
    const midLeft = c > 0 ? (colClusters[c - 1] + colClusters[c]) / 2 : colClusters[c] - 10;
    const midRight = c < colClusters.length - 1 ? (colClusters[c] + colClusters[c + 1]) / 2 : colClusters[c] + 100;
    colBounds.push({ x1: midLeft, x2: midRight });
  }

  // Build grid
  const grid: CellData[][] = [];
  for (let r = 0; r < rowClusters.length; r++) {
    const row: CellData[] = [];
    for (let c = 0; c < colClusters.length; c++) {
      // Find items in this cell
      const cellItems = items.filter((it) => {
        const ri = findCluster(Math.round(it.y), rowClusters, TOLERANCE);
        const ci = findCluster(Math.round(it.x), colClusters, TOLERANCE);
        return ri === r && ci === c;
      });

      const text = cellItems.map((it) => it.text).join(" ").trim();
      const mainItem = cellItems[0];
      const isBold = mainItem ? /bold/i.test(mainItem.fontName) : false;
      const fontSize = mainItem ? mainItem.fontSize : 10;

      // Check borders using PDF lines
      const cX1 = colBounds[c].x1;
      const cX2 = colBounds[c].x2;
      const cY1 = rowBounds[r].y2; // lower Y in PDF
      const cY2 = rowBounds[r].y1; // upper Y in PDF

      row.push({
        text,
        fontSize,
        fontName: mainItem?.fontName || "",
        bold: isBold,
        borderTop: hasBorder(cX1, cY1, cX2, cY2, lines, "top"),
        borderBottom: hasBorder(cX1, cY1, cX2, cY2, lines, "bottom"),
        borderLeft: hasBorder(cX1, cY1, cX2, cY2, lines, "left"),
        borderRight: hasBorder(cX1, cY1, cX2, cY2, lines, "right"),
      });
    }
    grid.push(row);
  }

  // Column widths in character units (approximate)
  const colWidths = colBounds.map((b) => Math.max(8, Math.round((b.x2 - b.x1) / 6)));
  const rowHeights = rowBounds.map((b) => Math.max(15, Math.round(Math.abs(b.y1 - b.y2))));

  return { name: `Page ${pageNum}`, grid, colWidths, rowHeights };
}
