import { useState, useCallback, useRef } from "react";
import { Upload, FileSpreadsheet, Download, X, Loader2, FileText } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { toast } from "sonner";
import * as pdfjsLib from "pdfjs-dist";
import workerUrl from "pdfjs-dist/build/pdf.worker.min.js?url";
import { extractPageData, type PageResult } from "@/lib/pdf-extract";
import { buildWorkbook, downloadWorkbook } from "@/lib/excel-writer";

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

interface ConvertedResult {
  fileName: string;
  pages: PageResult[];
}

const PdfToExcelConverter = () => {
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isConverting, setIsConverting] = useState(false);
  const [convertedData, setConvertedData] = useState<ConvertedResult | null>(null);
  const [progress, setProgress] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  const handleFile = useCallback((f: File) => {
    if (f.type !== "application/pdf") {
      toast.error("PDF 파일만 업로드할 수 있습니다.");
      return;
    }
    if (f.size > 20 * 1024 * 1024) {
      toast.error("파일 크기는 20MB 이하여야 합니다.");
      return;
    }
    setFile(f);
    setConvertedData(null);
  }, []);

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  }, [handleFile]);

  const handleConvert = async () => {
    if (!file) return;
    setIsConverting(true);
    try {
      const buffer = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
      const pages: PageResult[] = [];

      for (let i = 1; i <= pdf.numPages; i++) {
        setProgress(`페이지 ${i}/${pdf.numPages} 분석 중...`);
        const result = await extractPageData(pdf, i);
        if (result) pages.push(result);
      }

      if (pages.length === 0) {
        throw new Error("PDF에서 데이터를 추출할 수 없습니다.");
      }

      setConvertedData({
        fileName: file.name.replace(/\.pdf$/i, ""),
        pages,
      });
      toast.success("변환이 완료되었습니다! 서식과 테두리가 보존됩니다.");
    } catch (err: any) {
      toast.error(err.message || "변환 중 오류가 발생했습니다.");
    } finally {
      setIsConverting(false);
      setProgress("");
    }
  };

  const handleDownload = () => {
    if (!convertedData) return;
    const wb = buildWorkbook(convertedData.pages);
    downloadWorkbook(wb, convertedData.fileName);
    toast.success("파일이 다운로드되었습니다!");
  };

  const reset = () => {
    setFile(null);
    setConvertedData(null);
    if (inputRef.current) inputRef.current.value = "";
  };

  const previewPage = convertedData?.pages[0];
  const previewRows = previewPage?.grid.slice(0, 10) || [];
  const totalRows = previewPage?.grid.length || 0;
  const totalCols = previewPage?.grid[0]?.length || 0;

  return (
    <div className="min-h-screen bg-background flex flex-col items-center justify-center p-4 sm:p-8">
      <div className="w-full max-w-2xl space-y-8">
        {/* Header */}
        <div className="text-center space-y-3">
          <div className="inline-flex items-center gap-3 px-4 py-2 rounded-full bg-accent">
            <FileSpreadsheet className="w-5 h-5 text-accent-foreground" />
            <span className="text-sm font-semibold text-accent-foreground tracking-wide">
              PDF → Excel 변환기
            </span>
          </div>
          <h1 className="text-3xl sm:text-4xl font-bold text-foreground tracking-tight">
            PDF를 엑셀로 변환하세요
          </h1>
          <p className="text-muted-foreground text-lg">
            서식, 테두리, 텍스트 위치를 보존하여 Excel로 변환합니다.
          </p>
        </div>

        {/* Drop Zone */}
        {!file ? (
          <Card
            className={`relative border-2 border-dashed transition-all duration-200 cursor-pointer ${
              isDragging
                ? "border-primary bg-[hsl(var(--drop-zone-hover))] scale-[1.01]"
                : "border-border bg-[hsl(var(--drop-zone))] hover:border-primary/50 hover:bg-[hsl(var(--drop-zone-hover))]"
            }`}
            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={onDrop}
            onClick={() => inputRef.current?.click()}
          >
            <div className="flex flex-col items-center justify-center py-16 px-6 gap-4">
              <div className="w-16 h-16 rounded-2xl bg-primary/10 flex items-center justify-center">
                <Upload className="w-8 h-8 text-primary" />
              </div>
              <div className="text-center space-y-1">
                <p className="text-foreground font-semibold text-lg">
                  PDF 파일을 드래그하거나 클릭하세요
                </p>
                <p className="text-muted-foreground text-sm">최대 20MB</p>
              </div>
            </div>
            <input
              ref={inputRef}
              type="file"
              accept=".pdf"
              className="hidden"
              onChange={(e) => e.target.files?.[0] && handleFile(e.target.files[0])}
            />
          </Card>
        ) : (
          <Card className="p-6 space-y-5">
            {/* File Info */}
            <div className="flex items-center gap-4">
              <div className="w-12 h-12 rounded-xl bg-destructive/10 flex items-center justify-center shrink-0">
                <FileText className="w-6 h-6 text-destructive" />
              </div>
              <div className="flex-1 min-w-0">
                <p className="font-semibold text-foreground truncate">{file.name}</p>
                <p className="text-sm text-muted-foreground">
                  {(file.size / 1024 / 1024).toFixed(2)} MB
                </p>
              </div>
              <button
                onClick={reset}
                className="p-2 rounded-lg hover:bg-muted transition-colors text-muted-foreground hover:text-foreground"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Actions */}
            <div className="flex gap-3">
              {!convertedData ? (
                <Button
                  className="flex-1 h-12 text-base font-semibold"
                  onClick={handleConvert}
                  disabled={isConverting}
                >
                  {isConverting ? (
                    <>
                      <Loader2 className="w-5 h-5 animate-spin mr-2" />
                      {progress || "변환 중..."}
                    </>
                  ) : (
                    <>
                      <FileSpreadsheet className="w-5 h-5 mr-2" />
                      엑셀로 변환
                    </>
                  )}
                </Button>
              ) : (
                <>
                  <Button
                    className="flex-1 h-12 text-base font-semibold"
                    onClick={handleDownload}
                  >
                    <Download className="w-5 h-5 mr-2" />
                    다운로드 (.xlsx)
                  </Button>
                  <Button
                    variant="outline"
                    className="h-12"
                    onClick={reset}
                  >
                    새 파일
                  </Button>
                </>
              )}
            </div>

            {/* Preview */}
            {convertedData && previewPage && (
              <div className="space-y-3">
                <p className="text-sm font-medium text-muted-foreground">
                  {convertedData.pages.length}개 페이지 · {totalRows}행 × {totalCols}열 추출 완료
                </p>
                <div className="max-h-64 overflow-auto rounded-lg border bg-muted/30">
                  <table className="w-full text-sm">
                    <tbody>
                      {previewRows.map((row, i) => (
                        <tr key={i} className="border-t border-border">
                          {row.map((cell, j) => (
                            <td
                              key={j}
                              className={`px-3 py-2 whitespace-nowrap text-foreground ${
                                cell.bold ? "font-bold" : ""
                              } ${cell.borderBottom ? "border-b border-foreground/30" : ""} ${
                                cell.borderRight ? "border-r border-foreground/30" : ""
                              }`}
                              style={{ fontSize: `${Math.max(11, cell.fontSize * 0.75)}px` }}
                            >
                              {cell.text}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                {totalRows > 10 && (
                  <p className="text-xs text-muted-foreground text-center">
                    ... 외 {totalRows - 10}개 행
                  </p>
                )}
              </div>
            )}
          </Card>
        )}
      </div>
    </div>
  );
};

export default PdfToExcelConverter;
