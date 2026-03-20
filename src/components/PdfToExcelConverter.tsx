import { useState, useCallback, useRef } from "react";
import {
  Upload, FileSpreadsheet, Download, X, Loader2, FileText,
  FileCode, ArrowRight, Plus,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { toast } from "sonner";
import * as pdfjsLib from "pdfjs-dist";
import workerUrl from "pdfjs-dist/build/pdf.worker.min.js?url";
import { extractPageData, type PageResult } from "@/lib/pdf-extract";
import { buildWorkbook, downloadWorkbook } from "@/lib/excel-writer";
import {
  pagesToXml,
  parseTemplateXml,
  mergeJsonWithTemplate,
  generateSampleJson,
} from "@/lib/template-xml";

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

type Mode = "extract" | "generate";

const PdfToExcelConverter = () => {
  const [mode, setMode] = useState<Mode>("extract");

  // === Extract mode state ===
  const [pdfFile, setPdfFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isConverting, setIsConverting] = useState(false);
  const [progress, setProgress] = useState("");
  const [extractedPages, setExtractedPages] = useState<PageResult[] | null>(null);
  const pdfInputRef = useRef<HTMLInputElement>(null);

  // === Generate mode state ===
  const [xmlTemplate, setXmlTemplate] = useState<string | null>(null);
  const [xmlFileName, setXmlFileName] = useState("");
  const [jsonText, setJsonText] = useState("");
  const [sampleJson, setSampleJson] = useState("");
  const xmlInputRef = useRef<HTMLInputElement>(null);

  // ---- Extract mode handlers ----
  const handlePdfFile = useCallback((f: File) => {
    if (f.type !== "application/pdf") {
      toast.error("PDF 파일만 업로드할 수 있습니다.");
      return;
    }
    if (f.size > 20 * 1024 * 1024) {
      toast.error("파일 크기는 20MB 이하여야 합니다.");
      return;
    }
    setPdfFile(f);
    setExtractedPages(null);
  }, []);

  const onPdfDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragging(false);
      if (e.dataTransfer.files[0]) handlePdfFile(e.dataTransfer.files[0]);
    },
    [handlePdfFile]
  );

  const handleExtract = async () => {
    if (!pdfFile) return;
    setIsConverting(true);
    try {
      const buffer = await pdfFile.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
      const pages: PageResult[] = [];
      for (let i = 1; i <= pdf.numPages; i++) {
        setProgress(`페이지 ${i}/${pdf.numPages} 분석 중...`);
        const result = await extractPageData(pdf, i);
        if (result) pages.push(result);
      }
      if (pages.length === 0) throw new Error("PDF에서 데이터를 추출할 수 없습니다.");
      setExtractedPages(pages);
      toast.success("서식 추출 완료! XML을 다운로드하세요.");
    } catch (err: any) {
      toast.error(err.message || "추출 중 오류가 발생했습니다.");
    } finally {
      setIsConverting(false);
      setProgress("");
    }
  };

  const downloadXml = () => {
    if (!extractedPages) return;
    const xml = pagesToXml(extractedPages);
    const blob = new Blob([xml], { type: "application/xml" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${pdfFile?.name.replace(/\.pdf$/i, "") || "template"}_서식.xml`;
    a.click();
    URL.revokeObjectURL(url);
    toast.success("XML 서식 템플릿이 다운로드되었습니다!");
  };

  const downloadSampleJson = () => {
    if (!extractedPages) return;
    const json = generateSampleJson(extractedPages);
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "sample_data.json";
    a.click();
    URL.revokeObjectURL(url);
    toast.success("샘플 JSON이 다운로드되었습니다!");
  };

  const resetExtract = () => {
    setPdfFile(null);
    setExtractedPages(null);
    if (pdfInputRef.current) pdfInputRef.current.value = "";
  };

  // ---- Generate mode handlers ----
  const handleXmlFile = (f: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const text = e.target?.result as string;
      try {
        const sheets = parseTemplateXml(text);
        if (sheets.length === 0) throw new Error("유효한 서식 템플릿이 아닙니다.");
        setXmlTemplate(text);
        setXmlFileName(f.name);
        // 샘플 JSON 생성
        const sample: Record<string, string>[] = [];
        const sheet = sheets[0];
        for (let r = 0; r < Math.min(2, sheet.rows); r++) {
          const obj: Record<string, string> = {};
          for (let c = 0; c < sheet.cols; c++) {
            obj[`col${c}`] = `데이터${r + 1}-${c + 1}`;
          }
          sample.push(obj);
        }
        setSampleJson(JSON.stringify(sample, null, 2));
        toast.success(`서식 템플릿 로드 완료 (${sheets.length}개 시트, ${sheet.cols}개 열)`);
      } catch {
        toast.error("유효한 XML 서식 템플릿이 아닙니다.");
      }
    };
    reader.readAsText(f);
  };

  const handleGenerateExcel = () => {
    if (!xmlTemplate || !jsonText.trim()) {
      toast.error("XML 템플릿과 JSON 데이터가 모두 필요합니다.");
      return;
    }
    try {
      const data = JSON.parse(jsonText);
      if (!Array.isArray(data)) throw new Error("JSON은 배열 형태여야 합니다.");
      const sheets = parseTemplateXml(xmlTemplate);
      const pages = mergeJsonWithTemplate(sheets, data);
      const wb = buildWorkbook(pages);
      downloadWorkbook(wb, xmlFileName.replace(/\.xml$/i, "") || "output");
      toast.success("엑셀 파일이 다운로드되었습니다!");
    } catch (err: any) {
      toast.error(err.message || "엑셀 생성 중 오류가 발생했습니다.");
    }
  };

  const resetGenerate = () => {
    setXmlTemplate(null);
    setXmlFileName("");
    setJsonText("");
    setSampleJson("");
    if (xmlInputRef.current) xmlInputRef.current.value = "";
  };

  // ---- Preview ----
  const previewPage = extractedPages?.[0];
  const previewRows = previewPage?.grid.slice(0, 8) || [];

  return (
    <div className="min-h-screen bg-background flex flex-col items-center p-4 sm:p-8">
      <div className="w-full max-w-2xl space-y-6">
        {/* Header */}
        <div className="text-center space-y-3 pt-8">
          <div className="inline-flex items-center gap-3 px-4 py-2 rounded-full bg-accent">
            <FileSpreadsheet className="w-5 h-5 text-accent-foreground" />
            <span className="text-sm font-semibold text-accent-foreground tracking-wide">
              PDF 서식 → Excel 변환기
            </span>
          </div>
          <h1 className="text-3xl sm:text-4xl font-bold text-foreground tracking-tight">
            서식 보존 엑셀 생성
          </h1>
          <p className="text-muted-foreground">
            PDF에서 서식을 XML로 추출 → JSON 데이터를 넣어 Excel 다운로드
          </p>
        </div>

        {/* Mode Toggle */}
        <div className="flex rounded-xl bg-muted p-1 gap-1">
          <button
            onClick={() => setMode("extract")}
            className={`flex-1 py-3 px-4 rounded-lg text-sm font-semibold transition-all flex items-center justify-center gap-2 ${
              mode === "extract"
                ? "bg-card text-foreground shadow-sm"
                : "text-muted-foreground hover:text-foreground"
            }`}
          >
            <FileText className="w-4 h-4" />
            1. 서식 추출 (PDF → XML)
          </button>
          <button
            onClick={() => setMode("generate")}
            className={`flex-1 py-3 px-4 rounded-lg text-sm font-semibold transition-all flex items-center justify-center gap-2 ${
              mode === "generate"
                ? "bg-card text-foreground shadow-sm"
                : "text-muted-foreground hover:text-foreground"
            }`}
          >
            <FileSpreadsheet className="w-4 h-4" />
            2. 엑셀 생성 (JSON → Excel)
          </button>
        </div>

        {/* ===== EXTRACT MODE ===== */}
        {mode === "extract" && (
          <>
            {!pdfFile ? (
              <Card
                className={`relative border-2 border-dashed transition-all duration-200 cursor-pointer ${
                  isDragging
                    ? "border-primary bg-[hsl(var(--drop-zone-hover))] scale-[1.01]"
                    : "border-border bg-[hsl(var(--drop-zone))] hover:border-primary/50 hover:bg-[hsl(var(--drop-zone-hover))]"
                }`}
                onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                onDragLeave={() => setIsDragging(false)}
                onDrop={onPdfDrop}
                onClick={() => pdfInputRef.current?.click()}
              >
                <div className="flex flex-col items-center justify-center py-16 px-6 gap-4">
                  <div className="w-16 h-16 rounded-2xl bg-primary/10 flex items-center justify-center">
                    <Upload className="w-8 h-8 text-primary" />
                  </div>
                  <div className="text-center space-y-1">
                    <p className="text-foreground font-semibold text-lg">
                      PDF 파일을 드래그하거나 클릭하세요
                    </p>
                    <p className="text-muted-foreground text-sm">서식(테두리, 폰트, 레이아웃)을 XML로 추출합니다</p>
                  </div>
                </div>
                <input
                  ref={pdfInputRef}
                  type="file"
                  accept=".pdf"
                  className="hidden"
                  onChange={(e) => e.target.files?.[0] && handlePdfFile(e.target.files[0])}
                />
              </Card>
            ) : (
              <Card className="p-6 space-y-5">
                <div className="flex items-center gap-4">
                  <div className="w-12 h-12 rounded-xl bg-destructive/10 flex items-center justify-center shrink-0">
                    <FileText className="w-6 h-6 text-destructive" />
                  </div>
                  <div className="flex-1 min-w-0">
                    <p className="font-semibold text-foreground truncate">{pdfFile.name}</p>
                    <p className="text-sm text-muted-foreground">
                      {(pdfFile.size / 1024 / 1024).toFixed(2)} MB
                    </p>
                  </div>
                  <button
                    onClick={resetExtract}
                    className="p-2 rounded-lg hover:bg-muted transition-colors text-muted-foreground hover:text-foreground"
                  >
                    <X className="w-5 h-5" />
                  </button>
                </div>

                {!extractedPages ? (
                  <Button
                    className="w-full h-12 text-base font-semibold"
                    onClick={handleExtract}
                    disabled={isConverting}
                  >
                    {isConverting ? (
                      <>
                        <Loader2 className="w-5 h-5 animate-spin mr-2" />
                        {progress || "분석 중..."}
                      </>
                    ) : (
                      <>
                        <FileCode className="w-5 h-5 mr-2" />
                        서식 추출하기
                      </>
                    )}
                  </Button>
                ) : (
                  <div className="space-y-3">
                    <p className="text-sm font-medium text-muted-foreground">
                      ✅ {extractedPages.length}개 페이지에서 서식 추출 완료
                    </p>
                    <div className="flex gap-3">
                      <Button className="flex-1 h-12 font-semibold" onClick={downloadXml}>
                        <Download className="w-5 h-5 mr-2" />
                        XML 서식 다운로드
                      </Button>
                      <Button
                        variant="outline"
                        className="h-12 font-semibold"
                        onClick={downloadSampleJson}
                      >
                        샘플 JSON
                      </Button>
                    </div>
                    <Button
                      variant="secondary"
                      className="w-full"
                      onClick={() => {
                        // XML을 바로 Generate 모드로 전달
                        const xml = pagesToXml(extractedPages);
                        setXmlTemplate(xml);
                        setXmlFileName(pdfFile?.name.replace(/\.pdf$/i, "_서식.xml") || "template.xml");
                        const sample = generateSampleJson(extractedPages);
                        setSampleJson(sample);
                        setMode("generate");
                        toast.info("서식이 로드되었습니다. JSON 데이터를 입력하세요.");
                      }}
                    >
                      <ArrowRight className="w-4 h-4 mr-2" />
                      바로 엑셀 생성 모드로 이동
                    </Button>

                    {/* Preview */}
                    {previewPage && (
                      <div className="space-y-2">
                        <p className="text-xs font-medium text-muted-foreground">서식 미리보기</p>
                        <div className="max-h-48 overflow-auto rounded-lg border bg-muted/30">
                          <table className="w-full text-xs">
                            <tbody>
                              {previewRows.map((row, i) => (
                                <tr key={i} className="border-t border-border">
                                  {row.map((cell, j) => (
                                    <td
                                      key={j}
                                      className={`px-2 py-1 whitespace-nowrap text-foreground ${
                                        cell.bold ? "font-bold" : ""
                                      } ${cell.borderBottom ? "border-b border-foreground/30" : ""} ${
                                        cell.borderRight ? "border-r border-foreground/30" : ""
                                      }`}
                                    >
                                      {cell.text || <span className="text-muted-foreground/40">—</span>}
                                    </td>
                                  ))}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </Card>
            )}
          </>
        )}

        {/* ===== GENERATE MODE ===== */}
        {mode === "generate" && (
          <Card className="p-6 space-y-5">
            {/* XML Template */}
            <div className="space-y-3">
              <label className="text-sm font-semibold text-foreground">1. XML 서식 템플릿</label>
              {xmlTemplate ? (
                <div className="flex items-center gap-3 p-3 rounded-lg bg-muted">
                  <FileCode className="w-5 h-5 text-primary shrink-0" />
                  <span className="text-sm font-medium text-foreground flex-1 truncate">
                    {xmlFileName}
                  </span>
                  <button
                    onClick={resetGenerate}
                    className="p-1 rounded hover:bg-background text-muted-foreground"
                  >
                    <X className="w-4 h-4" />
                  </button>
                </div>
              ) : (
                <button
                  onClick={() => xmlInputRef.current?.click()}
                  className="w-full p-4 border-2 border-dashed border-border rounded-lg hover:border-primary/50 hover:bg-[hsl(var(--drop-zone-hover))] transition-colors flex items-center justify-center gap-2 text-muted-foreground"
                >
                  <Plus className="w-5 h-5" />
                  <span className="text-sm font-medium">XML 서식 파일 업로드</span>
                </button>
              )}
              <input
                ref={xmlInputRef}
                type="file"
                accept=".xml"
                className="hidden"
                onChange={(e) => e.target.files?.[0] && handleXmlFile(e.target.files[0])}
              />
            </div>

            {/* JSON Data */}
            <div className="space-y-3">
              <div className="flex items-center justify-between">
                <label className="text-sm font-semibold text-foreground">2. JSON 데이터</label>
                {sampleJson && (
                  <button
                    onClick={() => setJsonText(sampleJson)}
                    className="text-xs text-primary hover:underline"
                  >
                    샘플 데이터 넣기
                  </button>
                )}
              </div>
              <textarea
                className="w-full h-48 p-3 rounded-lg border border-input bg-background text-foreground text-sm font-mono resize-none focus:outline-none focus:ring-2 focus:ring-ring"
                placeholder={`[\n  { "col0": "값1", "col1": "값2", ... },\n  { "col0": "값3", "col1": "값4", ... }\n]`}
                value={jsonText}
                onChange={(e) => setJsonText(e.target.value)}
              />
              <p className="text-xs text-muted-foreground">
                배열 형태의 JSON을 입력하세요. 각 객체가 한 행이 됩니다.
              </p>
            </div>

            {/* Generate Button */}
            <Button
              className="w-full h-12 text-base font-semibold"
              onClick={handleGenerateExcel}
              disabled={!xmlTemplate || !jsonText.trim()}
            >
              <Download className="w-5 h-5 mr-2" />
              엑셀 다운로드 (.xlsx)
            </Button>
          </Card>
        )}
      </div>
    </div>
  );
};

export default PdfToExcelConverter;
