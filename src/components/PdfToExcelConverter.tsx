import { useState, useRef } from "react";
import { Upload, FileSpreadsheet, Download, X, FileCode, FileJson } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { toast } from "sonner";
import {
  parseSpreadsheetML,
  buildXlsxFromTemplate,
  getColumnHeaders,
  type TemplateInfo,
} from "@/lib/template-xml";
import { downloadWorkbook } from "@/lib/excel-writer";

const PdfToExcelConverter = () => {
  const [template, setTemplate] = useState<TemplateInfo | null>(null);
  const [xmlFileName, setXmlFileName] = useState("");
  const [columnHeaders, setColumnHeaders] = useState<string[]>([]);
  const [jsonData, setJsonData] = useState<any[] | null>(null);
  const [jsonFileName, setJsonFileName] = useState("");

  const xmlInputRef = useRef<HTMLInputElement>(null);
  const jsonInputRef = useRef<HTMLInputElement>(null);

  const handleXmlFile = (f: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const text = e.target?.result as string;
      try {
        const tmpl = parseSpreadsheetML(text);
        setTemplate(tmpl);
        setXmlFileName(f.name);
        const headers = getColumnHeaders(tmpl);
        setColumnHeaders(headers);
        const sheetInfo = tmpl.sheets[0];
        toast.success(
          `서식 로드 완료 (${tmpl.sheets.length}개 시트, ${tmpl.dataRowIndices.length}개 데이터 행 감지)`
        );
      } catch (err: any) {
        toast.error(err.message || "XML 파싱에 실패했습니다.");
      }
    };
    reader.readAsText(f);
  };

  const handleJsonFile = (f: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const text = e.target?.result as string;
      try {
        const data = JSON.parse(text);
        if (!Array.isArray(data)) throw new Error("JSON은 배열 형태여야 합니다.");
        setJsonData(data);
        setJsonFileName(f.name);
        toast.success(`JSON 로드 완료 (${data.length}개 행)`);
      } catch (err: any) {
        toast.error(err.message || "유효한 JSON 파일이 아닙니다.");
      }
    };
    reader.readAsText(f);
  };

  const handleGenerate = () => {
    if (!template || !jsonData) return;
    try {
      const wb = buildXlsxFromTemplate(template, jsonData);
      downloadWorkbook(wb, xmlFileName.replace(/\.xml$/i, "") || "output");
      toast.success("엑셀 파일이 다운로드되었습니다!");
    } catch (err: any) {
      toast.error(err.message || "엑셀 생성 중 오류가 발생했습니다.");
    }
  };

  const reset = () => {
    setTemplate(null);
    setXmlFileName("");
    setColumnHeaders([]);
    setJsonData(null);
    setJsonFileName("");
    if (xmlInputRef.current) xmlInputRef.current.value = "";
    if (jsonInputRef.current) jsonInputRef.current.value = "";
  };

  const ready = !!template && !!jsonData;

  return (
    <div className="min-h-screen bg-background flex flex-col items-center justify-center p-4 sm:p-8">
      <div className="w-full max-w-xl space-y-6">
        {/* Header */}
        <div className="text-center space-y-3">
          <div className="inline-flex items-center gap-3 px-4 py-2 rounded-full bg-accent">
            <FileSpreadsheet className="w-5 h-5 text-accent-foreground" />
            <span className="text-sm font-semibold text-accent-foreground tracking-wide">
              XML + JSON → Excel
            </span>
          </div>
          <h1 className="text-3xl sm:text-4xl font-bold text-foreground tracking-tight">
            서식 보존 엑셀 생성
          </h1>
          <p className="text-muted-foreground">
            Excel XML 서식과 JSON 데이터를 합쳐 엑셀 파일을 생성합니다.
          </p>
        </div>

        <Card className="p-6 space-y-5">
          {/* XML Upload */}
          <div className="space-y-2">
            <label className="text-sm font-semibold text-foreground flex items-center gap-2">
              <FileCode className="w-4 h-4 text-primary" />
              XML 서식 파일
            </label>
            {template ? (
              <div className="flex items-center gap-3 p-3 rounded-lg bg-muted">
                <FileCode className="w-5 h-5 text-primary shrink-0" />
                <div className="flex-1 min-w-0">
                  <p className="text-sm font-medium text-foreground truncate">{xmlFileName}</p>
                  <p className="text-xs text-muted-foreground">
                    {template.sheets.length}개 시트 · 데이터 행 {template.dataRowIndices.length}개 감지
                  </p>
                </div>
                <button
                  onClick={() => {
                    setTemplate(null);
                    setXmlFileName("");
                    setColumnHeaders([]);
                    if (xmlInputRef.current) xmlInputRef.current.value = "";
                  }}
                  className="p-1 rounded hover:bg-background text-muted-foreground hover:text-foreground"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
            ) : (
              <button
                onClick={() => xmlInputRef.current?.click()}
                className="w-full p-6 border-2 border-dashed border-border rounded-xl hover:border-primary/50 hover:bg-[hsl(var(--drop-zone-hover))] transition-all flex flex-col items-center gap-2 text-muted-foreground"
              >
                <Upload className="w-6 h-6" />
                <span className="text-sm font-medium">Excel XML 파일 업로드</span>
              </button>
            )}
            <input
              ref={xmlInputRef}
              type="file"
              accept=".xml"
              className="hidden"
              onChange={(e) => e.target.files?.[0] && handleXmlFile(e.target.files[0])}
            />

            {/* Column Headers Info */}
            {columnHeaders.length > 0 && (
              <div className="p-3 rounded-lg bg-accent/50 space-y-1">
                <p className="text-xs font-semibold text-accent-foreground">감지된 컬럼:</p>
                <div className="flex flex-wrap gap-1">
                  {columnHeaders.map(
                    (h, i) =>
                      h && (
                        <span
                          key={i}
                          className="px-2 py-0.5 rounded bg-accent text-accent-foreground text-xs font-medium"
                        >
                          {h}
                        </span>
                      )
                  )}
                </div>
                <p className="text-xs text-muted-foreground mt-1">
                  JSON 키 또는 배열 순서가 위 컬럼 순서와 일치해야 합니다.
                </p>
              </div>
            )}
          </div>

          {/* JSON Upload */}
          <div className="space-y-2">
            <label className="text-sm font-semibold text-foreground flex items-center gap-2">
              <FileJson className="w-4 h-4 text-primary" />
              JSON 데이터
            </label>
            {jsonData ? (
              <div className="flex items-center gap-3 p-3 rounded-lg bg-muted">
                <FileJson className="w-5 h-5 text-primary shrink-0" />
                <div className="flex-1 min-w-0">
                  <p className="text-sm font-medium text-foreground truncate">{jsonFileName}</p>
                  <p className="text-xs text-muted-foreground">{jsonData.length}개 행</p>
                </div>
                <button
                  onClick={() => {
                    setJsonData(null);
                    setJsonFileName("");
                    if (jsonInputRef.current) jsonInputRef.current.value = "";
                  }}
                  className="p-1 rounded hover:bg-background text-muted-foreground hover:text-foreground"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
            ) : (
              <button
                onClick={() => jsonInputRef.current?.click()}
                className="w-full p-6 border-2 border-dashed border-border rounded-xl hover:border-primary/50 hover:bg-[hsl(var(--drop-zone-hover))] transition-all flex flex-col items-center gap-2 text-muted-foreground"
              >
                <Upload className="w-6 h-6" />
                <span className="text-sm font-medium">JSON 파일 업로드</span>
              </button>
            )}
            <input
              ref={jsonInputRef}
              type="file"
              accept=".json"
              className="hidden"
              onChange={(e) => e.target.files?.[0] && handleJsonFile(e.target.files[0])}
            />
          </div>

          {/* Actions */}
          <div className="flex gap-3 pt-2">
            <Button
              className="flex-1 h-12 text-base font-semibold"
              onClick={handleGenerate}
              disabled={!ready}
            >
              <Download className="w-5 h-5 mr-2" />
              엑셀 다운로드 (.xlsx)
            </Button>
            {(template || jsonData) && (
              <Button variant="outline" className="h-12" onClick={reset}>
                초기화
              </Button>
            )}
          </div>
        </Card>
      </div>
    </div>
  );
};

export default PdfToExcelConverter;
