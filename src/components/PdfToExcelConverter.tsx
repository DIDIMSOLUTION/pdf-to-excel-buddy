import { useState, useEffect, useCallback } from "react";
import { Download, FileSpreadsheet, ChevronDown, ChevronRight } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { toast } from "sonner";
import {
  DEFAULT_STYLES_XML,
  DEFAULT_HEADER_XML,
  DEFAULT_ROW_XML,
  DEFAULT_SUMMARY_XML,
  DEFAULT_FOOTER_XML,
  DEFAULT_JSON,
  buildXlsxFromSections,
} from "@/lib/template-xml";
import { downloadWorkbook } from "@/lib/excel-writer";

interface SectionInputProps {
  label: string;
  value: string;
  onChange: (v: string) => void;
  rows?: number;
  placeholder?: string;
}

const SectionInput = ({ label, value, onChange, rows = 6, placeholder }: SectionInputProps) => (
  <div className="space-y-1.5">
    <label className="text-xs font-semibold text-muted-foreground uppercase tracking-wider">
      {label}
    </label>
    <textarea
      className="w-full rounded-lg border border-border bg-background px-3 py-2 text-xs font-mono text-foreground placeholder:text-muted-foreground focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring resize-y leading-relaxed"
      rows={rows}
      value={value}
      onChange={(e) => onChange(e.target.value)}
      placeholder={placeholder}
      spellCheck={false}
    />
  </div>
);

const PdfToExcelConverter = () => {
  const [stylesXml, setStylesXml] = useState(DEFAULT_STYLES_XML);
  const [headerXml, setHeaderXml] = useState(DEFAULT_HEADER_XML);
  const [rowXml, setRowXml] = useState(DEFAULT_ROW_XML);
  const [summaryXml, setSummaryXml] = useState(DEFAULT_SUMMARY_XML);
  const [footerXml, setFooterXml] = useState(DEFAULT_FOOTER_XML);
  const [jsonText, setJsonText] = useState(DEFAULT_JSON);
  const [showStyles, setShowStyles] = useState(false);
  const [autoDownloaded, setAutoDownloaded] = useState(false);

  const generateFromJson = useCallback((json: string, fileName = "output") => {
    const parsed = JSON.parse(json);
    let jsonData: any[];
    let templateVars: Record<string, string> = {};

    if (parsed && !Array.isArray(parsed) && parsed.data) {
      templateVars = parsed.vars || {};
      jsonData = Array.isArray(parsed.data) ? parsed.data : [parsed.data];
    } else {
      jsonData = Array.isArray(parsed) ? parsed : [parsed];
    }

    const wb = buildXlsxFromSections({
      stylesXml,
      headerXml,
      rowXml,
      summaryXml,
      footerXml,
      jsonData,
      templateVars,
    });
    downloadWorkbook(wb, fileName);
  }, [stylesXml, headerXml, rowXml, summaryXml, footerXml]);

  const handleGenerate = () => {
    try {
      generateFromJson(jsonText);
      toast.success("엑셀 파일이 다운로드되었습니다!");
    } catch (err: any) {
      toast.error(err.message || "생성 중 오류가 발생했습니다.");
    }
  };

  // URL 파라미터로 자동 다운로드: ?json=<encodedJSON>&filename=<name>
  useEffect(() => {
    if (autoDownloaded) return;
    const params = new URLSearchParams(window.location.search);
    const jsonParam = params.get("json");
    if (!jsonParam) return;

    setAutoDownloaded(true);
    try {
      const decoded = decodeURIComponent(jsonParam);
      const fileName = params.get("filename") || "output";
      setJsonText(decoded);
      // 약간의 딜레이 후 다운로드 (렌더링 완료 대기)
      setTimeout(() => {
        try {
          generateFromJson(decoded, fileName);
          toast.success("자동 다운로드 완료!");
        } catch (err: any) {
          toast.error("자동 생성 실패: " + (err.message || "오류"));
        }
      }, 500);
    } catch (err: any) {
      toast.error("URL 파라미터 파싱 실패: " + (err.message || "오류"));
    }
  }, [autoDownloaded, generateFromJson]);

  const handleReset = () => {
    setStylesXml(DEFAULT_STYLES_XML);
    setHeaderXml(DEFAULT_HEADER_XML);
    setRowXml(DEFAULT_ROW_XML);
    setSummaryXml(DEFAULT_SUMMARY_XML);
    setFooterXml(DEFAULT_FOOTER_XML);
    setJsonText(DEFAULT_JSON);
  };

  return (
    <div className="min-h-screen bg-background flex flex-col items-center p-4 sm:p-8">
      <div className="w-full max-w-2xl space-y-5">
        {/* Header */}
        <div className="text-center space-y-2 pt-4">
          <div className="inline-flex items-center gap-2 px-3 py-1.5 rounded-full bg-accent">
            <FileSpreadsheet className="w-4 h-4 text-accent-foreground" />
            <span className="text-xs font-semibold text-accent-foreground tracking-wide">
              XML 서식 + JSON → Excel
            </span>
          </div>
          <h1 className="text-2xl sm:text-3xl font-bold text-foreground tracking-tight">
            서식 보존 엑셀 생성
          </h1>
          <p className="text-sm text-muted-foreground">
            Header · Row · Summary · Footer XML 서식과 JSON 데이터를 합쳐 엑셀을 생성합니다.
          </p>
        </div>

        {/* XML Sections */}
        <Card className="p-5 space-y-4">
          <h2 className="text-sm font-bold text-foreground">서식 설정 (SpreadsheetML XML)</h2>

          {/* Styles (collapsible) */}
          <div className="border border-border rounded-lg">
            <button
              onClick={() => setShowStyles(!showStyles)}
              className="w-full flex items-center gap-2 px-3 py-2 text-xs font-semibold text-muted-foreground hover:text-foreground transition-colors"
            >
              {showStyles ? <ChevronDown className="w-3.5 h-3.5" /> : <ChevronRight className="w-3.5 h-3.5" />}
              스타일 정의 (고급)
            </button>
            {showStyles && (
              <div className="px-3 pb-3">
                <textarea
                  className="w-full rounded-lg border border-border bg-background px-3 py-2 text-xs font-mono text-foreground focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring resize-y leading-relaxed"
                  rows={10}
                  value={stylesXml}
                  onChange={(e) => setStylesXml(e.target.value)}
                  spellCheck={false}
                />
              </div>
            )}
          </div>

          <SectionInput label="Header" value={headerXml} onChange={setHeaderXml} rows={8} />
          <SectionInput label="Row (데이터 행 템플릿)" value={rowXml} onChange={setRowXml} rows={5} />
          <SectionInput label="Summary (소계 행 템플릿)" value={summaryXml} onChange={setSummaryXml} rows={4} />
          <SectionInput label="Footer" value={footerXml} onChange={setFooterXml} rows={5} />
        </Card>

        {/* JSON Data */}
        <Card className="p-5 space-y-3">
          <h2 className="text-sm font-bold text-foreground">데이터 입력 (JSON)</h2>
          <p className="text-xs text-muted-foreground">
            <code className="px-1 py-0.5 rounded bg-muted text-accent-foreground">{"{ \"vars\": {...}, \"data\": [[...]] }"}</code> 형식 — <code className="px-1 py-0.5 rounded bg-muted text-accent-foreground">vars</code>의 키가 Header XML 내 <code className="px-1 py-0.5 rounded bg-muted text-accent-foreground">{"{{키}}"}</code>를 대치합니다.
          </p>
          <textarea
            className="w-full rounded-lg border border-border bg-background px-3 py-2 text-xs font-mono text-foreground placeholder:text-muted-foreground focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring resize-y leading-relaxed"
            rows={10}
            value={jsonText}
            onChange={(e) => setJsonText(e.target.value)}
            placeholder={'{\n  "vars": {"경찰서":"서울강남","품목":"동복","연도차수":"2025년 2차"},\n  "data": [\n    [{"type":"row","No":"1","부서":"..."},{"type":"summary","label":"(부서계) : 1"}]\n  ]\n}'}
            spellCheck={false}
          />
        </Card>

        {/* Actions */}
        <div className="flex gap-3">
          <Button className="flex-1 h-11 text-sm font-semibold" onClick={handleGenerate}>
            <Download className="w-4 h-4 mr-2" />
            엑셀 다운로드 (.xlsx)
          </Button>
          <Button variant="outline" className="h-11" onClick={handleReset}>
            초기화
          </Button>
        </div>
      </div>
    </div>
  );
};

export default PdfToExcelConverter;
