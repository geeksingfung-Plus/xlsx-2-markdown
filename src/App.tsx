import { useState } from 'react';
import { useDropzone } from 'react-dropzone';
import * as XLSX from 'xlsx';
import { FileSpreadsheet, FileText, Download, Loader2, FileCode2, CheckCircle2 } from 'lucide-react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { cn } from './lib/utils';

// --- Components ---

function FileDropzone({
  accept,
  onDrop,
  icon: Icon,
  title,
  file,
  onRemove
}: {
  accept: Record<string, string[]>,
  onDrop: (files: File[]) => void,
  icon: any,
  title: string,
  file: File | null,
  onRemove: () => void
}) {
  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    accept,
    maxFiles: 1,
    onDrop: (accepted) => onDrop(accepted)
  });

  if (file) {
    return (
      <div className="relative flex items-center justify-between p-4 border border-gray-200 rounded-xl bg-white shadow-sm">
        <div className="flex items-center gap-3 overflow-hidden">
          <div className="p-2 bg-green-50 text-green-600 rounded-lg shrink-0">
            <CheckCircle2 className="w-5 h-5" />
          </div>
          <div className="min-w-0 pr-4">
            <p className="text-sm font-medium text-gray-900 truncate">{file.name}</p>
            <p className="text-xs text-gray-500">{(file.size / 1024).toFixed(1)} KB</p>
          </div>
        </div>
        <button
          onClick={(e) => { e.stopPropagation(); onRemove(); }}
          className="px-3 py-1 text-xs font-medium text-red-600 hover:text-red-700 bg-red-50 hover:bg-red-100 rounded-md transition-colors shrink-0 cursor-pointer"
        >
          Remove
        </button>
      </div>
    );
  }

  return (
    <div
      {...getRootProps()}
      className={cn(
        "flex flex-col items-center justify-center p-8 border-2 border-dashed rounded-xl cursor-pointer transition-all text-center group",
        isDragActive ? "border-blue-500 bg-blue-50/50" : "border-gray-200 hover:border-gray-300 hover:bg-gray-50/50"
      )}
    >
      <input {...getInputProps()} />
      <div className={cn(
        "p-3 rounded-full mb-4 shadow-sm ring-1 transition-colors duration-200",
        isDragActive ? "bg-blue-100 ring-blue-200 text-blue-600" : "bg-white ring-gray-100 text-gray-400 group-hover:text-gray-600"
      )}>
        <Icon className="w-6 h-6" />
      </div>
      <p className="text-sm font-medium text-gray-900 mb-1">{title}</p>
      <p className="text-xs text-gray-500">Drag & drop or click to select</p>
    </div>
  );
}

// --- Main App ---

export default function App() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [mdFile, setMdFile] = useState<File | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [resultMd, setResultMd] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<'preview' | 'code'>('preview');

  const processExcel = async (file: File): Promise<ArrayBuffer> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        if (e.target?.result instanceof ArrayBuffer) {
          resolve(e.target.result);
        } else {
          reject(new Error("Failed to read Excel file as ArrayBuffer"));
        }
      };
      reader.onerror = () => reject(new Error("Failed to read Excel file"));
      reader.readAsArrayBuffer(file);
    });
  };

  const processMarkdown = async (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        resolve(e.target?.result as string);
      };
      reader.onerror = () => reject(new Error("Failed to read Markdown file"));
      reader.readAsText(file);
    });
  };

  const handleGenerate = async () => {
    if (!excelFile || !mdFile) return;

    setIsLoading(true);
    setError(null);
    setResultMd(null);

    // Give a tiny delay to allow UI to show loading state
    await new Promise(resolve => setTimeout(resolve, 50));

    try {
      const excelBuffer = await processExcel(excelFile);
      const mdContent = await processMarkdown(mdFile);

      const workbook = XLSX.read(excelBuffer, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json<Record<string, any>>(worksheet);

      if (rows.length === 0) {
        throw new Error("No data found in the first sheet of the Excel file.");
      }

      let finalMarkdown = '';

      rows.forEach((row, index) => {
        let rowMarkdown = mdContent;
        
        // Match {{ColumnName}} or {ColumnName} in the markdown template
        Object.keys(row).forEach(key => {
          const value = row[key] !== undefined && row[key] !== null ? String(row[key]) : '';
          
          // Escape special regex characters in the column name just in case
          const safeKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          
          const regexDouble = new RegExp(`\\{\\{${safeKey}\\}\\}`, 'gi');
          const regexSingle = new RegExp(`\\{${safeKey}\\}`, 'gi');
          
          rowMarkdown = rowMarkdown.replace(regexDouble, value).replace(regexSingle, value);
        });
        
        finalMarkdown += rowMarkdown;
        
        // Add a separator between rows if not the last row
        if (index < rows.length - 1) {
            finalMarkdown += '\n\n---\n\n';
        }
      });

      setResultMd(finalMarkdown);
    } catch (err: any) {
      console.error(err);
      setError(err.message || "An error occurred during generation.");
    } finally {
      setIsLoading(false);
    }
  };

  const handleDownload = () => {
    if (!resultMd) return;
    const blob = new Blob([resultMd], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'generated_output.md';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-gray-50/50 flex flex-col font-sans">
      <header className="bg-white border-b border-gray-200 px-6 py-4 flex items-center justify-between sticky top-0 z-10">
        <div className="flex items-center justify-between w-full max-w-5xl mx-auto">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-blue-600 text-white rounded-lg">
              <FileCode2 className="w-5 h-5" />
            </div>
            <div>
              <h1 className="text-xl font-semibold text-gray-900 tracking-tight">Data to Markdown</h1>
              <p className="text-xs text-gray-500 font-medium">Map Excel data to Markdown templates</p>
            </div>
          </div>
          <div className="text-sm font-medium text-gray-400">
            Power by Geeksing Fung
          </div>
        </div>
      </header>

      <main className="flex-1 w-full max-w-5xl mx-auto p-6 md:p-8 space-y-8">
        
        {/* Input Section */}
        <section className="bg-white rounded-2xl shadow-sm border border-gray-200/60 p-6 md:p-8">
          <div className="grid md:grid-cols-2 gap-6">
            <div className="space-y-3">
              <h2 className="text-sm font-semibold text-gray-900 flex items-center gap-2">
                <span className="flex items-center justify-center w-6 h-6 rounded-full bg-gray-100 text-gray-600 text-xs">1</span>
                Source Data (.xlsx)
              </h2>
              <FileDropzone
                accept={{
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                  'application/vnd.ms-excel': ['.xls']
                }}
                onDrop={(files) => setExcelFile(files[0])}
                icon={FileSpreadsheet}
                title="Drop Excel File"
                file={excelFile}
                onRemove={() => setExcelFile(null)}
              />
            </div>

            <div className="space-y-3">
              <h2 className="text-sm font-semibold text-gray-900 flex items-center gap-2">
                <span className="flex items-center justify-center w-6 h-6 rounded-full bg-gray-100 text-gray-600 text-xs">2</span>
                Reference Template (.md)
              </h2>
              <FileDropzone
                accept={{
                  'text/markdown': ['.md'],
                  'text/plain': ['.txt']
                }}
                onDrop={(files) => setMdFile(files[0])}
                icon={FileText}
                title="Drop Markdown Template"
                file={mdFile}
                onRemove={() => setMdFile(null)}
              />
            </div>
          </div>

          <div className="mt-8 pt-6 border-t border-gray-100 flex items-center justify-between">
            <p className="text-sm text-gray-500">
              Your Markdown template should use <code className="bg-blue-50 px-1.5 py-0.5 rounded text-blue-700 font-mono">{"{ColumnName}"}</code> tags to map data automatically.
            </p>
            <button
              onClick={handleGenerate}
              disabled={!excelFile || !mdFile || isLoading}
              className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed text-white text-sm font-medium rounded-xl shadow-sm transition-all flex items-center gap-2 cursor-pointer"
            >
              {isLoading ? (
                <>
                  <Loader2 className="w-4 h-4 animate-spin" />
                  Processing...
                </>
              ) : (
                <>
                  <FileCode2 className="w-4 h-4" />
                  Generate Output
                </>
              )}
            </button>
          </div>

          {error && (
            <div className="mt-4 p-4 bg-red-50 text-red-700 rounded-xl text-sm border border-red-100">
              <p className="font-semibold mb-1">Could not generate markdown</p>
              <p>{error}</p>
            </div>
          )}
        </section>

        {/* Result Section */}
        {resultMd && (
          <section className="bg-white rounded-2xl shadow-sm border border-gray-200/60 overflow-hidden flex flex-col h-[600px] animate-in fade-in slide-in-from-bottom-4 duration-500">
            <div className="bg-gray-50/50 border-b border-gray-200 px-4 py-3 flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2 bg-gray-100 p-1 rounded-lg">
                <button
                  onClick={() => setActiveTab('preview')}
                  className={cn(
                    "px-3 py-1.5 text-sm font-medium rounded-md transition-all cursor-pointer",
                    activeTab === 'preview' ? "bg-white text-gray-900 shadow-sm" : "text-gray-500 hover:text-gray-900"
                  )}
                >
                  Preview
                </button>
                <button
                  onClick={() => setActiveTab('code')}
                  className={cn(
                    "px-3 py-1.5 text-sm font-medium rounded-md transition-all cursor-pointer",
                    activeTab === 'code' ? "bg-white text-gray-900 shadow-sm" : "text-gray-500 hover:text-gray-900"
                  )}
                >
                  Raw Markdown
                </button>
              </div>

              <button
                onClick={handleDownload}
                className="px-4 py-2 bg-gray-900 hover:bg-gray-800 text-white text-sm font-medium rounded-lg shadow-sm transition-all flex items-center gap-2 cursor-pointer"
              >
                <Download className="w-4 h-4" />
                Download .md
              </button>
            </div>

            <div className="flex-1 overflow-auto p-6 md:p-8">
              {activeTab === 'preview' ? (
                <article className="prose prose-blue max-w-none">
                  <ReactMarkdown remarkPlugins={[remarkGfm]}>
                    {resultMd}
                  </ReactMarkdown>
                </article>
              ) : (
                <pre className="font-mono text-sm text-gray-800 whitespace-pre-wrap break-words">
                  {resultMd}
                </pre>
              )}
            </div>
          </section>
        )}
      </main>
    </div>
  );
}

