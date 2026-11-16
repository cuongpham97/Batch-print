import { useMemo, useState } from 'react';
import { Check, ChevronDown, ChevronUp, Copy, Download } from 'lucide-react';
import { downloadBatFile } from '../utils/batGenerator';

interface BatCodePanelProps {
  script: string;
  scriptFileName: string;
  onFileNameChange: (name: string) => void;
}

export function BatCodePanel({
  script,
  scriptFileName,
  onFileNameChange,
}: BatCodePanelProps) {
  const [copied, setCopied] = useState(false);
  const [showCode, setShowCode] = useState(false);

  const lines = useMemo(() => script.split(/\r?\n/), [script]);

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(script);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1800);
    } catch {
      const area = document.createElement('textarea');
      area.value = script;
      document.body.appendChild(area);
      area.select();
      document.execCommand('copy');
      document.body.removeChild(area);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1800);
    }
  };

  const handleDownload = () => {
    downloadBatFile(scriptFileName, script);
  };

  return (
    <section
      id="section-script"
      className="bg-white border border-slate-200 rounded-xl overflow-hidden flex flex-col"
    >
      <div className="px-4 py-2.5 border-b border-rose-800 bg-rose-700 flex items-center justify-between gap-2">
        <div className="flex items-center gap-2">
          <span className="w-2 h-2 rounded-full bg-rose-200 shrink-0" aria-hidden="true" />
          <span className="text-xs font-semibold uppercase tracking-wide text-white">
            Xuất file .bat
          </span>
        </div>
        <span className="text-[11px] font-mono font-medium text-rose-100 tabular-nums">
          {lines.length} dòng
        </span>
      </div>

      <div className="p-4 space-y-3">
        <div className="flex flex-col sm:flex-row gap-2">
          <input
            id="bat-filename-input"
            type="text"
            aria-label="Tên file .bat"
            value={scriptFileName}
            onChange={(e) => onFileNameChange(e.target.value)}
            placeholder="InHangLoat.bat"
            className="flex-1 h-9 px-2.5 text-xs font-mono text-slate-900 bg-white border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600/20 focus:border-blue-600"
          />

          <div className="flex items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={handleCopy}
              className="inline-flex items-center gap-1 h-9 px-3 text-xs font-medium text-slate-700 bg-white border border-slate-200 rounded-lg hover:bg-slate-50 transition-colors whitespace-nowrap cursor-pointer"
            >
              {copied ? (
                <>
                  <Check className="w-3.5 h-3.5 text-emerald-600" aria-hidden="true" />
                  <span>Đã chép</span>
                </>
              ) : (
                <>
                  <Copy className="w-3.5 h-3.5" aria-hidden="true" />
                  <span>Chép</span>
                </>
              )}
            </button>

            <button
              type="button"
              onClick={handleDownload}
              className="inline-flex items-center gap-1.5 h-9 px-3.5 text-xs font-semibold text-white bg-blue-600 rounded-lg hover:bg-blue-700 transition-colors whitespace-nowrap cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" aria-hidden="true" />
              <span>Tải .bat</span>
            </button>
          </div>
        </div>

        <button
          type="button"
          onClick={() => setShowCode((v) => !v)}
          className="inline-flex items-center gap-1 text-xs font-medium text-slate-600 hover:text-slate-900 cursor-pointer"
        >
          {showCode ? (
            <>
              <ChevronUp className="w-3.5 h-3.5" aria-hidden="true" />
              <span>Ẩn mã nguồn</span>
            </>
          ) : (
            <>
              <ChevronDown className="w-3.5 h-3.5" aria-hidden="true" />
              <span>Xem mã nguồn</span>
            </>
          )}
        </button>
      </div>

      {showCode && (
        <div className="bg-slate-950 text-slate-100 border-t border-slate-200 overflow-auto max-h-[380px] code-scroll">
          <table className="w-full border-collapse font-mono text-[11px] leading-4.5">
            <tbody>
              {lines.map((line, idx) => {
                const isComment =
                  line.trimStart().startsWith('::') ||
                  line.trimStart().startsWith('#');
                const isMarker = line.trim() === '###PS_WORKER_START###';
                return (
                  <tr key={idx} className="hover:bg-slate-900/80">
                    <td className="select-none text-right pr-2.5 pl-2.5 py-0 text-slate-500 tabular-nums w-9 align-top">
                      {idx + 1}
                    </td>
                    <td
                      className={`pr-3 py-0 whitespace-pre ${
                        isMarker
                          ? 'text-amber-300 font-semibold'
                          : isComment
                            ? 'text-slate-400'
                            : 'text-slate-100'
                      }`}
                    >
                      {line || ' '}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </section>
  );
}
