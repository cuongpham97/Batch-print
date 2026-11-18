import { ReactNode, useEffect, useRef, useState } from 'react';
import { Check, CircleAlert, Copy, Download, RotateCcw, Upload, X } from 'lucide-react';
import {
  DOC_TYPE_TABS,
  DefaultSides,
  DocTypeId,
  DocTypeProfile,
  DuplexOrientationStrategy,
  ExcelFitMode,
  ExcelPrintScope,
  ExcelSheetMode,
  FolderSourceMode,
  MarginBoundaryMode,
  MarginEdgesCm,
  MarginMode,
  OrientationMode,
  OverPageAction,
  PageMargins,
  PaperSize,
  PostPrintAction,
  PrintConfig,
  PrinterMode,
  ScaleMode,
  SortOrder,
  cloneDefaultPrintConfig,
  parseImportedPrintConfig,
} from '../types/printConfig';

interface ConfigFormProps {
  config: PrintConfig;
  activeTab: DocTypeId;
  onTabChange: (tab: DocTypeId) => void;
  onChange: (updater: (prev: PrintConfig) => PrintConfig) => void;
}

const PAPER_SIZES: { value: PaperSize; label: string }[] = [
  { value: 'A4', label: 'A4 (21 × 29.7 cm)' },
  { value: 'A3', label: 'A3 (29.7 × 42 cm)' },
  { value: 'A5', label: 'A5 (14.8 × 21 cm)' },
  { value: 'Letter', label: 'Letter (21.6 × 27.9 cm)' },
  { value: 'Legal', label: 'Legal (21.6 × 35.6 cm)' },
  { value: 'B5', label: 'B5 (17.6 × 25 cm)' },
];

const DUPLEX_STRATEGY_OPTIONS: {
  value: DuplexOrientationStrategy;
  label: string;
}[] = [
  { value: 'FileDefault', label: 'Mặc định của file' },
  { value: 'FixedLongEdge', label: 'Lật cạnh dài' },
  { value: 'FixedShortEdge', label: 'Lật cạnh ngắn' },
  { value: 'PortraitLongLandscapeShort', label: 'Dọc lật dài · Ngang lật ngắn' },
  { value: 'PortraitShortLandscapeLong', label: 'Dọc lật ngắn · Ngang lật dài' },
];

const MARGIN_OPTIONS: { value: MarginMode; label: string }[] = [
  { value: 'Default', label: 'Mặc định của file' },
  { value: 'Zero', label: 'Không lề (0 cm)' },
  { value: 'Narrow', label: 'Hẹp (0.5 cm)' },
  { value: 'Normal', label: 'Thường (1.5 cm)' },
  { value: 'Custom', label: 'Tùy chỉnh (cm)' },
  { value: 'CustomByOrientation', label: 'Tùy chỉnh theo chiều trang' },
];

function getEngineOptions(
  tab: DocTypeId
): { value: string; label: string }[] {
  switch (tab) {
    case 'word':
      return [
        { value: 'OfficeCom', label: 'Microsoft Word COM' },
        { value: 'ShellFallback', label: 'Windows Shell' },
      ];
    case 'excel':
      return [
        { value: 'OfficeCom', label: 'Microsoft Excel COM' },
        { value: 'ShellFallback', label: 'Windows Shell' },
      ];
    case 'pdf':
      return [
        { value: 'SmartAuto', label: 'Tự động (WinRT / Sumatra)' },
        { value: 'WindowsNative', label: 'Windows WinRT' },
        { value: 'SumatraPDF', label: 'SumatraPDF' },
        { value: 'AdobeFoxit', label: 'Adobe / Foxit' },
        { value: 'WordCom', label: 'Microsoft Word COM (dự phòng)' },
      ];
    case 'images':
      return [{ value: 'DotNetPrint', label: '.NET PrintDocument (.webp qua Chrome)' }];
    case 'text':
      return [{ value: 'DotNetText', label: '.NET PrintDocument' }];
    case 'other':
      return [{ value: 'ShellOrChrome', label: 'Chrome Kiosk (.html/.htm) · Shell' }];
  }
}

const selectClass =
  'w-full h-9 px-2.5 text-xs text-slate-900 bg-white border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600/20 focus:border-blue-600 disabled:bg-slate-50 disabled:text-slate-400';

const DOC_TAB_THEMES: Record<
  DocTypeId,
  {
    activeTab: string;
    inactiveTab: string;
    activeDot: string;
    inactiveDot: string;
    panelBg: string;
    headerDivider: string;
    checkboxAccent: string;
  }
> = {
  word: {
    activeTab: 'bg-white text-slate-900 border-blue-600',
    inactiveTab: 'bg-slate-50 text-slate-500 border-transparent hover:text-slate-900 hover:bg-slate-100/70',
    activeDot: 'bg-blue-600',
    inactiveDot: 'bg-slate-400',
    panelBg: 'bg-white',
    headerDivider: 'border-slate-100',
    checkboxAccent: 'text-blue-600 focus:ring-blue-600',
  },
  excel: {
    activeTab: 'bg-white text-slate-900 border-blue-600',
    inactiveTab: 'bg-slate-50 text-slate-500 border-transparent hover:text-slate-900 hover:bg-slate-100/70',
    activeDot: 'bg-blue-600',
    inactiveDot: 'bg-slate-400',
    panelBg: 'bg-white',
    headerDivider: 'border-slate-100',
    checkboxAccent: 'text-blue-600 focus:ring-blue-600',
  },
  images: {
    activeTab: 'bg-white text-slate-900 border-blue-600',
    inactiveTab: 'bg-slate-50 text-slate-500 border-transparent hover:text-slate-900 hover:bg-slate-100/70',
    activeDot: 'bg-blue-600',
    inactiveDot: 'bg-slate-400',
    panelBg: 'bg-white',
    headerDivider: 'border-slate-100',
    checkboxAccent: 'text-blue-600 focus:ring-blue-600',
  },
  pdf: {
    activeTab: 'bg-white text-slate-900 border-blue-600',
    inactiveTab: 'bg-slate-50 text-slate-500 border-transparent hover:text-slate-900 hover:bg-slate-100/70',
    activeDot: 'bg-blue-600',
    inactiveDot: 'bg-slate-400',
    panelBg: 'bg-white',
    headerDivider: 'border-slate-100',
    checkboxAccent: 'text-blue-600 focus:ring-blue-600',
  },
  text: {
    activeTab: 'bg-white text-slate-900 border-blue-600',
    inactiveTab: 'bg-slate-50 text-slate-500 border-transparent hover:text-slate-900 hover:bg-slate-100/70',
    activeDot: 'bg-blue-600',
    inactiveDot: 'bg-slate-400',
    panelBg: 'bg-white',
    headerDivider: 'border-slate-100',
    checkboxAccent: 'text-blue-600 focus:ring-blue-600',
  },
  other: {
    activeTab: 'bg-white text-slate-900 border-blue-600',
    inactiveTab: 'bg-slate-50 text-slate-500 border-transparent hover:text-slate-900 hover:bg-slate-100/70',
    activeDot: 'bg-blue-600',
    inactiveDot: 'bg-slate-400',
    panelBg: 'bg-white',
    headerDivider: 'border-slate-100',
    checkboxAccent: 'text-blue-600 focus:ring-blue-600',
  },
};

const TOOLTIP_ALIGN_CLASSES = {
  col1: 'left-0',
  col2: 'right-0 sm:right-auto sm:left-1/2 sm:-translate-x-1/2',
  col3: 'left-0 sm:left-1/2 sm:-translate-x-1/2',
  col4: 'right-0',
  scaleTip: '-left-12 sm:left-1/2 sm:-translate-x-1/2',
} as const;

function HoverTip({
  text,
  content,
  align,
  wide = false,
}: {
  text: string;
  content?: ReactNode;
  align: keyof typeof TOOLTIP_ALIGN_CLASSES;
  wide?: boolean;
}) {
  return (
    <span className="relative inline-flex items-center group/tip shrink-0 z-10 hover:z-[100] focus-within:z-[100]">
      <button
        type="button"
        aria-label={text}
        className="inline-flex items-center justify-center p-0.5 rounded-full text-slate-300 hover:text-slate-600 hover:bg-slate-100 focus:text-blue-600 focus:outline-none transition-colors cursor-help"
      >
        <CircleAlert className="w-3 h-3" strokeWidth={1.6} aria-hidden="true" />
      </button>
      <div
        role="tooltip"
        className={`pointer-events-none invisible opacity-0 group-hover/tip:visible group-hover/tip:opacity-100 group-focus-within/tip:visible group-focus-within/tip:opacity-100 transition-opacity duration-150 absolute bottom-full ${TOOLTIP_ALIGN_CLASSES[align]} mb-1.5 z-[100] ${
          wide ? 'w-[min(27rem,calc(100vw-2.5rem))] sm:w-[430px]' : 'w-52'
        } p-2.5 rounded-md bg-slate-900 text-slate-50 text-[11px] font-normal leading-snug shadow-xl ring-1 ring-slate-700/80`}
      >
        {content ?? text}
      </div>
    </span>
  );
}

const LIST_PRINTERS_CMD = 'Get-Printer | Select-Object -ExpandProperty Name';

function PrinterHelpTip() {
  const [open, setOpen] = useState(false);
  const [copied, setCopied] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    const handlePointerDown = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(false);
    };
    document.addEventListener('mousedown', handlePointerDown);
    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('mousedown', handlePointerDown);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [open]);

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(LIST_PRINTERS_CMD);
      setCopied(true);
      setTimeout(() => setCopied(false), 1800);
    } catch {
      // Fallback: user can select text directly from the code block
    }
  };

  return (
    <div ref={containerRef} className="relative inline-flex items-center shrink-0">
      <button
        type="button"
        onClick={() => setOpen((v) => !v)}
        aria-expanded={open}
        aria-label="Hướng dẫn lấy danh sách tên máy in bằng PowerShell"
        title="Xem lệnh PowerShell lấy danh sách tên máy in"
        className={`inline-flex items-center justify-center p-0.5 rounded-full transition-colors cursor-pointer focus:outline-none ${
          open
            ? 'text-blue-600 bg-blue-50'
            : 'text-slate-400 hover:text-blue-600 hover:bg-slate-100'
        }`}
      >
        <CircleAlert className="w-3.5 h-3.5" strokeWidth={1.75} aria-hidden="true" />
      </button>

      {open && (
        <div
          role="dialog"
          aria-label="Hướng dẫn lấy tên máy in"
          className="absolute -left-16 sm:left-1/2 sm:-translate-x-1/2 top-full mt-1.5 z-40 w-[min(19rem,calc(100vw-3rem))] sm:w-80 rounded-lg border border-slate-700 bg-slate-900 text-slate-100 p-3 shadow-xl space-y-2"
        >
          <div className="flex items-center justify-between gap-2">
            <span className="text-[11px] font-semibold text-white">
              Lấy danh sách máy in bằng PowerShell
            </span>
            <button
              type="button"
              onClick={() => setOpen(false)}
              aria-label="Đóng hướng dẫn"
              className="p-0.5 rounded text-slate-400 hover:text-white hover:bg-slate-800 cursor-pointer"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          </div>

          <ol className="text-[11px] leading-relaxed text-slate-300 space-y-1 list-decimal list-inside">
            <li>
              Nhấn <kbd className="px-1 py-0.5 rounded bg-slate-800 border border-slate-700 text-slate-100 font-mono text-[10px]">Win + R</kbd>, gõ{' '}
              <code className="text-amber-300 font-mono">powershell</code> rồi <kbd className="px-1 py-0.5 rounded bg-slate-800 border border-slate-700 text-slate-100 font-mono text-[10px]">Enter</kbd>.
            </li>
            <li>Dán lệnh dưới đây rồi Enter để hiện đủ tên máy in:</li>
          </ol>

          <div className="flex items-center justify-between gap-1.5 rounded-md bg-slate-950 border border-slate-800 px-2 py-1.5">
            <code className="text-[11px] font-mono text-emerald-300 select-all break-all">
              {LIST_PRINTERS_CMD}
            </code>
            <button
              type="button"
              onClick={handleCopy}
              className="inline-flex items-center gap-1 px-1.5 py-1 rounded bg-slate-800 hover:bg-slate-700 text-[10px] font-medium text-slate-200 shrink-0 cursor-pointer transition-colors"
            >
              {copied ? (
                <>
                  <Check className="w-3 h-3 text-emerald-400" />
                  <span>Đã copy</span>
                </>
              ) : (
                <>
                  <Copy className="w-3 h-3" />
                  <span>Copy</span>
                </>
              )}
            </button>
          </div>

          <p className="text-[11px] leading-relaxed text-slate-300">
            3. Copy đúng dòng tên máy in bạn cần và dán vào ô <strong className="text-white font-medium">Tên máy in</strong>.
          </p>
        </div>
      )}
    </div>
  );
}

export function ConfigForm({
  config,
  activeTab,
  onTabChange,
  onChange,
}: ConfigFormProps) {
  const activeMeta = DOC_TYPE_TABS.find((t) => t.id === activeTab) ?? DOC_TYPE_TABS[0];
  const activeTabTheme = DOC_TAB_THEMES[activeTab];
  const profile = config.docProfiles[activeTab];
  const engineOptions = getEngineOptions(activeTab);

  const jsonInputRef = useRef<HTMLInputElement>(null);
  const statusTimerRef = useRef<number | null>(null);
  const [resetModalOpen, setResetModalOpen] = useState(false);
  const [otherStatus, setOtherStatus] = useState<{
    type: 'success' | 'error';
    message: string;
  } | null>(null);

  useEffect(() => {
    return () => {
      if (statusTimerRef.current !== null) {
        window.clearTimeout(statusTimerRef.current);
      }
    };
  }, []);

  useEffect(() => {
    if (!resetModalOpen) return;
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setResetModalOpen(false);
    };
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [resetModalOpen]);

  const flashOtherStatus = (type: 'success' | 'error', message: string) => {
    setOtherStatus({ type, message });
    if (statusTimerRef.current !== null) {
      window.clearTimeout(statusTimerRef.current);
    }
    statusTimerRef.current = window.setTimeout(() => {
      setOtherStatus(null);
      statusTimerRef.current = null;
    }, 3000);
  };

  const handleExportJson = () => {
    const baseName =
      config.scriptFileName.trim().replace(/\.bat$/i, '') || 'batch-print';
    const fileName = `${baseName}-config.json`;
    const blob = new Blob([JSON.stringify(config, null, 2)], {
      type: 'application/json;charset=utf-8',
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    setResetModalOpen(false);
    flashOtherStatus('success', `Đã xuất file ${fileName}`);
  };

  const handleImportJson = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    try {
      const text = await file.text();
      const parsed = JSON.parse(text) as unknown;
      const imported = parseImportedPrintConfig(parsed);
      if (!imported) {
        flashOtherStatus('error', 'File JSON không đúng cấu trúc cài đặt.');
        return;
      }
      onChange(() => imported);
      setResetModalOpen(false);
      flashOtherStatus('success', `Đã nạp cài đặt từ ${file.name}`);
    } catch {
      flashOtherStatus('error', 'Không đọc được file JSON.');
    }
  };

  const handleConfirmResetDefaults = () => {
    onChange(() => cloneDefaultPrintConfig());
    setResetModalOpen(false);
    flashOtherStatus('success', 'Đã khôi phục toàn bộ cài đặt mặc định.');
  };

  const updateField = <K extends keyof PrintConfig>(key: K, value: PrintConfig[K]) => {
    onChange((prev) => ({ ...prev, [key]: value }));
  };

  const updateProfile = <K extends keyof DocTypeProfile>(
    key: K,
    value: DocTypeProfile[K]
  ) => {
    onChange((prev) => ({
      ...prev,
      docProfiles: {
        ...prev.docProfiles,
        [activeTab]: {
          ...prev.docProfiles[activeTab],
          [key]: value,
        },
      },
    }));
  };

  const updateMarginField = <K extends keyof PageMargins>(
    key: K,
    value: PageMargins[K]
  ) => {
    updateProfile('margins', {
      ...profile.margins,
      [key]: value,
    });
  };

  const updateOrientationMarginEdge = (
    orient: 'portrait' | 'landscape',
    edge: keyof MarginEdgesCm,
    value: number
  ) => {
    updateProfile('margins', {
      ...profile.margins,
      [orient]: {
        ...profile.margins[orient],
        [edge]: value,
      },
    });
  };

  return (
    <div className="space-y-4">
      {/* 1. Cài đặt chung */}
      <section
        id="section-source"
        className="bg-white border border-slate-200 rounded-xl"
      >
        <div className="px-4 py-2.5 border-b border-blue-700 bg-blue-600 rounded-t-[11px] flex items-center gap-2">
          <span className="w-2 h-2 rounded-full bg-blue-200 shrink-0" aria-hidden="true" />
          <h2 className="text-xs font-semibold uppercase tracking-wide text-white">
            1. Cài đặt chung
          </h2>
        </div>

        <div className="p-4 space-y-3.5">
          {/* Quick Config Actions */}
          <div className="pb-3 border-b border-slate-100 space-y-2">
            <div className="flex flex-wrap items-center gap-2">
              <input
                ref={jsonInputRef}
                type="file"
                accept=".json,application/json"
                onChange={handleImportJson}
                className="hidden"
                aria-label="Nhập cài đặt từ file JSON"
              />
              <button
                type="button"
                onClick={() => jsonInputRef.current?.click()}
                className="inline-flex items-center gap-1.5 h-8 px-3 text-xs font-medium text-slate-700 bg-white border border-slate-200 rounded-lg hover:bg-slate-50 transition-colors whitespace-nowrap cursor-pointer"
              >
                <Upload className="w-3.5 h-3.5" aria-hidden="true" />
                <span>Nhập từ JSON</span>
              </button>

              <button
                type="button"
                onClick={handleExportJson}
                className="inline-flex items-center gap-1.5 h-8 px-3 text-xs font-medium text-slate-700 bg-white border border-slate-200 rounded-lg hover:bg-slate-50 transition-colors whitespace-nowrap cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" aria-hidden="true" />
                <span>Xuất cài đặt</span>
              </button>

              <button
                type="button"
                onClick={() => setResetModalOpen(true)}
                className="inline-flex items-center gap-1.5 h-8 px-3 text-xs font-medium rounded-lg border text-rose-700 bg-rose-50/60 border-rose-200 hover:bg-rose-100/70 transition-colors whitespace-nowrap cursor-pointer sm:ml-auto"
              >
                <RotateCcw className="w-3.5 h-3.5" aria-hidden="true" />
                <span>Khôi phục cài đặt mặc định</span>
              </button>
            </div>

            {otherStatus && (
              <p
                role="status"
                className={`text-xs font-medium flex items-center gap-1.5 ${
                  otherStatus.type === 'success' ? 'text-emerald-700' : 'text-rose-600'
                }`}
              >
                {otherStatus.type === 'success' && (
                  <Check className="w-3.5 h-3.5 shrink-0" aria-hidden="true" />
                )}
                <span>{otherStatus.message}</span>
              </p>
            )}
          </div>

          {/* 2-Column Layout for General Settings */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-3">
            {/* Column 1: Thư mục */}
            <div className="space-y-3">
              <div className="space-y-1">
                <label
                  htmlFor="folder-source-select"
                  className="block text-xs font-medium text-slate-700"
                >
                  Thư mục nguồn
                </label>
                <select
                  id="folder-source-select"
                  value={config.folderSourceMode}
                  onChange={(e) =>
                    updateField('folderSourceMode', e.target.value as FolderSourceMode)
                  }
                  className={selectClass}
                >
                  <option value="ScriptDir">Thư mục chứa file .bat (%~dp0)</option>
                  <option value="CustomPath">Đường dẫn cố định</option>
                </select>
              </div>

              {config.folderSourceMode === 'CustomPath' && (
                <div className="space-y-1">
                  <label
                    htmlFor="custom-folder-input"
                    className="block text-xs font-medium text-slate-700"
                  >
                    Đường dẫn thư mục
                  </label>
                  <input
                    id="custom-folder-input"
                    type="text"
                    value={config.customFolderPath}
                    onChange={(e) => updateField('customFolderPath', e.target.value)}
                    placeholder="VD: D:\TaiLieuCanIn"
                    className={`${selectClass} font-mono`}
                  />
                </div>
              )}

              <div className="space-y-1">
                <label
                  htmlFor="sort-order"
                  className="block text-xs font-medium text-slate-700"
                >
                  Thứ tự in
                </label>
                <select
                  id="sort-order"
                  value={config.sortOrder}
                  onChange={(e) => updateField('sortOrder', e.target.value as SortOrder)}
                  className={selectClass}
                >
                  <option value="FolderOrder">Theo thứ tự tìm thấy trong thư mục</option>
                  <option value="NameAsc">Tên A → Z</option>
                  <option value="NameDesc">Tên Z → A</option>
                  <option value="DateOldest">Cũ nhất trước</option>
                  <option value="DateNewest">Mới nhất trước</option>
                </select>
              </div>

              <div className="space-y-1">
                <div className="flex items-center gap-1">
                  <label
                    htmlFor="margin-boundary-select"
                    className="block text-xs font-medium text-slate-700"
                  >
                    Giới hạn lề khi in
                  </label>
                  <HoverTip
                    align="col1"
                    wide
                    text="Luôn theo lề đã cài đặt: giữ đúng số cm bạn đặt (có thể bị cắt sát mép nếu nhỏ hơn giới hạn vật lý). Nhỏ nhất theo lề giới hạn vật lý của máy in: tự nâng lên mức tối thiểu máy in in được để không mất chữ/ảnh."
                    content={
                      <table className="w-full border-collapse text-[10.5px] leading-snug border border-slate-600/80 rounded overflow-hidden">
                        <thead>
                          <tr className="bg-slate-900/90 text-slate-200">
                            <th className="text-left font-semibold px-2 py-1 border-b border-r border-slate-600/80">
                              Lựa chọn
                            </th>
                            <th className="text-left font-semibold px-2 py-1 border-b border-slate-600/80">
                              Mô tả
                            </th>
                          </tr>
                        </thead>
                        <tbody className="text-slate-200">
                          <tr className="border-b border-slate-700/80">
                            <td className="px-2 py-1.5 font-semibold text-white border-r border-slate-700/80 whitespace-nowrap align-top">
                              Luôn theo lề đã cài đặt
                            </td>
                            <td className="px-2 py-1.5 align-top">
                              Giữ đúng số cm bạn đã đặt (VD <strong className="text-white">0 cm</strong> sát mép giấy); phần nằm ngoài vùng in vật lý của máy in thường sẽ bị cắt.
                            </td>
                          </tr>
                          <tr>
                            <td className="px-2 py-1.5 font-semibold text-white border-r border-slate-700/80 align-top">
                              Nhỏ nhất theo lề giới hạn vật lý của máy in
                            </td>
                            <td className="px-2 py-1.5 align-top">
                              Nếu lề cài đặt nhỏ hơn giới hạn vật lý, tự nâng lên mức tối thiểu máy in in được để không mất chữ/ảnh; nếu lớn hơn vẫn giữ đúng số cm đã cài.
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    }
                  />
                </div>
                <select
                  id="margin-boundary-select"
                  value={config.marginBoundaryMode}
                  onChange={(e) =>
                    updateField('marginBoundaryMode', e.target.value as MarginBoundaryMode)
                  }
                  className={selectClass}
                >
                  <option value="StrictConfigured">Luôn theo lề đã cài đặt</option>
                  <option value="ClampToPrinterMin">
                    Nhỏ nhất theo lề giới hạn vật lý của máy in
                  </option>
                </select>
              </div>
            </div>

            {/* Column 2: Máy in & Thực thi */}
            <div className="space-y-3">
              <div className="space-y-1">
                <label
                  htmlFor="printer-mode-select"
                  className="block text-xs font-medium text-slate-700"
                >
                  Máy in
                </label>
                <select
                  id="printer-mode-select"
                  value={config.printerMode}
                  onChange={(e) =>
                    updateField('printerMode', e.target.value as PrinterMode)
                  }
                  className={selectClass}
                >
                  <option value="WindowsDefault">Mặc định của Windows</option>
                  <option value="SpecificPrinter">Chỉ định tên máy in</option>
                  <option value="AskOnRun">Hỏi khi chạy</option>
                </select>
              </div>

              {(config.printerMode === 'SpecificPrinter' ||
                config.printerMode === 'AskOnRun') && (
                <div className="space-y-1.5">
                  {config.printerMode === 'SpecificPrinter' ? (
                    <>
                      <div className="flex items-center gap-1">
                        <label
                          htmlFor="printer-name-input"
                          className="block text-xs font-medium text-slate-700"
                        >
                          Tên máy in
                        </label>
                        <PrinterHelpTip />
                      </div>
                      <input
                        id="printer-name-input"
                        type="text"
                        value={config.printerName}
                        onChange={(e) => updateField('printerName', e.target.value)}
                        placeholder="VD: Canon LBP2900 hoặc \\PC\HP_M404"
                        className={`${selectClass} font-mono`}
                      />
                    </>
                  ) : (
                    <p className="text-[11px] leading-snug text-slate-500">
                      Khi chạy file .bat, cửa sổ sẽ liệt kê các máy in trên Windows để bạn chọn theo số thứ tự.
                    </p>
                  )}
                  <label className="inline-flex items-center gap-2 text-xs text-slate-700 cursor-pointer select-none">
                    <input
                      type="checkbox"
                      checked={config.setAsWindowsDefault}
                      onChange={(e) =>
                        updateField('setAsWindowsDefault', e.target.checked)
                      }
                      className="w-3.5 h-3.5 rounded border-slate-300 text-blue-600 focus:ring-blue-600"
                    />
                    <span>Đặt làm mặc định Windows khi chạy</span>
                  </label>
                </div>
              )}

              <div className="grid grid-cols-2 gap-2.5">
                <div className="space-y-1">
                  <label
                    htmlFor="post-action-select"
                    className="block text-xs font-medium text-slate-700"
                  >
                    Sau khi in
                  </label>
                  <select
                    id="post-action-select"
                    value={config.postPrintAction}
                    onChange={(e) =>
                      updateField('postPrintAction', e.target.value as PostPrintAction)
                    }
                    className={selectClass}
                  >
                    <option value="Keep">Giữ nguyên file</option>
                    <option value="MoveToPrinted">Chuyển vào _DaIn</option>
                    <option value="Delete">Xóa file</option>
                  </select>
                </div>

                <div className="space-y-1">
                  <label
                    htmlFor="delay-select"
                    className="block text-xs font-medium text-slate-700"
                  >
                    Trễ mỗi file
                  </label>
                  <select
                    id="delay-select"
                    value={config.jobDelaySeconds}
                    onChange={(e) => updateField('jobDelaySeconds', Number(e.target.value))}
                    className={`${selectClass} tabular-nums`}
                  >
                    <option value={0}>0 giây</option>
                    <option value={1}>1 giây</option>
                    <option value={2}>2 giây</option>
                    <option value={3}>3 giây</option>
                    <option value={5}>5 giây</option>
                  </select>
                </div>
              </div>
            </div>
          </div>

          {/* Checkboxes row */}
          <div className="pt-2.5 border-t border-slate-100 grid grid-cols-2 sm:grid-cols-4 gap-2">
            <div className="inline-flex items-center gap-1 min-w-0">
              <label className="inline-flex items-center gap-1.5 text-xs text-slate-700 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={config.includeSubfolders}
                  onChange={(e) => updateField('includeSubfolders', e.target.checked)}
                  className="w-3.5 h-3.5 rounded border-slate-300 text-blue-600 focus:ring-blue-600 shrink-0"
                />
                <span>Quét thư mục con</span>
              </label>
              <HoverTip
                align="col1"
                text="Quét và in cả các file nằm trong các thư mục con lồng bên trong."
              />
            </div>

            <div className="inline-flex items-center gap-1 min-w-0">
              <label className="inline-flex items-center gap-1.5 text-xs text-slate-700 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={config.restorePrinterConfig}
                  onChange={(e) => updateField('restorePrinterConfig', e.target.checked)}
                  className="w-3.5 h-3.5 rounded border-slate-300 text-blue-600 focus:ring-blue-600 shrink-0"
                />
                <span>Khôi phục máy in</span>
              </label>
              <HoverTip
                align="col2"
                text="Tự trả khổ giấy và chế độ in 1/2 mặt của máy in về trạng thái ban đầu sau khi in xong."
              />
            </div>

            <div className="inline-flex items-center gap-1 min-w-0">
              <label className="inline-flex items-center gap-1.5 text-xs text-slate-700 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={config.autoCloseConsole}
                  onChange={(e) => updateField('autoCloseConsole', e.target.checked)}
                  className="w-3.5 h-3.5 rounded border-slate-300 text-blue-600 focus:ring-blue-600 shrink-0"
                />
                <span>Tự đóng cửa sổ</span>
              </label>
              <HoverTip
                align="col3"
                text="Tự tắt cửa sổ CMD ngay khi chạy xong (bỏ tích để dừng lại xem kết quả in)."
              />
            </div>

            <div className="inline-flex items-center gap-1 min-w-0">
              <label className="inline-flex items-center gap-1.5 text-xs text-slate-700 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={config.exportLogTxt}
                  onChange={(e) => updateField('exportLogTxt', e.target.checked)}
                  className="w-3.5 h-3.5 rounded border-slate-300 text-blue-600 focus:ring-blue-600 shrink-0"
                />
                <span>Xuất kết quả ra .txt</span>
              </label>
              <HoverTip
                align="col4"
                text="Tự động tạo file log .txt ghi lại toàn bộ kết quả in ấn trong thư mục nguồn khi chạy xong."
              />
            </div>
          </div>
        </div>
      </section>

      {/* 2. Cài đặt theo loại tài liệu */}
      <section
        id="section-doctypes"
        className="bg-white border border-slate-200 rounded-xl"
      >
        <div className="px-4 py-2.5 border-b border-amber-800 bg-amber-700 rounded-t-[11px] flex items-center gap-2">
          <span className="w-2 h-2 rounded-full bg-amber-200 shrink-0" aria-hidden="true" />
          <h2 className="text-xs font-semibold uppercase tracking-wide text-white">
            2. Cài đặt theo loại tài liệu
          </h2>
        </div>

        {/* 6 Document Type Tabs */}
        <div
          role="tablist"
          aria-label="Loại tài liệu"
          className="grid grid-cols-3 sm:grid-cols-6 border-b border-slate-200 bg-slate-50"
        >
          {DOC_TYPE_TABS.map((tab) => {
            const isSelected = tab.id === activeTab;
            const isEnabled = config.docProfiles[tab.id].enabled;
            const tabTheme = DOC_TAB_THEMES[tab.id];
            return (
              <button
                key={tab.id}
                type="button"
                role="tab"
                aria-selected={isSelected}
                onClick={() => onTabChange(tab.id)}
                className={`px-2 py-2.5 text-xs font-semibold border-b-2 transition-colors whitespace-nowrap shrink-0 cursor-pointer flex items-center justify-center ${
                  isSelected ? tabTheme.activeTab : tabTheme.inactiveTab
                }`}
              >
                <span className={isEnabled ? '' : 'line-through opacity-60'}>
                  {tab.label}
                </span>
              </button>
            );
          })}
        </div>

        {/* Active Tab Profile Panel */}
        <div role="tabpanel" className={`p-4 space-y-3.5 rounded-b-[11px] transition-colors ${activeTabTheme.panelBg}`}>
          {/* Enable row */}
          <div className={`flex flex-wrap items-center justify-between gap-2 pb-3 border-b ${activeTabTheme.headerDivider}`}>
            <label className="inline-flex items-center gap-2 text-xs font-semibold text-slate-900 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={profile.enabled}
                onChange={(e) => updateProfile('enabled', e.target.checked)}
                className={`w-4 h-4 rounded border-slate-300 ${activeTabTheme.checkboxAccent}`}
              />
              <span>Bật in {activeMeta.label}</span>
            </label>
            <span className="text-[11px] font-mono text-slate-500">
              {activeTab === 'other'
                ? profile.customExtensions?.trim() || 'Chưa nhập đuôi'
                : activeMeta.extensionsLabel}
            </span>
          </div>

          {activeTab === 'other' && (
            <div className="space-y-1">
              <label
                htmlFor="other-extensions-input"
                className="block text-xs font-medium text-slate-700"
              >
                Đuôi mở rộng tùy chỉnh
              </label>
              <input
                id="other-extensions-input"
                type="text"
                value={profile.customExtensions || ''}
                onChange={(e) => updateProfile('customExtensions', e.target.value)}
                placeholder=".html, .htm, .xps"
                className={`${selectClass} font-mono`}
              />
            </div>
          )}

          {/* Two clean columns for all document settings */}
          <div
            className={
              profile.enabled ? 'space-y-3.5' : 'space-y-3.5 opacity-45 pointer-events-none'
            }
          >
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-3">
              {/* Column 1: Giấy & Mặt in */}
              <div className="space-y-3">
                <div className="space-y-1">
                  <label
                    htmlFor={`paper-${activeTab}`}
                    className="block text-xs font-medium text-slate-700"
                  >
                    Khổ giấy mặc định
                  </label>
                  <select
                    id={`paper-${activeTab}`}
                    value={profile.defaultPaperSize}
                    onChange={(e) =>
                      updateProfile('defaultPaperSize', e.target.value as PaperSize)
                    }
                    className={selectClass}
                  >
                    {PAPER_SIZES.map((p) => (
                      <option key={p.value} value={p.value}>
                        {p.label}
                      </option>
                    ))}
                  </select>
                </div>

                <div className="space-y-1">
                  <label
                    htmlFor={`sides-${activeTab}`}
                    className="block text-xs font-medium text-slate-700"
                  >
                    Số mặt mặc định
                  </label>
                  <select
                    id={`sides-${activeTab}`}
                    value={profile.defaultSides}
                    onChange={(e) =>
                      updateProfile('defaultSides', e.target.value as DefaultSides)
                    }
                    className={selectClass}
                  >
                    <option value="OneSided">1 mặt</option>
                    <option value="TwoSided">2 mặt</option>
                  </select>
                </div>

                <div className="space-y-1">
                  <label
                    htmlFor={`duplex-strategy-${activeTab}`}
                    className="block text-xs font-medium text-slate-700"
                  >
                    Kiểu lật cạnh dài / ngắn
                  </label>
                  <select
                    id={`duplex-strategy-${activeTab}`}
                    disabled={profile.defaultSides !== 'TwoSided'}
                    value={
                      activeTab !== 'pdf' && profile.duplexStrategy === 'FileDefault'
                        ? 'PortraitLongLandscapeShort'
                        : profile.duplexStrategy
                    }
                    onChange={(e) =>
                      updateProfile(
                        'duplexStrategy',
                        e.target.value as DuplexOrientationStrategy
                      )
                    }
                    className={selectClass}
                  >
                    {DUPLEX_STRATEGY_OPTIONS.filter(
                      (s) => activeTab === 'pdf' || s.value !== 'FileDefault'
                    ).map((s) => (
                      <option key={s.value} value={s.value}>
                        {s.label}
                      </option>
                    ))}
                  </select>
                </div>

                <div className="space-y-1">
                  <label
                    htmlFor={`copies-${activeTab}`}
                    className="block text-xs font-medium text-slate-700"
                  >
                    Số bản
                  </label>
                  <input
                    id={`copies-${activeTab}`}
                    type="number"
                    min={1}
                    max={99}
                    value={profile.copies}
                    onChange={(e) =>
                      updateProfile(
                        'copies',
                        Math.max(1, Math.min(99, Number(e.target.value) || 1))
                      )
                    }
                    className={`${selectClass} font-mono tabular-nums`}
                  />
                </div>
              </div>

              {/* Column 2: Chiều, Tỷ lệ, Lề & Động cơ */}
              <div className="space-y-3">
                <div className="space-y-1">
                  <label
                    htmlFor={`orient-${activeTab}`}
                    className="block text-xs font-medium text-slate-700"
                  >
                    Chiều trang
                  </label>
                  <select
                    id={`orient-${activeTab}`}
                    value={profile.orientationMode}
                    onChange={(e) =>
                      updateProfile('orientationMode', e.target.value as OrientationMode)
                    }
                    className={selectClass}
                  >
                    <option value="AutoDetect">Tự nhận diện chiều</option>
                    <option value="ForcePortrait">Ép Dọc</option>
                    <option value="ForceLandscape">Ép Ngang</option>
                  </select>
                </div>

                {activeTab !== 'word' && activeTab !== 'excel' && (
                  <div className="space-y-1">
                    <div className="flex items-center gap-1">
                      <label
                        htmlFor={`scale-${activeTab}`}
                        className="block text-xs font-medium text-slate-700"
                      >
                        Tỷ lệ in
                      </label>
                      <HoverTip
                        align="scaleTip"
                        wide
                        text="Vừa trang: phóng to khi nhỏ hơn, thu nhỏ khi lớn hơn. Thu nếu tràn: giữ 100% khi nhỏ hơn, thu nhỏ khi lớn hơn. 100% gốc: giữ 100%."
                        content={
                          <table className="w-full border-collapse text-[10.5px] leading-snug border border-slate-600/80 rounded overflow-hidden">
                            <thead>
                              <tr className="bg-slate-900/90 text-slate-200">
                                <th className="text-left font-semibold px-1.5 py-1 border-b border-r border-slate-600/80">
                                  Chế độ
                                </th>
                                <th className="text-left font-semibold px-1.5 py-1 border-b border-r border-slate-600/80">
                                  Khi file nhỏ hơn giấy
                                </th>
                                <th className="text-left font-semibold px-1.5 py-1 border-b border-slate-600/80">
                                  Khi file lớn hơn giấy
                                </th>
                              </tr>
                            </thead>
                            <tbody className="text-slate-200">
                              <tr className="border-b border-slate-700/80">
                                <td className="px-1.5 py-1 font-semibold text-white border-r border-slate-700/80 whitespace-nowrap align-top">
                                  Vừa trang
                                </td>
                                <td className="px-1.5 py-1 border-r border-slate-700/80 align-top">
                                  <strong className="text-white">Phóng to</strong> đều cả 2 chiều đến khi chạm lề
                                </td>
                                <td className="px-1.5 py-1 align-top">
                                  <strong className="text-white">Thu nhỏ</strong> đều cả 2 chiều cho lọt vừa trang
                                </td>
                              </tr>
                              <tr className="border-b border-slate-700/80">
                                <td className="px-1.5 py-1 font-semibold text-white border-r border-slate-700/80 whitespace-nowrap align-top">
                                  Thu nếu tràn
                                </td>
                                <td className="px-1.5 py-1 border-r border-slate-700/80 align-top">
                                  <strong className="text-white">Giữ nguyên 100%</strong> (không phóng to)
                                </td>
                                <td className="px-1.5 py-1 align-top">
                                  <strong className="text-white">Thu nhỏ</strong> đều cả 2 chiều cho lọt vừa trang
                                </td>
                              </tr>
                              <tr>
                                <td className="px-1.5 py-1 font-semibold text-white border-r border-slate-700/80 whitespace-nowrap align-top">
                                  100% kích thước gốc
                                </td>
                                <td className="px-1.5 py-1 border-r border-slate-700/80 align-top">
                                  <strong className="text-white">Giữ nguyên 100%</strong>
                                </td>
                                <td className="px-1.5 py-1 align-top">
                                  <strong className="text-white">Giữ nguyên 100%</strong> (phần tràn bị cắt)
                                </td>
                              </tr>
                            </tbody>
                          </table>
                        }
                      />
                    </div>
                    <select
                      id={`scale-${activeTab}`}
                      value={profile.scaleMode}
                      onChange={(e) => updateProfile('scaleMode', e.target.value as ScaleMode)}
                      className={selectClass}
                    >
                      <option value="Fit">Vừa trang</option>
                      <option value="Shrink">Thu nếu tràn</option>
                      <option value="NoScale">100% kích thước gốc</option>
                    </select>
                  </div>
                )}

                <div className="space-y-1">
                  <label
                    htmlFor={`margin-${activeTab}`}
                    className="block text-xs font-medium text-slate-700"
                  >
                    Lề trang
                  </label>
                  <select
                    id={`margin-${activeTab}`}
                    value={profile.margins.mode}
                    onChange={(e) =>
                      updateMarginField('mode', e.target.value as MarginMode)
                    }
                    className={selectClass}
                  >
                    {MARGIN_OPTIONS.map((m) => (
                      <option key={m.value} value={m.value}>
                        {m.label}
                      </option>
                    ))}
                  </select>
                </div>

                <div className="space-y-1">
                  <label
                    htmlFor={`engine-${activeTab}`}
                    className="block text-xs font-medium text-slate-700"
                  >
                    Động cơ in
                  </label>
                  <select
                    id={`engine-${activeTab}`}
                    value={profile.engine}
                    disabled={engineOptions.length <= 1}
                    onChange={(e) => updateProfile('engine', e.target.value)}
                    className={selectClass}
                  >
                    {engineOptions.map((eng) => (
                      <option key={eng.value} value={eng.value}>
                        {eng.label}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-3">
              <div className="space-y-1">
                <label
                  htmlFor={`over-page-action-${activeTab}`}
                  className="block text-xs font-medium text-slate-700"
                >
                  Khi vượt quá số trang quy định
                </label>
                <select
                  id={`over-page-action-${activeTab}`}
                  value={profile.overPageAction}
                  onChange={(e) =>
                    updateProfile('overPageAction', e.target.value as OverPageAction)
                  }
                  className={selectClass}
                >
                  <option value="AlwaysPrint">Luôn in</option>
                  <option value="AskBeforePrint">Hỏi trước khi in</option>
                  <option value="Skip">Bỏ qua</option>
                </select>
              </div>

              <div className="space-y-1">
                <label
                  htmlFor={`max-pages-limit-${activeTab}`}
                  className="block text-xs font-medium text-slate-700"
                >
                  Số trang quy định
                </label>
                <input
                  id={`max-pages-limit-${activeTab}`}
                  type="number"
                  min={1}
                  max={9999}
                  disabled={profile.overPageAction === 'AlwaysPrint'}
                  value={profile.maxPagesLimit}
                  onChange={(e) =>
                    updateProfile(
                      'maxPagesLimit',
                      Math.max(1, Math.min(9999, Number(e.target.value) || 1))
                    )
                  }
                  className={`${selectClass} font-mono tabular-nums`}
                />
              </div>
            </div>

            {activeTab === 'word' && (
              <div className="p-3 bg-slate-50/70 border border-slate-200/80 rounded-lg">
                <div className="inline-flex items-center gap-1.5">
                  <label className="inline-flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer select-none">
                    <input
                      type="checkbox"
                      checked={Boolean(profile.wordSkipBlankPages)}
                      onChange={(e) =>
                        updateProfile('wordSkipBlankPages', e.target.checked)
                      }
                      className="w-3.5 h-3.5 rounded border-slate-300 text-blue-600 focus:ring-blue-600"
                    />
                    <span>Bỏ qua trang trắng</span>
                  </label>
                  <HoverTip
                    align="col1"
                    text="Tự động kiểm tra từng trang qua Microsoft Word COM và bỏ qua các trang không có chữ, bảng hay hình ảnh."
                  />
                </div>
              </div>
            )}

            {activeTab === 'excel' && (
              <div className="p-3 bg-slate-50/70 border border-slate-200/80 rounded-lg space-y-3">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-3">
                  <div className="space-y-1">
                    <label
                      htmlFor="excel-sheet-mode"
                      className="block text-xs font-medium text-slate-700"
                    >
                      Sheet cần in
                    </label>
                    <select
                      id="excel-sheet-mode"
                      value={profile.excelSheetMode || 'AllSheets'}
                      onChange={(e) =>
                        updateProfile('excelSheetMode', e.target.value as ExcelSheetMode)
                      }
                      className={selectClass}
                    >
                      <option value="AllSheets">In tất cả sheet</option>
                      <option value="ActiveSheet">Chỉ sheet đang chọn</option>
                      <option value="FirstSheet">Sheet đầu tiên</option>
                      <option value="AskSheets">Hỏi khi in (báo số trang từng sheet)</option>
                    </select>
                  </div>

                  <div className="space-y-1">
                    <label
                      htmlFor="excel-print-scope"
                      className="block text-xs font-medium text-slate-700"
                    >
                      Phạm vi in trên sheet
                    </label>
                    <select
                      id="excel-print-scope"
                      value={profile.excelPrintScope || 'EntireSheet'}
                      onChange={(e) =>
                        updateProfile('excelPrintScope', e.target.value as ExcelPrintScope)
                      }
                      className={selectClass}
                    >
                      <option value="EntireSheet">In cả sheet</option>
                      <option value="SelectedRange">Vùng được chọn</option>
                    </select>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-3">
                  <div className="space-y-1">
                    <label
                      htmlFor="excel-fit-mode"
                      className="block text-xs font-medium text-slate-700"
                    >
                      Co giãn trang (Fit)
                    </label>
                    <select
                      id="excel-fit-mode"
                      value={profile.excelFitMode || 'None'}
                      onChange={(e) =>
                        updateProfile('excelFitMode', e.target.value as ExcelFitMode)
                      }
                      className={selectClass}
                    >
                      <option value="None">Mặc định của file</option>
                      <option value="FitAllColumnsOnePage">Fit tất cả cột vào một trang</option>
                      <option value="FitAllRowsOnePage">Fit tất cả hàng vào một trang</option>
                      <option value="FitSheetOnePage">Cả sheet vào một trang</option>
                    </select>
                  </div>

                  <div className="space-y-1">
                    <label className="inline-flex items-center gap-1.5 text-xs font-medium text-slate-700 cursor-pointer select-none">
                      <input
                        type="checkbox"
                        checked={Boolean(profile.excelAutoLandscapeByCols)}
                        onChange={(e) =>
                          updateProfile('excelAutoLandscapeByCols', e.target.checked)
                        }
                        className="w-3.5 h-3.5 rounded border-slate-300 text-blue-600 focus:ring-blue-600"
                      />
                      <span>Ép ngang nếu nhiều cột</span>
                    </label>
                    <div className="flex items-center gap-2">
                      <span className="text-xs text-slate-500 whitespace-nowrap shrink-0">
                        Từ (cột):
                      </span>
                      <input
                        id="excel-col-threshold"
                        type="number"
                        min={1}
                        max={200}
                        disabled={!profile.excelAutoLandscapeByCols}
                        value={profile.excelLandscapeColThreshold ?? 7}
                        onChange={(e) =>
                          updateProfile(
                            'excelLandscapeColThreshold',
                            Math.max(1, Math.min(200, Number(e.target.value) || 1))
                          )
                        }
                        className={`${selectClass} font-mono tabular-nums`}
                      />
                    </div>
                  </div>
                </div>

                <div className="pt-2.5 border-t border-slate-200/70 flex flex-wrap items-center gap-x-5 gap-y-2">
                  <div className="inline-flex items-center gap-1.5">
                    <label className="inline-flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer select-none">
                      <input
                        type="checkbox"
                        checked={Boolean(profile.excelSkipBlankAreas)}
                        onChange={(e) =>
                          updateProfile('excelSkipBlankAreas', e.target.checked)
                        }
                        className="w-3.5 h-3.5 rounded border-slate-300 text-blue-600 focus:ring-blue-600"
                      />
                      <span>In bỏ qua vùng ô trống</span>
                    </label>
                    <HoverTip
                      align="col1"
                      text="Tự động thu gọn vùng in vừa khít phần ô có dữ liệu qua Excel COM, bỏ qua hàng/cột trống thừa."
                    />
                  </div>

                  <div className="inline-flex items-center gap-1.5">
                    <label className="inline-flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer select-none">
                      <input
                        type="checkbox"
                        checked={Boolean(profile.excelSkipBlankSheets)}
                        onChange={(e) =>
                          updateProfile('excelSkipBlankSheets', e.target.checked)
                        }
                        className="w-3.5 h-3.5 rounded border-slate-300 text-blue-600 focus:ring-blue-600"
                      />
                      <span>Bỏ qua sheet trống</span>
                    </label>
                    <HoverTip
                      align="col3"
                      text="Không in các sheet không có dữ liệu hoặc hình vẽ; bỏ qua cả file nếu mọi sheet đều trống."
                    />
                  </div>
                </div>
              </div>
            )}

            {profile.margins.mode === 'Custom' && (
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 p-3 bg-white/85 border border-slate-200/90 rounded-lg">
                {(
                  [
                    { key: 'topCm', label: 'Lề trên (cm)' },
                    { key: 'leftCm', label: 'Lề trái (cm)' },
                    { key: 'bottomCm', label: 'Lề dưới (cm)' },
                    { key: 'rightCm', label: 'Lề phải (cm)' },
                  ] as const
                ).map((edge) => (
                  <div key={edge.key} className="space-y-1">
                    <label
                      htmlFor={`margin-${activeTab}-${edge.key}`}
                      className="block text-[11px] font-medium text-slate-600"
                    >
                      {edge.label}
                    </label>
                    <input
                      id={`margin-${activeTab}-${edge.key}`}
                      type="number"
                      min={0}
                      max={8}
                      step={0.1}
                      value={profile.margins[edge.key]}
                      onChange={(e) => {
                        const val = parseFloat(e.target.value);
                        updateMarginField(
                          edge.key,
                          Math.max(0, Math.min(8, Number.isFinite(val) ? val : 0))
                        );
                      }}
                      className={`${selectClass} font-mono tabular-nums`}
                    />
                  </div>
                ))}
              </div>
            )}

            {profile.margins.mode === 'CustomByOrientation' && (
              <div className="space-y-2.5 p-3 bg-white/85 border border-slate-200/90 rounded-lg">
                {(
                  [
                    { orient: 'portrait' as const, title: 'Trang Dọc (Portrait)' },
                    { orient: 'landscape' as const, title: 'Trang Ngang (Landscape)' },
                  ]
                ).map((group) => (
                  <div key={group.orient} className="space-y-1.5">
                    <span className="block text-[11px] font-semibold text-slate-700">
                      {group.title}
                    </span>
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                      {(
                        [
                          { key: 'topCm', label: 'Trên (cm)' },
                          { key: 'leftCm', label: 'Trái (cm)' },
                          { key: 'bottomCm', label: 'Dưới (cm)' },
                          { key: 'rightCm', label: 'Phải (cm)' },
                        ] as const
                      ).map((edge) => (
                        <div key={edge.key} className="space-y-1">
                          <label
                            htmlFor={`margin-${activeTab}-${group.orient}-${edge.key}`}
                            className="block text-[11px] text-slate-600"
                          >
                            {edge.label}
                          </label>
                          <input
                            id={`margin-${activeTab}-${group.orient}-${edge.key}`}
                            type="number"
                            min={0}
                            max={8}
                            step={0.1}
                            value={profile.margins[group.orient][edge.key]}
                            onChange={(e) => {
                              const val = parseFloat(e.target.value);
                              updateOrientationMarginEdge(
                                group.orient,
                                edge.key,
                                Math.max(0, Math.min(8, Number.isFinite(val) ? val : 0))
                              );
                            }}
                            className={`${selectClass} font-mono tabular-nums`}
                          />
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            )}

            {activeTab === 'pdf' &&
              (profile.engine === 'SmartAuto' || profile.engine === 'SumatraPDF') && (
                <div className="space-y-1">
                  <label
                    htmlFor="sumatra-path-input"
                    className="block text-xs font-medium text-slate-700"
                  >
                    Đường dẫn SumatraPDF.exe (tùy chọn)
                  </label>
                  <input
                    id="sumatra-path-input"
                    type="text"
                    value={profile.sumatraCustomPath || ''}
                    onChange={(e) => updateProfile('sumatraCustomPath', e.target.value)}
                    placeholder="Để trống để tự tìm trên máy"
                    className={`${selectClass} font-mono`}
                  />
                </div>
              )}
          </div>
        </div>
      </section>

      {resetModalOpen && (
        <div
          role="dialog"
          aria-modal="true"
          aria-labelledby="reset-modal-title"
          onClick={() => setResetModalOpen(false)}
          className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-[1px]"
        >
          <div
            onClick={(e) => e.stopPropagation()}
            className="w-full max-w-sm bg-white border border-slate-200 rounded-xl shadow-2xl overflow-hidden"
          >
            <div className="p-4 space-y-1.5">
              <h3
                id="reset-modal-title"
                className="text-sm font-semibold text-slate-900"
              >
                Khôi phục cài đặt mặc định?
              </h3>
              <p className="text-xs leading-relaxed text-slate-600">
                Toàn bộ thông số ở mục <strong>Cài đặt chung</strong> và tất cả loại tài liệu
                sẽ được đặt lại về trạng thái ban đầu.
              </p>
            </div>

            <div className="px-4 py-3 bg-slate-50 border-t border-slate-100 flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setResetModalOpen(false)}
                className="inline-flex items-center h-8 px-3 text-xs font-medium text-slate-700 bg-white border border-slate-200 rounded-lg hover:bg-slate-100 transition-colors cursor-pointer"
              >
                Hủy
              </button>
              <button
                type="button"
                onClick={handleConfirmResetDefaults}
                className="inline-flex items-center gap-1.5 h-8 px-3.5 text-xs font-semibold text-white bg-rose-600 rounded-lg hover:bg-rose-700 transition-colors cursor-pointer"
              >
                <RotateCcw className="w-3.5 h-3.5" aria-hidden="true" />
                <span>Khôi phục mặc định</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
