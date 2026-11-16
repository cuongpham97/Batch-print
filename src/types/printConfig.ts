export type PaperSize = 'A4' | 'A3' | 'A5' | 'Letter' | 'Legal' | 'B5';

export type DefaultSides = 'OneSided' | 'TwoSided';

export type DuplexMode = 'OneSided' | 'TwoSidedLongEdge' | 'TwoSidedShortEdge' | 'FileDefault';

export type OrientationMode = 'AutoDetect' | 'ForcePortrait' | 'ForceLandscape';

export type DuplexOrientationStrategy =
  | 'FixedLongEdge'
  | 'FixedShortEdge'
  | 'PortraitLongLandscapeShort'
  | 'PortraitShortLandscapeLong'
  | 'FileDefault';

export type ScaleMode = 'Fit' | 'NoScale' | 'Shrink';

export type MarginMode =
  | 'Default'
  | 'Zero'
  | 'Narrow'
  | 'Normal'
  | 'Custom'
  | 'CustomByOrientation';

export interface MarginEdgesCm {
  topCm: number;
  bottomCm: number;
  leftCm: number;
  rightCm: number;
}

export interface PageMargins extends MarginEdgesCm {
  mode: MarginMode;
  portrait: MarginEdgesCm;
  landscape: MarginEdgesCm;
}

export type DocTypeId = 'word' | 'excel' | 'images' | 'pdf' | 'text' | 'other';

export type PdfEngine = 'SmartAuto' | 'WindowsNative' | 'SumatraPDF' | 'AdobeFoxit';
export type OfficeEngine = 'OfficeCom' | 'ShellFallback';
export type ImageEngine = 'DotNetPrint' | 'ChromeWebp';
export type TextEngine = 'DotNetText';
export type OtherEngine = 'ShellOrChrome';

export type FolderSourceMode = 'ScriptDir' | 'CustomPath';
export type PostPrintAction = 'Keep' | 'MoveToPrinted' | 'Delete';
export type SortOrder = 'FolderOrder' | 'NameAsc' | 'NameDesc' | 'DateNewest' | 'DateOldest';
export type OverPageAction = 'AlwaysPrint' | 'AskBeforePrint' | 'Skip';

export type ExcelSheetMode = 'AllSheets' | 'ActiveSheet' | 'FirstSheet' | 'AskSheets';
export type ExcelPrintScope = 'EntireSheet' | 'SelectedRange';
export type ExcelFitMode = 'None' | 'FitAllColumnsOnePage' | 'FitAllRowsOnePage' | 'FitSheetOnePage';

export interface DocTypeProfile {
  enabled: boolean;
  defaultPaperSize: PaperSize;
  defaultSides: DefaultSides;
  orientationMode: OrientationMode;
  duplexStrategy: DuplexOrientationStrategy;
  scaleMode: ScaleMode;
  margins: PageMargins;
  copies: number;
  engine: string;
  overPageAction: OverPageAction;
  maxPagesLimit: number;
  wordSkipBlankPages?: boolean;
  excelSheetMode?: ExcelSheetMode;
  excelPrintScope?: ExcelPrintScope;
  excelFitMode?: ExcelFitMode;
  excelAutoLandscapeByCols?: boolean;
  excelLandscapeColThreshold?: number;
  excelSkipBlankAreas?: boolean;
  excelSkipBlankSheets?: boolean;
  sumatraCustomPath?: string;
  customExtensions?: string;
}

export type PrinterMode = 'WindowsDefault' | 'SpecificPrinter' | 'AskOnRun';

export interface PrintConfig {
  scriptFileName: string;
  folderSourceMode: FolderSourceMode;
  customFolderPath: string;
  includeSubfolders: boolean;
  sortOrder: SortOrder;

  printerMode: PrinterMode;
  printerName: string;
  setAsWindowsDefault: boolean;
  restorePrinterConfig: boolean;

  jobDelaySeconds: number;
  postPrintAction: PostPrintAction;
  autoCloseConsole: boolean;
  exportLogTxt: boolean;

  docProfiles: Record<DocTypeId, DocTypeProfile>;
}

export interface DocTypeMeta {
  id: DocTypeId;
  label: string;
  extensionsLabel: string;
}

export const DOC_TYPE_TABS: DocTypeMeta[] = [
  { id: 'word', label: 'Word', extensionsLabel: '.docx, .doc, .rtf' },
  { id: 'excel', label: 'Excel', extensionsLabel: '.xlsx, .xls, .xlsm' },
  { id: 'images', label: 'Ảnh', extensionsLabel: '.jpg, .png, .tif, .webp' },
  { id: 'pdf', label: 'PDF', extensionsLabel: '.pdf' },
  { id: 'text', label: 'Văn bản', extensionsLabel: '.txt, .log, .csv' },
  { id: 'other', label: 'Khác', extensionsLabel: 'Tùy chỉnh' },
];

export interface PaperSpec {
  id: PaperSize;
  label: string;
  widthMm: number;
  heightMm: number;
  dimensionsLabel: string;
}

export const PAPER_SPECS: Record<PaperSize, PaperSpec> = {
  A4: { id: 'A4', label: 'A4', widthMm: 210, heightMm: 297, dimensionsLabel: '21 × 29.7 cm' },
  A3: { id: 'A3', label: 'A3', widthMm: 297, heightMm: 420, dimensionsLabel: '29.7 × 42 cm' },
  A5: { id: 'A5', label: 'A5', widthMm: 148, heightMm: 210, dimensionsLabel: '14.8 × 21 cm' },
  Letter: { id: 'Letter', label: 'Letter', widthMm: 215.9, heightMm: 279.4, dimensionsLabel: '21.6 × 27.9 cm' },
  Legal: { id: 'Legal', label: 'Legal', widthMm: 215.9, heightMm: 355.6, dimensionsLabel: '21.6 × 35.6 cm' },
  B5: { id: 'B5', label: 'B5', widthMm: 176, heightMm: 250, dimensionsLabel: '17.6 × 25 cm' },
};

export const DUPLEX_LABELS: Record<DuplexMode, string> = {
  OneSided: '1 mặt',
  TwoSidedLongEdge: '2 mặt · Lật cạnh dài',
  TwoSidedShortEdge: '2 mặt · Lật cạnh ngắn',
  FileDefault: '2 mặt · Mặc định của file',
};

function createBaseProfile(overrides: Partial<DocTypeProfile>): DocTypeProfile {
  return {
    enabled: true,
    defaultPaperSize: 'A4',
    defaultSides: 'TwoSided',
    orientationMode: 'AutoDetect',
    duplexStrategy: 'PortraitLongLandscapeShort',
    scaleMode: 'Shrink',
    margins: {
      mode: 'Default',
      topCm: 2,
      leftCm: 3,
      bottomCm: 2,
      rightCm: 2,
      portrait: { topCm: 2, leftCm: 3, bottomCm: 2, rightCm: 2 },
      landscape: { topCm: 2, leftCm: 3, bottomCm: 2, rightCm: 2 },
    },
    copies: 1,
    engine: 'WindowsNative',
    overPageAction: 'AskBeforePrint',
    maxPagesLimit: 15,
    ...overrides,
  };
}

export function resolveEffectiveMarginsCm(
  margins?: PageMargins,
  orientation: 'Portrait' | 'Landscape' = 'Portrait'
): { top: number; bottom: number; left: number; right: number } | null {
  if (!margins || margins.mode === 'Default') return null;
  if (margins.mode === 'Zero') return { top: 0, bottom: 0, left: 0, right: 0 };
  if (margins.mode === 'Narrow') return { top: 0.5, bottom: 0.5, left: 0.5, right: 0.5 };
  if (margins.mode === 'Normal') return { top: 1.5, bottom: 1.5, left: 1.5, right: 1.5 };
  const clamp = (v: number) => Math.max(0, Math.min(8, Number.isFinite(v) ? v : 0));
  const source: MarginEdgesCm =
    margins.mode === 'CustomByOrientation'
      ? orientation === 'Landscape'
        ? margins.landscape
        : margins.portrait
      : margins;
  return {
    top: clamp(source.topCm),
    bottom: clamp(source.bottomCm),
    left: clamp(source.leftCm),
    right: clamp(source.rightCm),
  };
}

export function resolveEffectiveMarginsMm(
  margins?: PageMargins,
  orientation: 'Portrait' | 'Landscape' = 'Portrait'
): { top: number; bottom: number; left: number; right: number } | null {
  const cm = resolveEffectiveMarginsCm(margins, orientation);
  if (!cm) return null;
  return {
    top: Math.round(cm.top * 100) / 10,
    bottom: Math.round(cm.bottom * 100) / 10,
    left: Math.round(cm.left * 100) / 10,
    right: Math.round(cm.right * 100) / 10,
  };
}

export const DEFAULT_PRINT_CONFIG: PrintConfig = {
  scriptFileName: 'InHangLoat.bat',
  folderSourceMode: 'ScriptDir',
  customFolderPath: 'D:\\TaiLieuCanIn',
  includeSubfolders: true,
  sortOrder: 'FolderOrder',

  printerMode: 'WindowsDefault',
  printerName: 'Canon LBP2900',
  setAsWindowsDefault: false,
  restorePrinterConfig: true,

  jobDelaySeconds: 1,
  postPrintAction: 'MoveToPrinted',
  autoCloseConsole: false,
  exportLogTxt: false,

  docProfiles: {
    word: createBaseProfile({ enabled: true, engine: 'OfficeCom', wordSkipBlankPages: true }),
    excel: createBaseProfile({
      enabled: true,
      engine: 'OfficeCom',
      excelSheetMode: 'AllSheets',
      excelPrintScope: 'EntireSheet',
      excelFitMode: 'FitAllColumnsOnePage',
      excelAutoLandscapeByCols: true,
      excelLandscapeColThreshold: 10,
      excelSkipBlankAreas: true,
      excelSkipBlankSheets: true,
    }),
    images: createBaseProfile({ enabled: true, engine: 'DotNetPrint' }),
    pdf: createBaseProfile({ enabled: true, engine: 'WindowsNative', sumatraCustomPath: '' }),
    text: createBaseProfile({ enabled: true, engine: 'DotNetText' }),
    other: createBaseProfile({
      enabled: true,
      engine: 'ShellOrChrome',
      customExtensions: '.html, .htm',
    }),
  },
};

export function cloneDefaultPrintConfig(): PrintConfig {
  return JSON.parse(JSON.stringify(DEFAULT_PRINT_CONFIG)) as PrintConfig;
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function mergeEdges(base: MarginEdgesCm, raw: unknown): MarginEdgesCm {
  if (!isPlainObject(raw)) return { ...base };
  return {
    topCm: typeof raw.topCm === 'number' && Number.isFinite(raw.topCm) ? raw.topCm : base.topCm,
    bottomCm:
      typeof raw.bottomCm === 'number' && Number.isFinite(raw.bottomCm)
        ? raw.bottomCm
        : base.bottomCm,
    leftCm:
      typeof raw.leftCm === 'number' && Number.isFinite(raw.leftCm) ? raw.leftCm : base.leftCm,
    rightCm:
      typeof raw.rightCm === 'number' && Number.isFinite(raw.rightCm) ? raw.rightCm : base.rightCm,
  };
}

export function parseImportedPrintConfig(raw: unknown): PrintConfig | null {
  if (!isPlainObject(raw)) return null;

  const hasKnownKey =
    'docProfiles' in raw ||
    'printerMode' in raw ||
    'folderSourceMode' in raw ||
    'scriptFileName' in raw;
  if (!hasKnownKey) return null;

  const base = cloneDefaultPrintConfig();
  const mergedProfiles = { ...base.docProfiles };

  if (isPlainObject(raw.docProfiles)) {
    for (const tab of DOC_TYPE_TABS) {
      const defProf = base.docProfiles[tab.id];
      const rawProf = raw.docProfiles[tab.id];
      if (!isPlainObject(rawProf)) continue;

      const rawMargins = isPlainObject(rawProf.margins) ? rawProf.margins : undefined;
      mergedProfiles[tab.id] = {
        ...defProf,
        ...(rawProf as Partial<DocTypeProfile>),
        margins: {
          ...defProf.margins,
          ...(rawMargins ? (rawMargins as Partial<PageMargins>) : {}),
          portrait: mergeEdges(defProf.margins.portrait, rawMargins?.portrait),
          landscape: mergeEdges(defProf.margins.landscape, rawMargins?.landscape),
        },
      };
    }
  }

  return {
    scriptFileName:
      typeof raw.scriptFileName === 'string' && raw.scriptFileName.trim()
        ? raw.scriptFileName
        : base.scriptFileName,
    folderSourceMode:
      raw.folderSourceMode === 'ScriptDir' || raw.folderSourceMode === 'CustomPath'
        ? raw.folderSourceMode
        : base.folderSourceMode,
    customFolderPath:
      typeof raw.customFolderPath === 'string' ? raw.customFolderPath : base.customFolderPath,
    includeSubfolders:
      typeof raw.includeSubfolders === 'boolean'
        ? raw.includeSubfolders
        : base.includeSubfolders,
    sortOrder:
      typeof raw.sortOrder === 'string'
        ? (raw.sortOrder as SortOrder)
        : base.sortOrder,
    printerMode:
      raw.printerMode === 'WindowsDefault' ||
      raw.printerMode === 'SpecificPrinter' ||
      raw.printerMode === 'AskOnRun'
        ? raw.printerMode
        : base.printerMode,
    printerName: typeof raw.printerName === 'string' ? raw.printerName : base.printerName,
    setAsWindowsDefault:
      typeof raw.setAsWindowsDefault === 'boolean'
        ? raw.setAsWindowsDefault
        : base.setAsWindowsDefault,
    restorePrinterConfig:
      typeof raw.restorePrinterConfig === 'boolean'
        ? raw.restorePrinterConfig
        : base.restorePrinterConfig,
    jobDelaySeconds:
      typeof raw.jobDelaySeconds === 'number' && Number.isFinite(raw.jobDelaySeconds)
        ? raw.jobDelaySeconds
        : base.jobDelaySeconds,
    postPrintAction:
      raw.postPrintAction === 'Keep' ||
      raw.postPrintAction === 'MoveToPrinted' ||
      raw.postPrintAction === 'Delete'
        ? raw.postPrintAction
        : base.postPrintAction,
    autoCloseConsole:
      typeof raw.autoCloseConsole === 'boolean' ? raw.autoCloseConsole : base.autoCloseConsole,
    exportLogTxt: typeof raw.exportLogTxt === 'boolean' ? raw.exportLogTxt : base.exportLogTxt,
    docProfiles: mergedProfiles,
  };
}
