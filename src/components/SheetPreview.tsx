import {
  DOC_TYPE_TABS,
  DUPLEX_LABELS,
  DocTypeId,
  DocTypeProfile,
  PAPER_SPECS,
  PrintConfig,
  resolveEffectiveMarginsCm,
  resolveEffectiveMarginsMm,
} from '../types/printConfig';
import { resolveEffectiveRule } from '../utils/batGenerator';

interface SheetPreviewProps {
  config: PrintConfig;
  activeTab: DocTypeId;
}

interface SheetCardProps {
  sampleLabel: string;
  sampleOrientation: 'Portrait' | 'Landscape';
  profile: DocTypeProfile;
  patternId: string;
}

function SheetCard({
  sampleLabel,
  sampleOrientation,
  profile,
  patternId,
}: SheetCardProps) {
  const decision = resolveEffectiveRule(profile, sampleOrientation);
  const spec = PAPER_SPECS[decision.effectivePaperSize];
  const isLandscape = decision.effectiveOrientation === 'Landscape';

  const rawW = isLandscape ? spec.heightMm : spec.widthMm;
  const rawH = isLandscape ? spec.widthMm : spec.heightMm;
  const scale = Math.min(120 / rawW, 80 / rawH);
  const sheetW = Math.round(rawW * scale);
  const sheetH = Math.round(rawH * scale);
  const sheetX = Math.round((200 - sheetW) / 2);
  const sheetY = Math.round((132 - sheetH) / 2);

  const isTwoSided = decision.effectiveDuplex !== 'OneSided';
  const hasExplicitFlipEdge =
    decision.effectiveDuplex === 'TwoSidedLongEdge' ||
    decision.effectiveDuplex === 'TwoSidedShortEdge';
  const flipOnLongEdge = decision.effectiveDuplex === 'TwoSidedLongEdge';
  const flipAxisIsVertical =
    hasExplicitFlipEdge && ((flipOnLongEdge && !isLandscape) || (!flipOnLongEdge && isLandscape));

  const effMargins = resolveEffectiveMarginsMm(
    profile.margins,
    decision.effectiveOrientation
  );
  const effMarginsCm = resolveEffectiveMarginsCm(
    profile.margins,
    decision.effectiveOrientation
  );

  const mPadTop = effMargins ? Math.min(sheetH * 0.35, Math.round(effMargins.top * scale)) : 0;
  const mPadBottom = effMargins
    ? Math.min(sheetH * 0.35, Math.round(effMargins.bottom * scale))
    : 0;
  const mPadLeft = effMargins ? Math.min(sheetW * 0.35, Math.round(effMargins.left * scale)) : 0;
  const mPadRight = effMargins ? Math.min(sheetW * 0.35, Math.round(effMargins.right * scale)) : 0;

  const isAutoDuplex = decision.effectiveDuplex === 'FileDefault';
  const shortDuplexLabel =
    decision.effectiveDuplex === 'OneSided'
      ? '1 mặt'
      : decision.effectiveDuplex === 'TwoSidedLongEdge'
        ? 'Lật cạnh dài'
        : decision.effectiveDuplex === 'TwoSidedShortEdge'
          ? 'Lật cạnh ngắn'
          : 'auto';

  return (
    <div className="bg-slate-50 border border-slate-200 rounded-lg p-2.5 space-y-2">
      <div className="flex items-center justify-between gap-1 text-[11px]">
        <span className="font-semibold text-slate-800">{sampleLabel}</span>
        <span
          className={
            isAutoDuplex
              ? 'font-semibold text-emerald-600'
              : 'font-medium text-blue-700'
          }
        >
          {shortDuplexLabel}
        </span>
      </div>

      <svg
        viewBox="0 0 200 132"
        className="w-full h-auto"
        role="img"
        aria-label={`${sampleLabel} ${spec.id} ${isLandscape ? 'Ngang' : 'Dọc'} ${DUPLEX_LABELS[decision.effectiveDuplex]}`}
      >
        <defs>
          <pattern
            id={patternId}
            width="12"
            height="12"
            patternUnits="userSpaceOnUse"
          >
            <path
              d="M 12 0 L 0 0 0 12"
              fill="none"
              stroke="#e2e8f0"
              strokeWidth="0.7"
            />
          </pattern>
        </defs>
        <rect width="200" height="132" fill={`url(#${patternId})`} rx="4" />

        {isTwoSided && (
          <rect
            x={sheetX + 4}
            y={sheetY + 4}
            width={sheetW}
            height={sheetH}
            rx="2"
            fill="#e2e8f0"
            stroke="#cbd5e1"
            strokeWidth="1"
          />
        )}

        <rect
          x={sheetX}
          y={sheetY}
          width={sheetW}
          height={sheetH}
          rx="2"
          fill={effMargins ? '#f1f5f9' : '#ffffff'}
          stroke="#475569"
          strokeWidth="1.2"
        />

        {effMargins && (
          <rect
            x={sheetX + mPadLeft}
            y={sheetY + mPadTop}
            width={Math.max(8, sheetW - mPadLeft - mPadRight)}
            height={Math.max(8, sheetH - mPadTop - mPadBottom)}
            fill="#ffffff"
            stroke="#94a3b8"
            strokeWidth="0.7"
            strokeDasharray="2 1.5"
          />
        )}

        <line
          x1={sheetX + Math.max(6, mPadLeft + 4)}
          y1={sheetY + Math.max(8, mPadTop + 6)}
          x2={sheetX + Math.min(sheetW - mPadRight - 6, Math.max(14, mPadLeft + 28))}
          y2={sheetY + Math.max(8, mPadTop + 6)}
          stroke="#2563eb"
          strokeWidth="2"
          strokeLinecap="round"
        />
        <line
          x1={sheetX + Math.max(6, mPadLeft + 4)}
          y1={sheetY + Math.max(15, mPadTop + 13)}
          x2={sheetX + Math.max(16, sheetW - mPadRight - 5)}
          y2={sheetY + Math.max(15, mPadTop + 13)}
          stroke="#cbd5e1"
          strokeWidth="1.2"
          strokeLinecap="round"
        />
        <line
          x1={sheetX + Math.max(6, mPadLeft + 4)}
          y1={sheetY + Math.max(21, mPadTop + 19)}
          x2={sheetX + Math.max(12, sheetW - mPadRight - 10)}
          y2={sheetY + Math.max(21, mPadTop + 19)}
          stroke="#cbd5e1"
          strokeWidth="1.2"
          strokeLinecap="round"
        />

        {hasExplicitFlipEdge && flipAxisIsVertical && (
          <g>
            <line
              x1={sheetX}
              y1={sheetY - 4}
              x2={sheetX}
              y2={sheetY + sheetH + 4}
              stroke="#2563eb"
              strokeWidth="2"
              strokeDasharray="3 2"
            />
            <circle cx={sheetX} cy={sheetY + sheetH / 2} r="3" fill="#2563eb" />
          </g>
        )}

        {hasExplicitFlipEdge && !flipAxisIsVertical && (
          <g>
            <line
              x1={sheetX - 4}
              y1={sheetY}
              x2={sheetX + sheetW + 4}
              y2={sheetY}
              stroke="#2563eb"
              strokeWidth="2"
              strokeDasharray="3 2"
            />
            <circle cx={sheetX + sheetW / 2} cy={sheetY} r="3" fill="#2563eb" />
          </g>
        )}

        {effMarginsCm && (
          <g
            fill="#475569"
            fontSize="7.5"
            fontWeight="600"
            fontFamily="ui-monospace, SFMono-Regular, Menlo, monospace"
          >
            {/* Lề trên */}
            <text x={sheetX + sheetW / 2} y={sheetY - 8} textAnchor="middle">
              {effMarginsCm.top}
            </text>
            {/* Lề dưới */}
            <text x={sheetX + sheetW / 2} y={sheetY + sheetH + 13} textAnchor="middle">
              {effMarginsCm.bottom}
            </text>
            {/* Lề trái */}
            <text
              x={sheetX - 7}
              y={sheetY + sheetH / 2 + 2.5}
              textAnchor="end"
            >
              {effMarginsCm.left}
            </text>
            {/* Lề phải */}
            <text
              x={sheetX + sheetW + 7}
              y={sheetY + sheetH / 2 + 2.5}
              textAnchor="start"
            >
              {effMarginsCm.right}
            </text>
          </g>
        )}
      </svg>

      <div className="text-center text-[10.5px] font-mono tabular-nums text-slate-600">
        {spec.id} (
        {isLandscape
          ? `${Math.round(spec.heightMm) / 10}×${Math.round(spec.widthMm) / 10}`
          : `${Math.round(spec.widthMm) / 10}×${Math.round(spec.heightMm) / 10}`}
        )
      </div>
    </div>
  );
}

export function SheetPreview({ config, activeTab }: SheetPreviewProps) {
  const tabMeta = DOC_TYPE_TABS.find((t) => t.id === activeTab) ?? DOC_TYPE_TABS[0];
  const profile = config.docProfiles[activeTab];

  const isAutoPrinter = config.printerMode === 'WindowsDefault';
  const printerDisplay = isAutoPrinter
    ? 'auto'
    : config.printerMode === 'AskOnRun'
      ? 'Hỏi khi chạy'
      : config.printerName.trim() || '(Chưa nhập)';

  const scaleLabel =
    profile.scaleMode === 'Fit'
      ? 'Vừa trang'
      : profile.scaleMode === 'Shrink'
        ? 'Thu nếu tràn'
        : '100%';

  const portCm = resolveEffectiveMarginsCm(profile.margins, 'Portrait');
  const landCm = resolveEffectiveMarginsCm(profile.margins, 'Landscape');

  const isAutoMargin = profile.margins.mode === 'Default';
  const marginOverview = isAutoMargin
    ? 'auto'
    : profile.margins.mode === 'CustomByOrientation'
      ? `Dọc ${portCm?.top}/${portCm?.left}/${portCm?.bottom}/${portCm?.right} · Ngang ${landCm?.top}/${landCm?.left}/${landCm?.bottom}/${landCm?.right} cm`
      : `${portCm?.top}/${portCm?.left}/${portCm?.bottom}/${portCm?.right} cm`;

  const excelSheetLabel =
    profile.excelSheetMode === 'ActiveSheet'
      ? 'Sheet đang chọn'
      : profile.excelSheetMode === 'FirstSheet'
        ? 'Sheet đầu tiên'
        : profile.excelSheetMode === 'AskSheets'
          ? 'Hỏi khi in'
          : 'Tất cả sheet';

  const excelScopeLabel =
    profile.excelPrintScope === 'SelectedRange' ? 'Vùng được chọn' : 'Cả sheet';

  const isAutoExcelFit = !profile.excelFitMode || profile.excelFitMode === 'None';
  const excelFitLabel =
    profile.excelFitMode === 'FitAllColumnsOnePage'
      ? 'Fit tất cả cột vào 1 trang'
      : profile.excelFitMode === 'FitAllRowsOnePage'
        ? 'Fit tất cả hàng vào 1 trang'
        : profile.excelFitMode === 'FitSheetOnePage'
          ? 'Cả sheet vào 1 trang'
          : 'auto';

  return (
    <section
      id="section-preview"
      className="bg-white border border-slate-200 rounded-xl overflow-hidden scroll-mt-16"
    >
      <div className="px-4 py-2.5 border-b border-emerald-800 bg-emerald-700 flex items-center gap-2">
        <span className="w-2 h-2 rounded-full bg-emerald-200 shrink-0" aria-hidden="true" />
        <h2 className="text-xs font-semibold uppercase tracking-wide text-white">
          Mô phỏng ({tabMeta.label})
        </h2>
      </div>

      <div className={`p-4 space-y-3 ${profile.enabled ? '' : 'opacity-50'}`}>
        <div className="grid grid-cols-2 gap-3">
          <SheetCard
            sampleLabel="Trang Dọc"
            sampleOrientation="Portrait"
            profile={profile}
            patternId="sheet-grid-portrait"
          />
          <SheetCard
            sampleLabel="Trang Ngang"
            sampleOrientation="Landscape"
            profile={profile}
            patternId="sheet-grid-landscape"
          />
        </div>

        <dl className="grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-2 pt-2 border-t border-slate-100 text-xs">
          <div className="flex items-baseline gap-1.5 pb-1.5 border-b border-slate-100">
            <dt className="text-slate-500 shrink-0">Trạng thái:</dt>
            <dd className="font-medium text-slate-900">
              {profile.enabled ? 'Đang bật' : 'Đang tắt'}
            </dd>
          </div>
          <div className="flex items-baseline gap-1.5 pb-1.5 border-b border-slate-100 min-w-0">
            <dt className="text-slate-500 shrink-0">Máy in:</dt>
            <dd
              className={`truncate ${
                isAutoPrinter
                  ? 'font-semibold text-emerald-600'
                  : 'font-medium text-slate-900'
              }`}
            >
              {printerDisplay}
            </dd>
          </div>
          {activeTab !== 'word' && activeTab !== 'excel' && (
            <div className="flex items-baseline gap-1.5 pb-1.5 border-b border-slate-100">
              <dt className="text-slate-500 shrink-0">Tỷ lệ:</dt>
              <dd className="font-medium text-slate-900">{scaleLabel}</dd>
            </div>
          )}
          <div
            className={`flex items-baseline gap-1.5 pb-1.5 border-b border-slate-100 ${
              activeTab === 'word' || activeTab === 'excel' ? 'sm:col-span-2' : ''
            }`}
          >
            <dt className="text-slate-500 shrink-0">Số bản:</dt>
            <dd className="font-medium text-slate-900 tabular-nums">
              {profile.copies} bản
            </dd>
          </div>
          <div className="sm:col-span-2 flex items-baseline gap-1.5 pb-1.5 border-b border-slate-100">
            <dt className="text-slate-500 shrink-0">Vượt số trang:</dt>
            <dd className="font-medium text-slate-900 tabular-nums">
              {profile.overPageAction === 'AlwaysPrint'
                ? 'Luôn in'
                : profile.overPageAction === 'AskBeforePrint'
                  ? `Hỏi trước nếu > ${profile.maxPagesLimit} trang`
                  : `Bỏ qua nếu > ${profile.maxPagesLimit} trang`}
            </dd>
          </div>
          <div
            className={`sm:col-span-2 flex items-baseline gap-1.5 ${
              activeTab === 'excel' || activeTab === 'word'
                ? 'pb-1.5 border-b border-slate-100'
                : ''
            }`}
          >
            <dt className="text-slate-500 shrink-0">Lề trang:</dt>
            <dd
              className={`tabular-nums ${
                isAutoMargin
                  ? 'font-semibold text-emerald-600'
                  : 'font-medium text-slate-900'
              }`}
            >
              {profile.margins.mode === 'CustomByOrientation' ? (
                <div className="flex flex-wrap gap-x-3 gap-y-0.5">
                  <span>
                    Dọc: {portCm?.top}/{portCm?.left}/{portCm?.bottom}/{portCm?.right} cm
                  </span>
                  <span>
                    Ngang: {landCm?.top}/{landCm?.left}/{landCm?.bottom}/{landCm?.right} cm
                  </span>
                </div>
              ) : (
                marginOverview
              )}
            </dd>
          </div>
          {activeTab === 'word' && (
            <div className="sm:col-span-2 flex items-baseline gap-1.5">
              <dt className="text-slate-500 shrink-0">Trang trắng:</dt>
              <dd className="font-medium text-slate-900">
                {profile.wordSkipBlankPages ? 'Bỏ qua trang trắng' : 'Giữ nguyên'}
              </dd>
            </div>
          )}
          {activeTab === 'excel' && (
            <>
              <div className="flex items-baseline gap-1.5 pb-1.5 border-b border-slate-100">
                <dt className="text-slate-500 shrink-0">Sheet cần in:</dt>
                <dd className="font-medium text-slate-900">{excelSheetLabel}</dd>
              </div>
              <div className="flex items-baseline gap-1.5 pb-1.5 border-b border-slate-100">
                <dt className="text-slate-500 shrink-0">Phạm vi sheet:</dt>
                <dd className="font-medium text-slate-900">{excelScopeLabel}</dd>
              </div>
              <div className="sm:col-span-2 flex items-baseline gap-1.5 pb-1.5 border-b border-slate-100">
                <dt className="text-slate-500 shrink-0">Co giãn (Fit):</dt>
                <dd
                  className={
                    isAutoExcelFit
                      ? 'font-semibold text-emerald-600'
                      : 'font-medium text-slate-900'
                  }
                >
                  {excelFitLabel}
                </dd>
              </div>
              <div className="sm:col-span-2 flex items-baseline gap-1.5 pb-1.5 border-b border-slate-100">
                <dt className="text-slate-500 shrink-0">Ép ngang nếu nhiều cột:</dt>
                <dd className="font-medium text-slate-900 tabular-nums">
                  {profile.excelAutoLandscapeByCols
                    ? `Từ ≥ ${profile.excelLandscapeColThreshold ?? 7} cột`
                    : 'Tắt'}
                </dd>
              </div>
              <div className="flex items-baseline gap-1.5">
                <dt className="text-slate-500 shrink-0">Vùng ô trống:</dt>
                <dd className="font-medium text-slate-900">
                  {profile.excelSkipBlankAreas ? 'Bỏ qua' : 'Giữ nguyên'}
                </dd>
              </div>
              <div className="flex items-baseline gap-1.5">
                <dt className="text-slate-500 shrink-0">Sheet trống:</dt>
                <dd className="font-medium text-slate-900">
                  {profile.excelSkipBlankSheets ? 'Bỏ qua' : 'Giữ nguyên'}
                </dd>
              </div>
            </>
          )}
        </dl>
      </div>
    </section>
  );
}
