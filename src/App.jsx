import { useEffect, useState } from 'react';
import Icons from './components/Icons.jsx';
import Banner from './components/Banner.jsx';
import UploadTab from './components/UploadTab.jsx';
import PreviewTab from './components/PreviewTab.jsx';
import SummaryTab from './components/SummaryTab.jsx';
import MatchingTab from './components/MatchingTab.jsx';
import ChangelogTab from './components/ChangelogTab.jsx';
import ChartDashboard from './components/ChartDashboard.jsx';
import ExcelPreview from './components/ExcelPreview.jsx';
import useConsolidator from './state/useConsolidator.js';
import { FEATURES } from './features/index.js';
import { formatMonth, todayStamp } from './utils/format.js';

const APP_TITLE = 'รีบทำเดียวอดเลนเกม';

// ==================== MAIN APP ====================
export default function App() {
  const { state, actions: core } = useConsolidator();
  const [tab, setTab] = useState('upload');
  const [showPreview, setShowPreview] = useState(false);
  const [showChart, setShowChart] = useState(false);

  const actions = {
    ...core,
    goTo: setTab,
    openPreview: () => setShowPreview(true),
    openChart: () => setShowChart(true),
    processAll: async () => {
      const ok = await core.processAll();
      if (ok) {
        setTab('preview');
        FEATURES.forEach((f) => f.onProcessed?.(ctx));
      }
      return ok;
    },
    exportExcel: (options = {}) => {
      const extraSheets = [
        ...FEATURES.flatMap((f) => (f.exportSheets ? f.exportSheets(ctx) || [] : [])),
        ...(options.extraSheets || []),
      ];
      return core.exportExcel({ ...options, extraSheets });
    },
    reset: () => {
      core.reset();
      setTab('upload');
      setShowPreview(false);
      setShowChart(false);
    },
  };
  const ctx = { state, actions, monthLabels: state.monthLabels };

  const tabs = [
    { id: 'upload', order: 10, label: 'อัปโหลดไฟล์', Icon: Icons.Upload, enabled: true, Component: UploadTab },
    { id: 'preview', order: 20, label: 'ตรวจสอบข้อมูล', Icon: Icons.Eye, enabled: !!state.processedData, tooltip: 'ต้องประมวลผลไฟล์ก่อน', Component: PreviewTab },
    { id: 'matching', order: 30, label: 'รายละเอียด Matching', Icon: Icons.Link, enabled: !!state.matchingDetails, tooltip: 'ต้องประมวลผลไฟล์ก่อน', Component: MatchingTab },
    { id: 'changelog', order: 35, label: 'ประวัติการแก้ไข', Icon: Icons.History, enabled: state.changeLog.length > 0, tooltip: 'ยังไม่มีรายการ', Component: ChangelogTab },
    { id: 'summary', order: 40, label: 'สรุปรวม', Icon: Icons.Chart, enabled: !!state.summaryData, tooltip: 'ต้องประมวลผลไฟล์ก่อน', Component: SummaryTab },
    ...FEATURES.filter((f) => f.tab).map((f) => ({
      id: f.id,
      order: f.order ?? 100,
      ...f.tab,
      enabled: f.tab.enabled ? !!f.tab.enabled(ctx) : true,
    })),
  ].sort((a, b) => a.order - b.order);

  const activeTab = tabs.find((t) => t.id === tab && t.enabled) || tabs[0];
  useEffect(() => {
    if (activeTab.id !== tab) setTab(activeTab.id);
  }, [activeTab.id, tab]);

  const previewFileName = `Production_${state.template?.kind === 'xlsx' ? 'Updated' : 'Summary'}_${todayStamp()}.xlsx`;

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50 text-slate-900">
      <header className="bg-white border-b border-slate-200 sticky top-0 z-40 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 py-3 sm:py-4 flex items-center justify-between gap-3">
          <div className="flex items-center gap-3 min-w-0">
            <div className="w-11 h-11 sm:w-12 sm:h-12 p-2.5 bg-gradient-to-br from-blue-600 to-blue-700 rounded-xl text-white shadow-lg flex-shrink-0"><Icons.File /></div>
            <div className="min-w-0">
              <h1 className="text-lg sm:text-xl font-bold bg-gradient-to-r from-blue-700 to-blue-500 bg-clip-text text-transparent truncate">{APP_TITLE}</h1>
              <p className="text-xs sm:text-sm text-slate-500 truncate">
                TMT Camera Production Plan Summary
                {state.productionMonth && <span className="text-blue-600"> • เดือนผลิต {formatMonth(state.productionMonth)} (N)</span>}
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2 flex-shrink-0">
            {FEATURES.map((f) => (f.headerActions ? <span key={f.id}>{f.headerActions(ctx)}</span> : null))}
            <button onClick={actions.reset} className="flex items-center gap-2 px-3 sm:px-4 py-2 bg-slate-100 hover:bg-slate-200 rounded-lg text-slate-600">
              <span className="w-4 h-4"><Icons.Reset /></span><span className="hidden sm:inline">รีเซ็ต</span>
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-6">
        <nav className="flex flex-wrap gap-2 mb-6" aria-label="แท็บ">
          {tabs.map((item) => (
            <div key={item.id} className="relative group">
              <button
                onClick={() => item.enabled && setTab(item.id)}
                disabled={!item.enabled}
                aria-current={activeTab.id === item.id ? 'page' : undefined}
                className={`flex items-center gap-2 px-3 sm:px-4 py-2.5 rounded-xl font-medium transition-all border text-sm sm:text-base ${
                  activeTab.id === item.id ? 'bg-blue-600 text-white border-blue-600 shadow-md'
                    : !item.enabled ? 'bg-slate-50 text-slate-300 border-slate-200 cursor-not-allowed'
                    : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'}`}
              >
                <span className="w-5 h-5"><item.Icon /></span>{item.label}
                {item.badge?.(ctx) ? <span className="ml-1 px-1.5 py-0.5 rounded-full text-xs bg-white/20">{item.badge(ctx)}</span> : null}
              </button>
              {item.tooltip && !item.enabled && (
                <div className="absolute bottom-full left-1/2 -translate-x-1/2 mb-2 px-3 py-1.5 bg-slate-800 text-white text-xs rounded-lg whitespace-nowrap opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all z-10">
                  {item.tooltip}
                  <div className="absolute top-full left-1/2 -translate-x-1/2 border-4 border-transparent border-t-slate-800" />
                </div>
              )}
            </div>
          ))}
        </nav>

        <Banner tone="error" message={state.error} onDismiss={actions.dismissError} />
        <Banner tone={state.notice?.tone || 'warning'} message={state.notice?.message} onDismiss={actions.dismissNotice} />
        {state.isStale && (
          <Banner tone="warning" message="รายการไฟล์เปลี่ยนไปหลังจากประมวลผลครั้งล่าสุด กด “รวมข้อมูลและคำนวณ” อีกครั้งเพื่ออัปเดตผลลัพธ์" />
        )}

        <activeTab.Component state={state} actions={actions} monthLabels={state.monthLabels} />

        {showChart && <ChartDashboard data={state.processedData} summaryData={state.summaryData} monthLabels={state.monthLabels} onClose={() => setShowChart(false)} />}
        {showPreview && (
          <ExcelPreview
            data={state.processedData}
            summaryData={state.summaryData}
            fileName={previewFileName}
            onClose={() => setShowPreview(false)}
            onDownload={() => actions.exportExcel()}
          />
        )}
      </main>

      <footer className="border-t border-slate-200 mt-12 py-6 bg-white text-center text-slate-400 text-sm">
        moowhan v{__APP_VERSION__}
      </footer>
    </div>
  );
}
