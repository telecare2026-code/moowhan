import { useMemo, useState } from 'react';
import Icons from './Icons.jsx';
import { PLANTS, PLANT_META } from '../constants.js';
import { formatNumber } from '../utils/format.js';

const PAGE = 50;

// ==================== PREVIEW TAB ("ตรวจสอบข้อมูล") ====================
export default function PreviewTab({ state, actions, monthLabels }) {
  const { processedData, sourceFiles, template, fileCounts } = state;
  const [searchTerm, setSearchTerm] = useState('');
  const [filterPlant, setFilterPlant] = useState('all');
  const [expanded, setExpanded] = useState({});
  const [limits, setLimits] = useState({});

  const filtered = useMemo(() => {
    const out = {};
    const q = searchTerm.trim().toLowerCase();
    Object.entries(processedData || {}).forEach(([sheet, rows]) => {
      const plant = sheet.split(' ')[0];
      if (filterPlant !== 'all' && plant !== filterPlant) return;
      out[sheet] = q ? rows.filter((r) => r.partNumber.toLowerCase().includes(q) || String(r.partCode).toLowerCase().includes(q)) : rows;
    });
    return out;
  }, [processedData, searchTerm, filterPlant]);

  const missingPlants = PLANTS.filter((p) => fileCounts[p] === 0);
  const filesWithError = sourceFiles.filter((f) => f.status === 'error').length;
  const filesProcessed = sourceFiles.filter((f) => f.status === 'done').length;
  const toggle = (sheet) => setExpanded((prev) => ({ ...prev, [sheet]: !prev[sheet] }));

  return (
    <div className="space-y-6">
      <div className="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm">
        <div className="flex flex-col sm:flex-row gap-3 sm:gap-4">
          <input
            type="text"
            placeholder="ค้นหา Part Number / Part Code..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="flex-1 px-4 py-2 border border-slate-200 rounded-xl"
          />
          <select value={filterPlant} onChange={(e) => setFilterPlant(e.target.value)} className="px-4 py-2 border border-slate-200 rounded-xl">
            <option value="all">ทุกโรงงาน</option>
            {PLANTS.map((p) => <option key={p} value={p}>{p} – {PLANT_META[p].label}</option>)}
          </select>
        </div>
      </div>

      <div className="bg-gradient-to-br from-emerald-50 to-teal-50 border border-emerald-200 rounded-2xl p-4 sm:p-6">
        <div className="flex flex-wrap justify-between items-center gap-3 mb-4">
          <h3 className="text-lg font-semibold text-emerald-900">ตรวจสอบข้อมูล</h3>
          <button onClick={actions.openPreview} className="px-4 py-2 bg-emerald-600 text-white rounded-lg flex items-center gap-2">
            <span className="w-4 h-4"><Icons.Eye /></span>Excel Preview
          </button>
        </div>
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
          <div className="bg-white/70 rounded-xl p-3"><p className="text-sm text-emerald-600">ไฟล์หลัก</p><p className="font-semibold">{template ? '✓ พร้อม' : 'ไม่มี'}</p></div>
          <div className="bg-white/70 rounded-xl p-3"><p className="text-sm text-emerald-600">ประมวลผลสำเร็จ</p><p className="font-semibold">{filesProcessed} / {sourceFiles.length}</p></div>
          <div className="bg-white/70 rounded-xl p-3"><p className="text-sm text-emerald-600">โรงงานขาด</p><p className={`font-semibold ${missingPlants.length ? 'text-amber-600' : ''}`}>{missingPlants.length > 0 ? missingPlants.join(', ') : 'ครบ'}</p></div>
          <div className="bg-white/70 rounded-xl p-3"><p className="text-sm text-emerald-600">ข้อผิดพลาด</p><p className={`font-semibold ${filesWithError ? 'text-red-600' : ''}`}>{filesWithError > 0 ? `${filesWithError} ไฟล์` : 'ไม่มี'}</p></div>
        </div>
      </div>

      {Object.entries(filtered).map(([sheet, rows]) => {
        const plant = sheet.split(' ')[0];
        const limit = limits[sheet] || PAGE;
        return (
          <div key={sheet} className="bg-white border border-slate-200 rounded-xl overflow-hidden">
            <button onClick={() => toggle(sheet)} className="w-full px-4 sm:px-5 py-4 flex items-center justify-between hover:bg-slate-50" aria-expanded={!!expanded[sheet]}>
              <div className="flex items-center gap-3">
                <span className={`px-3 py-1 rounded-lg text-sm font-medium ${PLANT_META[plant]?.badge}`}>{plant}</span>
                <span className="font-semibold">{sheet}</span>
                <span className="text-slate-400">({rows.length})</span>
              </div>
              <span className="w-5 h-5 text-slate-400">{expanded[sheet] ? <Icons.ChevronUp /> : <Icons.ChevronDown />}</span>
            </button>
            {expanded[sheet] && (
              <div className="px-4 sm:px-5 pb-4 overflow-x-auto">
                {rows.length === 0 ? (
                  <p className="text-sm text-slate-400 py-3">ไม่มีข้อมูล</p>
                ) : (
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="text-slate-500 border-b">
                        <th className="text-left py-2">Part Number</th>
                        <th className="text-left">Dock</th>
                        <th className="text-right">Pack</th>
                        {monthLabels.map((m) => <th key={m} className="text-right whitespace-nowrap pl-3">{m}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.slice(0, limit).map((row) => (
                        <tr key={row.id} className="border-b hover:bg-slate-50">
                          <td className="py-2 font-mono text-blue-600">{row.partNumber}</td>
                          <td className="text-slate-500 text-xs">{row.dockCode}</td>
                          <td className="text-right text-slate-500">{formatNumber(row.packingSize)}</td>
                          <td className="text-right">{formatNumber(row.n)}</td>
                          <td className="text-right">{formatNumber(row.n1)}</td>
                          <td className="text-right">{formatNumber(row.n2)}</td>
                          <td className="text-right">{formatNumber(row.n3)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
                {rows.length > limit && (
                  <button onClick={() => setLimits((p) => ({ ...p, [sheet]: limit + PAGE }))} className="mt-3 text-sm text-blue-600 hover:underline">
                    แสดงเพิ่ม ({rows.length - limit} แถวที่เหลือ)
                  </button>
                )}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
