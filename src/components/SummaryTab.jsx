import { useMemo, useState } from 'react';
import Icons from './Icons.jsx';
import { MONTH_META, PLANT_META } from '../constants.js';
import { formatNumber } from '../utils/format.js';

// ==================== SUMMARY TAB ("สรุปรวม") ====================
export default function SummaryTab({ state, actions, monthLabels }) {
  const { summaryData, totals, exporting } = state;
  const [searchTerm, setSearchTerm] = useState('');

  const filtered = useMemo(() => {
    const q = searchTerm.trim().toLowerCase();
    return q ? (summaryData || []).filter((r) => r.partNumber.toLowerCase().includes(q)) : summaryData || [];
  }, [summaryData, searchTerm]);

  const filteredTotal = useMemo(
    () => filtered.reduce((acc, r) => ({ n: acc.n + r.n, n1: acc.n1 + r.n1, n2: acc.n2 + r.n2, n3: acc.n3 + r.n3 }), { n: 0, n1: 0, n2: 0, n3: 0 }),
    [filtered],
  );

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 sm:gap-4">
        {MONTH_META.map((m, i) => (
          <div key={m.key} className={`bg-gradient-to-br ${m.gradient} rounded-2xl p-4 sm:p-5 text-white shadow-lg`}>
            <p className="text-white/80 text-sm">{monthLabels[i]}</p>
            <p className="text-2xl sm:text-3xl font-bold">{formatNumber(totals[m.key])}</p>
          </div>
        ))}
      </div>

      <div className="flex flex-col sm:flex-row gap-3 sm:gap-4">
        <button onClick={actions.openChart} className="flex-1 py-3 bg-gradient-to-r from-purple-600 to-pink-600 text-white rounded-xl flex items-center justify-center gap-2">
          <span className="w-5 h-5"><Icons.PieChart /></span>ดูกราฟวิเคราะห์
        </button>
        <button onClick={actions.openPreview} className="flex-1 py-3 bg-white border-2 border-blue-600 text-blue-600 rounded-xl flex items-center justify-center gap-2">
          <span className="w-5 h-5"><Icons.Eye /></span>Preview Excel
        </button>
        <button onClick={() => actions.exportExcel()} disabled={exporting} className="flex-1 py-3 bg-gradient-to-r from-blue-600 to-blue-700 text-white rounded-xl flex items-center justify-center gap-2 disabled:opacity-60">
          <span className="w-5 h-5">{exporting ? <Icons.Refresh /> : <Icons.Download />}</span>{exporting ? 'กำลังสร้างไฟล์...' : 'ดาวน์โหลด'}
        </button>
      </div>

      <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden">
        <div className="px-4 sm:px-6 py-4 border-b bg-slate-50 flex flex-wrap items-center justify-between gap-3">
          <h3 className="font-semibold">สรุปยอดรวม ({filtered.length} รายการ)</h3>
          <div className="relative">
            <span className="absolute left-2 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400"><Icons.Search /></span>
            <input type="text" placeholder="ค้นหา Part Number..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="pl-8 pr-3 py-1.5 border border-slate-200 rounded-lg text-sm w-56" />
          </div>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="bg-slate-50">
                <th className="text-left py-3 px-4">Part Number</th>
                <th className="text-left">Plants</th>
                {monthLabels.map((m) => <th key={m} className="text-right whitespace-nowrap px-2">{m}</th>)}
                <th className="text-right pr-4">Total</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((row) => (
                <tr key={row.partNumber} className="border-b hover:bg-slate-50">
                  <td className="py-3 px-4 font-mono text-blue-600">{row.partNumber}</td>
                  <td>
                    <div className="flex flex-wrap gap-1">
                      {row.plants.split(', ').map((p) => <span key={p} className={`px-2 py-0.5 rounded text-xs ${PLANT_META[p]?.badge}`}>{p}</span>)}
                    </div>
                  </td>
                  <td className="text-right px-2">{formatNumber(row.n)}</td>
                  <td className="text-right px-2">{formatNumber(row.n1)}</td>
                  <td className="text-right px-2">{formatNumber(row.n2)}</td>
                  <td className="text-right px-2">{formatNumber(row.n3)}</td>
                  <td className="text-right font-bold pr-4">{formatNumber(row.n + row.n1 + row.n2 + row.n3)}</td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr><td colSpan={7} className="py-8 text-center text-slate-400">ไม่พบข้อมูล</td></tr>
              )}
            </tbody>
            {filtered.length > 0 && (
              <tfoot>
                <tr className="bg-slate-50 font-semibold">
                  <td className="py-3 px-4" colSpan={2}>รวม ({filtered.length} รายการ)</td>
                  <td className="text-right px-2">{formatNumber(filteredTotal.n)}</td>
                  <td className="text-right px-2">{formatNumber(filteredTotal.n1)}</td>
                  <td className="text-right px-2">{formatNumber(filteredTotal.n2)}</td>
                  <td className="text-right px-2">{formatNumber(filteredTotal.n3)}</td>
                  <td className="text-right pr-4">{formatNumber(filteredTotal.n + filteredTotal.n1 + filteredTotal.n2 + filteredTotal.n3)}</td>
                </tr>
              </tfoot>
            )}
          </table>
        </div>
      </div>
    </div>
  );
}
