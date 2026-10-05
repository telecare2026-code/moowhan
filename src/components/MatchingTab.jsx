import { Fragment, useMemo, useState } from 'react';
import { PLANT_META } from '../constants.js';
import { formatNumber } from '../utils/format.js';

// ==================== MATCHING TAB ====================
// Which plant files contributed to each part number.
export default function MatchingTab({ state, monthLabels }) {
  const { matchingDetails } = state;
  const [onlyMulti, setOnlyMulti] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');

  const entries = useMemo(() => {
    const q = searchTerm.trim().toLowerCase();
    return Object.entries(matchingDetails || {})
      .filter(([part, d]) => (!onlyMulti || d.plants.size > 1) && (!q || part.toLowerCase().includes(q)))
      .sort((a, b) => a[0].localeCompare(b[0]));
  }, [matchingDetails, onlyMulti, searchTerm]);

  return (
    <div className="space-y-6">
      <div className="bg-blue-50 border border-blue-200 rounded-xl p-4">
        <p className="text-blue-800 text-sm">
          <span className="font-semibold">รายละเอียด Matching:</span> แสดงว่าแต่ละ Part Number ถูกดึงมาจากไฟล์ของโรงงานใดบ้าง (แถวซ้ำกันในโรงงานเดียวคือ Dock ต่างกัน)
        </p>
      </div>

      <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden">
        <div className="px-4 sm:px-6 py-4 border-b bg-slate-50 flex flex-wrap justify-between items-center gap-3">
          <h3 className="font-semibold">รายละเอียดการ Matching ({entries.length} Part Numbers)</h3>
          <div className="flex items-center gap-4">
            <label className="flex items-center gap-2 text-sm text-slate-600 cursor-pointer">
              <input type="checkbox" checked={onlyMulti} onChange={(e) => setOnlyMulti(e.target.checked)} className="rounded" />
              เฉพาะที่มีหลายโรงงาน
            </label>
            <input type="text" placeholder="ค้นหา..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="px-3 py-1.5 border border-slate-200 rounded-lg text-sm w-40" />
          </div>
        </div>
        <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
          <table className="w-full text-sm">
            <thead className="sticky top-0 bg-slate-50 z-10">
              <tr>
                <th className="text-left py-3 px-4 border-b">Part Number</th>
                <th className="text-left border-b">Plant</th>
                <th className="text-left border-b">Part Code</th>
                <th className="text-left border-b">Dock</th>
                {monthLabels.map((m) => <th key={m} className="text-right border-b whitespace-nowrap px-2">{m}</th>)}
              </tr>
            </thead>
            <tbody>
              {entries.map(([partNumber, d]) => (
                <Fragment key={partNumber}>
                  {d.sources.map((source, idx) => (
                    <tr key={`${partNumber}-${idx}`} className="border-b hover:bg-slate-50">
                      {idx === 0 && (
                        <td rowSpan={d.sources.length} className="py-3 px-4 font-mono text-blue-600 font-medium align-top bg-slate-50/50">
                          {partNumber}
                          <div className="text-xs text-slate-400 font-sans font-normal mt-1">{d.sources.length} แหล่ง</div>
                        </td>
                      )}
                      <td className="py-3"><span className={`px-2 py-0.5 rounded text-xs ${PLANT_META[source.plant]?.badge || 'bg-gray-100 text-gray-700'}`}>{source.plant}</span></td>
                      <td className="font-mono text-xs text-slate-600">{source.partCode}</td>
                      <td className="text-xs text-slate-500">{source.dockCode}</td>
                      <td className="text-right px-2">{formatNumber(source.n)}</td>
                      <td className="text-right px-2">{formatNumber(source.n1)}</td>
                      <td className="text-right px-2">{formatNumber(source.n2)}</td>
                      <td className="text-right px-2">{formatNumber(source.n3)}</td>
                    </tr>
                  ))}
                </Fragment>
              ))}
              {entries.length === 0 && <tr><td colSpan={8} className="py-8 text-center text-slate-400">ไม่พบข้อมูล</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
