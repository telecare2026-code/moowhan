import React, { useState } from 'react';

const Icons = {
  Upload: () => <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeWidth="2" d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>,
  File: () => <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeWidth="2" d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><path strokeWidth="2" d="M14 2v6h6M8 13h8M8 17h8"/></svg>,
  Check: () => <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeWidth="2" d="M22 11.08V12a10 10 0 11-5.93-9.14"/><path strokeWidth="2" d="M22 4L12 14.01l-3-3"/></svg>,
  Trash: () => <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeWidth="2" d="M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"/></svg>,
  Play: () => <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24"><polygon points="5 3 19 12 5 21 5 3" fill="currentColor"/></svg>,
  Download: () => <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeWidth="2" d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3"/></svg>,
  Chart: () => <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeWidth="2" d="M3 3v18h18"/><path strokeWidth="2" d="M18 17V9M13 17V5M8 17v-3"/></svg>,
  Eye: () => <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeWidth="2" d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3" strokeWidth="2"/></svg>,
  Refresh: () => <svg className="w-full h-full animate-spin" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeWidth="2" d="M23 4v6h-6M1 20v-6h6"/><path strokeWidth="2" d="M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15"/></svg>,
  ChevronDown: () => <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeWidth="2" d="M6 9l6 6 6-6"/></svg>,
  ChevronUp: () => <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeWidth="2" d="M18 15l-6-6-6 6"/></svg>,
};

const mockData = {
  'BP Daily': [
    { part: '86790-0K051-00', code: 'I338', pack: 25, n: 1891, n1: 1677, n2: 919, n3: 0 },
    { part: '86790-0K080-00', code: 'A369', pack: 42, n: 299, n1: 67, n2: 0, n3: 0 },
  ],
  'BPK Daily': [
    { part: '86790-0K051-00', code: 'W036', pack: 30, n: 5790, n1: 7980, n2: 5490, n3: 3720 },
    { part: '86790-0K110-00', code: 'M523', pack: 20, n: 6680, n1: 6740, n2: 5120, n3: 4220 },
  ],
  'GW Daily': [
    { part: '86790-BZ271-00', code: 'A055', pack: 36, n: 2088, n1: 1980, n2: 1800, n3: 1836 },
    { part: '86790-BZ310-00', code: 'B800', pack: 15, n: 3303, n1: 3128, n2: 3128, n3: 2584 },
  ],
  'SR Daily': [
    { part: '86790-0K051-00', code: 'L342', pack: 25, n: 1352, n1: 1787, n2: 2037, n3: 0 },
    { part: '86790-0K080-00', code: 'G871', pack: 42, n: 676, n1: 1170, n2: 1084, n3: 451 },
  ],
};

export default function App() {
  const [tab, setTab] = useState('upload');
  const [mainFile, setMainFile] = useState(null);
  const [files, setFiles] = useState([]);
  const [processing, setProcessing] = useState(false);
  const [processed, setProcessed] = useState(false);
  const [expanded, setExpanded] = useState({});

  const addFiles = (plant) => {
    const mockFiles = {
      BP: [{ name: 'BP_veh_481D.xls', cat: 'BP' }, { name: 'BP_veh_581D.xls', cat: 'BP' }],
      BPK: [{ name: 'BPK_packing_481D.xls', cat: 'BPK' }, { name: 'BPK_packing_581D.xls', cat: 'BPK' }],
      GW: [{ name: 'GW_veh_BSUV.xls', cat: 'GW' }, { name: 'GW_veh_DG7.xls', cat: 'GW' }],
      SR: [{ name: 'SR_veh_481D.xls', cat: 'SR' }, { name: 'SR_veh_581D.xls', cat: 'SR' }],
    };
    setFiles(prev => [...prev, ...mockFiles[plant].filter(f => !prev.find(p => p.name === f.name))]);
  };

  const removeFile = (name) => setFiles(prev => prev.filter(f => f.name !== name));

  const process = async () => {
    setProcessing(true);
    await new Promise(r => setTimeout(r, 2000));
    setProcessing(false);
    setProcessed(true);
    setTab('preview');
  };

  const catColors = {
    BP: 'bg-blue-500/20 text-blue-400 border-blue-500/50',
    BPK: 'bg-green-500/20 text-green-400 border-green-500/50',
    GW: 'bg-purple-500/20 text-purple-400 border-purple-500/50',
    SR: 'bg-orange-500/20 text-orange-400 border-orange-500/50',
  };

  const summary = processed ? Object.values(mockData).flat().reduce((acc, row) => {
    if (!acc[row.part]) acc[row.part] = { n: 0, n1: 0, n2: 0, n3: 0, plants: new Set() };
    acc[row.part].n += row.n;
    acc[row.part].n1 += row.n1;
    acc[row.part].n2 += row.n2;
    acc[row.part].n3 += row.n3;
    return acc;
  }, {}) : {};

  const totals = processed ? {
    n: Object.values(summary).reduce((s, r) => s + r.n, 0),
    n1: Object.values(summary).reduce((s, r) => s + r.n1, 0),
    n2: Object.values(summary).reduce((s, r) => s + r.n2, 0),
    n3: Object.values(summary).reduce((s, r) => s + r.n3, 0),
  } : { n: 0, n1: 0, n2: 0, n3: 0 };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900 text-white">
      {/* Header */}
      <header className="bg-slate-800/60 backdrop-blur border-b border-slate-700 sticky top-0 z-50">
        <div className="max-w-6xl mx-auto px-4 py-3">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 p-2 bg-gradient-to-br from-emerald-400 to-cyan-500 rounded-xl">
              <Icons.File />
            </div>
            <div>
              <h1 className="text-lg font-bold">Production Data Consolidator</h1>
              <p className="text-xs text-slate-400">TMT Camera Production Plan</p>
            </div>
          </div>
        </div>
      </header>

      {/* Tabs */}
      <div className="max-w-6xl mx-auto px-4 py-4">
        <div className="flex gap-2">
          {[
            { id: 'upload', label: 'อัปโหลด', Icon: Icons.Upload },
            { id: 'preview', label: 'ตรวจสอบ', Icon: Icons.Eye },
            { id: 'summary', label: 'สรุป', Icon: Icons.Chart },
          ].map(t => (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg font-medium transition-all ${
                tab === t.id
                  ? 'bg-emerald-500 text-white shadow-lg shadow-emerald-500/30'
                  : 'bg-slate-700/50 text-slate-300 hover:bg-slate-700'
              }`}
            >
              <div className="w-4 h-4"><t.Icon /></div>
              {t.label}
            </button>
          ))}
        </div>
      </div>

      <main className="max-w-6xl mx-auto px-4 pb-8">
        {/* Upload Tab */}
        {tab === 'upload' && (
          <div className="space-y-4">
            {/* Main File */}
            <div className="bg-slate-800/50 rounded-2xl border border-slate-700 p-5">
              <h2 className="text-base font-semibold mb-3 flex items-center gap-2">
                <span className="w-5 h-5 text-amber-400"><Icons.File /></span>
                ไฟล์หลัก (Template)
              </h2>
              <button
                onClick={() => setMainFile({ name: 'ไฟล์_หลัก.xlsx', size: '196.8 KB' })}
                className={`w-full border-2 border-dashed rounded-xl p-6 text-center transition-all ${
                  mainFile ? 'border-emerald-500 bg-emerald-500/10' : 'border-slate-600 hover:border-cyan-500'
                }`}
              >
                {mainFile ? (
                  <div className="flex items-center justify-center gap-3">
                    <div className="w-8 h-8 text-emerald-500"><Icons.Check /></div>
                    <div className="text-left">
                      <p className="font-medium">{mainFile.name}</p>
                      <p className="text-sm text-slate-400">{mainFile.size}</p>
                    </div>
                  </div>
                ) : (
                  <div>
                    <div className="w-10 h-10 mx-auto text-slate-500 mb-2"><Icons.Upload /></div>
                    <p className="text-slate-300">คลิกเพื่อเลือกไฟล์หลัก</p>
                  </div>
                )}
              </button>
            </div>

            {/* Source Files */}
            <div className="bg-slate-800/50 rounded-2xl border border-slate-700 p-5">
              <h2 className="text-base font-semibold mb-3">ไฟล์ข้อมูลรายโรงงาน</h2>
              
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
                {['BP', 'BPK', 'GW', 'SR'].map(plant => (
                  <button
                    key={plant}
                    onClick={() => addFiles(plant)}
                    className={`p-3 rounded-xl border-2 border-dashed text-center hover:border-cyan-500 transition-all ${catColors[plant]}`}
                  >
                    <div className="text-lg font-bold">{plant}</div>
                    <div className="text-xs opacity-75">+ เพิ่มไฟล์</div>
                  </button>
                ))}
              </div>

              {files.length > 0 && (
                <div className="space-y-2 max-h-48 overflow-y-auto">
                  {files.map((file, i) => (
                    <div key={i} className="flex items-center justify-between bg-slate-700/50 rounded-lg px-4 py-2">
                      <div className="flex items-center gap-3">
                        <div className="w-5 h-5 text-green-500"><Icons.Check /></div>
                        <span className="text-sm">{file.name}</span>
                      </div>
                      <div className="flex items-center gap-2">
                        <span className={`px-2 py-0.5 rounded text-xs border ${catColors[file.cat]}`}>{file.cat}</span>
                        <button onClick={() => removeFile(file.name)} className="w-4 h-4 text-red-400 hover:text-red-300">
                          <Icons.Trash />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>

            {/* Process Button */}
            <button
              onClick={process}
              disabled={!mainFile || files.length === 0 || processing}
              className={`w-full py-4 rounded-xl font-semibold text-lg flex items-center justify-center gap-3 transition-all ${
                mainFile && files.length > 0 && !processing
                  ? 'bg-gradient-to-r from-emerald-500 to-cyan-500 shadow-lg shadow-emerald-500/30 hover:shadow-emerald-500/50'
                  : 'bg-slate-700 text-slate-400 cursor-not-allowed'
              }`}
            >
              {processing ? (
                <>
                  <div className="w-5 h-5"><Icons.Refresh /></div>
                  กำลังประมวลผล...
                </>
              ) : (
                <>
                  <div className="w-5 h-5"><Icons.Play /></div>
                  รวมข้อมูลและคำนวณ
                </>
              )}
            </button>
          </div>
        )}

        {/* Preview Tab */}
        {tab === 'preview' && processed && (
          <div className="space-y-3">
            {Object.entries(mockData).map(([sheet, rows]) => (
              <div key={sheet} className="bg-slate-800/50 rounded-xl border border-slate-700 overflow-hidden">
                <button
                  onClick={() => setExpanded(p => ({ ...p, [sheet]: !p[sheet] }))}
                  className="w-full px-4 py-3 flex items-center justify-between hover:bg-slate-700/30"
                >
                  <div className="flex items-center gap-2">
                    <span className={`px-2 py-0.5 rounded text-xs border ${catColors[sheet.split(' ')[0]]}`}>
                      {sheet.split(' ')[0]}
                    </span>
                    <span className="font-medium">{sheet}</span>
                    <span className="text-slate-400 text-sm">({rows.length} รายการ)</span>
                  </div>
                  <div className="w-5 h-5 text-slate-400">
                    {expanded[sheet] ? <Icons.ChevronUp /> : <Icons.ChevronDown />}
                  </div>
                </button>
                
                {expanded[sheet] && (
                  <div className="px-4 pb-3 overflow-x-auto">
                    <table className="w-full text-sm">
                      <thead>
                        <tr className="text-slate-400 border-b border-slate-700">
                          <th className="text-left py-2 px-2">Part Number</th>
                          <th className="text-right py-2 px-2">N</th>
                          <th className="text-right py-2 px-2">N+1</th>
                          <th className="text-right py-2 px-2">N+2</th>
                          <th className="text-right py-2 px-2">N+3</th>
                        </tr>
                      </thead>
                      <tbody>
                        {rows.map((row, i) => (
                          <tr key={i} className="border-b border-slate-700/50">
                            <td className="py-2 px-2 font-mono text-cyan-400">{row.part}</td>
                            <td className="py-2 px-2 text-right">{row.n.toLocaleString()}</td>
                            <td className="py-2 px-2 text-right">{row.n1.toLocaleString()}</td>
                            <td className="py-2 px-2 text-right">{row.n2.toLocaleString()}</td>
                            <td className="py-2 px-2 text-right">{row.n3.toLocaleString()}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            ))}
          </div>
        )}

        {/* Summary Tab */}
        {tab === 'summary' && processed && (
          <div className="space-y-4">
            {/* Summary Cards */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              {[
                { label: 'N (Feb)', value: totals.n, color: 'from-cyan-500 to-blue-500' },
                { label: 'N+1 (Mar)', value: totals.n1, color: 'from-emerald-500 to-green-500' },
                { label: 'N+2 (Apr)', value: totals.n2, color: 'from-amber-500 to-orange-500' },
                { label: 'N+3 (May)', value: totals.n3, color: 'from-purple-500 to-pink-500' },
              ].map((card, i) => (
                <div key={i} className={`bg-gradient-to-br ${card.color} rounded-xl p-4 shadow-lg`}>
                  <p className="text-white/80 text-sm">{card.label}</p>
                  <p className="text-2xl font-bold">{card.value.toLocaleString()}</p>
                </div>
              ))}
            </div>

            {/* Summary Table */}
            <div className="bg-slate-800/50 rounded-xl border border-slate-700 overflow-hidden">
              <div className="px-4 py-3 border-b border-slate-700">
                <h3 className="font-semibold flex items-center gap-2">
                  <div className="w-5 h-5 text-emerald-400"><Icons.Chart /></div>
                  สรุปยอดรวมตาม Part Number
                </h3>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="bg-slate-700/50 text-slate-300">
                      <th className="text-left py-3 px-4">Part Number</th>
                      <th className="text-right py-3 px-4 text-cyan-400">N</th>
                      <th className="text-right py-3 px-4 text-emerald-400">N+1</th>
                      <th className="text-right py-3 px-4 text-amber-400">N+2</th>
                      <th className="text-right py-3 px-4 text-purple-400">N+3</th>
                      <th className="text-right py-3 px-4">Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {Object.entries(summary).map(([part, data], i) => (
                      <tr key={i} className="border-b border-slate-700/50 hover:bg-slate-700/30">
                        <td className="py-2 px-4 font-mono text-cyan-400">{part}</td>
                        <td className="py-2 px-4 text-right">{data.n.toLocaleString()}</td>
                        <td className="py-2 px-4 text-right">{data.n1.toLocaleString()}</td>
                        <td className="py-2 px-4 text-right">{data.n2.toLocaleString()}</td>
                        <td className="py-2 px-4 text-right">{data.n3.toLocaleString()}</td>
                        <td className="py-2 px-4 text-right font-bold">{(data.n + data.n1 + data.n2 + data.n3).toLocaleString()}</td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot>
                    <tr className="bg-slate-700/50 font-bold">
                      <td className="py-3 px-4">Grand Total</td>
                      <td className="py-3 px-4 text-right text-cyan-400">{totals.n.toLocaleString()}</td>
                      <td className="py-3 px-4 text-right text-emerald-400">{totals.n1.toLocaleString()}</td>
                      <td className="py-3 px-4 text-right text-amber-400">{totals.n2.toLocaleString()}</td>
                      <td className="py-3 px-4 text-right text-purple-400">{totals.n3.toLocaleString()}</td>
                      <td className="py-3 px-4 text-right">{(totals.n + totals.n1 + totals.n2 + totals.n3).toLocaleString()}</td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>

            {/* Download Button */}
            <button className="w-full py-4 bg-gradient-to-r from-emerald-500 to-cyan-500 rounded-xl font-semibold flex items-center justify-center gap-3 shadow-lg shadow-emerald-500/30">
              <div className="w-5 h-5"><Icons.Download /></div>
              ดาวน์โหลดไฟล์ Excel สรุป
            </button>
          </div>
        )}

        {/* Empty State */}
        {(tab === 'preview' || tab === 'summary') && !processed && (
          <div className="text-center py-16">
            <p className="text-slate-400 mb-4">กรุณาอัปโหลดไฟล์และประมวลผลก่อน</p>
            <button onClick={() => setTab('upload')} className="px-6 py-2 bg-slate-700 rounded-lg hover:bg-slate-600">
              ไปหน้าอัปโหลด
            </button>
          </div>
        )}
      </main>
    </div>
  );
}