import React, { useState } from 'react';
import * as XLSX from 'xlsx';

const Icons = {
  Upload: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12" />
    </svg>
  ),
  File: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" />
      <path strokeWidth="2" d="M14 2v6h6M8 13h8M8 17h8" />
    </svg>
  ),
  Check: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M22 11.08V12a10 10 0 11-5.93-9.14" />
      <path strokeWidth="2" d="M22 4L12 14.01l-3-3" />
    </svg>
  ),
  Trash: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2" />
    </svg>
  ),
  Play: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <polygon points="5 3 19 12 5 21 5 3" fill="currentColor" />
    </svg>
  ),
  Chart: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M3 3v18h18" />
      <path strokeWidth="2" d="M18 17V9M13 17V5M8 17v-3" />
    </svg>
  ),
  Eye: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
      <circle cx="12" cy="12" r="3" strokeWidth="2" />
    </svg>
  ),
  Refresh: () => (
    <svg className="w-full h-full animate-spin" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M23 4v6h-6M1 20v-6h6" />
      <path strokeWidth="2" d="M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15" />
    </svg>
  ),
  ChevronDown: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M6 9l6 6 6-6" />
    </svg>
  ),
  ChevronUp: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M18 15l-6-6-6 6" />
    </svg>
  ),
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

const plantMeta = {
  BP: { label: 'Ban Pho', badge: 'bg-blue-100 text-blue-700', border: 'border-blue-200' },
  BPK: { label: 'Ban Pho Kaeng Khoi', badge: 'bg-emerald-100 text-emerald-700', border: 'border-emerald-200' },
  GW: { label: 'Gateway', badge: 'bg-purple-100 text-purple-700', border: 'border-purple-200' },
  SR: { label: 'Samrong', badge: 'bg-orange-100 text-orange-700', border: 'border-orange-200' },
};

const formatNumber = (value) => (value ?? 0).toLocaleString();
const formatSize = (bytes) => `${(bytes / 1024).toFixed(1)} KB`;

export default function App() {
  const [tab, setTab] = useState('upload');
  const [mainFile, setMainFile] = useState(null);
  const [sourceFiles, setSourceFiles] = useState({ BP: [], BPK: [], GW: [], SR: [] });
  const [processing, setProcessing] = useState(false);
  const [processed, setProcessed] = useState(false);
  const [expanded, setExpanded] = useState({});

  const totalFiles = Object.values(sourceFiles).reduce((sum, files) => sum + files.length, 0);
  const canProcess = Boolean(mainFile) && totalFiles > 0 && !processing;
  const plantList = Object.keys(plantMeta);
  const missingPlants = plantList.filter((plant) => sourceFiles[plant].length === 0);

  const resetAll = () => {
    setTab('upload');
    setMainFile(null);
    setSourceFiles({ BP: [], BPK: [], GW: [], SR: [] });
    setProcessing(false);
    setProcessed(false);
    setExpanded({});
  };

  const handleMainFileChange = (event) => {
    const file = event.target.files?.[0];
    if (file) {
      setMainFile({ name: file.name, size: file.size });
    }
  };

  const handlePlantFiles = (plant, event) => {
    const files = Array.from(event.target.files || []);
    if (files.length === 0) return;

    setSourceFiles((prev) => ({
      ...prev,
      [plant]: [
        ...prev[plant],
        ...files.map((file) => ({
          name: file.name,
          size: file.size,
        })),
      ],
    }));
  };

  const removePlantFile = (plant, index) => {
    setSourceFiles((prev) => ({
      ...prev,
      [plant]: prev[plant].filter((_, idx) => idx !== index),
    }));
  };

  const process = async () => {
    if (!canProcess) return;
    setProcessing(true);
    await new Promise((resolve) => setTimeout(resolve, 1500));
    setProcessing(false);
    setProcessed(true);
    setTab('preview');
  };

  const summary = processed
    ? Object.values(mockData).flat().reduce((acc, row) => {
        if (!acc[row.part]) acc[row.part] = { n: 0, n1: 0, n2: 0, n3: 0 };
        acc[row.part].n += row.n;
        acc[row.part].n1 += row.n1;
        acc[row.part].n2 += row.n2;
        acc[row.part].n3 += row.n3;
        return acc;
      }, {})
    : {};

  const totals = processed
    ? {
        n: Object.values(summary).reduce((sum, row) => sum + row.n, 0),
        n1: Object.values(summary).reduce((sum, row) => sum + row.n1, 0),
        n2: Object.values(summary).reduce((sum, row) => sum + row.n2, 0),
        n3: Object.values(summary).reduce((sum, row) => sum + row.n3, 0),
      }
    : { n: 0, n1: 0, n2: 0, n3: 0 };

  const partPlantMap = processed
    ? Object.entries(mockData).reduce((acc, [sheet, rows]) => {
        const plant = sheet.split(' ')[0];
        rows.forEach((row) => {
          if (!acc[row.part]) acc[row.part] = new Set();
          acc[row.part].add(plant);
        });
        return acc;
      }, {})
    : {};

  const duplicateParts = Object.keys(partPlantMap).filter((part) => partPlantMap[part].size > 1);

  const buildSheet = (rows) =>
    rows.length > 0 ? XLSX.utils.json_to_sheet(rows) : XLSX.utils.aoa_to_sheet([['No data']]);

  const handleDownload = () => {
    if (!processed) return;

    const summaryRows = Object.entries(summary).map(([part, data]) => ({
      PartNumber: part,
      N: data.n,
      N1: data.n1,
      N2: data.n2,
      N3: data.n3,
      Total: data.n + data.n1 + data.n2 + data.n3,
    }));

    const verificationRows = [
      {
        Item: 'ไฟล์หลัก',
        Status: mainFile ? 'พร้อม' : 'ไม่พบ',
        Detail: mainFile?.name || '-',
      },
      {
        Item: 'จำนวนไฟล์รวม',
        Status: totalFiles > 0 ? 'พร้อม' : 'ไม่พบ',
        Detail: `${totalFiles} ไฟล์`,
      },
      {
        Item: 'โรงงานที่ไม่มีไฟล์',
        Status: missingPlants.length > 0 ? 'พบ' : 'ครบ',
        Detail: missingPlants.length > 0 ? missingPlants.join(', ') : '-',
      },
      {
        Item: 'Part ที่ซ้ำหลายโรงงาน',
        Status: duplicateParts.length > 0 ? 'พบ' : 'ไม่พบ',
        Detail: duplicateParts.length > 0 ? `${duplicateParts.length} รายการ` : '-',
      },
    ];

    const duplicateRows = duplicateParts.map((part) => ({
      PartNumber: part,
      Plants: Array.from(partPlantMap[part]).join(', '),
    }));

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, buildSheet(summaryRows), 'Summary');
    XLSX.utils.book_append_sheet(workbook, buildSheet(verificationRows), 'Verification');
    XLSX.utils.book_append_sheet(workbook, buildSheet(duplicateRows), 'DuplicateParts');

    Object.entries(mockData).forEach(([sheet, rows]) => {
      const sheetRows = rows.map((row) => ({
        PartNumber: row.part,
        PartCode: row.code,
        PackingSize: row.pack,
        N: row.n,
        N1: row.n1,
        N2: row.n2,
        N3: row.n3,
      }));
      XLSX.utils.book_append_sheet(workbook, buildSheet(sheetRows), sheet.replace(' ', '_'));
    });

    XLSX.writeFile(workbook, 'production-consolidated.xlsx');
  };

  return (
    <div className="min-h-screen bg-slate-100 text-slate-900">
      <header className="bg-white border-b border-slate-200">
        <div className="max-w-6xl mx-auto px-4 py-4 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 p-2 bg-blue-600 rounded-lg text-white">
              <Icons.File />
            </div>
            <div>
              <h1 className="text-lg font-semibold">Production Data Consolidator</h1>
              <p className="text-sm text-slate-500">TMT Camera Production Plan</p>
            </div>
          </div>
          <button
            onClick={resetAll}
            className="px-4 py-2 text-sm font-medium text-slate-600 border border-slate-200 rounded-lg hover:bg-slate-50"
          >
            ล้างข้อมูล
          </button>
        </div>
      </header>

      <div className="max-w-6xl mx-auto px-4 py-4">
        <div className="flex flex-wrap gap-2">
          {[
            { id: 'upload', label: 'อัปโหลดไฟล์', Icon: Icons.Upload },
            { id: 'preview', label: 'ตรวจสอบข้อมูล', Icon: Icons.Eye },
            { id: 'summary', label: 'สรุปผลรวม', Icon: Icons.Chart },
          ].map((item) => (
            <button
              key={item.id}
              onClick={() => setTab(item.id)}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium border ${
                tab === item.id
                  ? 'bg-blue-600 text-white border-blue-600'
                  : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'
              }`}
            >
              <span className="w-4 h-4">
                <item.Icon />
              </span>
              {item.label}
            </button>
          ))}
        </div>
      </div>

      <main className="max-w-6xl mx-auto px-4 pb-10 space-y-6">
        {tab === 'upload' && (
          <>
            <section className="bg-white border border-slate-200 rounded-2xl p-5">
              <h2 className="text-base font-semibold mb-3 flex items-center gap-2">
                <span className="w-5 h-5 text-blue-600">
                  <Icons.File />
                </span>
                ไฟล์หลัก (Template)
              </h2>
              <label className="block">
                <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleMainFileChange} />
                <div
                  className={`border-2 border-dashed rounded-xl p-6 text-center transition ${
                    mainFile ? 'border-blue-300 bg-blue-50' : 'border-slate-200 hover:bg-slate-50'
                  }`}
                >
                  {mainFile ? (
                    <div className="flex items-center justify-center gap-3">
                      <div className="w-8 h-8 text-blue-600">
                        <Icons.Check />
                      </div>
                      <div className="text-left">
                        <p className="font-medium">{mainFile.name}</p>
                        <p className="text-sm text-slate-500">{formatSize(mainFile.size)}</p>
                      </div>
                    </div>
                  ) : (
                    <div>
                      <div className="w-10 h-10 mx-auto text-slate-400 mb-2">
                        <Icons.Upload />
                      </div>
                      <p className="text-slate-600">คลิกเพื่อเลือกไฟล์หลัก (.xlsx หรือ .xls)</p>
                    </div>
                  )}
                </div>
              </label>
            </section>

            <section className="bg-white border border-slate-200 rounded-2xl p-5">
              <h2 className="text-base font-semibold mb-3">ไฟล์ข้อมูลรายโรงงาน</h2>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {Object.entries(plantMeta).map(([plant, meta]) => (
                  <div key={plant} className={`border rounded-xl p-4 ${meta.border}`}>
                    <div className="flex items-center justify-between mb-3">
                      <div>
                        <p className="text-sm font-semibold">{plant}</p>
                        <p className="text-xs text-slate-500">{meta.label}</p>
                      </div>
                      <span className={`px-2 py-1 text-xs rounded-full ${meta.badge}`}>
                        {sourceFiles[plant].length} ไฟล์
                      </span>
                    </div>
                    <label className="block">
                      <input
                        type="file"
                        multiple
                        accept=".xlsx,.xls"
                        className="hidden"
                        onChange={(event) => handlePlantFiles(plant, event)}
                      />
                      <div className="border border-slate-200 rounded-lg px-3 py-2 text-center text-sm text-slate-600 hover:bg-slate-50">
                        เลือกไฟล์ {plant}
                      </div>
                    </label>

                    {sourceFiles[plant].length > 0 && (
                      <div className="mt-3 space-y-2 max-h-36 overflow-y-auto">
                        {sourceFiles[plant].map((file, index) => (
                          <div
                            key={`${file.name}-${index}`}
                            className="flex items-center justify-between bg-slate-50 border border-slate-200 rounded-lg px-3 py-2"
                          >
                            <div className="text-sm">
                              <p className="font-medium text-slate-700 truncate">{file.name}</p>
                              <p className="text-xs text-slate-400">{formatSize(file.size)}</p>
                            </div>
                            <button
                              onClick={() => removePlantFile(plant, index)}
                              className="w-4 h-4 text-slate-400 hover:text-red-500"
                            >
                              <Icons.Trash />
                            </button>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </section>

            <button
              onClick={process}
              disabled={!canProcess}
              className={`w-full py-4 rounded-xl font-semibold flex items-center justify-center gap-3 ${
                canProcess
                  ? 'bg-blue-600 text-white hover:bg-blue-700'
                  : 'bg-slate-200 text-slate-400 cursor-not-allowed'
              }`}
            >
              {processing ? (
                <>
                  <span className="w-5 h-5">
                    <Icons.Refresh />
                  </span>
                  กำลังประมวลผล...
                </>
              ) : (
                <>
                  <span className="w-5 h-5">
                    <Icons.Play />
                  </span>
                  รวมข้อมูลและคำนวณ
                </>
              )}
            </button>
          </>
        )}

        {tab === 'preview' && processed && (
          <div className="space-y-4">
            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <div className="flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
                <div>
                  <h3 className="text-base font-semibold">ตรวจสอบข้อมูลเบื้องต้น</h3>
                  <p className="text-sm text-slate-500">เช็คไฟล์ที่อัปโหลดและข้อมูลซ้ำก่อนสรุปผล</p>
                </div>
                <button
                  onClick={handleDownload}
                  className="px-4 py-2 text-sm font-medium bg-blue-600 text-white rounded-lg hover:bg-blue-700"
                >
                  ดาวน์โหลดไฟล์สรุป
                </button>
              </div>
              <div className="mt-4 grid grid-cols-1 md:grid-cols-3 gap-3 text-sm">
                <div className="border border-slate-200 rounded-lg p-3">
                  <p className="text-slate-500">ไฟล์หลัก</p>
                  <p className={`font-semibold ${mainFile ? 'text-emerald-600' : 'text-red-600'}`}>
                    {mainFile ? 'พร้อม' : 'ไม่พบ'}
                  </p>
                  <p className="text-xs text-slate-400">{mainFile?.name || '-'}</p>
                </div>
                <div className="border border-slate-200 rounded-lg p-3">
                  <p className="text-slate-500">ไฟล์รวมทั้งหมด</p>
                  <p className={`font-semibold ${totalFiles > 0 ? 'text-emerald-600' : 'text-red-600'}`}>
                    {totalFiles > 0 ? `${totalFiles} ไฟล์` : 'ไม่พบไฟล์'}
                  </p>
                  <p className="text-xs text-slate-400">ครอบคลุม {plantList.length - missingPlants.length} โรงงาน</p>
                </div>
                <div className="border border-slate-200 rounded-lg p-3">
                  <p className="text-slate-500">ข้อมูลซ้ำหลายโรงงาน</p>
                  <p className={`font-semibold ${duplicateParts.length > 0 ? 'text-amber-600' : 'text-emerald-600'}`}>
                    {duplicateParts.length > 0 ? `${duplicateParts.length} รายการ` : 'ไม่พบ'}
                  </p>
                  <p className="text-xs text-slate-400">
                    {missingPlants.length > 0 ? `ยังขาด: ${missingPlants.join(', ')}` : 'ครบทุกโรงงาน'}
                  </p>
                </div>
              </div>
            </div>

            {Object.entries(mockData).map(([sheet, rows]) => (
              <div key={sheet} className="bg-white border border-slate-200 rounded-xl overflow-hidden">
                <button
                  onClick={() => setExpanded((prev) => ({ ...prev, [sheet]: !prev[sheet] }))}
                  className="w-full px-4 py-3 flex items-center justify-between hover:bg-slate-50"
                >
                  <div className="flex items-center gap-2">
                    <span className={`px-2 py-0.5 rounded text-xs ${plantMeta[sheet.split(' ')[0]].badge}`}>
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
                        <tr className="text-slate-500 border-b border-slate-200">
                          <th className="text-left py-2 px-2">Part Number</th>
                          <th className="text-right py-2 px-2">N</th>
                          <th className="text-right py-2 px-2">N+1</th>
                          <th className="text-right py-2 px-2">N+2</th>
                          <th className="text-right py-2 px-2">N+3</th>
                        </tr>
                      </thead>
                      <tbody>
                        {rows.map((row, index) => (
                          <tr key={index} className="border-b border-slate-100">
                            <td className="py-2 px-2 font-mono text-blue-600">{row.part}</td>
                            <td className="py-2 px-2 text-right">{formatNumber(row.n)}</td>
                            <td className="py-2 px-2 text-right">{formatNumber(row.n1)}</td>
                            <td className="py-2 px-2 text-right">{formatNumber(row.n2)}</td>
                            <td className="py-2 px-2 text-right">{formatNumber(row.n3)}</td>
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

        {tab === 'summary' && processed && (
          <div className="space-y-4">
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              {[
                { label: 'N (Feb)', value: totals.n },
                { label: 'N+1 (Mar)', value: totals.n1 },
                { label: 'N+2 (Apr)', value: totals.n2 },
                { label: 'N+3 (May)', value: totals.n3 },
              ].map((card) => (
                <div key={card.label} className="bg-white border border-slate-200 rounded-xl p-4">
                  <p className="text-sm text-slate-500">{card.label}</p>
                  <p className="text-2xl font-semibold">{formatNumber(card.value)}</p>
                </div>
              ))}
            </div>

            <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
              <div className="px-4 py-3 border-b border-slate-200">
                <h3 className="font-semibold flex items-center gap-2">
                  <span className="w-5 h-5 text-blue-600">
                    <Icons.Chart />
                  </span>
                  สรุปยอดรวมตาม Part Number
                </h3>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="bg-slate-50 text-slate-500">
                      <th className="text-left py-3 px-4">Part Number</th>
                      <th className="text-right py-3 px-4">N</th>
                      <th className="text-right py-3 px-4">N+1</th>
                      <th className="text-right py-3 px-4">N+2</th>
                      <th className="text-right py-3 px-4">N+3</th>
                      <th className="text-right py-3 px-4">Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {Object.entries(summary).map(([part, data]) => (
                      <tr key={part} className="border-b border-slate-100">
                        <td className="py-2 px-4 font-mono text-blue-600">{part}</td>
                        <td className="py-2 px-4 text-right">{formatNumber(data.n)}</td>
                        <td className="py-2 px-4 text-right">{formatNumber(data.n1)}</td>
                        <td className="py-2 px-4 text-right">{formatNumber(data.n2)}</td>
                        <td className="py-2 px-4 text-right">{formatNumber(data.n3)}</td>
                        <td className="py-2 px-4 text-right font-semibold">
                          {formatNumber(data.n + data.n1 + data.n2 + data.n3)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot>
                    <tr className="bg-slate-50 font-semibold">
                      <td className="py-3 px-4">Grand Total</td>
                      <td className="py-3 px-4 text-right">{formatNumber(totals.n)}</td>
                      <td className="py-3 px-4 text-right">{formatNumber(totals.n1)}</td>
                      <td className="py-3 px-4 text-right">{formatNumber(totals.n2)}</td>
                      <td className="py-3 px-4 text-right">{formatNumber(totals.n3)}</td>
                      <td className="py-3 px-4 text-right">
                        {formatNumber(totals.n + totals.n1 + totals.n2 + totals.n3)}
                      </td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>

            <button
              onClick={handleDownload}
              className="w-full py-4 rounded-xl font-semibold bg-blue-600 text-white hover:bg-blue-700"
            >
              ดาวน์โหลดไฟล์สรุปรวม (Excel)
            </button>
          </div>
        )}

        {(tab === 'preview' || tab === 'summary') && !processed && (
          <div className="bg-white border border-slate-200 rounded-xl p-6 text-center">
            <p className="text-slate-500 mb-4">กรุณาอัปโหลดไฟล์และประมวลผลก่อน</p>
            <button onClick={() => setTab('upload')} className="px-5 py-2 bg-blue-600 text-white rounded-lg">
              ไปหน้าอัปโหลด
            </button>
          </div>
        )}
      </main>
    </div>
  );
}
