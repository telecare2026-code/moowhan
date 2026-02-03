import React, { useState, useCallback } from 'react';
import { Upload, FileSpreadsheet, CheckCircle, AlertCircle, Download, Trash2, Play, FolderOpen, BarChart3, Table, RefreshCw, ChevronDown, ChevronUp, Eye } from 'lucide-react';

// Mock data for demonstration
const generateMockData = () => ({
  'BP Daily': [
    { partNumber: '86790-0K051-00', partCode: 'I338', packingSize: 25, n: 1891, n1: 1677, n2: 919, n3: 0 },
    { partNumber: '86790-0K080-00', partCode: 'A369', packingSize: 42, n: 299, n1: 67, n2: 0, n3: 0 },
    { partNumber: '86790-0K110-00', partCode: 'A370', packingSize: 18, n: 1201, n1: 958, n2: 547, n3: 1185 },
  ],
  'BPK Daily': [
    { partNumber: '86790-0K051-00', partCode: 'W036', packingSize: 30, n: 5790, n1: 7980, n2: 5490, n3: 3720 },
    { partNumber: '86790-0K110-00', partCode: 'M523', packingSize: 20, n: 6680, n1: 6740, n2: 5120, n3: 4220 },
  ],
  'GW Daily': [
    { partNumber: '86790-BZ271-00', partCode: 'A055', packingSize: 36, n: 2088, n1: 1980, n2: 1800, n3: 1836 },
    { partNumber: '86790-BZ310-00', partCode: 'B800', packingSize: 15, n: 3303, n1: 3128, n2: 3128, n3: 2584 },
  ],
  'SR Daily': [
    { partNumber: '86790-0K051-00', partCode: 'L342', packingSize: 25, n: 1352, n1: 1787, n2: 2037, n3: 0 },
    { partNumber: '86790-0K080-00', partCode: 'G871', packingSize: 42, n: 676, n1: 1170, n2: 1084, n3: 451 },
  ],
});

export default function ExcelConsolidator() {
  const [files, setFiles] = useState([]);
  const [mainFile, setMainFile] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [processed, setProcessed] = useState(false);
  const [activeTab, setActiveTab] = useState('upload');
  const [previewData, setPreviewData] = useState(null);
  const [summaryData, setSummaryData] = useState(null);
  const [expandedSections, setExpandedSections] = useState({});

  const fileCategories = {
    'BP': { color: 'bg-blue-500', label: 'Ban Pho', files: [] },
    'BPK': { color: 'bg-green-500', label: 'Ban Pho Kaeng Khoi', files: [] },
    'GW': { color: 'bg-purple-500', label: 'Gateway', files: [] },
    'SR': { color: 'bg-orange-500', label: 'Samrong', files: [] },
  };

  const categorizeFile = (filename) => {
    if (filename.startsWith('BP_') && !filename.startsWith('BPK')) return 'BP';
    if (filename.startsWith('BPK')) return 'BPK';
    if (filename.startsWith('GW')) return 'GW';
    if (filename.startsWith('SR')) return 'SR';
    return null;
  };

  const handleFileDrop = useCallback((e) => {
    e.preventDefault();
    const droppedFiles = Array.from(e.dataTransfer?.files || e.target.files || []);
    
    droppedFiles.forEach(file => {
      const category = categorizeFile(file.name);
      if (category) {
        setFiles(prev => [...prev, { 
          name: file.name, 
          category, 
          size: file.size,
          status: 'ready',
          file 
        }]);
      } else if (file.name.includes('หลัก') || file.name.includes('main')) {
        setMainFile({ name: file.name, size: file.size, file });
      }
    });
  }, []);

  const removeFile = (index) => {
    setFiles(prev => prev.filter((_, i) => i !== index));
  };

  const processFiles = async () => {
    setProcessing(true);
    
    // Simulate processing
    for (let i = 0; i < files.length; i++) {
      await new Promise(r => setTimeout(r, 500));
      setFiles(prev => prev.map((f, idx) => 
        idx === i ? { ...f, status: 'processing' } : f
      ));
      await new Promise(r => setTimeout(r, 300));
      setFiles(prev => prev.map((f, idx) => 
        idx === i ? { ...f, status: 'done' } : f
      ));
    }

    // Generate mock preview data
    const mockData = generateMockData();
    setPreviewData(mockData);

    // Calculate summary
    const summary = {};
    Object.entries(mockData).forEach(([sheet, rows]) => {
      rows.forEach(row => {
        if (!summary[row.partNumber]) {
          summary[row.partNumber] = { n: 0, n1: 0, n2: 0, n3: 0, plants: new Set() };
        }
        summary[row.partNumber].n += row.n;
        summary[row.partNumber].n1 += row.n1;
        summary[row.partNumber].n2 += row.n2;
        summary[row.partNumber].n3 += row.n3;
        summary[row.partNumber].plants.add(sheet.split(' ')[0]);
      });
    });

    setSummaryData(Object.entries(summary).map(([part, data]) => ({
      partNumber: part,
      plants: Array.from(data.plants).join(', '),
      ...data
    })));

    setProcessing(false);
    setProcessed(true);
    setActiveTab('preview');
  };

  const toggleSection = (section) => {
    setExpandedSections(prev => ({ ...prev, [section]: !prev[section] }));
  };

  const formatNumber = (num) => num?.toLocaleString() || '0';

  const getStatusIcon = (status) => {
    switch (status) {
      case 'done': return <CheckCircle className="w-5 h-5 text-green-500" />;
      case 'processing': return <RefreshCw className="w-5 h-5 text-blue-500 animate-spin" />;
      default: return <FileSpreadsheet className="w-5 h-5 text-gray-400" />;
    }
  };

  const getCategoryColor = (category) => {
    const colors = {
      'BP': 'bg-blue-100 text-blue-700 border-blue-300',
      'BPK': 'bg-green-100 text-green-700 border-green-300',
      'GW': 'bg-purple-100 text-purple-700 border-purple-300',
      'SR': 'bg-orange-100 text-orange-700 border-orange-300',
    };
    return colors[category] || 'bg-gray-100 text-gray-700';
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900">
      {/* Header */}
      <header className="bg-slate-800/50 backdrop-blur-lg border-b border-slate-700 sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="p-2 bg-gradient-to-br from-emerald-400 to-cyan-500 rounded-xl">
                <FileSpreadsheet className="w-6 h-6 text-white" />
              </div>
              <div>
                <h1 className="text-xl font-bold text-white">รีบทำเดียวอดเลนเกม</h1>
                <p className="text-sm text-slate-400">TMT Camera Production Plan Summary</p>
              </div>
            </div>
            <div className="flex items-center gap-2">
              <span className="px-3 py-1 bg-emerald-500/20 text-emerald-400 rounded-full text-sm font-medium">
                v1.0
              </span>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-6">
        {/* Tab Navigation */}
        <div className="flex gap-2 mb-6">
          {[
            { id: 'upload', label: 'อัปโหลดไฟล์', icon: Upload },
            { id: 'preview', label: 'ตรวจสอบข้อมูล', icon: Eye },
            { id: 'summary', label: 'สรุปรวม', icon: BarChart3 },
          ].map(tab => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id)}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg font-medium transition-all ${
                activeTab === tab.id
                  ? 'bg-emerald-500 text-white shadow-lg shadow-emerald-500/25'
                  : 'bg-slate-700/50 text-slate-300 hover:bg-slate-700'
              }`}
            >
              <tab.icon className="w-4 h-4" />
              {tab.label}
            </button>
          ))}
        </div>

        {/* Upload Tab */}
        {activeTab === 'upload' && (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            {/* Main File Upload */}
            <div className="lg:col-span-1">
              <div className="bg-slate-800/50 backdrop-blur rounded-2xl border border-slate-700 p-6">
                <h2 className="text-lg font-semibold text-white mb-4 flex items-center gap-2">
                  <FolderOpen className="w-5 h-5 text-amber-400" />
                  ไฟล์หลัก (Template)
                </h2>
                
                <div
                  onDragOver={(e) => e.preventDefault()}
                  onDrop={(e) => {
                    e.preventDefault();
                    const file = e.dataTransfer.files[0];
                    if (file) setMainFile({ name: file.name, size: file.size, file });
                  }}
                  className={`border-2 border-dashed rounded-xl p-6 text-center transition-all cursor-pointer ${
                    mainFile 
                      ? 'border-emerald-500 bg-emerald-500/10' 
                      : 'border-slate-600 hover:border-slate-500 hover:bg-slate-700/30'
                  }`}
                >
                  {mainFile ? (
                    <div className="flex flex-col items-center gap-2">
                      <CheckCircle className="w-10 h-10 text-emerald-500" />
                      <p className="text-white font-medium text-sm">{mainFile.name}</p>
                      <p className="text-slate-400 text-xs">{(mainFile.size / 1024).toFixed(1)} KB</p>
                      <button 
                        onClick={() => setMainFile(null)}
                        className="text-red-400 text-sm hover:text-red-300 mt-2"
                      >
                        ลบไฟล์
                      </button>
                    </div>
                  ) : (
                    <div className="flex flex-col items-center gap-2">
                      <Upload className="w-10 h-10 text-slate-500" />
                      <p className="text-slate-300">ลากไฟล์มาวางที่นี่</p>
                      <p className="text-slate-500 text-sm">หรือคลิกเพื่อเลือก</p>
                    </div>
                  )}
                  <input
                    type="file"
                    accept=".xlsx,.xls"
                    className="hidden"
                    onChange={(e) => {
                      const file = e.target.files[0];
                      if (file) setMainFile({ name: file.name, size: file.size, file });
                    }}
                  />
                </div>
              </div>
            </div>

            {/* Source Files Upload */}
            <div className="lg:col-span-2">
              <div className="bg-slate-800/50 backdrop-blur rounded-2xl border border-slate-700 p-6">
                <h2 className="text-lg font-semibold text-white mb-4 flex items-center gap-2">
                  <Table className="w-5 h-5 text-cyan-400" />
                  ไฟล์ข้อมูลย่อย
                </h2>

                {/* Drop Zone */}
                <div
                  onDragOver={(e) => e.preventDefault()}
                  onDrop={handleFileDrop}
                  className="border-2 border-dashed border-slate-600 rounded-xl p-8 text-center mb-4 hover:border-cyan-500 hover:bg-cyan-500/5 transition-all cursor-pointer"
                >
                  <Upload className="w-12 h-12 text-slate-500 mx-auto mb-3" />
                  <p className="text-slate-300 mb-1">ลากไฟล์ .xls หรือ .xlsx มาวางที่นี่</p>
                  <p className="text-slate-500 text-sm">รองรับไฟล์: BP_*, BPK_*, GW_*, SR_*</p>
                  <input
                    type="file"
                    multiple
                    accept=".xlsx,.xls"
                    className="hidden"
                    onChange={handleFileDrop}
                  />
                </div>

                {/* Plant Categories */}
                <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
                  {Object.entries(fileCategories).map(([key, { label }]) => {
                    const count = files.filter(f => f.category === key).length;
                    return (
                      <div key={key} className={`rounded-lg p-3 border ${getCategoryColor(key)}`}>
                        <div className="text-xs opacity-75">{label}</div>
                        <div className="text-lg font-bold">{count} ไฟล์</div>
                      </div>
                    );
                  })}
                </div>

                {/* File List */}
                {files.length > 0 && (
                  <div className="space-y-2 max-h-64 overflow-y-auto">
                    {files.map((file, index) => (
                      <div
                        key={index}
                        className="flex items-center justify-between bg-slate-700/50 rounded-lg px-4 py-3"
                      >
                        <div className="flex items-center gap-3">
                          {getStatusIcon(file.status)}
                          <div>
                            <p className="text-white text-sm font-medium">{file.name}</p>
                            <p className="text-slate-400 text-xs">{(file.size / 1024).toFixed(1)} KB</p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2">
                          <span className={`px-2 py-1 rounded text-xs font-medium ${getCategoryColor(file.category)}`}>
                            {file.category}
                          </span>
                          <button
                            onClick={() => removeFile(index)}
                            className="p-1 text-slate-400 hover:text-red-400 transition-colors"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>

            {/* Process Button */}
            <div className="lg:col-span-3">
              <button
                onClick={processFiles}
                disabled={files.length === 0 || processing}
                className={`w-full py-4 rounded-xl font-semibold text-lg flex items-center justify-center gap-3 transition-all ${
                  files.length > 0 && !processing
                    ? 'bg-gradient-to-r from-emerald-500 to-cyan-500 text-white shadow-lg shadow-emerald-500/25 hover:shadow-emerald-500/40'
                    : 'bg-slate-700 text-slate-400 cursor-not-allowed'
                }`}
              >
                {processing ? (
                  <>
                    <RefreshCw className="w-5 h-5 animate-spin" />
                    กำลังประมวลผล...
                  </>
                ) : (
                  <>
                    <Play className="w-5 h-5" />
                    รวมข้อมูลและคำนวณ
                  </>
                )}
              </button>
            </div>
          </div>
        )}

        {/* Preview Tab */}
        {activeTab === 'preview' && previewData && (
          <div className="space-y-4">
            {Object.entries(previewData).map(([sheet, rows]) => (
              <div key={sheet} className="bg-slate-800/50 backdrop-blur rounded-2xl border border-slate-700 overflow-hidden">
                <button
                  onClick={() => toggleSection(sheet)}
                  className="w-full px-6 py-4 flex items-center justify-between hover:bg-slate-700/30 transition-colors"
                >
                  <div className="flex items-center gap-3">
                    <span className={`px-3 py-1 rounded-lg text-sm font-medium ${getCategoryColor(sheet.split(' ')[0])}`}>
                      {sheet.split(' ')[0]}
                    </span>
                    <h3 className="text-white font-semibold">{sheet}</h3>
                    <span className="text-slate-400 text-sm">({rows.length} รายการ)</span>
                  </div>
                  {expandedSections[sheet] ? (
                    <ChevronUp className="w-5 h-5 text-slate-400" />
                  ) : (
                    <ChevronDown className="w-5 h-5 text-slate-400" />
                  )}
                </button>
                
                {expandedSections[sheet] && (
                  <div className="px-6 pb-4">
                    <div className="overflow-x-auto">
                      <table className="w-full">
                        <thead>
                          <tr className="border-b border-slate-700">
                            <th className="text-left py-3 px-4 text-slate-400 font-medium text-sm">Part Number</th>
                            <th className="text-left py-3 px-4 text-slate-400 font-medium text-sm">Part Code</th>
                            <th className="text-right py-3 px-4 text-slate-400 font-medium text-sm">Packing</th>
                            <th className="text-right py-3 px-4 text-slate-400 font-medium text-sm">N (Feb)</th>
                            <th className="text-right py-3 px-4 text-slate-400 font-medium text-sm">N+1 (Mar)</th>
                            <th className="text-right py-3 px-4 text-slate-400 font-medium text-sm">N+2 (Apr)</th>
                            <th className="text-right py-3 px-4 text-slate-400 font-medium text-sm">N+3 (May)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {rows.map((row, idx) => (
                            <tr key={idx} className="border-b border-slate-700/50 hover:bg-slate-700/30">
                              <td className="py-3 px-4 text-white font-mono text-sm">{row.partNumber}</td>
                              <td className="py-3 px-4 text-slate-300 text-sm">{row.partCode}</td>
                              <td className="py-3 px-4 text-slate-300 text-sm text-right">{row.packingSize}</td>
                              <td className="py-3 px-4 text-cyan-400 font-medium text-right">{formatNumber(row.n)}</td>
                              <td className="py-3 px-4 text-emerald-400 font-medium text-right">{formatNumber(row.n1)}</td>
                              <td className="py-3 px-4 text-amber-400 font-medium text-right">{formatNumber(row.n2)}</td>
                              <td className="py-3 px-4 text-purple-400 font-medium text-right">{formatNumber(row.n3)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            ))}
          </div>
        )}

        {/* Summary Tab */}
        {activeTab === 'summary' && summaryData && (
          <div className="space-y-6">
            {/* Summary Cards */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              {[
                { label: 'N (Feb)', value: summaryData.reduce((s, r) => s + r.n, 0), color: 'from-cyan-500 to-blue-500' },
                { label: 'N+1 (Mar)', value: summaryData.reduce((s, r) => s + r.n1, 0), color: 'from-emerald-500 to-green-500' },
                { label: 'N+2 (Apr)', value: summaryData.reduce((s, r) => s + r.n2, 0), color: 'from-amber-500 to-orange-500' },
                { label: 'N+3 (May)', value: summaryData.reduce((s, r) => s + r.n3, 0), color: 'from-purple-500 to-pink-500' },
              ].map((card, idx) => (
                <div key={idx} className={`bg-gradient-to-br ${card.color} rounded-2xl p-6 shadow-lg`}>
                  <p className="text-white/80 text-sm mb-1">{card.label}</p>
                  <p className="text-white text-3xl font-bold">{formatNumber(card.value)}</p>
                  <p className="text-white/60 text-xs mt-1">ชิ้น</p>
                </div>
              ))}
            </div>

            {/* Summary Table */}
            <div className="bg-slate-800/50 backdrop-blur rounded-2xl border border-slate-700 overflow-hidden">
              <div className="px-6 py-4 border-b border-slate-700">
                <h3 className="text-white font-semibold flex items-center gap-2">
                  <BarChart3 className="w-5 h-5 text-emerald-400" />
                  สรุปยอดรวมตาม Part Number
                </h3>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full">
                  <thead>
                    <tr className="bg-slate-700/50">
                      <th className="text-left py-3 px-6 text-slate-300 font-medium">Part Number</th>
                      <th className="text-left py-3 px-6 text-slate-300 font-medium">Plants</th>
                      <th className="text-right py-3 px-6 text-cyan-400 font-medium">N (Feb)</th>
                      <th className="text-right py-3 px-6 text-emerald-400 font-medium">N+1 (Mar)</th>
                      <th className="text-right py-3 px-6 text-amber-400 font-medium">N+2 (Apr)</th>
                      <th className="text-right py-3 px-6 text-purple-400 font-medium">N+3 (May)</th>
                      <th className="text-right py-3 px-6 text-slate-300 font-medium">Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {summaryData.map((row, idx) => (
                      <tr key={idx} className="border-b border-slate-700/50 hover:bg-slate-700/30">
                        <td className="py-3 px-6 text-white font-mono">{row.partNumber}</td>
                        <td className="py-3 px-6">
                          <div className="flex gap-1">
                            {row.plants.split(', ').map(plant => (
                              <span key={plant} className={`px-2 py-0.5 rounded text-xs font-medium ${getCategoryColor(plant)}`}>
                                {plant}
                              </span>
                            ))}
                          </div>
                        </td>
                        <td className="py-3 px-6 text-cyan-400 font-medium text-right">{formatNumber(row.n)}</td>
                        <td className="py-3 px-6 text-emerald-400 font-medium text-right">{formatNumber(row.n1)}</td>
                        <td className="py-3 px-6 text-amber-400 font-medium text-right">{formatNumber(row.n2)}</td>
                        <td className="py-3 px-6 text-purple-400 font-medium text-right">{formatNumber(row.n3)}</td>
                        <td className="py-3 px-6 text-white font-bold text-right">
                          {formatNumber(row.n + row.n1 + row.n2 + row.n3)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot>
                    <tr className="bg-slate-700/50 font-bold">
                      <td className="py-4 px-6 text-white">Grand Total</td>
                      <td className="py-4 px-6 text-slate-400">{summaryData.length} Parts</td>
                      <td className="py-4 px-6 text-cyan-400 text-right">{formatNumber(summaryData.reduce((s, r) => s + r.n, 0))}</td>
                      <td className="py-4 px-6 text-emerald-400 text-right">{formatNumber(summaryData.reduce((s, r) => s + r.n1, 0))}</td>
                      <td className="py-4 px-6 text-amber-400 text-right">{formatNumber(summaryData.reduce((s, r) => s + r.n2, 0))}</td>
                      <td className="py-4 px-6 text-purple-400 text-right">{formatNumber(summaryData.reduce((s, r) => s + r.n3, 0))}</td>
                      <td className="py-4 px-6 text-white text-right">
                        {formatNumber(summaryData.reduce((s, r) => s + r.n + r.n1 + r.n2 + r.n3, 0))}
                      </td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>

            {/* Download Button */}
            <button className="w-full py-4 bg-gradient-to-r from-emerald-500 to-cyan-500 rounded-xl font-semibold text-lg text-white flex items-center justify-center gap-3 shadow-lg shadow-emerald-500/25 hover:shadow-emerald-500/40 transition-all">
              <Download className="w-5 h-5" />
              ดาวน์โหลดไฟล์ Excel สรุป
            </button>
          </div>
        )}

        {/* Empty State */}
        {activeTab !== 'upload' && !previewData && (
          <div className="text-center py-20">
            <AlertCircle className="w-16 h-16 text-slate-600 mx-auto mb-4" />
            <h3 className="text-xl text-slate-400 mb-2">ยังไม่มีข้อมูล</h3>
            <p className="text-slate-500">กรุณาอัปโหลดไฟล์และประมวลผลก่อน</p>
            <button
              onClick={() => setActiveTab('upload')}
              className="mt-4 px-6 py-2 bg-slate-700 text-white rounded-lg hover:bg-slate-600 transition-colors"
            >
              ไปหน้าอัปโหลด
            </button>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="border-t border-slate-700 mt-12 py-6">
        <div className="max-w-7xl mx-auto px-4 text-center text-slate-500 text-sm">
          ชอน • TMT Camera Production Plan Summary System
        </div>
      </footer>
    </div>
  );
}
