import React, { useState, useCallback } from 'react';
import { Upload, FileSpreadsheet, CheckCircle, AlertCircle, Download, Trash2, Play, BarChart3, Building2, Package, TrendingUp, RefreshCw } from 'lucide-react';

// Main Application Component
export default function ExcelConsolidationTool() {
  const [mainFile, setMainFile] = useState(null);
  const [sourceFiles, setSourceFiles] = useState({
    BP: [],
    BPK: [],
    GW: [],
    SR: []
  });
  const [processing, setProcessing] = useState(false);
  const [results, setResults] = useState(null);
  const [activeTab, setActiveTab] = useState('upload');
  const [error, setError] = useState(null);

  // Plant configurations
  const plantConfig = {
    BP: { name: 'Ban Pho', color: 'blue', icon: Building2, description: 'Vehicle Assembly' },
    BPK: { name: 'Ban Pho Kaeng Khoi', color: 'green', icon: Package, description: 'Packing & Export' },
    GW: { name: 'Gateway', color: 'purple', icon: TrendingUp, description: 'Vehicle & Packing' },
    SR: { name: 'Samrong', color: 'orange', icon: Building2, description: 'Vehicle Assembly' }
  };

  // Handle main file upload
  const handleMainFileUpload = (e) => {
    const file = e.target.files[0];
    if (file) {
      setMainFile({
        name: file.name,
        size: (file.size / 1024).toFixed(1) + ' KB',
        file: file
      });
      setError(null);
    }
  };

  // Handle source files upload
  const handleSourceFilesUpload = (plant, e) => {
    const files = Array.from(e.target.files);
    const newFiles = files.map(file => ({
      name: file.name,
      size: (file.size / 1024).toFixed(1) + ' KB',
      file: file
    }));
    
    setSourceFiles(prev => ({
      ...prev,
      [plant]: [...prev[plant], ...newFiles]
    }));
    setError(null);
  };

  // Remove file
  const removeSourceFile = (plant, index) => {
    setSourceFiles(prev => ({
      ...prev,
      [plant]: prev[plant].filter((_, i) => i !== index)
    }));
  };

  // Process files (simulation)
  const processFiles = async () => {
    setProcessing(true);
    setError(null);
    
    // Simulate processing
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    // Generate mock results based on uploaded files
    const totalSourceFiles = Object.values(sourceFiles).flat().length;
    
    if (!mainFile || totalSourceFiles === 0) {
      setError('กรุณาอัปโหลดไฟล์หลักและไฟล์ข้อมูลอย่างน้อย 1 ไฟล์');
      setProcessing(false);
      return;
    }

    // Mock results
    const mockResults = {
      summary: {
        totalParts: 20,
        totalRecords: 42,
        grandTotal: {
          N: 67252,
          N1: 79486,
          N2: 66744,
          N3: 60180
        }
      },
      byPlant: {
        BP: { records: 12, total: 9079 },
        BPK: { records: 16, total: 30850 },
        GW: { records: 8, total: 24902 },
        SR: { records: 6, total: 2421 }
      },
      pivotData: [
        { part: '86790-0K051-00', n: 9673, n1: 12094, n2: 8906, n3: 3950, plants: 'BP, BPK, SR' },
        { part: '86790-0K110-00', n: 8051, n1: 8078, n2: 5797, n3: 5615, plants: 'BP, BPK' },
        { part: '86790-BZ310-00', n: 12003, n1: 11815, n2: 12076, n3: 10271, plants: 'BPK, GW' },
        { part: '86790-BZ280-00', n: 8210, n1: 5430, n2: 5520, n3: 5170, plants: 'GW' },
        { part: '86790-BZ210-00', n: 5651, n1: 7044, n2: 5224, n3: 7912, plants: 'GW' },
        { part: '86790-0K190-00', n: 4745, n1: 5243, n2: 3468, n3: 3284, plants: 'BP, BPK' },
        { part: '86790-0K200-00', n: 3020, n1: 2820, n2: 2880, n3: 2620, plants: 'BPK' },
        { part: '86790-0K141-00', n: 1946, n1: 2404, n2: 1451, n3: 912, plants: 'BP, BPK, SR' },
      ],
      sheets: ['BP Daily', 'BPK Daily', 'GW Daily', 'SR Daily', 'Total', 'Sheet2', 'Analyze']
    };
    
    setResults(mockResults);
    setActiveTab('results');
    setProcessing(false);
  };

  // Reset all
  const resetAll = () => {
    setMainFile(null);
    setSourceFiles({ BP: [], BPK: [], GW: [], SR: [] });
    setResults(null);
    setActiveTab('upload');
    setError(null);
  };

  // File upload card component
  const FileUploadCard = ({ plant }) => {
    const config = plantConfig[plant];
    const Icon = config.icon;
    const files = sourceFiles[plant];
    const colorClasses = {
      blue: 'border-blue-300 bg-blue-50 hover:bg-blue-100',
      green: 'border-green-300 bg-green-50 hover:bg-green-100',
      purple: 'border-purple-300 bg-purple-50 hover:bg-purple-100',
      orange: 'border-orange-300 bg-orange-50 hover:bg-orange-100'
    };
    const badgeClasses = {
      blue: 'bg-blue-500',
      green: 'bg-green-500',
      purple: 'bg-purple-500',
      orange: 'bg-orange-500'
    };

    return (
      <div className={`border-2 border-dashed rounded-xl p-4 transition-all ${colorClasses[config.color]}`}>
        <div className="flex items-center gap-3 mb-3">
          <div className={`p-2 rounded-lg ${badgeClasses[config.color]} text-white`}>
            <Icon size={20} />
          </div>
          <div>
            <h3 className="font-semibold text-gray-800">{plant} - {config.name}</h3>
            <p className="text-xs text-gray-500">{config.description}</p>
          </div>
          {files.length > 0 && (
            <span className={`ml-auto px-2 py-1 rounded-full text-xs text-white ${badgeClasses[config.color]}`}>
              {files.length} ไฟล์
            </span>
          )}
        </div>
        
        <label className="block cursor-pointer">
          <input
            type="file"
            multiple
            accept=".xls,.xlsx"
            onChange={(e) => handleSourceFilesUpload(plant, e)}
            className="hidden"
          />
          <div className="border border-gray-300 rounded-lg p-3 text-center hover:border-gray-400 bg-white">
            <Upload size={20} className="mx-auto text-gray-400 mb-1" />
            <span className="text-sm text-gray-600">เลือกไฟล์ {plant}</span>
          </div>
        </label>

        {files.length > 0 && (
          <div className="mt-3 space-y-2 max-h-32 overflow-y-auto">
            {files.map((file, index) => (
              <div key={index} className="flex items-center justify-between bg-white rounded-lg px-3 py-2 text-sm">
                <div className="flex items-center gap-2 truncate">
                  <FileSpreadsheet size={14} className="text-green-600 flex-shrink-0" />
                  <span className="truncate">{file.name}</span>
                </div>
                <button
                  onClick={() => removeSourceFile(plant, index)}
                  className="text-red-400 hover:text-red-600 flex-shrink-0 ml-2"
                >
                  <Trash2 size={14} />
                </button>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  };

  // Stats card component
  const StatsCard = ({ title, value, subtitle, icon: Icon, color }) => {
    const colorClasses = {
      blue: 'bg-blue-500',
      green: 'bg-green-500',
      purple: 'bg-purple-500',
      orange: 'bg-orange-500'
    };
    
    return (
      <div className="bg-white rounded-xl p-4 shadow-sm border">
        <div className="flex items-center gap-3">
          <div className={`p-3 rounded-xl ${colorClasses[color]} text-white`}>
            <Icon size={24} />
          </div>
          <div>
            <p className="text-2xl font-bold text-gray-800">{value.toLocaleString()}</p>
            <p className="text-sm text-gray-500">{title}</p>
            {subtitle && <p className="text-xs text-gray-400">{subtitle}</p>}
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50">
      {/* Header */}
      <header className="bg-white shadow-sm border-b">
        <div className="max-w-6xl mx-auto px-4 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="p-2 bg-gradient-to-br from-blue-500 to-purple-600 rounded-xl text-white">
                <BarChart3 size={28} />
              </div>
              <div>
                <h1 className="text-xl font-bold text-gray-800">Production Forecast Consolidator</h1>
                <p className="text-sm text-gray-500">TMT Camera Production Planning Tool</p>
              </div>
            </div>
            <button
              onClick={resetAll}
              className="flex items-center gap-2 px-4 py-2 text-gray-600 hover:text-gray-800 hover:bg-gray-100 rounded-lg transition-colors"
            >
              <RefreshCw size={18} />
              รีเซ็ต
            </button>
          </div>
        </div>
      </header>

      {/* Tabs */}
      <div className="max-w-6xl mx-auto px-4 py-4">
        <div className="flex gap-2 bg-white rounded-xl p-1 shadow-sm w-fit">
          <button
            onClick={() => setActiveTab('upload')}
            className={`px-6 py-2 rounded-lg font-medium transition-all ${
              activeTab === 'upload'
                ? 'bg-blue-500 text-white shadow-md'
                : 'text-gray-600 hover:bg-gray-100'
            }`}
          >
            📁 อัปโหลดไฟล์
          </button>
          <button
            onClick={() => setActiveTab('results')}
            disabled={!results}
            className={`px-6 py-2 rounded-lg font-medium transition-all ${
              activeTab === 'results'
                ? 'bg-blue-500 text-white shadow-md'
                : results
                ? 'text-gray-600 hover:bg-gray-100'
                : 'text-gray-300 cursor-not-allowed'
            }`}
          >
            📊 ผลลัพธ์
          </button>
        </div>
      </div>

      {/* Main Content */}
      <main className="max-w-6xl mx-auto px-4 pb-8">
        {activeTab === 'upload' && (
          <div className="space-y-6">
            {/* Error Alert */}
            {error && (
              <div className="bg-red-50 border border-red-200 rounded-xl p-4 flex items-center gap-3">
                <AlertCircle className="text-red-500" size={20} />
                <span className="text-red-700">{error}</span>
              </div>
            )}

            {/* Main File Upload */}
            <div className="bg-white rounded-2xl shadow-sm border p-6">
              <h2 className="text-lg font-semibold text-gray-800 mb-4 flex items-center gap-2">
                <FileSpreadsheet className="text-blue-500" />
                ไฟล์หลัก (Template)
              </h2>
              
              <label className="block cursor-pointer">
                <input
                  type="file"
                  accept=".xlsx"
                  onChange={handleMainFileUpload}
                  className="hidden"
                />
                <div className={`border-2 border-dashed rounded-xl p-8 text-center transition-all ${
                  mainFile 
                    ? 'border-green-300 bg-green-50' 
                    : 'border-gray-300 hover:border-blue-400 hover:bg-blue-50'
                }`}>
                  {mainFile ? (
                    <div className="flex items-center justify-center gap-3">
                      <CheckCircle className="text-green-500" size={32} />
                      <div className="text-left">
                        <p className="font-medium text-gray-800">{mainFile.name}</p>
                        <p className="text-sm text-gray-500">{mainFile.size}</p>
                      </div>
                    </div>
                  ) : (
                    <>
                      <Upload size={40} className="mx-auto text-gray-400 mb-3" />
                      <p className="text-gray-600 font-medium">คลิกเพื่อเลือกไฟล์หลัก (.xlsx)</p>
                      <p className="text-sm text-gray-400 mt-1">ไฟล์ที่มีชีท Summary, Total, BP Daily, BPK Daily, GW Daily, SR Daily</p>
                    </>
                  )}
                </div>
              </label>
            </div>

            {/* Source Files Upload */}
            <div className="bg-white rounded-2xl shadow-sm border p-6">
              <h2 className="text-lg font-semibold text-gray-800 mb-4 flex items-center gap-2">
                <Package className="text-green-500" />
                ไฟล์ข้อมูลรายโรงงาน
              </h2>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {Object.keys(plantConfig).map(plant => (
                  <FileUploadCard key={plant} plant={plant} />
                ))}
              </div>
            </div>

            {/* Process Button */}
            <div className="flex justify-center">
              <button
                onClick={processFiles}
                disabled={processing || !mainFile}
                className={`flex items-center gap-3 px-8 py-4 rounded-xl font-semibold text-lg shadow-lg transition-all ${
                  processing || !mainFile
                    ? 'bg-gray-300 cursor-not-allowed text-gray-500'
                    : 'bg-gradient-to-r from-blue-500 to-purple-600 hover:from-blue-600 hover:to-purple-700 text-white hover:shadow-xl hover:scale-105'
                }`}
              >
                {processing ? (
                  <>
                    <div className="w-6 h-6 border-3 border-white border-t-transparent rounded-full animate-spin" />
                    กำลังประมวลผล...
                  </>
                ) : (
                  <>
                    <Play size={24} />
                    ประมวลผลและรวมข้อมูล
                  </>
                )}
              </button>
            </div>
          </div>
        )}

        {activeTab === 'results' && results && (
          <div className="space-y-6">
            {/* Summary Stats */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <StatsCard
                title="ยอดรวมเดือน N"
                value={results.summary.grandTotal.N}
                subtitle="Feb 2026"
                icon={BarChart3}
                color="blue"
              />
              <StatsCard
                title="ยอดรวมเดือน N+1"
                value={results.summary.grandTotal.N1}
                subtitle="Mar 2026"
                icon={TrendingUp}
                color="green"
              />
              <StatsCard
                title="ยอดรวมเดือน N+2"
                value={results.summary.grandTotal.N2}
                subtitle="Apr 2026"
                icon={TrendingUp}
                color="purple"
              />
              <StatsCard
                title="ยอดรวมเดือน N+3"
                value={results.summary.grandTotal.N3}
                subtitle="May 2026"
                icon={TrendingUp}
                color="orange"
              />
            </div>

            {/* Plant Summary */}
            <div className="bg-white rounded-2xl shadow-sm border p-6">
              <h2 className="text-lg font-semibold text-gray-800 mb-4">📊 สรุปรายโรงงาน</h2>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                {Object.entries(results.byPlant).map(([plant, data]) => {
                  const config = plantConfig[plant];
                  const colorClasses = {
                    blue: 'bg-blue-100 border-blue-200',
                    green: 'bg-green-100 border-green-200',
                    purple: 'bg-purple-100 border-purple-200',
                    orange: 'bg-orange-100 border-orange-200'
                  };
                  return (
                    <div key={plant} className={`rounded-xl p-4 border ${colorClasses[config.color]}`}>
                      <div className="font-semibold text-gray-800">{plant}</div>
                      <div className="text-sm text-gray-500">{config.name}</div>
                      <div className="mt-2 text-2xl font-bold">{data.records}</div>
                      <div className="text-xs text-gray-500">รายการ</div>
                      <div className="mt-1 text-lg font-semibold text-gray-700">{data.total.toLocaleString()}</div>
                      <div className="text-xs text-gray-500">ยอดรวม N</div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Pivot Table */}
            <div className="bg-white rounded-2xl shadow-sm border p-6">
              <h2 className="text-lg font-semibold text-gray-800 mb-4">📋 Pivot Summary (Sheet2)</h2>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="bg-gray-50">
                      <th className="px-4 py-3 text-left font-semibold text-gray-700 rounded-l-lg">Part Number</th>
                      <th className="px-4 py-3 text-right font-semibold text-gray-700">Sum of N</th>
                      <th className="px-4 py-3 text-right font-semibold text-gray-700">Sum of N+1</th>
                      <th className="px-4 py-3 text-right font-semibold text-gray-700">Sum of N+2</th>
                      <th className="px-4 py-3 text-right font-semibold text-gray-700">Sum of N+3</th>
                      <th className="px-4 py-3 text-left font-semibold text-gray-700 rounded-r-lg">Plants</th>
                    </tr>
                  </thead>
                  <tbody>
                    {results.pivotData.map((row, index) => (
                      <tr key={index} className="border-b hover:bg-gray-50">
                        <td className="px-4 py-3 font-mono text-blue-600">{row.part}</td>
                        <td className="px-4 py-3 text-right font-medium">{row.n.toLocaleString()}</td>
                        <td className="px-4 py-3 text-right">{row.n1.toLocaleString()}</td>
                        <td className="px-4 py-3 text-right">{row.n2.toLocaleString()}</td>
                        <td className="px-4 py-3 text-right">{row.n3.toLocaleString()}</td>
                        <td className="px-4 py-3">
                          <div className="flex flex-wrap gap-1">
                            {row.plants.split(', ').map(p => (
                              <span key={p} className="px-2 py-0.5 bg-gray-100 rounded text-xs">{p}</span>
                            ))}
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Output Sheets */}
            <div className="bg-white rounded-2xl shadow-sm border p-6">
              <h2 className="text-lg font-semibold text-gray-800 mb-4">📑 ชีทที่อัปเดต</h2>
              <div className="flex flex-wrap gap-2">
                {results.sheets.map(sheet => (
                  <span key={sheet} className="px-4 py-2 bg-green-100 text-green-700 rounded-lg flex items-center gap-2">
                    <CheckCircle size={16} />
                    {sheet}
                  </span>
                ))}
              </div>
            </div>

            {/* Download Button */}
            <div className="flex justify-center gap-4">
              <button className="flex items-center gap-2 px-6 py-3 bg-gradient-to-r from-green-500 to-emerald-600 hover:from-green-600 hover:to-emerald-700 text-white rounded-xl font-semibold shadow-lg hover:shadow-xl transition-all">
                <Download size={20} />
                ดาวน์โหลดไฟล์ Excel
              </button>
            </div>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="bg-white border-t mt-8">
        <div className="max-w-6xl mx-auto px-4 py-4 text-center text-sm text-gray-500">
          Production Forecast Consolidator v1.0 | TMT Camera Production Planning
        </div>
      </footer>
    </div>
  );
}
