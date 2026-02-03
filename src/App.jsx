import React, { useState, useCallback, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';

// ==================== ICONS ====================
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
    <svg className="w-full h-full" fill="currentColor" viewBox="0 0 24 24">
      <polygon points="5 3 19 12 5 21 5 3" />
    </svg>
  ),
  Download: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3" />
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
  Alert: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <circle cx="12" cy="12" r="10" strokeWidth="2" />
      <path strokeWidth="2" d="M12 8v4M12 16h.01" />
    </svg>
  ),
  Reset: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M3 12a9 9 0 109-9 9.75 9.75 0 00-6.74 2.74L3 8" />
      <path strokeWidth="2" d="M3 3v5h5" />
    </svg>
  ),
  Search: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <circle cx="11" cy="11" r="8" strokeWidth="2" />
      <path strokeWidth="2" d="M21 21l-4.35-4.35" />
    </svg>
  ),
  Link: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M10 13a5 5 0 007.54.54l3-3a5 5 0 00-7.07-7.07l-1.72 1.71" />
      <path strokeWidth="2" d="M14 11a5 5 0 00-7.54-.54l-3 3a5 5 0 007.07 7.07l1.71-1.71" />
    </svg>
  ),
  History: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <circle cx="12" cy="12" r="10" strokeWidth="2" />
      <path strokeWidth="2" d="M12 6v6l4 2" />
    </svg>
  ),
  Compare: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" />
      <path strokeWidth="2" d="M9 12l2 2 4-4" />
    </svg>
  ),
  Filter: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <polygon strokeWidth="2" points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3" />
    </svg>
  ),
  Info: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <circle cx="12" cy="12" r="10" strokeWidth="2" />
      <path strokeWidth="2" d="M12 16v-4M12 8h.01" />
    </svg>
  ),
  Comment: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M21 11.5a8.38 8.38 0 01-.9 3.8 8.5 8.5 0 01-7.6 4.7 8.38 8.38 0 01-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 01-.9-3.8 8.5 8.5 0 014.7-7.6 8.38 8.38 0 013.8-.9h.5a8.48 8.48 0 018 8v.5z" />
    </svg>
  ),
  Share: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <circle cx="18" cy="5" r="3" strokeWidth="2" />
      <circle cx="6" cy="12" r="3" strokeWidth="2" />
      <circle cx="18" cy="19" r="3" strokeWidth="2" />
      <path strokeWidth="2" d="M8.59 13.51l6.83 3.98M8.59 10.49l6.83-3.98" />
    </svg>
  ),
  Template: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <rect x="3" y="3" width="18" height="18" rx="2" ry="2" strokeWidth="2" />
      <path strokeWidth="2" d="M3 9h18M9 21V9" />
    </svg>
  ),
  Sort: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M11 5h10M11 12h7M11 19h4M3 5l4 4M7 9L3 5M3 19l4-4M7 15l-4 4" />
    </svg>
  ),
  Function: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M4 7V4h16v3M9 20h6M12 4v16" />
    </svg>
  ),
  PieChart: () => (
    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeWidth="2" d="M21.21 15.89A10 10 0 118 2.83" />
      <path strokeWidth="2" d="M22 12A10 10 0 0012 2v10z" />
    </svg>
  ),
};

// ==================== CONSTANTS ====================
const PLANTS = ['BP', 'BPK', 'GW', 'SR'];

const PLANT_META = {
  BP: { 
    label: 'Ban Pho', 
    badge: 'bg-blue-100 text-blue-700', 
    border: 'border-blue-300',
    gradient: 'from-blue-500 to-blue-600',
    lightBg: 'bg-blue-50',
    iconColor: 'text-blue-600',
    color: '#3B82F6'
  },
  BPK: { 
    label: 'Ban Pho Kaeng Khoi', 
    badge: 'bg-emerald-100 text-emerald-700', 
    border: 'border-emerald-300',
    gradient: 'from-emerald-500 to-emerald-600',
    lightBg: 'bg-emerald-50',
    iconColor: 'text-emerald-600',
    color: '#10B981'
  },
  GW: { 
    label: 'Gateway', 
    badge: 'bg-purple-100 text-purple-700', 
    border: 'border-purple-300',
    gradient: 'from-purple-500 to-purple-600',
    lightBg: 'bg-purple-50',
    iconColor: 'text-purple-600',
    color: '#8B5CF6'
  },
  SR: { 
    label: 'Samrong', 
    badge: 'bg-orange-100 text-orange-700', 
    border: 'border-orange-300',
    gradient: 'from-orange-500 to-orange-600',
    lightBg: 'bg-orange-50',
    iconColor: 'text-orange-600',
    color: '#F59E0B'
  },
};

// ==================== UTILITY FUNCTIONS ====================
const formatNumber = (num) => (num ?? 0).toLocaleString();
const formatSize = (bytes) => `${(bytes / 1024).toFixed(1)} KB`;

const getColumnLetter = (num) => {
  let result = '';
  while (num > 0) {
    num--;
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  return result || 'A';
};

// Read file with xlsx (for source files - no format needed)
const readExcelFileXLSX = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        resolve(workbook);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

// Read file with ExcelJS (for template - needs format preservation)
const readExcelFileExcelJS = async (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const buffer = e.target.result;
        const workbook = new ExcelJS.Workbook();
        try {
          await workbook.xlsx.load(buffer);
        } catch (loadErr) {
          try {
            const data = new Uint8Array(buffer);
            XLSX.read(data, { type: 'array' });
            const workbook2 = new ExcelJS.Workbook();
            await workbook2.xlsx.load(buffer);
            resolve(workbook2);
            return;
          } catch (xlsxErr) {
            throw loadErr;
          }
        }
        if (!workbook || workbook.worksheets.length === 0) {
          throw new Error('ไฟล์ Excel ไม่มี worksheet');
        }
        resolve(workbook);
      } catch (err) {
        const errorMsg = err.message || String(err);
        if (errorMsg.includes('comments') || errorMsg.includes('undefined') || errorMsg.includes('Cannot read')) {
          reject(new Error('ไฟล์ Excel มีโครงสร้างไม่สมบูรณ์ กรุณาลองเปิดและบันทึกไฟล์ใหม่ใน Excel'));
        } else {
          reject(new Error(`ไม่สามารถอ่านไฟล์ Excel: ${errorMsg}`));
        }
      }
    };
    reader.onerror = () => reject(new Error('ไม่สามารถอ่านไฟล์ได้'));
    reader.readAsArrayBuffer(file);
  });
};

// Extract data from xlsx workbook (for source files)
const extractDataFromSourceXLSX = (workbook) => {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const headerRowIndex = 12;
  const dataStartIndex = 13;

  if (jsonData.length <= dataStartIndex) return [];

  const headers = jsonData[headerRowIndex] || [];
  const data = [];

  let nCol = -1, n1Col = -1, n2Col = -1, n3Col = -1;
  headers.forEach((h, i) => {
    const val = String(h || '').trim().toUpperCase();
    if (val === 'N' && nCol === -1) nCol = i;
    else if (val === 'N+1' && n1Col === -1) n1Col = i;
    else if (val === 'N+2' && n2Col === -1) n2Col = i;
    else if (val === 'N+3' && n3Col === -1) n3Col = i;
  });

  if (nCol === -1) nCol = 39;
  if (n1Col === -1) n1Col = 71;
  if (n2Col === -1) n2Col = 103;
  if (n3Col === -1) n3Col = 135;

  for (let i = dataStartIndex; i < jsonData.length; i++) {
    const row = jsonData[i];
    if (!row || !row[0] || row[0] === '<EOF>') continue;

    const partNumber = String(row[0] || '').trim();
    if (!partNumber || partNumber.length < 5) continue;

    const maxCol = Math.max(row.length, 150);
    
    data.push({
      partNumber,
      partCode: row[1] || '',
      partDesc: row[2] || '',
      suppCode: row[3] || '',
      shippingDock: row[4] || '',
      dockCode: row[5] || '',
      carFamily: row[6] || '',
      packingSize: row[7] || 0,
      n: Number(row[nCol]) || 0,
      n1: Number(row[n1Col]) || 0,
      n2: Number(row[n2Col]) || 0,
      n3: Number(row[n3Col]) || 0,
      rawRow: row.slice(0, maxCol),
      colPositions: { nCol, n1Col, n2Col, n3Col },
    });
  }

  return data;
};

const categorizeFile = (filename) => {
  const name = filename.toUpperCase();
  if (name.startsWith('BPK_') || name.startsWith('BPK ') || name.startsWith('BPK-') || name.match(/^BPK[^A-Z]/)) return 'BPK';
  if (name.startsWith('BP_') || name.startsWith('BP ') || name.startsWith('BP-') || name.match(/^BP[^A-Z]/)) return 'BP';
  if (name.startsWith('GW_') || name.startsWith('GW ') || name.startsWith('GW-') || name.match(/^GW[^A-Z]/)) return 'GW';
  if (name.startsWith('SR_') || name.startsWith('SR ') || name.startsWith('SR-') || name.match(/^SR[^A-Z]/)) return 'SR';
  return null;
};

// ==================== CHART COMPONENT ====================
const ChartDashboard = ({ data, summaryData, onClose }) => {
  const [activeChart, setActiveChart] = useState('bar');
  
  const totals = useMemo(() => {
    if (!summaryData) return { n: 0, n1: 0, n2: 0, n3: 0 };
    return summaryData.reduce((acc, row) => ({
      n: acc.n + row.n,
      n1: acc.n1 + row.n1,
      n2: acc.n2 + row.n2,
      n3: acc.n3 + row.n3,
    }), { n: 0, n1: 0, n2: 0, n3: 0 });
  }, [summaryData]);

  const plantData = useMemo(() => {
    if (!data) return [];
    return Object.entries(data).map(([sheet, rows]) => {
      const plant = sheet.split(' ')[0];
      const total = rows.reduce((sum, row) => sum + row.n + row.n1 + row.n2 + row.n3, 0);
      return { plant, total, count: rows.length };
    });
  }, [data]);

  const maxValue = Math.max(totals.n, totals.n1, totals.n2, totals.n3, 1);

  const BarChart = () => (
    <div className="space-y-6">
      <h3 className="text-lg font-semibold text-slate-800">ยอดรวมรายเดือน</h3>
      <div className="space-y-4">
        {[
          { label: 'N (Feb)', value: totals.n, color: 'bg-blue-500', textColor: 'text-blue-600' },
          { label: 'N+1 (Mar)', value: totals.n1, color: 'bg-emerald-500', textColor: 'text-emerald-600' },
          { label: 'N+2 (Apr)', value: totals.n2, color: 'bg-amber-500', textColor: 'text-amber-600' },
          { label: 'N+3 (May)', value: totals.n3, color: 'bg-purple-500', textColor: 'text-purple-600' },
        ].map((item) => (
          <div key={item.label} className="flex items-center gap-4">
            <span className="w-20 text-sm text-slate-600">{item.label}</span>
            <div className="flex-1 bg-slate-100 rounded-full h-8 overflow-hidden">
              <div 
                className={`h-full ${item.color} rounded-full transition-all duration-500 flex items-center justify-end pr-2`}
                style={{ width: `${(item.value / maxValue) * 100}%` }}
              >
                {item.value > maxValue * 0.15 && (
                  <span className="text-white text-sm font-medium">{formatNumber(item.value)}</span>
                )}
              </div>
            </div>
            <span className={`w-24 text-right font-semibold ${item.textColor}`}>{formatNumber(item.value)}</span>
          </div>
        ))}
      </div>
    </div>
  );

  const PieChart = () => (
    <div className="space-y-6">
      <h3 className="text-lg font-semibold text-slate-800">สัดส่วนรายโรงงาน</h3>
      <div className="flex items-center justify-center">
        <svg viewBox="0 0 200 200" className="w-64 h-64">
          {plantData.reduce((acc, item, idx) => {
            const total = plantData.reduce((sum, p) => sum + p.total, 0);
            const startAngle = acc.prevAngle;
            const angle = (item.total / total) * 360;
            const endAngle = startAngle + angle;
            
            const startRad = (startAngle - 90) * Math.PI / 180;
            const endRad = (endAngle - 90) * Math.PI / 180;
            
            const x1 = 100 + 80 * Math.cos(startRad);
            const y1 = 100 + 80 * Math.sin(startRad);
            const x2 = 100 + 80 * Math.cos(endRad);
            const y2 = 100 + 80 * Math.sin(endRad);
            
            const largeArc = angle > 180 ? 1 : 0;
            
            const path = `M 100 100 L ${x1} ${y1} A 80 80 0 ${largeArc} 1 ${x2} ${y2} Z`;
            
            acc.elements.push(
              <path
                key={item.plant}
                d={path}
                fill={PLANT_META[item.plant].color}
                stroke="white"
                strokeWidth="2"
              />
            );
            acc.prevAngle = endAngle;
            return acc;
          }, { elements: [], prevAngle: 0 }).elements}
          <circle cx="100" cy="100" r="40" fill="white" />
          <text x="100" y="95" textAnchor="middle" className="text-sm fill-slate-600">Total</text>
          <text x="100" y="115" textAnchor="middle" className="text-lg font-bold fill-slate-800">
            {formatNumber(plantData.reduce((sum, p) => sum + p.total, 0))}
          </text>
        </svg>
      </div>
      <div className="flex flex-wrap gap-3 justify-center">
        {plantData.map(item => (
          <div key={item.plant} className="flex items-center gap-2">
            <div 
              className="w-4 h-4 rounded"
              style={{ backgroundColor: PLANT_META[item.plant].color }}
            />
            <span className="text-sm text-slate-600">{item.plant}: {formatNumber(item.total)}</span>
          </div>
        ))}
      </div>
    </div>
  );

  const TrendChart = () => {
    const months = ['N (Feb)', 'N+1 (Mar)', 'N+2 (Apr)', 'N+3 (May)'];
    const values = [totals.n, totals.n1, totals.n2, totals.n3];
    const maxVal = Math.max(...values, 1);
    const points = values.map((v, i) => ({
      x: 50 + i * 150,
      y: 200 - (v / maxVal) * 150
    }));

    return (
      <div className="space-y-6">
        <h3 className="text-lg font-semibold text-slate-800">แนวโน้มการผลิต</h3>
        <svg viewBox="0 0 500 250" className="w-full h-64">
          {/* Grid lines */}
          {[0, 50, 100, 150, 200].map(y => (
            <line key={y} x1="50" y1={y + 50} x2="450" y2={y + 50} stroke="#E2E8F0" strokeWidth="1" />
          ))}
          {/* Line */}
          <polyline
            points={points.map(p => `${p.x},${p.y}`).join(' ')}
            fill="none"
            stroke="#3B82F6"
            strokeWidth="3"
          />
          {/* Points */}
          {points.map((p, i) => (
            <g key={i}>
              <circle cx={p.x} cy={p.y} r="6" fill="#3B82F6" stroke="white" strokeWidth="2" />
              <text x={p.x} y={p.y - 15} textAnchor="middle" className="text-sm fill-slate-700 font-medium">
                {formatNumber(values[i])}
              </text>
              <text x={p.x} y={230} textAnchor="middle" className="text-xs fill-slate-500">
                {months[i]}
              </text>
            </g>
          ))}
        </svg>
      </div>
    );
  };

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col">
        <div className="bg-gradient-to-r from-blue-600 to-purple-600 px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-white/20 rounded-xl flex items-center justify-center text-white">
              <Icons.PieChart />
            </div>
            <h2 className="text-white font-semibold text-lg">วิเคราะห์ข้อมูล (Charts)</h2>
          </div>
          <button onClick={onClose} className="w-10 h-10 bg-white/20 hover:bg-white/30 rounded-lg text-white">✕</button>
        </div>

        <div className="flex gap-2 p-4 border-b border-slate-200">
          {[
            { id: 'bar', label: 'แท่ง', icon: Icons.Chart },
            { id: 'pie', label: 'วงกลม', icon: Icons.PieChart },
            { id: 'trend', label: 'เส้น', icon: Icons.Chart },
          ].map(({ id, label, icon: Icon }) => (
            <button
              key={id}
              onClick={() => setActiveChart(id)}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg transition-colors ${
                activeChart === id ? 'bg-blue-100 text-blue-700' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              <span className="w-4 h-4"><Icon /></span>
              {label}
            </button>
          ))}
        </div>

        <div className="flex-1 overflow-auto p-6">
          {activeChart === 'bar' && <BarChart />}
          {activeChart === 'pie' && <PieChart />}
          {activeChart === 'trend' && <TrendChart />}
        </div>
      </div>
    </div>
  );
};

// ==================== EXCEL PREVIEW COMPONENT ====================
const ExcelPreview = ({ data, summaryData, fileName, onClose, onDownload, formulas = {} }) => {
  const [activeSheet, setActiveSheet] = useState('Summary');
  const [zoom, setZoom] = useState(100);
  const [showGridlines, setShowGridlines] = useState(true);
  const [showFormulas, setShowFormulas] = useState(false);
  const [highlightedCells, setHighlightedCells] = useState(new Set());
  const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });
  const [filterValue, setFilterValue] = useState('');
  const scrollRef = useRef(null);

  const sheets = useMemo(() => {
    const availableSheets = [];
    if (summaryData) availableSheets.push('Summary');
    if (data) {
      Object.keys(data).forEach(sheet => {
        if (data[sheet]?.length > 0) availableSheets.push(sheet);
      });
    }
    return availableSheets;
  }, [data, summaryData]);

  const getSheetData = (sheetName) => {
    let headers, rows, colWidths, highlightedCols;
    
    if (sheetName === 'Summary') {
      headers = ['Part Number', 'Plants', 'Sum of N', 'Sum of N+1', 'Sum of N+2', 'Sum of N+3', 'Total'];
      colWidths = [180, 120, 80, 80, 80, 80, 80];
      highlightedCols = [2, 3, 4, 5];
      
      rows = summaryData?.map(row => [
        row.partNumber,
        row.plants,
        row.n,
        row.n1,
        row.n2,
        row.n3,
        row.n + row.n1 + row.n2 + row.n3
      ]) || [];
    } else {
      headers = ['PART NUMBER', 'PART CODE', 'PART DESC', 'SUPP CODE', 'SHIPPING DOCK', 'DOCK CODE', 'CAR FAMILY', 'PACKING SIZE', 'N', 'N+1', 'N+2', 'N+3'];
      colWidths = [160, 90, 150, 90, 100, 80, 100, 90, 70, 70, 70, 70];
      highlightedCols = [8, 9, 10, 11];
      
      rows = data?.[sheetName]?.map(row => [
        row.partNumber,
        row.partCode,
        row.partDesc,
        row.suppCode,
        row.shippingDock,
        row.dockCode,
        row.carFamily,
        row.packingSize,
        row.n,
        row.n1,
        row.n2,
        row.n3
      ]) || [];
    }

    // Apply sorting
    if (sortConfig.key !== null) {
      rows = [...rows].sort((a, b) => {
        const aVal = a[sortConfig.key];
        const bVal = b[sortConfig.key];
        if (typeof aVal === 'number' && typeof bVal === 'number') {
          return sortConfig.direction === 'asc' ? aVal - bVal : bVal - aVal;
        }
        return sortConfig.direction === 'asc' 
          ? String(aVal).localeCompare(String(bVal))
          : String(bVal).localeCompare(String(aVal));
      });
    }

    // Apply filtering
    if (filterValue) {
      rows = rows.filter(row => 
        row.some(cell => String(cell).toLowerCase().includes(filterValue.toLowerCase()))
      );
    }

    return { headers, rows, colWidths, highlightedCols };
  };

  const currentSheet = getSheetData(activeSheet);
  const totalRows = currentSheet.rows.length;
  const totalCols = currentSheet.headers.length;

  const toggleHighlight = (rowIdx, colIdx) => {
    const key = `${activeSheet}-${rowIdx}-${colIdx}`;
    const newSet = new Set(highlightedCells);
    if (newSet.has(key)) newSet.delete(key);
    else newSet.add(key);
    setHighlightedCells(newSet);
  };

  const handleSort = (colIdx) => {
    setSortConfig(prev => ({
      key: colIdx,
      direction: prev.key === colIdx && prev.direction === 'asc' ? 'desc' : 'asc'
    }));
  };

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-7xl h-[90vh] flex flex-col overflow-hidden">
        {/* Header */}
        <div className="bg-gradient-to-r from-emerald-600 to-teal-600 px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-white/20 rounded-xl flex items-center justify-center">
              <Icons.File />
            </div>
            <div>
              <h2 className="text-white font-semibold text-lg">Excel Preview</h2>
              <p className="text-emerald-100 text-sm">{fileName || 'Production_Summary.xlsx'}</p>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <button onClick={onDownload} className="px-4 py-2 bg-white text-emerald-700 rounded-lg font-medium hover:bg-emerald-50 flex items-center gap-2">
              <span className="w-4 h-4"><Icons.Download /></span>
              ดาวน์โหลด
            </button>
            <button onClick={onClose} className="w-10 h-10 bg-white/20 hover:bg-white/30 rounded-lg text-white">✕</button>
          </div>
        </div>

        {/* Toolbar */}
        <div className="bg-slate-50 border-b border-slate-200 px-4 py-2 flex flex-wrap items-center gap-4">
          <div className="flex items-center gap-2">
            <span className="text-sm text-slate-500">ซูม:</span>
            <select value={zoom} onChange={(e) => setZoom(Number(e.target.value))} className="px-2 py-1 border border-slate-300 rounded text-sm">
              <option value={75}>75%</option>
              <option value={100}>100%</option>
              <option value={125}>125%</option>
              <option value={150}>150%</option>
            </select>
          </div>
          
          <div className="h-6 w-px bg-slate-300" />
          
          <label className="flex items-center gap-2 cursor-pointer">
            <input type="checkbox" checked={showGridlines} onChange={(e) => setShowGridlines(e.target.checked)} className="rounded" />
            <span className="text-sm text-slate-600">เส้นตาราง</span>
          </label>
          
          <label className="flex items-center gap-2 cursor-pointer">
            <input type="checkbox" checked={showFormulas} onChange={(e) => setShowFormulas(e.target.checked)} className="rounded" />
            <span className="text-sm text-slate-600 flex items-center gap-1">
              <span className="w-4 h-4"><Icons.Function /></span>
              แสดงสูตร
            </span>
          </label>

          <div className="h-6 w-px bg-slate-300" />

          <div className="relative">
            <span className="absolute left-2 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400"><Icons.Search /></span>
            <input
              type="text"
              placeholder="ค้นหา..."
              value={filterValue}
              onChange={(e) => setFilterValue(e.target.value)}
              className="pl-8 pr-3 py-1 border border-slate-300 rounded text-sm w-40"
            />
          </div>

          <div className="ml-auto flex items-center gap-3 text-sm text-slate-500">
            <span>{totalRows.toLocaleString()} แถว</span>
            <span>•</span>
            <span>{totalCols} คอลัมน์</span>
          </div>
        </div>

        {/* Sheet Tabs */}
        <div className="bg-slate-100 border-b border-slate-200 flex">
          {sheets.map(sheet => (
            <button
              key={sheet}
              onClick={() => setActiveSheet(sheet)}
              className={`px-6 py-2.5 text-sm font-medium border-r border-slate-200 transition-colors ${
                activeSheet === sheet ? 'bg-white text-emerald-700 border-t-2 border-t-emerald-500' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              {sheet}
              {sheet !== 'Summary' && data?.[sheet]?.length > 0 && (
                <span className="ml-2 text-xs text-slate-400">({data[sheet].length})</span>
              )}
            </button>
          ))}
        </div>

        {/* Spreadsheet Grid */}
        <div ref={scrollRef} className="flex-1 overflow-auto bg-white" style={{ transform: `scale(${zoom / 100})`, transformOrigin: 'top left', width: `${10000 / zoom}%`, height: `${10000 / zoom}%` }}>
          <div className="inline-block min-w-full">
            <table className="border-collapse">
              <thead className="sticky top-0 z-20">
                <tr>
                  <th className="w-12 h-8 bg-slate-200 border border-slate-300 sticky left-0 z-30" />
                  {currentSheet.headers.map((_, idx) => (
                    <th key={idx} className="h-8 bg-slate-100 border border-slate-300 text-xs text-slate-500 font-medium text-center" style={{ width: currentSheet.colWidths[idx] || 100 }}>
                      {getColumnLetter(idx + 1)}
                    </th>
                  ))}
                </tr>
                <tr>
                  <th className="w-12 h-10 bg-slate-200 border border-slate-300 sticky left-0 z-30 text-xs text-slate-600">{activeSheet === 'Summary' ? 'Sum' : '1'}</th>
                  {currentSheet.headers.map((header, idx) => (
                    <th 
                      key={idx}
                      onClick={() => handleSort(idx)}
                      className={`h-10 bg-emerald-50 border border-slate-300 text-xs font-semibold text-slate-700 px-2 text-left whitespace-nowrap cursor-pointer hover:bg-emerald-100 ${currentSheet.highlightedCols.includes(idx) ? 'bg-blue-50' : ''}`}
                      style={{ width: currentSheet.colWidths[idx] || 100 }}
                    >
                      <div className="flex items-center justify-between">
                        {header}
                        {sortConfig.key === idx && (
                          <span className="text-emerald-600">{sortConfig.direction === 'asc' ? '↑' : '↓'}</span>
                        )}
                      </div>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {currentSheet.rows.map((row, rowIdx) => (
                  <tr key={rowIdx}>
                    <td className="w-12 h-8 bg-slate-50 border border-slate-300 sticky left-0 z-10 text-xs text-slate-500 text-center font-medium">{rowIdx + 2}</td>
                    {row.map((cell, colIdx) => {
                      const cellKey = `${activeSheet}-${rowIdx}-${colIdx}`;
                      const isHighlighted = highlightedCells.has(cellKey);
                      const isNumberCol = currentSheet.highlightedCols.includes(colIdx);
                      const cellFormula = formulas[`${activeSheet}-${rowIdx}-${colIdx}`];
                      
                      return (
                        <td
                          key={colIdx}
                          onClick={() => toggleHighlight(rowIdx, colIdx)}
                          className={`h-8 border border-slate-300 text-xs px-2 whitespace-nowrap cursor-pointer transition-all ${isHighlighted ? 'bg-yellow-200 ring-2 ring-yellow-400 ring-inset' : isNumberCol ? 'bg-blue-50/50' : 'bg-white hover:bg-slate-50'}`}
                          style={{ width: currentSheet.colWidths[colIdx] || 100, textAlign: typeof cell === 'number' ? 'right' : 'left' }}
                          title={cellFormula && showFormulas ? cellFormula : ''}
                        >
                          {showFormulas && cellFormula ? cellFormula : (typeof cell === 'number' ? formatNumber(cell) : cell || '')}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        <div className="bg-emerald-600 text-white px-4 py-2 text-sm flex items-center justify-between">
          <div className="flex items-center gap-4">
            <span>Sheet: {activeSheet}</span>
            <span>|</span>
            <span>Ready</span>
          </div>
          <div className="flex items-center gap-4">
            <span>Zoom: {zoom}%</span>
            <span>|</span>
            <span>{highlightedCells.size} cells highlighted</span>
          </div>
        </div>
      </div>
    </div>
  );
};

// [Rest of the file continues with the main App component...]
// Due to length, let me create the complete file in chunks


// ==================== MAIN APP COMPONENT ====================
export default function App() {
  const [tab, setTab] = useState('upload');
  const [mainFile, setMainFile] = useState(null);
  const [mainWorkbook, setMainWorkbook] = useState(null);
  const [sourceFiles, setSourceFiles] = useState([]);
  const [processing, setProcessing] = useState(false);
  const [processStatus, setProcessStatus] = useState('');
  const [currentProcessIndex, setCurrentProcessIndex] = useState(-1);
  const [processedData, setProcessedData] = useState(null);
  const [summaryData, setSummaryData] = useState(null);
  const [matchingDetails, setMatchingDetails] = useState(null);
  const [changeLog, setChangeLog] = useState([]);
  const [expanded, setExpanded] = useState({});
  const [error, setError] = useState(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterPlant, setFilterPlant] = useState('all');
  const [showPreview, setShowPreview] = useState(false);
  const [showChart, setShowChart] = useState(false);
  const [previewFileName, setPreviewFileName] = useState('');

  // Stats
  const fileCounts = PLANTS.reduce((acc, p) => {
    acc[p] = sourceFiles.filter((f) => f.category === p).length;
    return acc;
  }, {});

  const totalSourceFiles = sourceFiles.length;
  const canProcess = totalSourceFiles > 0 && !processing;

  const totals = summaryData
    ? summaryData.reduce(
        (acc, r) => ({
          n: acc.n + r.n,
          n1: acc.n1 + r.n1,
          n2: acc.n2 + r.n2,
          n3: acc.n3 + r.n3,
        }),
        { n: 0, n1: 0, n2: 0, n3: 0 }
      )
    : { n: 0, n1: 0, n2: 0, n3: 0 };

  // Reset all
  const resetAll = () => {
    setTab('upload');
    setMainFile(null);
    setMainWorkbook(null);
    setSourceFiles([]);
    setProcessing(false);
    setProcessStatus('');
    setCurrentProcessIndex(-1);
    setProcessedData(null);
    setSummaryData(null);
    setMatchingDetails(null);
    setChangeLog([]);
    setExpanded({});
    setError(null);
    setSearchTerm('');
    setFilterPlant('all');
    setShowPreview(false);
    setShowChart(false);
  };

  // Handle main file upload
  const handleMainFileSelect = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      try {
        const workbook = await readExcelFileExcelJS(file);
        setMainFile({
          name: file.name,
          size: file.size,
          sheets: workbook.worksheets.map((ws) => ws.name),
        });
        setMainWorkbook(workbook);
        setError(null);
      } catch (excelJSErr) {
        console.warn('ExcelJS failed, using xlsx fallback:', excelJSErr);
        const workbook = await readExcelFileXLSX(file);
        setMainFile({
          name: file.name,
          size: file.size,
          sheets: workbook.SheetNames,
        });
        setMainWorkbook(null);
        setError('ไฟล์ถูกอ่านด้วย xlsx (อาจไม่สามารถรักษาฟอร์แมตได้ 100%)');
      }
    } catch (err) {
      setError('ไม่สามารถอ่านไฟล์หลักได้: ' + err.message);
    }
  };

  // Handle source files upload
  const handleSourceFilesSelect = (e) => {
    const files = Array.from(e.target.files || []);
    const newFiles = [];
    const rejectedFiles = [];

    files.forEach((file) => {
      const category = categorizeFile(file.name);
      if (category) {
        if (!sourceFiles.find((f) => f.name === file.name)) {
          newFiles.push({
            name: file.name,
            size: file.size,
            category,
            file,
            status: 'ready',
            rowCount: 0,
            error: null,
          });
        }
      } else {
        rejectedFiles.push(file.name);
      }
    });

    if (newFiles.length > 0) {
      setSourceFiles((prev) => [...prev, ...newFiles]);
    }

    if (rejectedFiles.length > 0) {
      setError(`ไม่สามารถจัดประเภทไฟล์ได้ ${rejectedFiles.length} ไฟล์: ${rejectedFiles.join(', ')} \nชื่อไฟล์ต้องขึ้นต้นด้วย BP_, BPK_, GW_, SR_`);
    } else if (newFiles.length > 0) {
      setError(null);
    }
  };

  // Remove source file
  const removeSourceFile = (index) => {
    setSourceFiles((prev) => prev.filter((_, i) => i !== index));
  };

  // Process all files
  const processAllFiles = async () => {
    if (!canProcess) return;

    setProcessing(true);
    setProcessStatus('กำลังเริ่มต้นประมวลผล...');
    setCurrentProcessIndex(-1);
    setError(null);

    try {
      const data = {
        'BP Daily': [],
        'BPK Daily': [],
        'GW Daily': [],
        'SR Daily': [],
      };

      const updatedFiles = [...sourceFiles];
      const logs = [];

      for (let i = 0; i < updatedFiles.length; i++) {
        const fileInfo = updatedFiles[i];
        setCurrentProcessIndex(i);
        setProcessStatus(`กำลังอ่านไฟล์: ${fileInfo.name}`);

        try {
          updatedFiles[i] = { ...fileInfo, status: 'processing' };
          setSourceFiles([...updatedFiles]);

          const workbook = await readExcelFileXLSX(fileInfo.file);
          const extracted = extractDataFromSourceXLSX(workbook);
          const sheetName = `${fileInfo.category} Daily`;

          extracted.forEach(row => {
            logs.push({
              timestamp: new Date().toISOString(),
              fileName: fileInfo.name,
              plant: fileInfo.category,
              partNumber: row.partNumber,
              action: 'extracted',
              details: `ดึงข้อมูล: N=${row.n}, N+1=${row.n1}, N+2=${row.n2}, N+3=${row.n3}`
            });
          });

          if (data[sheetName]) {
            data[sheetName].push(...extracted);
          }

          updatedFiles[i] = { ...fileInfo, status: 'done', rowCount: extracted.length };
          setSourceFiles([...updatedFiles]);
        } catch (err) {
          updatedFiles[i] = { ...fileInfo, status: 'error', error: err.message };
          setSourceFiles([...updatedFiles]);
          logs.push({
            timestamp: new Date().toISOString(),
            fileName: fileInfo.name,
            plant: fileInfo.category,
            action: 'error',
            details: err.message
          });
        }
      }

      setProcessedData(data);
      setProcessStatus('กำลังคำนวณข้อมูลสรุป...');

      const summary = {};
      const matchDetails = {};
      
      Object.entries(data).forEach(([sheet, rows]) => {
        const plant = sheet.split(' ')[0];
        rows.forEach((row) => {
          if (!summary[row.partNumber]) {
            summary[row.partNumber] = {
              n: 0, n1: 0, n2: 0, n3: 0,
              plants: new Set(),
              sources: []
            };
          }
          
          summary[row.partNumber].sources.push({
            plant, partCode: row.partCode, packingSize: row.packingSize,
            n: row.n, n1: row.n1, n2: row.n2, n3: row.n3
          });
          
          summary[row.partNumber].n += row.n;
          summary[row.partNumber].n1 += row.n1;
          summary[row.partNumber].n2 += row.n2;
          summary[row.partNumber].n3 += row.n3;
          summary[row.partNumber].plants.add(plant);
        });
      });

      const summaryArray = Object.entries(summary)
        .map(([part, d]) => ({
          partNumber: part,
          plants: Array.from(d.plants).sort().join(', '),
          n: d.n, n1: d.n1, n2: d.n2, n3: d.n3,
          sources: d.sources
        }))
        .sort((a, b) => a.partNumber.localeCompare(b.partNumber));

      setSummaryData(summaryArray);
      setMatchingDetails(summary);
      setChangeLog(logs);
      setTab('preview');
    } catch (err) {
      setError('เกิดข้อผิดพลาดในการประมวลผล: ' + err.message);
    }

    setProcessing(false);
    setProcessStatus('');
    setCurrentProcessIndex(-1);
  };

  // Export to Excel
  const exportToExcel = async () => {
    if (!processedData) return;

    let workbook;
    let fileName;
    const dateStr = new Date().toISOString().slice(0, 10);
    const highlightFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6F7FF' } };
    const highlightBorder = {
      top: { style: 'thin', color: { argb: 'FF1E40AF' } },
      left: { style: 'thin', color: { argb: 'FF1E40AF' } },
      bottom: { style: 'thin', color: { argb: 'FF1E40AF' } },
      right: { style: 'thin', color: { argb: 'FF1E40AF' } },
    };
    const applyHighlight = (cell) => {
      cell.style = { ...cell.style, fill: highlightFill, border: highlightBorder };
    };
    const isFormulaCell = (cell) => cell?.value && typeof cell.value === 'object' && (cell.value.formula || cell.value.sharedFormula);
    const safeClearCell = (cell) => { if (!cell || isFormulaCell(cell)) return; if (cell.value !== null && cell.value !== undefined) cell.value = null; };
    const safeSetCellValue = (cell, value) => { if (!cell || isFormulaCell(cell)) return; if (value !== undefined && value !== null && value !== '') cell.value = value; };

    if (mainWorkbook && mainWorkbook.worksheets) {
      workbook = mainWorkbook;
      Object.entries(processedData).forEach(([sheetName, rows]) => {
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) return;
        const dataStartRow = 14;
        const templateRow = worksheet.getRow(14);

        for (let r = dataStartRow; r <= Math.min(worksheet.rowCount, dataStartRow + 100); r++) {
          const row = worksheet.getRow(r);
          for (let c = 1; c <= 200; c++) safeClearCell(row.getCell(c));
        }

        rows.forEach((rowData, idx) => {
          const targetRow = worksheet.getRow(dataStartRow + idx);
          if (rowData.rawRow && Array.isArray(rowData.rawRow)) {
            rowData.rawRow.forEach((value, colIndex) => {
              const colNumber = colIndex + 1;
              if (colNumber <= 200 && value !== undefined && value !== null && value !== '') {
                const cell = targetRow.getCell(colNumber);
                safeSetCellValue(cell, value);
                applyHighlight(cell);
              }
            });
          }
          targetRow.commit();
        });
      });

      // Update Analyze sheet (if exists in template) - Copy ALL columns A-EG
      const analyzeSheet = workbook.getWorksheet('Analyze');
      if (analyzeSheet) {
        const analyzeStartRow = 3;
        const analyzeRows = [];
        
        Object.entries(processedData).forEach(([sheetName, rows]) => {
          const plant = sheetName.split(' ')[0];
          rows.forEach((row) => analyzeRows.push({ plant, ...row }));
        });
        
        analyzeRows.sort((a, b) => {
          const plantSort = a.plant.localeCompare(b.plant);
          if (plantSort !== 0) return plantSort;
          return a.partNumber.localeCompare(b.partNumber);
        });

        const MAX_COLS = 140; // Column EG
        const SOURCE_TO_ANALYZE_OFFSET = 2; // Analyze col 2 = Source col 0

        // Clear old data (up to EG column)
        for (let r = analyzeStartRow; r <= analyzeStartRow + 300; r++) {
          const row = analyzeSheet.getRow(r);
          for (let c = 1; c <= MAX_COLS + 2; c++) safeClearCell(row.getCell(c));
        }

        // Write new data with all columns
        analyzeRows.forEach((rowData, idx) => {
          const row = analyzeSheet.getRow(analyzeStartRow + idx);
          // Column A: Plant
          const cellA = row.getCell(1);
          safeSetCellValue(cellA, rowData.plant);
          applyHighlight(cellA);
          
          // Copy all columns from rawRow starting at column B
          if (rowData.rawRow && Array.isArray(rowData.rawRow)) {
            for (let srcCol = 0; srcCol < Math.min(rowData.rawRow.length, MAX_COLS); srcCol++) {
              const value = rowData.rawRow[srcCol];
              const destCol = srcCol + SOURCE_TO_ANALYZE_OFFSET;
              if (value !== undefined && value !== null && value !== '') {
                const cell = row.getCell(destCol);
                safeSetCellValue(cell, value);
                applyHighlight(cell);
              }
            }
          }
          row.commit();
        });
      }

      fileName = `Production_Updated_${dateStr}.xlsx`;
    } else {
      workbook = new ExcelJS.Workbook();
      Object.entries(processedData).forEach(([sheetName, rows]) => {
        if (rows.length === 0) return;
        const worksheet = workbook.addWorksheet(sheetName.replace(' ', '_'));
        const headerRow = worksheet.addRow(['PART NUMBER', 'PART CODE', 'PART DESC', 'SUPP CODE', 'SHIPPING DOCK', 'DOCK CODE', 'CAR FAMILY', 'PACKING SIZE', 'N', 'N+1', 'N+2', 'N+3']);
        headerRow.font = { bold: true };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
        rows.forEach((row) => worksheet.addRow([row.partNumber, row.partCode, row.partDesc, row.suppCode, row.shippingDock, row.dockCode, row.carFamily, row.packingSize, row.n, row.n1, row.n2, row.n3]));
      });

      if (summaryData) {
        const worksheet = workbook.addWorksheet('Summary');
        const headerRow = worksheet.addRow(['Part Number', 'Plants', 'Sum of N', 'Sum of N+1', 'Sum of N+2', 'Sum of N+3', 'Total']);
        headerRow.font = { bold: true };
        summaryData.forEach((row) => worksheet.addRow([row.partNumber, row.plants, row.n, row.n1, row.n2, row.n3, row.n + row.n1 + row.n2 + row.n3]));
      }

      // Add Analyze sheet (summary of all 4 plants) - Copy ALL columns A-EG
      const analyzeSheet = workbook.addWorksheet('Analyze');
      
      const analyzeRows = [];
      Object.entries(processedData).forEach(([sheetName, rows]) => {
        const plant = sheetName.split(' ')[0];
        rows.forEach((row) => {
          analyzeRows.push({ plant, ...row });
        });
      });
      analyzeRows.sort((a, b) => {
        const plantSort = a.plant.localeCompare(b.plant);
        if (plantSort !== 0) return plantSort;
        return a.partNumber.localeCompare(b.partNumber);
      });
      
      // Write rows with all columns from rawRow (up to 140 columns = EG)
      const MAX_COLS = 140; // Column EG
      const SOURCE_TO_ANALYZE_OFFSET = 2; // Analyze col 2 = Source col 0
      
      analyzeRows.forEach((rowData, idx) => {
        const targetRow = analyzeSheet.getRow(idx + 1);
        // Column A: Plant
        targetRow.getCell(1).value = rowData.plant;
        // Copy all columns from rawRow starting at column B
        if (rowData.rawRow && Array.isArray(rowData.rawRow)) {
          for (let srcCol = 0; srcCol < Math.min(rowData.rawRow.length, MAX_COLS); srcCol++) {
            const value = rowData.rawRow[srcCol];
            const destCol = srcCol + SOURCE_TO_ANALYZE_OFFSET; // 0+2=2 (B), 1+2=3 (C), etc.
            if (value !== undefined && value !== null && value !== '') {
              targetRow.getCell(destCol).value = value;
            }
          }
        }
        targetRow.commit();
      });
      
      fileName = `Production_Summary_${dateStr}.xlsx`;
    }

    try {
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
    } catch (err) {
      setError('เกิดข้อผิดพลาดในการดาวน์โหลดไฟล์: ' + err.message);
    }
  };

  const toggleSection = (section) => setExpanded((prev) => ({ ...prev, [section]: !prev[section] }));
  const missingPlants = PLANTS.filter((p) => fileCounts[p] === 0);
  const filesWithError = sourceFiles.filter((f) => f.status === 'error').length;
  const filesProcessed = sourceFiles.filter((f) => f.status === 'done').length;

  const filteredProcessedData = useMemo(() => {
    if (!processedData) return null;
    const filtered = {};
    Object.entries(processedData).forEach(([sheet, rows]) => {
      const plant = sheet.split(' ')[0];
      if (filterPlant !== 'all' && plant !== filterPlant) return;
      filtered[sheet] = rows.filter(row => row.partNumber.toLowerCase().includes(searchTerm.toLowerCase()));
    });
    return filtered;
  }, [processedData, searchTerm, filterPlant]);

  const filteredSummaryData = useMemo(() => {
    if (!summaryData) return null;
    return summaryData.filter(row => row.partNumber.toLowerCase().includes(searchTerm.toLowerCase()));
  }, [summaryData, searchTerm]);

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50 text-slate-900">
      <header className="bg-white border-b border-slate-200 sticky top-0 z-50 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-12 h-12 p-2.5 bg-gradient-to-br from-blue-600 to-blue-700 rounded-xl text-white shadow-lg">
              <Icons.File />
            </div>
            <div>
              <h1 className="text-xl font-bold bg-gradient-to-r from-blue-700 to-blue-500 bg-clip-text text-transparent">รีบทำเดียวอดเลนเกม</h1>
              <p className="text-sm text-slate-500">TMT Camera Production Plan Summary</p>
            </div>
          </div>
          <button onClick={resetAll} className="flex items-center gap-2 px-4 py-2 bg-slate-100 hover:bg-slate-200 rounded-lg text-slate-600">
            <span className="w-4 h-4"><Icons.Reset /></span>รีเซ็ต
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-6">
        <div className="flex flex-wrap gap-2 mb-6">
          {[
            { id: 'upload', label: 'อัปโหลดไฟล์', Icon: Icons.Upload },
            { id: 'preview', label: 'ตรวจสอบข้อมูล', Icon: Icons.Eye, disabled: !processedData, tooltip: 'ต้องประมวลผลไฟล์ก่อน' },
            { id: 'matching', label: 'รายละเอียด Matching', Icon: Icons.Link, disabled: !matchingDetails, tooltip: 'ต้องประมวลผลไฟล์ก่อน' },
            { id: 'changelog', label: 'ประวัติการแก้ไข', Icon: Icons.History, disabled: changeLog.length === 0, tooltip: changeLog.length === 0 ? 'ยังไม่มีการแก้ไข' : undefined },
            { id: 'summary', label: 'สรุปรวม', Icon: Icons.Chart, disabled: !summaryData, tooltip: 'ต้องประมวลผลไฟล์ก่อน' },
          ].map((item) => (
            <div key={item.id} className="relative group">
              <button 
                onClick={() => !item.disabled && setTab(item.id)} 
                disabled={item.disabled}
                className={`flex items-center gap-2 px-4 py-2.5 rounded-xl font-medium transition-all border ${
                  tab === item.id ? 'bg-blue-600 text-white border-blue-600 shadow-md' : 
                  item.disabled ? 'bg-slate-50 text-slate-300 border-slate-200 cursor-not-allowed' : 
                  'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'}`}
              >
                <span className="w-5 h-5"><item.Icon /></span>{item.label}
              </button>
              {item.tooltip && item.disabled && (
                <div className="absolute bottom-full left-1/2 -translate-x-1/2 mb-2 px-3 py-1.5 bg-slate-800 text-white text-xs rounded-lg whitespace-nowrap opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all z-10">
                  {item.tooltip}
                  <div className="absolute top-full left-1/2 -translate-x-1/2 border-4 border-transparent border-t-slate-800"></div>
                </div>
              )}
            </div>
          ))}
        </div>

        {error && (
          <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-xl flex items-start gap-3">
            <span className="w-6 h-6 text-red-500 flex-shrink-0"><Icons.Alert /></span>
            <span className="text-red-700 whitespace-pre-line">{error}</span>
          </div>
        )}

        {tab === 'upload' && (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            <div className="lg:col-span-1">
              <div className="bg-white border border-slate-200 rounded-2xl p-6 h-full shadow-sm">
                <h2 className="text-lg font-semibold mb-4 flex items-center gap-2">
                  <span className="w-5 h-5 text-amber-500"><Icons.File /></span>ไฟล์หลัก (Template)
                </h2>
                <label className="block cursor-pointer">
                  <input type="file" accept=".xlsx,.xls" onChange={handleMainFileSelect} className="hidden" />
                  <div className={`border-2 border-dashed rounded-xl p-8 text-center transition-all ${mainFile ? 'border-blue-400 bg-blue-50' : 'border-slate-200 hover:border-blue-300'}`}>
                    {mainFile ? (
                      <div className="flex flex-col items-center gap-3">
                        <span className="w-12 h-12 text-blue-600"><Icons.Check /></span>
                        <div>
                          <p className="font-medium">{mainFile.name}</p>
                          <p className="text-sm text-slate-500">{formatSize(mainFile.size)}</p>
                        </div>
                      </div>
                    ) : (
                      <>
                        <span className="w-12 h-12 mx-auto text-slate-400 block mb-3"><Icons.Upload /></span>
                        <p className="text-slate-600 mb-1">คลิกเพื่อเลือกไฟล์หลัก</p>
                        <p className="text-sm text-slate-400">.xlsx หรือ .xls</p>
                      </>
                    )}
                  </div>
                </label>
                <p className="text-xs text-slate-400 mt-3 text-center">* ไม่จำเป็นต้องมี</p>
              </div>
            </div>

            <div className="lg:col-span-2">
              <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
                <h2 className="text-lg font-semibold mb-4 flex items-center gap-2">
                  <span className="w-5 h-5 text-blue-600"><Icons.Chart /></span>ไฟล์ข้อมูลรายโรงงาน
                </h2>
                <label className="block cursor-pointer mb-4">
                  <input type="file" multiple accept=".xlsx,.xls" onChange={handleSourceFilesSelect} className="hidden" />
                  <div className="border-2 border-dashed border-slate-200 rounded-xl p-6 text-center hover:border-blue-400 hover:bg-blue-50 transition-all">
                    <span className="w-10 h-10 mx-auto text-blue-400 block mb-2"><Icons.Upload /></span>
                    <p className="text-slate-600 mb-1">ลากไฟล์มาวางหรือคลิกเลือก</p>
                    <p className="text-sm text-slate-400">BP_*, BPK_*, GW_*, SR_*</p>
                  </div>
                </label>

                <div className="grid grid-cols-4 gap-3 mb-4">
                  {PLANTS.map((cat) => (
                    <div key={cat} className={`p-3 rounded-xl border-2 ${PLANT_META[cat].border} ${PLANT_META[cat].lightBg} text-center`}>
                      <div className={`text-sm font-bold ${PLANT_META[cat].iconColor}`}>{cat}</div>
                      <div className="text-2xl font-bold">{fileCounts[cat]}</div>
                      <div className="text-xs text-slate-500">ไฟล์</div>
                    </div>
                  ))}
                </div>

                {sourceFiles.length > 0 && (
                  <div className="space-y-2 max-h-64 overflow-y-auto border border-slate-100 rounded-xl p-2">
                    {sourceFiles.map((file, i) => (
                      <div key={file.name} className="flex items-center justify-between bg-slate-50 border border-slate-200 rounded-lg px-4 py-3">
                        <div className="flex items-center gap-3">
                          <span className={`w-5 h-5 ${file.status === 'done' ? 'text-emerald-500' : file.status === 'error' ? 'text-red-500' : 'text-slate-400'}`}>
                            {file.status === 'done' ? <Icons.Check /> : <Icons.File />}
                          </span>
                          <div>
                            <p className="text-sm font-medium">{file.name}</p>
                            <p className="text-xs text-slate-400">{formatSize(file.size)} {file.rowCount > 0 && `• ${file.rowCount} แถว`}</p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2">
                          <span className={`px-2 py-1 rounded-lg text-xs font-medium ${PLANT_META[file.category].badge}`}>{file.category}</span>
                          <button onClick={() => removeSourceFile(i)} className="w-5 h-5 text-slate-400 hover:text-red-500"><Icons.Trash /></button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>

            <div className="lg:col-span-3">
              <button onClick={processAllFiles} disabled={!canProcess}
                className={`w-full py-4 rounded-xl font-semibold text-lg flex items-center justify-center gap-3 transition-all ${
                  canProcess ? 'bg-gradient-to-r from-blue-600 to-blue-700 text-white hover:from-blue-700 hover:to-blue-800 shadow-lg' : 'bg-slate-200 text-slate-400 cursor-not-allowed'}`}>
                <span className="w-6 h-6"><Icons.Play /></span>รวมข้อมูลและคำนวณ
              </button>
            </div>
          </div>
        )}

        {tab === 'preview' && processedData && (
          <div className="space-y-6">
            <div className="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm">
              <div className="flex gap-4">
                <input type="text" placeholder="ค้นหา Part Number..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)}
                  className="flex-1 px-4 py-2 border border-slate-200 rounded-xl" />
                <select value={filterPlant} onChange={(e) => setFilterPlant(e.target.value)} className="px-4 py-2 border border-slate-200 rounded-xl">
                  <option value="all">ทุกโรงงาน</option>
                  {PLANTS.map(p => <option key={p} value={p}>{p}</option>)}
                </select>
              </div>
            </div>

            <div className="bg-gradient-to-br from-emerald-50 to-teal-50 border border-emerald-200 rounded-2xl p-6">
              <div className="flex justify-between items-center mb-4">
                <h3 className="text-lg font-semibold text-emerald-900">ตรวจสอบข้อมูล</h3>
                <button onClick={() => { setPreviewFileName(`Production_${new Date().toISOString().slice(0, 10)}.xlsx`); setShowPreview(true); }}
                  className="px-4 py-2 bg-emerald-600 text-white rounded-lg flex items-center gap-2">
                  <span className="w-4 h-4"><Icons.Eye /></span>Excel Preview
                </button>
              </div>
              <div className="grid grid-cols-4 gap-3">
                <div className="bg-white/70 rounded-xl p-3"><p className="text-sm text-emerald-600">ไฟล์หลัก</p><p className="font-semibold">{mainFile ? '✓ พร้อม' : 'ไม่มี'}</p></div>
                <div className="bg-white/70 rounded-xl p-3"><p className="text-sm text-emerald-600">ประมวลผลสำเร็จ</p><p className="font-semibold">{filesProcessed} / {totalSourceFiles}</p></div>
                <div className="bg-white/70 rounded-xl p-3"><p className="text-sm text-emerald-600">โรงงานขาด</p><p className="font-semibold">{missingPlants.length > 0 ? missingPlants.join(', ') : 'ครบ'}</p></div>
                <div className="bg-white/70 rounded-xl p-3"><p className="text-sm text-emerald-600">ข้อผิดพลาด</p><p className="font-semibold">{filesWithError > 0 ? `${filesWithError} ไฟล์` : 'ไม่มี'}</p></div>
              </div>
            </div>

            {Object.entries(filteredProcessedData || {}).map(([sheet, rows]) => (
              <div key={sheet} className="bg-white border border-slate-200 rounded-xl overflow-hidden">
                <button onClick={() => toggleSection(sheet)} className="w-full px-5 py-4 flex items-center justify-between hover:bg-slate-50">
                  <div className="flex items-center gap-3">
                    <span className={`px-3 py-1 rounded-lg text-sm font-medium ${PLANT_META[sheet.split(' ')[0]].badge}`}>{sheet.split(' ')[0]}</span>
                    <span className="font-semibold">{sheet}</span>
                    <span className="text-slate-400">({rows.length})</span>
                  </div>
                  <span className="w-5 h-5 text-slate-400">{expanded[sheet] ? <Icons.ChevronUp /> : <Icons.ChevronDown />}</span>
                </button>
                {expanded[sheet] && rows.length > 0 && (
                  <div className="px-5 pb-4 overflow-x-auto">
                    <table className="w-full text-sm">
                      <thead><tr className="text-slate-500 border-b"><th className="text-left py-2">Part Number</th><th className="text-right">N</th><th className="text-right">N+1</th><th className="text-right">N+2</th><th className="text-right">N+3</th></tr></thead>
                      <tbody>{rows.slice(0, 50).map((row, idx) => (
                        <tr key={idx} className="border-b"><td className="py-2 font-mono text-blue-600">{row.partNumber}</td>
                          <td className="text-right">{formatNumber(row.n)}</td><td className="text-right">{formatNumber(row.n1)}</td>
                          <td className="text-right">{formatNumber(row.n2)}</td><td className="text-right">{formatNumber(row.n3)}</td></tr>
                      ))}</tbody>
                    </table>
                  </div>
                )}
              </div>
            ))}
          </div>
        )}

        {tab === 'summary' && summaryData && (
          <div className="space-y-6">
            <div className="grid grid-cols-4 gap-4">
              {[{label: 'N (Feb)', value: totals.n, color: 'from-blue-500 to-blue-600'}, {label: 'N+1 (Mar)', value: totals.n1, color: 'from-emerald-500 to-emerald-600'}, {label: 'N+2 (Apr)', value: totals.n2, color: 'from-amber-500 to-amber-600'}, {label: 'N+3 (May)', value: totals.n3, color: 'from-purple-500 to-purple-600'}].map(card => (
                <div key={card.label} className={`bg-gradient-to-br ${card.color} rounded-2xl p-5 text-white shadow-lg`}>
                  <p className="text-white/80 text-sm">{card.label}</p>
                  <p className="text-3xl font-bold">{formatNumber(card.value)}</p>
                </div>
              ))}
            </div>

            <div className="flex gap-4">
              <button onClick={() => setShowChart(true)} className="flex-1 py-3 bg-gradient-to-r from-purple-600 to-pink-600 text-white rounded-xl flex items-center justify-center gap-2">
                <span className="w-5 h-5"><Icons.PieChart /></span>ดูกราฟวิเคราะห์
              </button>
              <button onClick={() => { setPreviewFileName(`Production_${new Date().toISOString().slice(0, 10)}.xlsx`); setShowPreview(true); }}
                className="flex-1 py-3 bg-white border-2 border-blue-600 text-blue-600 rounded-xl flex items-center justify-center gap-2">
                <span className="w-5 h-5"><Icons.Eye /></span>Preview Excel
              </button>
              <button onClick={exportToExcel} className="flex-1 py-3 bg-gradient-to-r from-blue-600 to-blue-700 text-white rounded-xl flex items-center justify-center gap-2">
                <span className="w-5 h-5"><Icons.Download /></span>ดาวน์โหลด
              </button>
            </div>

            <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden">
              <div className="px-6 py-4 border-b bg-slate-50">
                <h3 className="font-semibold">สรุปยอดรวม ({filteredSummaryData?.length} รายการ)</h3>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead><tr className="bg-slate-50"><th className="text-left py-3 px-4">Part Number</th><th className="text-left">Plants</th><th className="text-right">N</th><th className="text-right">N+1</th><th className="text-right">N+2</th><th className="text-right">N+3</th><th className="text-right">Total</th></tr></thead>
                  <tbody>{filteredSummaryData?.map(row => (
                    <tr key={row.partNumber} className="border-b hover:bg-slate-50">
                      <td className="py-3 px-4 font-mono text-blue-600">{row.partNumber}</td>
                      <td>{row.plants.split(', ').map(p => <span key={p} className={`px-2 py-0.5 rounded text-xs ${PLANT_META[p]?.badge}`}>{p}</span>)}</td>
                      <td className="text-right">{formatNumber(row.n)}</td><td className="text-right">{formatNumber(row.n1)}</td>
                      <td className="text-right">{formatNumber(row.n2)}</td><td className="text-right">{formatNumber(row.n3)}</td>
                      <td className="text-right font-bold">{formatNumber(row.n + row.n1 + row.n2 + row.n3)}</td>
                    </tr>
                  ))}</tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* Matching Details Tab */}
        {tab === 'matching' && matchingDetails && (
          <div className="space-y-6">
            <div className="bg-blue-50 border border-blue-200 rounded-xl p-4">
              <p className="text-blue-800 text-sm">
                <span className="font-semibold">รายละเอียด Matching:</span> แสดงข้อมูลว่าแต่ละ Part Number ถูกดึงมาจากไฟล์ของโรงงานใดบ้าง
              </p>
            </div>
            
            <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden">
              <div className="px-6 py-4 border-b bg-slate-50 flex justify-between items-center">
                <h3 className="font-semibold">รายละเอียดการ Matching ({Object.keys(matchingDetails).length} Part Numbers)</h3>
              </div>
              <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
                <table className="w-full text-sm">
                  <thead className="sticky top-0 bg-slate-50 z-10">
                    <tr>
                      <th className="text-left py-3 px-4 border-b">Part Number</th>
                      <th className="text-left border-b">แหล่งที่มา (Plant)</th>
                      <th className="text-left border-b">Part Code</th>
                      <th className="text-right border-b">N</th>
                      <th className="text-right border-b">N+1</th>
                      <th className="text-right border-b">N+2</th>
                      <th className="text-right border-b">N+3</th>
                    </tr>
                  </thead>
                  <tbody>
                    {Object.entries(matchingDetails).sort((a, b) => a[0].localeCompare(b[0])).map(([partNumber, data]) => (
                      <React.Fragment key={partNumber}>
                        {data.sources.map((source, idx) => (
                          <tr key={`${partNumber}-${idx}`} className="border-b hover:bg-slate-50">
                            {idx === 0 && (
                              <td rowSpan={data.sources.length} className="py-3 px-4 font-mono text-blue-600 font-medium align-top bg-slate-50/50">
                                {partNumber}
                              </td>
                            )}
                            <td className="py-3">
                              <span className={`px-2 py-0.5 rounded text-xs ${PLANT_META[source.plant]?.badge || 'bg-gray-100 text-gray-700'}`}>
                                {source.plant}
                              </span>
                            </td>
                            <td className="font-mono text-xs text-slate-600">{source.partCode}</td>
                            <td className="text-right">{formatNumber(source.n)}</td>
                            <td className="text-right">{formatNumber(source.n1)}</td>
                            <td className="text-right">{formatNumber(source.n2)}</td>
                            <td className="text-right">{formatNumber(source.n3)}</td>
                          </tr>
                        ))}
                      </React.Fragment>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* Changelog Tab */}
        {tab === 'changelog' && changeLog.length > 0 && (
          <div className="space-y-6">
            <div className="bg-amber-50 border border-amber-200 rounded-xl p-4">
              <p className="text-amber-800 text-sm">
                <span className="font-semibold">ประวัติการประมวลผล:</span> แสดงรายการที่เกิดขึ้นระหว่างการดึงข้อมูลจากไฟล์
              </p>
            </div>
            
            <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden">
              <div className="px-6 py-4 border-b bg-slate-50">
                <h3 className="font-semibold">รายการประมวลผล ({changeLog.length} รายการ)</h3>
              </div>
              <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
                <table className="w-full text-sm">
                  <thead className="sticky top-0 bg-slate-50 z-10">
                    <tr className="border-b">
                      <th className="text-left py-3 px-4">เวลา</th>
                      <th className="text-left">ไฟล์</th>
                      <th className="text-left">โรงงาน</th>
                      <th className="text-left">Part Number</th>
                      <th className="text-left">การดำเนินการ</th>
                      <th className="text-left">รายละเอียด</th>
                    </tr>
                  </thead>
                  <tbody>
                    {changeLog.map((log, idx) => (
                      <tr key={idx} className="border-b hover:bg-slate-50">
                        <td className="py-3 px-4 text-slate-500 text-xs whitespace-nowrap">
                          {log.timestamp ? new Date(log.timestamp).toLocaleString('th-TH') : '-'}
                        </td>
                        <td className="py-3 text-slate-700 text-xs">{log.fileName}</td>
                        <td className="py-3">
                          <span className={`px-2 py-0.5 rounded text-xs ${PLANT_META[log.plant]?.badge || 'bg-gray-100 text-gray-700'}`}>
                            {log.plant}
                          </span>
                        </td>
                        <td className="py-3 font-mono text-xs text-blue-600">{log.partNumber || '-'}</td>
                        <td className="py-3">
                          <span className={`px-2 py-0.5 rounded text-xs font-medium ${
                            log.action === 'error' ? 'bg-red-100 text-red-700' :
                            log.action === 'extracted' ? 'bg-emerald-100 text-emerald-700' :
                            'bg-blue-100 text-blue-700'
                          }`}>
                            {log.action === 'error' ? 'ข้อผิดพลาด' :
                             log.action === 'extracted' ? 'ดึงข้อมูล' : log.action}
                          </span>
                        </td>
                        <td className="py-3 text-slate-600 text-xs">{log.details || '-'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {showChart && <ChartDashboard data={processedData} summaryData={summaryData} onClose={() => setShowChart(false)} />}
        {showPreview && <ExcelPreview data={processedData} summaryData={summaryData} fileName={previewFileName} onClose={() => setShowPreview(false)} onDownload={exportToExcel} />}
      </main>

      <footer className="border-t border-slate-200 mt-12 py-6 bg-white text-center text-slate-400 text-sm">
        ชอน v1.0
      </footer>
    </div>
  );
}
