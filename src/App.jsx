import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';

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
};

// ==================== CONSTANTS ====================
const PLANTS = ['BP', 'BPK', 'GW', 'SR'];

const PLANT_META = {
  BP: { label: 'Ban Pho', badge: 'bg-blue-100 text-blue-700', border: 'border-blue-300' },
  BPK: { label: 'Ban Pho Kaeng Khoi', badge: 'bg-emerald-100 text-emerald-700', border: 'border-emerald-300' },
  GW: { label: 'Gateway', badge: 'bg-purple-100 text-purple-700', border: 'border-purple-300' },
  SR: { label: 'Samrong', badge: 'bg-orange-100 text-orange-700', border: 'border-orange-300' },
};

// ==================== UTILITY FUNCTIONS ====================
const formatNumber = (num) => (num ?? 0).toLocaleString();
const formatSize = (bytes) => `${(bytes / 1024).toFixed(1)} KB`;

const categorizeFile = (filename) => {
  const name = filename.toUpperCase();
  // Handle both underscore and space separators: BP_xxx, BP xxx, BP-xxx
  // Check BPK first (longer prefix) to avoid matching BP
  if (name.startsWith('BPK_') || name.startsWith('BPK ') || name.startsWith('BPK-') || name.match(/^BPK[^A-Z]/)) return 'BPK';
  if (name.startsWith('BP_') || name.startsWith('BP ') || name.startsWith('BP-') || name.match(/^BP[^A-Z]/)) return 'BP';
  if (name.startsWith('GW_') || name.startsWith('GW ') || name.startsWith('GW-') || name.match(/^GW[^A-Z]/)) return 'GW';
  if (name.startsWith('SR_') || name.startsWith('SR ') || name.startsWith('SR-') || name.match(/^SR[^A-Z]/)) return 'SR';
  return null;
};

// ==================== EXCEL FUNCTIONS ====================
const readExcelFile = (file) => {
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

const extractDataFromSource = (workbook) => {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  // Find header row (row 13, index 12)
  const headerRowIndex = 12;
  const dataStartIndex = 13;

  if (jsonData.length <= dataStartIndex) return [];

  const headers = jsonData[headerRowIndex] || [];
  const data = [];

  // Find N, N+1, N+2, N+3 total columns
  let nCol = -1,
    n1Col = -1,
    n2Col = -1,
    n3Col = -1;
  headers.forEach((h, i) => {
    const val = String(h).trim().toUpperCase();
    if (val === 'N' && nCol === -1) nCol = i;
    else if (val === 'N+1' && n1Col === -1) n1Col = i;
    else if (val === 'N+2' && n2Col === -1) n2Col = i;
    else if (val === 'N+3' && n3Col === -1) n3Col = i;
  });

  // Default positions if not found
  if (nCol === -1) nCol = 39;
  if (n1Col === -1) n1Col = 71;
  if (n2Col === -1) n2Col = 103;
  if (n3Col === -1) n3Col = 135;

  for (let i = dataStartIndex; i < jsonData.length; i++) {
    const row = jsonData[i];
    if (!row || !row[0] || row[0] === '<EOF>') continue;

    const partNumber = String(row[0] || '').trim();
    if (!partNumber || partNumber.length < 5) continue;

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
      // Keep raw row data for template export (up to column 136)
      rawRow: row.slice(0, Math.max(136, n3Col + 1)),
      colPositions: { nCol, n1Col, n2Col, n3Col },
    });
  }

  return data;
};

// ==================== MAIN COMPONENT ====================
export default function App() {
  const [tab, setTab] = useState('upload');
  const [mainFile, setMainFile] = useState(null);
  const [mainWorkbook, setMainWorkbook] = useState(null);
  const [sourceFiles, setSourceFiles] = useState([]);
  const [processing, setProcessing] = useState(false);
  const [processedData, setProcessedData] = useState(null);
  const [summaryData, setSummaryData] = useState(null);
  const [expanded, setExpanded] = useState({});
  const [error, setError] = useState(null);

  // Counts per plant
  const fileCounts = PLANTS.reduce((acc, p) => {
    acc[p] = sourceFiles.filter((f) => f.category === p).length;
    return acc;
  }, {});

  const totalSourceFiles = sourceFiles.length;
  const canProcess = totalSourceFiles > 0 && !processing;

  // Calculate totals from summary
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
    setProcessedData(null);
    setSummaryData(null);
    setExpanded({});
    setError(null);
  };

  // Handle main file upload
  const handleMainFileSelect = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const workbook = await readExcelFile(file);
      setMainFile({
        name: file.name,
        size: file.size,
        sheets: workbook.SheetNames,
      });
      setMainWorkbook(workbook);
      setError(null);
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
        // Check duplicate
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
        // Track files that couldn't be categorized
        rejectedFiles.push(file.name);
      }
    });

    if (newFiles.length > 0) {
      setSourceFiles((prev) => [...prev, ...newFiles]);
    }

    // Show warning for rejected files
    if (rejectedFiles.length > 0) {
      setError(`ไม่สามารถจัดประเภทไฟล์ได้ ${rejectedFiles.length} ไฟล์: ${rejectedFiles.join(', ')} \nชื่อไฟล์ต้องขึ้นต้นด้วย BP_, BPK_, GW_, SR_ (หรือใช้ช่องว่างแทน _)`);
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
    setError(null);

    try {
      const data = {
        'BP Daily': [],
        'BPK Daily': [],
        'GW Daily': [],
        'SR Daily': [],
      };

      // Process each source file
      const updatedFiles = [...sourceFiles];

      for (let i = 0; i < updatedFiles.length; i++) {
        const fileInfo = updatedFiles[i];

        try {
          // Update status to processing
          updatedFiles[i] = { ...fileInfo, status: 'processing' };
          setSourceFiles([...updatedFiles]);

          const workbook = await readExcelFile(fileInfo.file);
          const extracted = extractDataFromSource(workbook);
          const sheetName = `${fileInfo.category} Daily`;

          if (data[sheetName]) {
            data[sheetName].push(...extracted);
          }

          // Update status to done
          updatedFiles[i] = { ...fileInfo, status: 'done', rowCount: extracted.length };
          setSourceFiles([...updatedFiles]);
        } catch (err) {
          updatedFiles[i] = { ...fileInfo, status: 'error', error: err.message };
          setSourceFiles([...updatedFiles]);
        }
      }

      setProcessedData(data);

      // Calculate summary
      const summary = {};
      Object.entries(data).forEach(([sheet, rows]) => {
        const plant = sheet.split(' ')[0];
        rows.forEach((row) => {
          if (!summary[row.partNumber]) {
            summary[row.partNumber] = {
              n: 0,
              n1: 0,
              n2: 0,
              n3: 0,
              plants: new Set(),
            };
          }
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
          n: d.n,
          n1: d.n1,
          n2: d.n2,
          n3: d.n3,
        }))
        .sort((a, b) => a.partNumber.localeCompare(b.partNumber));

      setSummaryData(summaryArray);
      setTab('preview');
    } catch (err) {
      setError('เกิดข้อผิดพลาดในการประมวลผล: ' + err.message);
    }

    setProcessing(false);
  };

  // Export to Excel - either update template or create new file
  const exportToExcel = () => {
    if (!processedData) return;

    let wb;
    let fileName;
    const dateStr = new Date().toISOString().slice(0, 10);

    // ===== CASE 1: Update existing template =====
    if (mainWorkbook) {
      wb = mainWorkbook;

      // Data column limit - only clear/write columns A to EJ (0-139)
      // This preserves summary sections on the right side of the sheet
      const DATA_COL_LIMIT = 139;

      // Update Daily sheets (BP Daily, BPK Daily, GW Daily, SR Daily)
      Object.entries(processedData).forEach(([sheetName, rows]) => {
        const ws = wb.Sheets[sheetName];
        if (!ws) return; // Sheet doesn't exist in template

        const dataStartRow = 14; // Data starts at row 14 (index 13)

        // Clear old data in data columns only (preserve right side summary)
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        for (let r = dataStartRow - 1; r <= Math.min(range.e.r, dataStartRow + 100); r++) {
          for (let c = 0; c <= Math.min(DATA_COL_LIMIT, range.e.c); c++) {
            const cellRef = XLSX.utils.encode_cell({ r, c });
            if (ws[cellRef]) {
              delete ws[cellRef];
            }
          }
        }

        // Write new data starting at row 14
        rows.forEach((row, idx) => {
          const rowNum = dataStartRow + idx;

          // If we have raw row data, use it to preserve all columns
          if (row.rawRow && row.colPositions) {
            // Write raw row data (only up to DATA_COL_LIMIT)
            row.rawRow.forEach((val, colIdx) => {
              if (colIdx > DATA_COL_LIMIT) return;
              const cellRef = XLSX.utils.encode_cell({ r: rowNum - 1, c: colIdx });
              if (val !== undefined && val !== null && val !== '') {
                ws[cellRef] = { v: val, t: typeof val === 'number' ? 'n' : 's' };
              }
            });
          } else {
            // Write basic columns (A-H + N columns)
            const basicData = [
              row.partNumber,
              row.partCode,
              row.partDesc,
              row.suppCode,
              row.shippingDock,
              row.dockCode,
              row.carFamily,
              row.packingSize,
            ];
            basicData.forEach((val, colIdx) => {
              const cellRef = XLSX.utils.encode_cell({ r: rowNum - 1, c: colIdx });
              if (val !== undefined && val !== null && val !== '') {
                ws[cellRef] = { v: val, t: typeof val === 'number' ? 'n' : 's' };
              }
            });

            // Write N values at default positions (AN=39, BT=71, CZ=103, EF=135)
            const nPositions = [
              { col: 39, val: row.n },
              { col: 71, val: row.n1 },
              { col: 103, val: row.n2 },
              { col: 135, val: row.n3 },
            ];
            nPositions.forEach(({ col, val }) => {
              const cellRef = XLSX.utils.encode_cell({ r: rowNum - 1, c: col });
              ws[cellRef] = { v: val, t: 'n' };
            });
          }
        });

        // Update sheet range (preserve original range if larger)
        const newRange = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        newRange.e.r = Math.max(newRange.e.r, dataStartRow - 1 + rows.length);
        ws['!ref'] = XLSX.utils.encode_range(newRange);
      });

      // Update Sheet2 (Summary/Pivot) if exists - preserve existing structure
      if (summaryData && wb.Sheets['Sheet2']) {
        const ws = wb.Sheets['Sheet2'];
        const pivotStartRow = 4; // Pivot data typically starts at row 4

        // Only clear columns A-E for pivot data
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        for (let r = pivotStartRow - 1; r <= Math.min(range.e.r, pivotStartRow + summaryData.length + 5); r++) {
          for (let c = 0; c <= 4; c++) {
            const cellRef = XLSX.utils.encode_cell({ r, c });
            if (ws[cellRef]) delete ws[cellRef];
          }
        }

        // Write new summary data
        summaryData.forEach((row, idx) => {
          const rowNum = pivotStartRow + idx;
          const rowData = [row.partNumber, row.n, row.n1, row.n2, row.n3];
          rowData.forEach((val, colIdx) => {
            const cellRef = XLSX.utils.encode_cell({ r: rowNum - 1, c: colIdx });
            ws[cellRef] = { v: val, t: typeof val === 'number' ? 'n' : 's' };
          });
        });

        // Add Grand Total
        const grandTotalRow = pivotStartRow + summaryData.length;
        const grandTotalData = ['Grand Total', totals.n, totals.n1, totals.n2, totals.n3];
        grandTotalData.forEach((val, colIdx) => {
          const cellRef = XLSX.utils.encode_cell({ r: grandTotalRow - 1, c: colIdx });
          ws[cellRef] = { v: val, t: typeof val === 'number' ? 'n' : 's' };
        });

        // Preserve original range
        const newRange = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        newRange.e.r = Math.max(newRange.e.r, grandTotalRow);
        ws['!ref'] = XLSX.utils.encode_range(newRange);
      }

      fileName = `Production_Updated_${dateStr}.xlsx`;
    }
    // ===== CASE 2: Create new file from scratch =====
    else {
      wb = XLSX.utils.book_new();

      // Daily sheets
      Object.entries(processedData).forEach(([sheetName, rows]) => {
        if (rows.length === 0) return;

        const headers = [
          'PART NUMBER',
          'PART CODE',
          'PART DESC',
          'SUPP CODE',
          'SHIPPING DOCK',
          'DOCK CODE',
          'CAR FAMILY',
          'PACKING SIZE',
          'N',
          'N+1',
          'N+2',
          'N+3',
        ];

        const wsData = [headers];
        rows.forEach((row) => {
          wsData.push([
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
            row.n3,
          ]);
        });

        const ws = XLSX.utils.aoa_to_sheet(wsData);
        XLSX.utils.book_append_sheet(wb, ws, sheetName.replace(' ', '_'));
      });

      // Summary sheet
      if (summaryData) {
        const summaryHeaders = ['Part Number', 'Plants', 'Sum of N', 'Sum of N+1', 'Sum of N+2', 'Sum of N+3', 'Total'];
        const summaryRows = [summaryHeaders];

        summaryData.forEach((row) => {
          summaryRows.push([
            row.partNumber,
            row.plants,
            row.n,
            row.n1,
            row.n2,
            row.n3,
            row.n + row.n1 + row.n2 + row.n3,
          ]);
        });

        // Grand Total
        summaryRows.push([
          'Grand Total',
          `${summaryData.length} Parts`,
          totals.n,
          totals.n1,
          totals.n2,
          totals.n3,
          totals.n + totals.n1 + totals.n2 + totals.n3,
        ]);

        const ws = XLSX.utils.aoa_to_sheet(summaryRows);
        XLSX.utils.book_append_sheet(wb, ws, 'Summary');
      }

      // Verification sheet
      const missingPlantsExport = PLANTS.filter((p) => fileCounts[p] === 0);
      const verificationData = [
        ['Item', 'Status', 'Detail'],
        ['ไฟล์หลัก (Template)', mainFile ? 'พร้อม' : 'ไม่มี (ไม่บังคับ)', mainFile?.name || '-'],
        ['จำนวนไฟล์ทั้งหมด', totalSourceFiles > 0 ? 'พร้อม' : 'ไม่พบ', `${totalSourceFiles} ไฟล์`],
        ['โรงงานที่มีไฟล์', '', PLANTS.filter((p) => fileCounts[p] > 0).join(', ') || '-'],
        ['โรงงานที่ไม่มีไฟล์', missingPlantsExport.length > 0 ? 'ขาด' : 'ครบ', missingPlantsExport.join(', ') || '-'],
      ];

      PLANTS.forEach((p) => {
        verificationData.push([`ไฟล์ ${p}`, fileCounts[p] > 0 ? 'พร้อม' : 'ไม่พบ', `${fileCounts[p]} ไฟล์`]);
      });

      const wsVerify = XLSX.utils.aoa_to_sheet(verificationData);
      XLSX.utils.book_append_sheet(wb, wsVerify, 'Verification');

      fileName = `Production_Summary_${dateStr}.xlsx`;
    }

    // Download the file
    XLSX.writeFile(wb, fileName);
  };

  // Toggle section
  const toggleSection = (section) => {
    setExpanded((prev) => ({ ...prev, [section]: !prev[section] }));
  };

  // Verification stats
  const missingPlants = PLANTS.filter((p) => fileCounts[p] === 0);
  const filesWithError = sourceFiles.filter((f) => f.status === 'error').length;
  const filesProcessed = sourceFiles.filter((f) => f.status === 'done').length;

  return (
    <div className="min-h-screen bg-slate-100 text-slate-900">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-12 h-12 p-2.5 bg-blue-600 rounded-xl text-white">
                <Icons.File />
              </div>
              <div>
                <h1 className="text-xl font-bold">Production Data Consolidator</h1>
                <p className="text-sm text-slate-500">TMT Camera Production Plan Summary</p>
              </div>
            </div>
            <button
              onClick={resetAll}
              className="flex items-center gap-2 px-4 py-2 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors"
            >
              <span className="w-4 h-4">
                <Icons.Reset />
              </span>
              รีเซ็ต
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-6">
        {/* Tabs */}
        <div className="flex gap-2 mb-6">
          {[
            { id: 'upload', label: 'อัปโหลดไฟล์', Icon: Icons.Upload },
            { id: 'preview', label: 'ตรวจสอบข้อมูล', Icon: Icons.Eye, disabled: !processedData },
            { id: 'summary', label: 'สรุปรวม', Icon: Icons.Chart, disabled: !summaryData },
          ].map((item) => (
            <button
              key={item.id}
              onClick={() => !item.disabled && setTab(item.id)}
              disabled={item.disabled}
              className={`flex items-center gap-2 px-5 py-2.5 rounded-xl font-medium transition-all border ${
                tab === item.id
                  ? 'bg-blue-600 text-white border-blue-600'
                  : item.disabled
                  ? 'bg-slate-50 text-slate-300 border-slate-200 cursor-not-allowed'
                  : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'
              }`}
            >
              <span className="w-5 h-5">
                <item.Icon />
              </span>
              {item.label}
            </button>
          ))}
        </div>

        {/* Error Alert */}
        {error && (
          <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-xl flex items-center gap-3">
            <span className="w-6 h-6 text-red-500">
              <Icons.Alert />
            </span>
            <span className="text-red-700">{error}</span>
          </div>
        )}

        {/* Upload Tab */}
        {tab === 'upload' && (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            {/* Main File */}
            <div className="lg:col-span-1">
              <div className="bg-white border border-slate-200 rounded-2xl p-6 h-full">
                <h2 className="text-lg font-semibold mb-4 flex items-center gap-2">
                  <span className="w-5 h-5 text-amber-500">
                    <Icons.File />
                  </span>
                  ไฟล์หลัก (Template)
                </h2>

                <label className="block cursor-pointer">
                  <input
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={handleMainFileSelect}
                    className="hidden"
                  />
                  <div
                    className={`border-2 border-dashed rounded-xl p-8 text-center transition-all ${
                      mainFile ? 'border-blue-300 bg-blue-50' : 'border-slate-200 hover:border-blue-300 hover:bg-slate-50'
                    }`}
                  >
                    {mainFile ? (
                      <div className="flex flex-col items-center gap-3">
                        <span className="w-12 h-12 text-blue-600">
                          <Icons.Check />
                        </span>
                        <div>
                          <p className="font-medium">{mainFile.name}</p>
                          <p className="text-sm text-slate-500">{formatSize(mainFile.size)}</p>
                          <p className="text-xs text-slate-400 mt-1">{mainFile.sheets?.length || 0} sheets</p>
                        </div>
                      </div>
                    ) : (
                      <>
                        <span className="w-12 h-12 mx-auto text-slate-400 block mb-3">
                          <Icons.Upload />
                        </span>
                        <p className="text-slate-600 mb-1">คลิกเพื่อเลือกไฟล์หลัก</p>
                        <p className="text-sm text-slate-400">.xlsx หรือ .xls</p>
                      </>
                    )}
                  </div>
                </label>

                <p className="text-xs text-slate-400 mt-3 text-center">* ไม่จำเป็นต้องมี ระบบจะสร้างไฟล์ใหม่ได้</p>
              </div>
            </div>

            {/* Source Files */}
            <div className="lg:col-span-2">
              <div className="bg-white border border-slate-200 rounded-2xl p-6">
                <h2 className="text-lg font-semibold mb-4 flex items-center gap-2">
                  <span className="w-5 h-5 text-blue-600">
                    <Icons.Chart />
                  </span>
                  ไฟล์ข้อมูลรายโรงงาน
                </h2>

                {/* Drop Zone */}
                <label className="block cursor-pointer mb-4">
                  <input
                    type="file"
                    multiple
                    accept=".xlsx,.xls"
                    onChange={handleSourceFilesSelect}
                    className="hidden"
                  />
                  <div className="border-2 border-dashed border-slate-200 rounded-xl p-6 text-center hover:border-blue-300 hover:bg-slate-50 transition-all">
                    <span className="w-10 h-10 mx-auto text-slate-400 block mb-2">
                      <Icons.Upload />
                    </span>
                    <p className="text-slate-600 mb-1">ลากไฟล์มาวางที่นี่ หรือคลิกเพื่อเลือก</p>
                    <p className="text-sm text-slate-400">รองรับ: BP_*, BPK_*, GW_*, SR_* (.xls, .xlsx)</p>
                  </div>
                </label>

                {/* Category Cards */}
                <div className="grid grid-cols-4 gap-3 mb-4">
                  {PLANTS.map((cat) => (
                    <div key={cat} className={`p-3 rounded-xl border ${PLANT_META[cat].border} ${PLANT_META[cat].badge}`}>
                      <div className="text-sm font-bold">{cat}</div>
                      <div className="text-2xl font-bold">{fileCounts[cat]}</div>
                      <div className="text-xs opacity-75">ไฟล์</div>
                    </div>
                  ))}
                </div>

                {/* File List */}
                {sourceFiles.length > 0 && (
                  <div className="space-y-2 max-h-64 overflow-y-auto">
                    {sourceFiles.map((file, i) => (
                      <div
                        key={file.name}
                        className="flex items-center justify-between bg-slate-50 border border-slate-200 rounded-lg px-4 py-3"
                      >
                        <div className="flex items-center gap-3">
                          <span
                            className={`w-5 h-5 ${
                              file.status === 'done'
                                ? 'text-emerald-500'
                                : file.status === 'error'
                                ? 'text-red-500'
                                : file.status === 'processing'
                                ? 'text-blue-500'
                                : 'text-slate-400'
                            }`}
                          >
                            {file.status === 'done' ? (
                              <Icons.Check />
                            ) : file.status === 'processing' ? (
                              <Icons.Refresh />
                            ) : (
                              <Icons.File />
                            )}
                          </span>
                          <div>
                            <p className="text-sm font-medium">{file.name}</p>
                            <p className="text-xs text-slate-400">
                              {formatSize(file.size)} {file.rowCount > 0 && `• ${file.rowCount} rows`}
                              {file.error && <span className="text-red-500"> • {file.error}</span>}
                            </p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2">
                          <span className={`px-2 py-1 rounded text-xs font-medium ${PLANT_META[file.category].badge}`}>
                            {file.category}
                          </span>
                          <button
                            onClick={() => removeSourceFile(i)}
                            className="w-5 h-5 text-slate-400 hover:text-red-500 transition-colors"
                          >
                            <Icons.Trash />
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
                onClick={processAllFiles}
                disabled={!canProcess}
                className={`w-full py-4 rounded-xl font-semibold text-lg flex items-center justify-center gap-3 transition-all ${
                  canProcess ? 'bg-blue-600 text-white hover:bg-blue-700' : 'bg-slate-200 text-slate-400 cursor-not-allowed'
                }`}
              >
                {processing ? (
                  <>
                    <span className="w-6 h-6">
                      <Icons.Refresh />
                    </span>
                    กำลังประมวลผล...
                  </>
                ) : (
                  <>
                    <span className="w-6 h-6">
                      <Icons.Play />
                    </span>
                    รวมข้อมูลและคำนวณ
                  </>
                )}
              </button>
            </div>
          </div>
        )}

        {/* Preview Tab */}
        {tab === 'preview' && processedData && (
          <div className="space-y-4">
            {/* Verification Card */}
            <div className="bg-white border border-slate-200 rounded-xl p-5">
              <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between mb-4">
                <div>
                  <h3 className="text-lg font-semibold">ตรวจสอบข้อมูลเบื้องต้น</h3>
                  <p className="text-sm text-slate-500">เช็คไฟล์ที่อัปโหลดและผลการประมวลผล</p>
                </div>
                <button
                  onClick={exportToExcel}
                  className="px-5 py-2 text-sm font-medium bg-blue-600 text-white rounded-lg hover:bg-blue-700"
                >
                  {mainWorkbook ? 'ดาวน์โหลด (อัปเดตเทมเพลท)' : 'ดาวน์โหลดไฟล์สรุป'}
                </button>
              </div>

              <div className="grid grid-cols-2 md:grid-cols-4 gap-3 text-sm">
                <div className="border border-slate-200 rounded-lg p-3">
                  <p className="text-slate-500">ไฟล์หลัก (Template)</p>
                  <p className={`font-semibold ${mainFile ? 'text-emerald-600' : 'text-slate-400'}`}>
                    {mainFile ? 'พร้อมใช้เป็นเทมเพลท' : 'ไม่มี (สร้างไฟล์ใหม่)'}
                  </p>
                  {mainFile && <p className="text-xs text-slate-400 truncate">{mainFile.name}</p>}
                </div>
                <div className="border border-slate-200 rounded-lg p-3">
                  <p className="text-slate-500">ไฟล์ประมวลผลสำเร็จ</p>
                  <p className={`font-semibold ${filesProcessed > 0 ? 'text-emerald-600' : 'text-red-600'}`}>
                    {filesProcessed} / {totalSourceFiles} ไฟล์
                  </p>
                </div>
                <div className="border border-slate-200 rounded-lg p-3">
                  <p className="text-slate-500">โรงงานที่ไม่มีไฟล์</p>
                  <p className={`font-semibold ${missingPlants.length > 0 ? 'text-amber-600' : 'text-emerald-600'}`}>
                    {missingPlants.length > 0 ? missingPlants.join(', ') : 'ครบทุกโรงงาน'}
                  </p>
                </div>
                <div className="border border-slate-200 rounded-lg p-3">
                  <p className="text-slate-500">ไฟล์ที่มีข้อผิดพลาด</p>
                  <p className={`font-semibold ${filesWithError > 0 ? 'text-red-600' : 'text-emerald-600'}`}>
                    {filesWithError > 0 ? `${filesWithError} ไฟล์` : 'ไม่มี'}
                  </p>
                </div>
              </div>
            </div>

            {/* Data by Plant */}
            {Object.entries(processedData).map(([sheet, rows]) => {
              const cat = sheet.split(' ')[0];
              const isExpanded = expanded[sheet];

              return (
                <div key={sheet} className="bg-white border border-slate-200 rounded-xl overflow-hidden">
                  <button
                    onClick={() => toggleSection(sheet)}
                    className="w-full px-5 py-4 flex items-center justify-between hover:bg-slate-50 transition-colors"
                  >
                    <div className="flex items-center gap-3">
                      <span className={`px-3 py-1 rounded-lg text-sm font-medium ${PLANT_META[cat].badge}`}>{cat}</span>
                      <span className="font-semibold">{sheet}</span>
                      <span className="text-slate-400">({rows.length} รายการ)</span>
                    </div>
                    <span className="w-5 h-5 text-slate-400">
                      {isExpanded ? <Icons.ChevronUp /> : <Icons.ChevronDown />}
                    </span>
                  </button>

                  {isExpanded && (
                    <div className="px-5 pb-4 overflow-x-auto">
                      {rows.length > 0 ? (
                        <table className="w-full text-sm">
                          <thead>
                            <tr className="text-slate-500 border-b border-slate-200">
                              <th className="text-left py-3 px-3">Part Number</th>
                              <th className="text-left py-3 px-3">Part Code</th>
                              <th className="text-right py-3 px-3">Packing</th>
                              <th className="text-right py-3 px-3 text-blue-600">N</th>
                              <th className="text-right py-3 px-3 text-emerald-600">N+1</th>
                              <th className="text-right py-3 px-3 text-amber-600">N+2</th>
                              <th className="text-right py-3 px-3 text-purple-600">N+3</th>
                            </tr>
                          </thead>
                          <tbody>
                            {rows.slice(0, 50).map((row, idx) => (
                              <tr key={idx} className="border-b border-slate-100 hover:bg-slate-50">
                                <td className="py-2 px-3 font-mono text-blue-600">{row.partNumber}</td>
                                <td className="py-2 px-3 text-slate-600">{row.partCode}</td>
                                <td className="py-2 px-3 text-right text-slate-600">{row.packingSize}</td>
                                <td className="py-2 px-3 text-right font-medium">{formatNumber(row.n)}</td>
                                <td className="py-2 px-3 text-right">{formatNumber(row.n1)}</td>
                                <td className="py-2 px-3 text-right">{formatNumber(row.n2)}</td>
                                <td className="py-2 px-3 text-right">{formatNumber(row.n3)}</td>
                              </tr>
                            ))}
                            {rows.length > 50 && (
                              <tr className="text-slate-400 text-center">
                                <td colSpan="7" className="py-3">
                                  ... และอีก {rows.length - 50} รายการ
                                </td>
                              </tr>
                            )}
                          </tbody>
                        </table>
                      ) : (
                        <p className="py-4 text-center text-slate-400">ไม่มีข้อมูล</p>
                      )}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}

        {/* Summary Tab */}
        {tab === 'summary' && summaryData && (
          <div className="space-y-6">
            {/* Summary Cards */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              {[
                { label: 'N (Feb)', value: totals.n, color: 'bg-blue-600' },
                { label: 'N+1 (Mar)', value: totals.n1, color: 'bg-emerald-600' },
                { label: 'N+2 (Apr)', value: totals.n2, color: 'bg-amber-500' },
                { label: 'N+3 (May)', value: totals.n3, color: 'bg-purple-600' },
              ].map((card) => (
                <div key={card.label} className={`${card.color} rounded-2xl p-5 text-white`}>
                  <p className="text-white/80 text-sm">{card.label}</p>
                  <p className="text-3xl font-bold">{formatNumber(card.value)}</p>
                  <p className="text-white/60 text-xs mt-1">ชิ้น</p>
                </div>
              ))}
            </div>

            {/* Summary Table */}
            <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden">
              <div className="px-6 py-4 border-b border-slate-200">
                <h3 className="font-semibold flex items-center gap-2">
                  <span className="w-5 h-5 text-blue-600">
                    <Icons.Chart />
                  </span>
                  สรุปยอดรวมตาม Part Number ({summaryData.length} รายการ)
                </h3>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="bg-slate-50">
                      <th className="text-left py-3 px-4 font-medium">Part Number</th>
                      <th className="text-left py-3 px-4 font-medium">Plants</th>
                      <th className="text-right py-3 px-4 font-medium text-blue-600">N</th>
                      <th className="text-right py-3 px-4 font-medium text-emerald-600">N+1</th>
                      <th className="text-right py-3 px-4 font-medium text-amber-600">N+2</th>
                      <th className="text-right py-3 px-4 font-medium text-purple-600">N+3</th>
                      <th className="text-right py-3 px-4 font-medium">Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {summaryData.map((row) => (
                      <tr key={row.partNumber} className="border-b border-slate-100 hover:bg-slate-50">
                        <td className="py-3 px-4 font-mono text-blue-600">{row.partNumber}</td>
                        <td className="py-3 px-4">
                          <div className="flex gap-1 flex-wrap">
                            {row.plants.split(', ').map((p) => (
                              <span key={p} className={`px-2 py-0.5 rounded text-xs ${PLANT_META[p]?.badge || ''}`}>
                                {p}
                              </span>
                            ))}
                          </div>
                        </td>
                        <td className="py-3 px-4 text-right text-blue-600 font-medium">{formatNumber(row.n)}</td>
                        <td className="py-3 px-4 text-right text-emerald-600">{formatNumber(row.n1)}</td>
                        <td className="py-3 px-4 text-right text-amber-600">{formatNumber(row.n2)}</td>
                        <td className="py-3 px-4 text-right text-purple-600">{formatNumber(row.n3)}</td>
                        <td className="py-3 px-4 text-right font-bold">
                          {formatNumber(row.n + row.n1 + row.n2 + row.n3)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot>
                    <tr className="bg-slate-50 font-bold">
                      <td className="py-4 px-4">Grand Total</td>
                      <td className="py-4 px-4 text-slate-400">{summaryData.length} Parts</td>
                      <td className="py-4 px-4 text-right text-blue-600">{formatNumber(totals.n)}</td>
                      <td className="py-4 px-4 text-right text-emerald-600">{formatNumber(totals.n1)}</td>
                      <td className="py-4 px-4 text-right text-amber-600">{formatNumber(totals.n2)}</td>
                      <td className="py-4 px-4 text-right text-purple-600">{formatNumber(totals.n3)}</td>
                      <td className="py-4 px-4 text-right">
                        {formatNumber(totals.n + totals.n1 + totals.n2 + totals.n3)}
                      </td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>

            {/* Download Button */}
            <button
              onClick={exportToExcel}
              className="w-full py-4 bg-blue-600 rounded-xl font-semibold text-lg text-white flex items-center justify-center gap-3 hover:bg-blue-700 transition-all"
            >
              <span className="w-6 h-6">
                <Icons.Download />
              </span>
              {mainWorkbook ? 'ดาวน์โหลดไฟล์ (อัปเดตเทมเพลท)' : 'ดาวน์โหลดไฟล์ Excel สรุป'}
            </button>
            {mainWorkbook && (
              <p className="text-center text-sm text-slate-500 mt-2">
                * ข้อมูลจะถูกแทนที่ลงในไฟล์เทมเพลทที่อัปโหลด
              </p>
            )}
          </div>
        )}

        {/* Empty State */}
        {(tab === 'preview' || tab === 'summary') && !processedData && (
          <div className="text-center py-20 bg-white border border-slate-200 rounded-xl">
            <span className="w-20 h-20 mx-auto text-slate-300 block mb-4">
              <Icons.Alert />
            </span>
            <h3 className="text-xl text-slate-400 mb-2">ยังไม่มีข้อมูล</h3>
            <p className="text-slate-400 mb-4">กรุณาอัปโหลดไฟล์และประมวลผลก่อน</p>
            <button onClick={() => setTab('upload')} className="px-6 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700">
              ไปหน้าอัปโหลด
            </button>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="border-t border-slate-200 mt-12 py-6 bg-white">
        <div className="max-w-7xl mx-auto px-4 text-center text-slate-400 text-sm">
          Production Data Consolidator v1.0 • TMT Camera Production Plan Summary System
        </div>
      </footer>
    </div>
  );
}
