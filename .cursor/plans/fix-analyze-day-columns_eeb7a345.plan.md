---
name: fix-analyze-day-columns
overview: แก้ให้คอลัมน์รายวัน (1–31) ของ Dec/Jan/Feb/Mar ในชีท `Analyze` ถูกเติมครบ 100% โดยอ่านหัวตารางของ source และ Analyze แบบ robust (รองรับ merged/ช่องว่าง) แล้วคัดลอกค่าแบบ key-to-key พร้อม diagnostics ตรวจสอบ mapping.
todos: []
isProject: false
---

# แผนแก้: คอลัมน์วัน (1–31) ใน Analyze ไม่มา

## เป้าหมาย

- เติมข้อมูลในชีท `Analyze` ให้ครบตามเทมเพลท: **Dec/Jan/Feb/Mar** รวมทั้ง `N`/`N+1`/`N+2`/`N+3` และวัน `1..31`
- คัดลอกจากไฟล์รายโรงงานแบบ **1:1** (ไม่คำนวณเอง)
- ไม่กระทบสูตร/shared-formula ในเทมเพลท และยังคงไฮไลท์ + กรอบสำหรับเซลล์ที่ระบบเขียน

## อินพุตที่ต้องใช้ (ผู้ใช้ยืนยันว่าจะอัปโหลด)

- ไฟล์ผลลัพธ์ที่ดาวน์โหลดจากระบบ **.xlsx** (ไม่ใช่ .html)
- ไฟล์รายโรงงานอย่างน้อย 1 ไฟล์ (เช่น `BP veh 481D.xls/xlsx`)

## สาเหตุที่เป็นไปได้

- ระบบหา “แถวชื่อเดือน/แถววัน” ของไฟล์ source ผิดตำแหน่ง (เพราะ header มีหลายบรรทัด/merge ทำให้ช่องถัดไปว่าง)
- ระบบสร้าง mapping แล้วแต่ key ไม่ตรงกัน (เช่น เดือน/วันมีช่องว่าง, ตัวเลขเป็น string, หรือเดือนอยู่คนละแถว)
- ระบบไม่ควรไปเขียนทับเซลล์สูตร/clone (shared formula) ใน Analyze ทำให้ข้ามการเขียนบางคอลัมน์

## วิธีแก้ (โค้ด)

ไฟล์หลักที่จะปรับ: [`c:\Users\Administrator\Downloads\excell\src\App.jsx`](c:\Users\Administrator\Downloads\excell\src\App.jsx)

### 1) ทำให้การหา header ของ Analyze robust

- สแกนหา “แถวเดือน” และ “แถว subheader” ใน `Analyze` โดย **scoring**:
- แถวเดือน: มีคำ `Dec/Jan/Feb/Mar` มากที่สุด
- แถว subheader: มี `N/N+1/N+2/N+3` และเลข `1..31` มากที่สุด
- รองรับ merge: ใช้ `cell.master.value` เมื่อ cell ว่าง
- สร้าง `analyzeDestMap["MONTH|SUB"] = destCol`

### 2) ทำให้การหา header ของ source robust

- สแกน 1–30 แถวแรกของ source เพื่อหาแถวเดือน/แถว subheader แบบ scoring เหมือนกัน
- ทำ **forward-fill** เดือนจากซ้ายไปขวา (เพราะ merge ทำให้ช่องถัดไปว่าง)
- สร้าง `sourceMap["MONTH|SUB"] = srcColIndex`

### 3) Diagnostics เพื่อยืนยัน mapping ก่อนเขียน

- แสดง (ใน UI หรือ console) ค่า:
- จำนวน key ที่ Analyze ต้องการ
- จำนวน key ที่ source map หาได้
- รายการ key ที่หาย (เช่น `JAN|15`)
- ถ้า missing keys เยอะ ให้หยุดการเขียน Analyze และแจ้ง error ที่อ่านง่าย

### 4) คัดลอกค่าแบบ key-to-key ลง Analyze

- สำหรับแต่ละ row ใน `Analyze` (ต่อ plant/part):
- วนทุก key ใน `analyzeDestMap`
- หา `srcCol` จาก `sourceMap` แล้วอ่าน `rawRow[srcCol]`
- เขียนลง `Analyze[destCol] `ด้วย `safeSetCellValue` (ไม่ทับสูตร)
- apply highlight+border เฉพาะเซลล์ที่เขียนจริง

### 5) ความปลอดภัยสูตร/shared-formula

- คง helper `isFormulaCell` / `safeClearCell` / `safeSetCellValue`
- ล้างเฉพาะคอลัมน์ที่ระบบเขียน และข้ามเซลล์สูตรเสมอ

## Test plan

- ใช้ไฟล์ที่ผู้ใช้อัปโหลด:
- รันในเว็บ: อัปโหลดเทมเพลท + อัปโหลดไฟล์ BP/BPK/GW/SR → Process → Download
- เปิดผลลัพธ์ ตรวจชีท `Analyze`:
- Dec/Jan/Feb/Mar มีวัน 1..31 ครบและค่าตรงกับไฟล์ source
- เซลล์ที่ระบบเติมมีสี/กรอบ
- ไม่มี error shared-formula ตอนดาวน์โหลด