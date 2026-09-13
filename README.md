# DE TEAM — ระบบเช็คชื่อพนักงานและสรุปค่าจ้าง
### Employee Attendance & Payroll Management Web Application

เว็บแอปพลิเคชันเช็คชื่อพนักงานรายวันและคำนวณค่าจ้างรายเดือน ออกแบบตามหลัก **Mobile-First UX/UI** สำหรับร้านและธุรกิจขนาดเล็ก-กลาง เชื่อมต่อฐานข้อมูล **Google Sheets** ผ่าน **Google Apps Script (GAS)**

🌐 **ทดลองใช้งานหน้าเว็บจริง (Live Demo):**  
👉 [https://tleza123.github.io/employee-attendance-web/](https://tleza123.github.io/employee-attendance-web/)

---

## 📱 จุดเด่นของหน้าเว็บ (Web UI Highlights)

- **Mobile-First & Touch-Friendly:** ออกแบบปุ่มกดและเป้าสัมผัสขนาดใหญ่ (ขั้นต่ำ 44×44 pt / 3.467rem) ป้องกันการกดผิดบนสมาร์ทโฟน
- **Fluid Typography:** ตัวหนังสืออ่านง่ายชัดเจน ปรับขนาดอัตโนมัติตามหน้าจอ (15px – 17.5px) รองรับฟอนต์ Sarabun และ Tahoma แบบมีหัว
- **Safe Area Support:** รองรับขอบจอด้านล่างและรอยบากของ iPhone / Android (`env(safe-area-inset-bottom)`)
- **Client-Side Photo Compression:** ระบบย่อรูปถ่ายหน้างานผ่าน HTML5 Canvas อัตโนมัติ (ไม่เกิน 192px, ขนาดไฟล์ <= 24 KiB) ก่อนส่งขึ้นเซิร์ฟเวอร์
- **Interactive Single-Page Application (SPA):** สลับหน้าได้ลื่นไหล ไม่ต้องรีโหลดทั้งหน้า พร้อมระบบแจ้งเตือนแบบ Live Region สำหรับการเข้าถึง (Accessibility)

---

## 📂 โครงสร้างโปรเจกต์ (Repository Structure)

```text
employee-attendance-web/
├── index.html                   # หน้าเว็บหลักแบบ Standalone (พร้อมจำลอง Mock Server ใช้งานได้ทันทีบน GitHub Pages / Vercel)
├── web/                         # ไฟล์ฝั่งหน้าเว็บสำหรับติดตั้งบน Google Apps Script
│   ├── Index.html               # โครงร่าง HTML และแท็บการทำงานหลัก (Layout Shell)
│   ├── Styles.html              # สไตล์ CSS ทั้งหมด (Design System Tokens)
│   ├── Fonts.html               # การตั้งค่าฟอนต์และการแสดงผลภาษาไทย
│   ├── App.html                 # ตรรกะฝั่งไคลเอนต์, State Controller, Canvas Compressor
│   └── preview-local.html       # ตัวอย่างพรีวิวแบบไฟล์เดี่ยว
├── gas/                         # ไฟล์ระบบหลังบ้าน Google Apps Script (Backend)
│   ├── Code.gs                  # จุดเริ่มต้นเว็บแอปและฟังก์ชันดูหน้าเว็บ (doGet)
│   ├── Auth.gs                  # ตรวจสอบสิทธิ์การเข้าถึง (เฉพาะอีเมลเจ้าของระบบ)
│   ├── Config.gs                # การตั้งค่าระบบและกุญแจความปลอดภัย (Script Properties)
│   ├── Schema.gs                # กำหนดโครงสร้างตารางและหัวคอลัมน์ใน Google Sheets
│   ├── Repository.gs            # จัดการการอ่าน-เขียนข้อมูลใน Google Sheets พร้อม LockService
│   ├── AttendanceService.gs     # บริการเช็คชื่อพนักงานรายวัน (เต็มวัน / ครึ่งวัน / ไม่มา)
│   ├── EmployeeService.gs       # จัดการข้อมูลพนักงาน, บันทึกวันสิ้นสุดการทำงาน, ประวัติค่าแรง
│   ├── ExtrasService.gs         # จัดการเงินพิเศษเทมเพลต และเงินพิเศษประจำเดือน
│   ├── MonthService.gs          # จัดการรอบเดือน, ตรวจสอบเงื่อนไข, ปิดรอบและเปิดรอบใหม่
│   ├── Payroll.gs               # คำนวณค่าจ้างตามวันทำงานและประวัติอัตราค่าแรง
│   ├── PhotoService.gs          # จัดการรูปถ่ายพนักงานในชีต Photos
│   ├── CalendarService.gs       # จัดการวันทำงาน วันหยุดประจำสัปดาห์ และวันหยุดพิเศษ
│   ├── ReportService.gs         # สร้างรายงานสรุปรายเดือนและรายงานรายบุคคล
│   └── appsscript.json          # Manifest การตั้งค่าสิทธิ์ของ Apps Script
└── tests/                       # ชุดทดสอบระบบอัตโนมัติ (Test Suite)
    └── test-suite.cjs           # Unit Tests ครอบคลุม 38 กรณีทดสอบ (100% Pass)
```

---

## 🚀 การนำไปใช้งาน (Deployment)

### 1. ใช้งานเป็นหน้าเว็บสาธารณะ (GitHub Pages / Vercel)
Repository นี้มี `index.html` อยู่ที่ Root โฟลเดอร์ สามารถนำไปโฮสต์เพื่อเปิดดูหน้าตาและทดสอบระบบได้ทันที:
- **GitHub Pages:** เปิดใช้งานที่ Settings > Pages > Branch `main` > Save
- **Vercel / Netlify:** เชื่อมต่อ GitHub repo นี้แล้วกด Deploy ได้ทันทีโดยไม่ต้องตั้งค่า Build Command

### 2. ติดตั้งใช้งานจริงกับ Google Sheets & Google Apps Script
1. สร้าง **Google Sheets** เปล่า 1 ไฟล์
2. ไปที่เมนู **ส่วนขยาย (Extensions) > Apps Script**
3. คัดลอกไฟล์ทั้งหมดในโฟลเดอร์ `gas/` (`.gs` และ `appsscript.json`) ไปวางในโปรเจกต์ Apps Script
4. คัดลอกไฟล์ทั้งหมดในโฟลเดอร์ `web/` (`Index.html`, `Styles.html`, `Fonts.html`, `App.html`) ไปวางเป็นไฟล์ HTML ในโปรเจกต์
5. ไปที่ **การตั้งค่าโปรเจกต์ (Project Settings) > คุณสมบัติของสคริปต์ (Script Properties)**:
   - เพิ่ม `SPREADSHEET_ID` = ไอดีชีตของคุณ
   - เพิ่ม `OWNER_EMAIL` = อีเมล Google ของเจ้าของระบบ
   - เพิ่ม `ADMIN_KEY` = รหัสลับความปลอดภัย
6. กด **การทำให้ใช้งานได้ (Deploy) > การทำให้ใช้งานได้รายการใหม่ (New Deployment)**
   - เลือกประเภท: **เว็บแอป (Web App)**
   - Execute as: **ฉัน (Me)**
   - Who has access: **เฉพาะฉัน (Only myself)**
7. นำ URL เว็บแอปที่ได้ไปเปิดใช้งานบนมือถือหรือบันทึกเป็นไอคอนบนหน้าจอโฮม (Add to Home Screen)

---

## 🧪 การทดสอบระบบ (Automated Tests)

รันชุดทดสอบผ่าน Node.js ในเครื่องคอมพิวเตอร์:
```bash
node tests/test-suite.cjs
```
ผลลัพธ์ผ่านครบ 38/38 รายการ (ครอบคลุมการเช็คชื่อ, การคำนวณเงินสตางค์, กฎการปิดรอบเดือน, และการย่อรูป)

---

## 📄 ลิขสิทธิ์ (License)
MIT License — พัฒนาสำหรับใช้งานในระบบจัดการภายในองค์กร DE TEAM
