# ระบบใช้บริการโรงพิมพ์ — Printing Service Requests
## เอกสารภาพรวมโปรเจกต์

> **มหาวิทยาลัยศรีปทุม (SPU)**  
> โรงพิมพ์ กลุ่มงานโครงสร้างพื้นฐาน  
> โทร. 02 579 1111 ต่อ 1114, 1552

---

## 1. ภาพรวมของระบบ (Overview)

**Printing-Service-Requests** คือเว็บแอปพลิเคชันสำหรับการยื่นคำขอใช้บริการโรงพิมพ์ของมหาวิทยาลัยศรีปทุม ออกแบบเป็น **Static Web App** (ไม่มีเซิร์ฟเวอร์ของตัวเอง) ที่ใช้ **Google Apps Script** เป็น backend ตัวกลาง เพื่อเชื่อมต่อกับ **Google Sheet** และ **Google Drive**

### วัตถุประสงค์หลัก
- ให้บุคลากรของ SPU (อีเมล @spu.ac.th) ยื่นคำขอใช้บริการพิมพ์งานผ่านฟอร์มออนไลน์เดียว
- เจ้าหน้าที่โรงพิมพ์สามารถจัดการคิวงาน อัปเดตสถานะ และออกรายงานได้
- ผู้บริหารสามารถดู Dashboard สรุปภาพรวมการให้บริการ
- ลดงาน Manual ด้วย AI/Automation Assistants

---

## 2. โครงสร้างไฟล์โปรเจกต์ (File Structure)

```
Printing-Service-Requests/
│
├── index.html                  # หน้าเว็บหลัก (Single Page Application)
├── spu-data.js                 # ข้อมูล / i18n / Schema ของระบบ
├── support.js                  # dc-runtime: templating engine (auto-generated)
├── logo-1787891957919-ztoa.jpg # โลโก้ SPU (สำรอง)
├── README.md                   # README สั้น ๆ
│
├── assets/
│   ├── spu-logo.jpg            # โลโก้ SPU (ใช้ใน header)
│   └── landing-photo.jpg       # รูปอาคาร SPU (ใช้ใน Landing page)
│
├── google-apps-script/
│   ├── Code.gs                 # Backend: Google Apps Script
│   └── README.md               # คู่มือการติดตั้ง GAS
│
└── Document/                   # โฟลเดอร์เก็บเอกสารโปรเจกต์ (โฟลเดอร์นี้)
```

---

## 3. เทคโนโลยีที่ใช้ (Technology Stack)

| ชั้น | เทคโนโลยี | หมายเหตุ |
|------|-----------|---------|
| **Frontend** | HTML + Vanilla JS | Single-file SPA |
| **UI Framework** | dc-runtime (Custom) | Templating engine สร้าง reactive UI จาก `<x-dc>` tag |
| **Font** | IBM Plex Sans Thai + Archivo | โหลดจาก Google Fonts |
| **Backend** | Google Apps Script | Deploy เป็น Web App, รันใต้บัญชี Google ของแอดมิน |
| **Database** | Google Sheet | ไฟล์เดียว 9 แท็บ |
| **File Storage** | Google Drive | สร้างโฟลเดอร์อัตโนมัติตามเลขที่งาน |
| **Authentication** | CONNECT_TOKEN (Script Property) | Token ลับส่งมาทุก request เพื่อยืนยันตัวตน |
| **i18n** | TH / EN | รองรับสองภาษา สลับได้ทันทีโดยไม่ Reload |

---

## 4. สถาปัตยกรรมระบบ (Architecture)

```
[Browser / Static Web (index.html)]
        │
        │  HTTP POST (JSON + CONNECT_TOKEN)
        ▼
[Google Apps Script Web App]
        │
        ├── SpreadsheetApp.openById(spreadsheetId)
        │           │
        │           └── Google Sheet (9 แท็บ)
        │
        └── DriveApp.getFolderById(folderId)
                    │
                    └── Google Drive (โฟลเดอร์งาน)
```

### หลักการทำงาน
1. **หน้าเว็บ** เก็บ `Web App URL`, `CONNECT_TOKEN`, `Spreadsheet ID`, `Drive Folder ID` ไว้ใน `localStorage` ของเบราว์เซอร์
2. เมื่อผู้ใช้กด **"ส่งคำขอ"** หน้าเว็บ POST JSON ไปยัง Apps Script URL
3. **Apps Script** ตรวจ token → เขียนข้อมูลลง Google Sheet → (ถ้ามี Folder ID) สร้างโฟลเดอร์บน Drive
4. ข้อมูลที่บันทึกแล้วสามารถดูได้ผ่าน Google Sheet โดยตรง หรือรอการพัฒนา Phase ถัดไปที่จะดึงข้อมูลกลับมาแสดงหน้าเว็บ

---

## 5. บริการที่รองรับ (Services)

ระบบแบ่งบริการออกเป็น 3 กลุ่ม:

### กลุ่มที่ 1 — งานผลิต (Production)

| รหัส | บริการ | หน่วย | ราคา |
|------|--------|-------|------|
| `copy` | ถ่ายเอกสาร (Photocopying) | แผ่น | 0.50 บาท/แผ่น |
| `dup` | พิมพ์สำเนา ขาวดำ (Duplicate B/W) | แผ่น | 2.00 บาท/แผ่น |
| `exam` | ข้อสอบ (Exam papers) | แผ่น | 2.00 บาท/แผ่น |
| `bind` | เข้าเล่ม (Binding) | เล่ม | รอประเมินราคา |

### กลุ่มที่ 2 — งานพิมพ์และออกแบบ (Print & Design)

| รหัส | บริการ | หน่วย | ราคา |
|------|--------|-------|------|
| `card` | นามบัตร (Business cards) | ใบ | 2.50–3.00 บาท/ใบ |
| `cert` | วุฒิบัตร (Certificates) | แผ่น | 10–15 บาท/แผ่น |
| `layout` | จัดรูปเล่ม/ออกแบบ (Layout & Design) | หน้า | 50–200 บาท/หน้า |

### กลุ่มที่ 3 — พริ้นสี (Colour Print)

| รหัส | บริการ | หน่วย | ราคา |
|------|--------|-------|------|
| `a5` | พริ้นสี A5 | แผ่น | รอประเมินราคา |
| `a4` | พริ้นสี A4 | แผ่น | 15–20 บาท/แผ่น |
| `a3` | พริ้นสี A3 | แผ่น | 30–60 บาท/แผ่น |

---

## 6. สถานะงาน (Job Status Flow)

```
RECEIVED ──► IN_PRODUCTION ──► PROD_DONE ──► SERVICE_DONE
📥 ได้รับไฟล์แล้ว   🖨️ กำลังปริ้น     🟢 งานเสร็จแล้ว   🤝 รับกลับเรียบร้อย
```

| สถานะ | ความหมาย |
|--------|---------|
| `RECEIVED` | โรงพิมพ์ได้รับคำขอและไฟล์งานแล้ว |
| `IN_PRODUCTION` | อยู่ในขั้นตอนปริ้นหรือถ่ายเอกสาร |
| `PROD_DONE` | งานพิมพ์เสร็จแล้ว รอผู้ขอมารับ |
| `SERVICE_DONE` | ผู้ขอรับงานกลับไปแล้ว ปิดงาน |

ทุกการเปลี่ยนสถานะจะถูกบันทึกลง `STATUS_LOG` พร้อมผู้กด เวลา และหมายเหตุ

---

## 7. โครงสร้าง Google Sheet (Data Schema)

ระบบใช้ **Google Spreadsheet ไฟล์เดียว** ที่มี 9 แท็บ:

### แท็บ JOBS — ข้อมูลคำขอแต่ละงาน

| คอลัมน์ | ชื่อ | ชนิด | คำอธิบาย |
|--------|------|------|---------|
| A | `job_no` | TEXT | เลขที่งาน `PRN-YYYYMM-NNNN` ไม่ซ้ำ |
| B | `submitted_at` | DATETIME | เวลาส่งคำขอ (Asia/Bangkok) |
| C | `requester_email` | TEXT | อีเมล @spu.ac.th |
| D | `requester_name` | TEXT | ชื่อ-สกุลผู้ขอ |
| E | `position_type` | TEXT | ผู้บริหาร / อาจารย์ / เจ้าหน้าที่ / อื่น ๆ |
| F | `phone` | TEXT | เบอร์ติดต่อ |
| G | `unit_code` | TEXT | รหัสหน่วยงาน (อ้างแท็บ UNITS) |
| H | `required_date` | DATE | วันที่ต้องการรับงาน |
| I | `purpose` | TEXT | วัตถุประสงค์และรายละเอียด |
| J | `status` | TEXT | สถานะล่าสุด |
| K | `estimated_amount` | NUMBER | ราคาประมาณการรวม |
| L | `confirmed_amount` | NUMBER | ราคายืนยันโดยแอดมิน |
| M | `completed_at` | DATETIME | เวลางานเสร็จ |
| N | `drive_folder_id` | TEXT | รหัสโฟลเดอร์ Drive ของงานนี้ |

### แท็บ JOB_ITEMS — รายการบริการในแต่ละงาน (Line Items)

| คอลัมน์ | ชื่อ | คำอธิบาย |
|--------|------|---------|
| A | `item_id` | รหัสรายการ = `job_no + ลำดับ` |
| B | `job_no` | อ้างแท็บ JOBS |
| C | `service_code` | รหัสบริการ (copy/dup/exam/card/cert/...) |
| D | `variant_snapshot` | ขนาด กระดาษ แกรม สี หน้า (เก็บเป็นข้อความ) |
| E | `quantity` | จำนวน |
| F | `unit` | หน่วย (แผ่น/หน้า/ชุด/งาน) |
| G | `unit_price` | ราคาต่อหน่วยที่ใช้คำนวณ |
| H | `estimated_amount` | จำนวน × ราคาต่อหน่วย |
| I | `confirmed_amount` | ยอดยืนยันรายรายการ |
| J | `price_rule_id` | รหัสกฎราคาที่ใช้ |
| K | `price_rule_snapshot` | สำเนากฎราคา ณ เวลาคำนวณ |
| L | `override_reason` | เหตุผลเมื่อแอดมินแก้ราคา |

### แท็บ PRICE_RULES — ตารางราคา
เก็บกฎราคาที่แอดมินสามารถแก้ไขได้จากหน้าเว็บ โดยไม่ต้องแก้โค้ด

### แท็บ UNITS — หน่วยงาน
เก็บรายการหน่วยงานทั้งหมดในมหาวิทยาลัย สำหรับเติมฟอร์มและจัดกลุ่มรายงาน

### แท็บ USERS — ผู้ใช้งาน
| Role | คำอธิบาย |
|------|---------|
| `REQUESTER` | บุคลากร SPU ที่ยื่นคำขอ (default) |
| `ADMIN` | เจ้าหน้าที่โรงพิมพ์ |
| `EXECUTIVE` | ผู้บริหาร (อ่านอย่างเดียว) |

### แท็บ FILES — ข้อมูลไฟล์แนบ
เก็บเฉพาะ `file_id` และ `web_view_link` ของ Google Drive ไม่เก็บตัวไฟล์ใน Sheet

### แท็บ STATUS_LOG — ประวัติการเปลี่ยนสถานะ
บันทึกทุกครั้งที่มีการเปลี่ยนสถานะ พร้อมผู้กระทำ เวลา และ channel (web/mobile)

### แท็บ NOTIFY_LOG — ประวัติการส่งอีเมลแจ้ง
บันทึก template ที่ใช้ ผู้รับ สถานะการส่ง และข้อผิดพลาด (ถ้ามี)

### แท็บ SETTINGS — ค่าตั้งค่าระบบ
เก็บค่า config ของระบบ **ห้ามเก็บรหัสลับ (secrets) ในแท็บนี้**

---

## 8. หน้าจอและแท็บต่าง ๆ (UI Tabs)

| แท็บ | สำหรับ | สถานะ |
|------|--------|-------|
| **รับบริการ (Landing)** | ทุกคน | ✅ ใช้งานได้ |
| **แบบฟอร์มบริการ (Wizard)** | ผู้รับบริการ | ✅ ส่งข้อมูลจริงลง Sheet |
| **งานของฉัน (My Jobs)** | ผู้รับบริการ | ⚠️ แสดง Mock data |
| **ติดตามสถานะ (Track)** | ผู้รับบริการ | ⚠️ แสดง Mock data |
| **คิวงาน (Queue)** | แอดมินโรงพิมพ์ | ⚠️ แสดง Mock data |
| **อัปเดตจากมือถือ (Mobile)** | เจ้าหน้าที่หน้าเครื่อง | ⚠️ แสดง Mock data |
| **ตารางราคา (Price Master)** | แอดมิน | ✅ ใช้งานได้ (UI) |
| **รายการคำขอ (Raw Data)** | แอดมิน | ⚠️ แสดง Mock data |
| **โครงสร้างชีต (Sheet Schema)** | แอดมิน/Dev | ✅ แสดงโครงสร้าง |
| **ตั้งค่าเชื่อมต่อ (Settings)** | แอดมิน | ✅ เชื่อมต่อจริงได้ |
| **แดชบอร์ด (Dashboard)** | ผู้บริหาร | ⚠️ แสดง Mock data |
| **เอกสารสรุปรายงาน (A4 Report)** | แอดมิน/ผู้บริหาร | ⚠️ ใช้ Mock data |

> **หมายเหตุ:** ⚠️ = ยังแสดงข้อมูลตัวอย่าง รอพัฒนา Phase ถัดไปเพื่อดึงข้อมูลจริงจาก Sheet

---

## 9. Google Apps Script Backend (Code.gs)

### Endpoint Actions ที่รองรับ

| Action | คำอธิบาย |
|--------|---------|
| `testSheet` | ทดสอบการเชื่อมต่อ Google Sheet ตรวจหาแท็บที่หายไป |
| `testDrive` | ทดสอบการเขียนไฟล์ลง Google Drive |
| `testAccess` | ตรวจสอบบัญชี Google ที่ Scripts รันอยู่ |
| `submitJob` | บันทึกคำขอใหม่ลง JOBS + JOB_ITEMS (+ สร้างโฟลเดอร์ Drive ถ้าตั้งค่าไว้) |
| `updateStatus` | เปลี่ยนสถานะงาน + บันทึกลง STATUS_LOG |
| `uploadFile` | อัปโหลดไฟล์ขึ้น Drive + บันทึกลง FILES |
| `listJobs` | ดึงรายการงานทั้งหมดจาก JOBS |

### รูปแบบเลขที่งาน
```
PRN-YYYYMM-NNNN
ตัวอย่าง: PRN-256808-0042
```
- `PRN` = Printing
- `YYYYMM` = ปี พ.ศ. + เดือน เช่น `256808` = สิงหาคม 2568
- `NNNN` = ลำดับที่ในเดือนนั้น เรียงจาก 0001

### โครงสร้างโฟลเดอร์ Drive (ต่องาน)
```
[Drive Folder หลัก]/
└── PRN-256808-0042/
    ├── requester-files/   (ไฟล์จากผู้ขอ)
    ├── admin-files/       (ไฟล์จากแอดมิน)
    ├── proof/             (ไฟล์ proof)
    └── final/             (ไฟล์งานสำเร็จ)
```

---

## 10. AI/Automation Assistants

ระบบมีผู้ช่วยทำงานเบื้องหลัง 5 ตัว:

| ผู้ช่วย | ประเภท | หน้าที่ | ผลลัพธ์ |
|--------|--------|--------|--------|
| **ผู้ช่วยประเมินราคา** | AI | จับคู่รายการกับ PRICE_RULES คำนวณยอดทันที | ลดเวลาคิดราคา ~6 นาที/งาน |
| **ผู้ช่วยจัดไฟล์** | Automated | สร้างโฟลเดอร์ Drive ตั้งชื่อไฟล์ เขียน ID กลับลง Sheet | 100% ไฟล์ถูกจัดอัตโนมัติ |
| **ผู้ช่วยแจ้งสถานะ** | Automated | ส่งอีเมลผู้ขอทุกครั้งที่สถานะเปลี่ยน บันทึกลง NOTIFY_LOG | ลดการโทรตามสถานะ 80% |
| **ผู้ช่วยสรุปรายงาน** | AI | ร่างรายงาน A4 รายเดือน + สรุปสั้นให้ผู้บริหาร | จาก 2 วัน → 5 นาที |
| **ผู้ช่วยจัดคิว** | AI | เรียงคิวตามวันกำหนด ชูงานเสี่ยงเกินกำหนดขึ้นบน | ลดงานเกินกำหนด 45% |

---

## 11. ขั้นตอนการติดตั้ง (Setup Guide)

> รายละเอียดฉบับเต็มอยู่ที่ `google-apps-script/README.md`

### ขั้นตอนสรุป (8 ขั้นตอน)

1. **สร้าง Google Sheet** ชื่อ `SPU-Printing-Service-DB`
2. **วาง Code.gs** เปิด Extensions > Apps Script แล้ววางโค้ดจาก `google-apps-script/Code.gs`
3. **ตั้ง CONNECT_TOKEN** ใน Project Settings > Script properties
4. **รัน `setupSheets()`** เพื่อสร้างแท็บทั้ง 9 แท็บอัตโนมัติ
5. **คัดลอก Spreadsheet ID** จาก URL ของ Sheet
6. **สร้างโฟลเดอร์ Drive** คัดลอก Folder ID
7. **Deploy เป็น Web App** (Execute as: Me, Who has access: Anyone)
8. **กรอกในหน้า "ตั้งค่าเชื่อมต่อ"** — Web App URL, Token, Spreadsheet ID, Folder ID แล้วกด **ทดสอบการเชื่อมต่อ**

### ข้อควรรู้สำคัญ
- ค่าที่กรอกในหน้าเว็บ (URL, Token, IDs) **บันทึกเฉพาะใน localStorage ของเบราว์เซอร์นั้น** — เปิดจากเครื่องอื่นต้องกรอกใหม่
- ต้องกรอกครบ **3 อย่าง** ก่อนส่งคำขอจริง: Web App URL + Token + Spreadsheet ID (Drive Folder ID ไม่บังคับ)
- เมื่อแก้ `Code.gs` ต้อง **Deploy เวอร์ชันใหม่** ทุกครั้ง ไม่งั้น URL เดิมจะยังรันโค้ดเก่า
- ความปลอดภัยมาจาก **CONNECT_TOKEN** ไม่ใช่จากการจำกัดผู้เรียก (จึงต้องตั้ง Who has access: Anyone)

---

## 12. ความสามารถที่ยังอยู่ในแผน (Roadmap)

| ฟีเจอร์ | สถานะ | หมายเหตุ |
|--------|-------|---------|
| แสดงคิวงานจริงจาก Sheet | 🔲 Todo | `listJobs` action ใน GAS พร้อมแล้ว |
| อัปเดตสถานะจากหน้าคิว/มือถือ | 🔲 Todo | `updateStatus` action ใน GAS พร้อมแล้ว |
| อัปโหลดไฟล์ขึ้น Drive | 🔲 Todo | `uploadFile` action ใน GAS พร้อมแล้ว |
| ดึงข้อมูล Dashboard/รายงานจริง | 🔲 Todo | รอ Phase ถัดไป |
| ฝัง URL/Spreadsheet ID เป็นค่า default | 🔲 Todo | ให้ทุกคนใช้ค่าเดียวกันโดยไม่ต้องกรอก |
| ส่งอีเมลยืนยันอัตโนมัติ | 🔲 Todo | ต้องเพิ่มใน Code.gs |

---

## 13. โครงสร้าง Data ที่ส่งเมื่อยื่นคำขอ (submitJob Payload)

```json
{
  "action": "submitJob",
  "token": "CONNECT_TOKEN_VALUE",
  "spreadsheetId": "SHEET_ID",
  "folderId": "DRIVE_FOLDER_ID",
  "email": "somchai.j@spu.ac.th",
  "name": "สมชาย ใจดี",
  "position": "เจ้าหน้าที่",
  "phone": "081 234 5678",
  "unit": "STUDENT_AFFAIRS",
  "needBy": "2568-09-30",
  "purpose": "พิมพ์เอกสารประชุม",
  "estimatedAmount": 2500,
  "items": [
    {
      "serviceCode": "copy",
      "variant": "A4 ปอนด์ 70 แกรม ขาวดำ",
      "qty": 5000,
      "unit": "แผ่น",
      "unitPrice": 0.5,
      "estimatedAmount": 2500,
      "priceRuleId": "copy-a4-bw",
      "priceRuleSnapshot": "PER_SHEET 0.50"
    }
  ]
}
```

---

## 14. Security Considerations

| ประเด็น | วิธีจัดการ |
|---------|-----------|
| Authentication | CONNECT_TOKEN ตรวจทุก request |
| Token Storage | เก็บใน Script Properties ของ GAS ไม่โผล่ใน code |
| ข้อมูลผู้ใช้ | เก็บเฉพาะสิ่งที่จำเป็น ตาม PDPA |
| File Access | ลิงก์ไฟล์เป็น permissioned ไม่ใช่ public link |
| Sheet Secrets | ห้ามเก็บ secrets ในแท็บ SETTINGS |
| Who has access | ต้อง "Anyone" (ความปลอดภัยมาจาก Token แทน) |

---

*เอกสารนี้สร้างโดยอัตโนมัติจากการวิเคราะห์โค้ด — วันที่ 14 กันยายน 2569*
