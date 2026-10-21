# Printing Service Requests 🖨️

ระบบขอใช้บริการงานพิมพ์ของ SPU สำหรับผู้รับบริการ เจ้าหน้าที่ และผู้บริหาร มีหน้าเว็บภาษาไทย/อังกฤษ เชื่อมต่อ Google Apps Script ผ่าน Node.js Authentication Gateway เพื่อจัดเก็บข้อมูลงานใน Google Sheets และไฟล์ใน Google Drive

> **สถานะของ repository:** ยังติดตั้งระบบครบชุดจากการ clone เพียงอย่างเดียวไม่ได้ เพราะไม่มี `google-apps-script/Code.gs` ซึ่งถูกระบุใน `.gitignore` ต้องขอซอร์ส Apps Script รุ่นที่เข้ากันและวิธีตั้งค่าชีตจากผู้ดูแลก่อนใช้งานจริง การเปิดหน้าเว็บหรือการผ่าน Gateway tests ไม่ได้ยืนยันว่าเชื่อม Sheets/Drive ได้แล้ว

## ความสามารถ

- ผู้รับบริการเข้าสู่ระบบด้วย Google Workspace ตามโดเมนที่กำหนด (ค่าเริ่มต้น `spu.ac.th`)
- เจ้าหน้าที่ `ADMIN` และผู้บริหาร `EXECUTIVE` เข้าสู่ระบบด้วยอีเมล/รหัสผ่านที่ตรวจผ่าน Apps Script
- แบบฟอร์มขอบริการหลายรายการ พร้อมราคาประมาณการ ไฟล์งาน และสลิปโอนเงิน
- ติดตามสถานะงาน จัดการคิว และยืนยันรับงานตามสิทธิ์
- แชทรายงาน พร้อมไฟล์แนบและจำนวนข้อความที่ยังไม่อ่าน
- หน้ารายงาน การส่งรายงานทางอีเมล และการจัดการกฎราคาสำหรับบทบาทที่ได้รับอนุญาต

รายการนี้อธิบายส่วนติดต่อและ action ที่มีในซอร์ส ฝั่ง Apps Script ต้องรองรับ action เหล่านี้จึงจะทำงานได้ครบ

## สถาปัตยกรรมและเทคโนโลยี

```text
Browser ── Google Identity Services (ผู้รับบริการ)
   │
   │ /auth/* และ /api
   ▼
Node.js / Express Gateway
   │ POST + server-side tokens + authenticated actor
   ▼
Google Apps Script (ต้องจัดหาเพิ่มเติม)
   ├── Google Sheets
   └── Google Drive
```

- **Frontend:** HTML, CSS, JavaScript และ Design Component runtime ใน `support.js` ซึ่งโหลด React/ReactDOM 18.3.1 จาก CDN ไม่ใช่ Vanilla JS ล้วน
- **Gateway:** Node.js, Express 5 และ `google-auth-library`
- **ข้อมูลและไฟล์:** Google Sheets/Drive ผ่าน Apps Script ไม่ใช่การเรียก Sheets API/Drive API จากเบราว์เซอร์โดยตรง
- **ทดสอบ:** Node.js built-in test runner (`node --test`)
- **Deployment:** static hosting สำหรับหน้าเว็บ และ Node.js หรือ Docker สำหรับ Gateway

## โครงสร้างไฟล์

```text
index.html             หน้าเว็บ เทมเพลต และตรรกะการทำงานหลัก
auth-config.js         URL ของ Gateway ที่เปิดเผยต่อเบราว์เซอร์ได้
spu-data.js             ข้อความ ข้อมูลประกอบ และคำอธิบายโครงสร้างชีต
support.js             Design Component runtime
assets/                รูปภาพของหน้าเว็บ
_ds/                   ไฟล์ design system
auth-gateway/
  server.js            HTTP endpoints, Google login และ proxy ไป Apps Script
  lib/                 session และ role/action policy
  test/                Gateway unit/integration tests (mock บริการภายนอก)
  .env.example         ตัวอย่าง environment variables
  Dockerfile           image สำหรับ Gateway เท่านั้น
tests/                 contract tests ที่ต้องใช้ Code.gs ด้วย
```

## เตรียมระบบ

1. ติดตั้ง Git, Node.js และ npm โดย `auth-gateway/package.json` กำหนด Node.js `>=20` และ `start:local` ต้องใช้รุ่นที่รองรับ `--env-file` (Node.js 20.6 ขึ้นไป) ควรเลือกรุ่นที่ยังได้รับการสนับสนุนจาก Node.js
2. เตรียม Google Cloud OAuth Client ชนิด **Web application** และบัญชี Google Workspace ในโดเมนที่อนุญาต
3. ขอ Apps Script source/deployment ที่เข้ากัน พร้อมโครงสร้างชีตและโฟลเดอร์ Drive จากผู้ดูแล อย่าสร้างไฟล์ `Code.gs` ว่างเพื่อให้ข้ามข้อผิดพลาด
4. ตั้ง Script Properties ของ Apps Script ได้แก่ `CONNECT_TOKEN`, `AUTH_GATEWAY_TOKEN`, `SPREADSHEET_ID`, `FOLDER_ID` ตาม backend รุ่นที่ใช้
5. ตั้ง `Authorized JavaScript origins` ของ OAuth Client ให้ตรงกับ origin ของหน้าเว็บ

รายละเอียด Gateway อยู่ใน [auth-gateway/README.md](auth-gateway/README.md)

## เริ่มต้นในเครื่อง (PowerShell)

```powershell
git clone https://github.com/Ai-TeamSPU/Printing-Service-Requests.git
cd Printing-Service-Requests
cd auth-gateway
npm ci
Copy-Item .env.example .env
```

แก้ `.env` ก่อนรัน โดยค่าที่ต้องกรอกจริงคือ:

| ตัวแปร | ความหมาย |
|---|---|
| `GOOGLE_CLIENT_ID` | Web OAuth Client ID ของระบบ |
| `ALLOWED_GOOGLE_DOMAIN` | โดเมน Workspace ที่อนุญาต เช่น `spu.ac.th` |
| `ALLOWED_ORIGINS` | origin หน้าเว็บแบบตรงตัว หลายค่าแยกด้วย comma |
| `SESSION_SECRET` | secret สำหรับเซ็น session |
| `APPS_SCRIPT_URL` | HTTPS Web App URL บน `script.google.com` ที่ลงท้าย `/exec` |
| `APPS_SCRIPT_CONNECT_TOKEN` | ตรงกับ `CONNECT_TOKEN` ใน Script Properties |
| `APPS_SCRIPT_GATEWAY_TOKEN` | ตรงกับ `AUTH_GATEWAY_TOKEN` ใน Script Properties |
| `REQUIRE_PERSONNEL` | `true` บังคับพบผู้รับบริการในทะเบียนเมื่อ login/ตรวจ session; ค่าเริ่มต้น `false` |
| `SESSION_TTL_SECONDS` | อายุ token ที่ออกแต่ละครั้ง ค่าเริ่มต้น 28,800 วินาที; อนุญาต 300–86,400 |
| `COOKIE_SECURE`, `COOKIE_SAME_SITE` | นโยบาย cookie ตาม HTTP/HTTPS และการวาง frontend/Gateway |
| `PORT`, `NODE_ENV` | พอร์ต Gateway และโหมดทำงาน |

`SESSION_SECRET` และ token ทั้งสองต้องแตกต่างกัน แต่ละค่ายาวอย่างน้อย 32 bytes และไม่ใช่ placeholder สร้างแต่ละค่าแยกกันได้ด้วย:

```powershell
node -e "console.log(require('node:crypto').randomBytes(32).toString('hex'))"
```

สำหรับ local HTTP ให้เปลี่ยนค่าต่อไปนี้ใน `.env` เพราะตัวอย่างเดิมตั้งเป็น production:

```dotenv
NODE_ENV=development
PORT=8080
ALLOWED_ORIGINS=http://localhost:8000
COOKIE_SECURE=false
COOKIE_SAME_SITE=lax
```

เพิ่ม `http://localhost:8000` ใน OAuth Authorized JavaScript origins แล้วรัน Gateway:

```powershell
npm run start:local
```

เปิดอีก terminal ที่ root ของ repository แล้วเสิร์ฟหน้าเว็บผ่าน static server เช่น หากติดตั้ง Python ไว้:

```powershell
python -m http.server 8000 --bind 127.0.0.1
```

เปิด `http://localhost:8000` ผ่านเบราว์เซอร์ ไม่เปิด `index.html` แบบ `file://` ใช้ hostname เดียวกันตลอด หากเลือก `127.0.0.1` ต้องตั้ง origin ใน Gateway และ OAuth ให้ตรงกันด้วย

`auth-config.js` ปัจจุบันเลือก Gateway พอร์ต `8080` อัตโนมัติเมื่อหน้าเว็บอยู่บน `localhost` หรือ `127.0.0.1` สำหรับ hostname อื่นจะชี้ `https://spu-printing-gateway.onrender.com` จึงต้องเปลี่ยน URL นี้เป็น Gateway ของตนเองก่อนเผยแพร่สำเนาระบบ

ตรวจว่า process ทำงานได้ด้วย `http://localhost:8080/health` แต่ endpoint นี้ไม่ได้ทดสอบการเชื่อมต่อ Apps Script, Sheets หรือ Drive

## ทดสอบ

จาก root ของ repository:

```powershell
npm --prefix auth-gateway ci
npm --prefix auth-gateway test
npm --prefix auth-gateway run check
node --test tests/*.test.cjs
```

- Gateway tests ใช้ mock Google/Apps Script จึงไม่ต้องมี credentials จริง
- Contract tests ใน `tests/` อ่าน `google-apps-script/Code.gs` โดยตรง และจะจบด้วย `ENOENT` ในสำเนาปัจจุบันจนกว่าจะจัดหาไฟล์ดังกล่าว
- หลังเตรียม backend แล้ว ต้องทดสอบ login, สิทธิ์ข้ามบัญชี, สร้างงาน, ไฟล์แนบ, สลิป, แชท และรายงานกับระบบทดสอบจริงเพิ่มเติม

## Production

- แยกการ deploy static frontend ออกจาก Gateway; Dockerfile ไม่ได้เสิร์ฟ `index.html`
- ทางเลือกที่ลดปัญหา cross-site cookies คือ reverse proxy `/auth/*`, `/api`, `/health` ไป Gateway ภายใต้ origin เดียวกับหน้าเว็บ แล้วตั้ง `apiBaseUrl: ""` ใน `auth-config.js` เอง (ไม่ใช่ค่า default ปัจจุบัน)
- หากใช้คนละ origin ให้ตั้ง URL Gateway และ `ALLOWED_ORIGINS` ให้ตรง หากคนละ site และต้องใช้ cookie ให้ตั้ง `COOKIE_SAME_SITE=none` และ `COOKIE_SECURE=true`; browser อาจยังบล็อก third-party cookies
- ตั้ง `NODE_ENV=production`, ใช้ HTTPS และ inject environment variables ผ่านระบบ deployment/secret store ก่อนเรียก `npm start`
- เก็บ `.env` และ shared secrets ไว้ฝั่งเซิร์ฟเวอร์ ห้ามใส่ใน `index.html` หรือ `auth-config.js`

สร้าง Gateway image จาก root ได้ด้วย:

```powershell
docker build -t spu-printing-auth-gateway ./auth-gateway
```

### พฤติกรรม session ปัจจุบัน

Gateway ตั้ง HttpOnly cookie **และส่ง application session token กลับใน JSON** หน้าเว็บเก็บ token ใน `sessionStorage` และแนบ `Authorization: Bearer` จึงไม่ใช่ระบบ cookie-only และ JavaScript บนหน้าเว็บอ่าน application token ได้ ส่วน Google ID token ใช้ส่งให้ Gateway ตรวจสอบตอน login

`/auth/session` ออก application token ใหม่เมื่อ session เดิมยังใช้ได้ ส่วน logout ล้าง cookie และหน้าเว็บล้าง token ในเครื่อง แต่ไม่ได้เพิกถอนสำเนา bearer token ที่ออกไปแล้ว การเปลี่ยน `SESSION_SECRET` ทำให้ session เดิมทั้งหมดใช้ไม่ได้

## ข้อจำกัดที่พบจากการตรวจโค้ด

- ไม่มี Apps Script source ใน repository จึงยังยืนยันการตรวจเจ้าของงาน สิทธิ์เข้าถึงไฟล์ หรือการเขียนชีตฝั่ง backend ไม่ได้
- เมื่อสร้างงานสำเร็จแต่ upload ไฟล์งาน/สลิปล้มเหลว หน้าเว็บยังแสดงส่งสำเร็จและล้างรายการไฟล์ ควรแก้ให้รายงานไฟล์ที่ล้มเหลวและ retry กับเลขงานเดิมก่อนใช้ flow นี้กับงานสำคัญ
- หน้าเว็บ timeout ที่ 15 วินาที แต่ Gateway รอ Apps Script ได้ 30 วินาที งานที่บันทึกสำเร็จหลัง browser ยกเลิกจึงอาจถูกแสดงว่าล้มเหลว ควรจัดเวลา timeout ให้สอดคล้องและยืนยันกลไกป้องกันคำขอซ้ำใน backend
- Rate limit เก็บใน memory ต่อ Gateway instance ไม่ใช่ quota ส่วนกลาง

การปรับ README ไม่ได้แก้พฤติกรรม runtime ข้างต้น

## แก้ปัญหาเบื้องต้น

| อาการ | จุดตรวจสอบ |
|---|---|
| Gateway ไม่เริ่มทำงาน | `.env` มี placeholder, token สั้น/ซ้ำ หรือ production ใช้ HTTP origin |
| `origin_forbidden` | origin ใน address bar ต้องตรงกับ `ALLOWED_ORIGINS` รวม scheme/hostname/port แต่ไม่มี path |
| `workspace_account_required` | ใช้บัญชี Workspace ที่ verified และ hosted domain ตรงกับค่าที่อนุญาต |
| `personnel_not_authorized` | เปิด `REQUIRE_PERSONNEL=true` แต่ไม่พบอีเมลในทะเบียน |
| `gateway_unreachable` / `gateway_timeout` | URL ใน `auth-config.js`, process Gateway และเวลาตอบกลับ upstream |
| `backend_auth_misconfigured` | Apps Script deployment และ token ฝั่ง Gateway/Script Properties |
| `ENOENT` ขณะรัน contract tests | ต้องจัดหา `google-apps-script/Code.gs` รุ่นที่ตรงกัน |
