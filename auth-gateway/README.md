# SPU Printing Service — Authentication Gateway

Gateway นี้เป็นด่านกลางระหว่างหน้าเว็บกับ Google Apps Script เพื่อให้ระบบเชื่อได้ว่าอีเมลผู้รับบริการเป็นบัญชี Google Workspace จริง ไม่ใช่อีเมลที่พิมพ์เอง

สิ่งที่ Gateway ทำ:

- ตรวจ Google ID token ด้วย `google-auth-library` และ OAuth Client ID ที่กำหนด
- บังคับ `email_verified`, `hd=spu.ac.th`, audience, issuer, signature และวันหมดอายุผ่าน Google library
- ออก application session อายุสั้นใน cookie แบบ `HttpOnly` (ไม่เก็บ ID token ใน `localStorage`)
- กำหนดสิทธิ์ต่อ action สำหรับ `REQUESTER`, `ADMIN`, `EXECUTIVE`
- ลบ `email`, `role`, `by`, token และ resource ID ที่ browser พยายามส่งทับ
- เก็บ Apps Script tokens ไว้ใน environment เท่านั้น

## 1. สิ่งที่ต้องสร้างใน Google Cloud

1. ใช้ Google Cloud project ของมหาวิทยาลัย
2. ตั้ง OAuth consent screen เป็น **Internal** ถ้า policy ของ Google Workspace อนุญาต
3. สร้าง OAuth Client ID ชนิด **Web application**
4. ใส่ origin ของหน้าเว็บจริงใน **Authorized JavaScript origins** เช่น `https://printing.spu.ac.th`
5. นำ Client ID ไปตั้งเป็น `GOOGLE_CLIENT_ID` ใน Gateway

Client ID เป็นค่าฝั่ง public ได้ แต่ client secret ไม่ได้ใช้ใน Google Identity Services flow นี้และห้ามนำมาใส่ในหน้าเว็บ

## 2. ตั้ง Script Properties ใน Apps Script

ไปที่ **Project Settings > Script properties** แล้วตั้งค่า:

| Property | ความหมาย |
|---|---|
| `CONNECT_TOKEN` | token ใหม่สำหรับเชื่อม Gateway กับ Apps Script |
| `AUTH_GATEWAY_TOKEN` | token แยกอีกตัวสำหรับพิสูจน์ว่าคำขอมาจาก Gateway |
| `SPREADSHEET_ID` | ID ของฐานข้อมูล Google Sheet |
| `FOLDER_ID` | ID โฟลเดอร์แม่บน Drive |

`CONNECT_TOKEN` และ `AUTH_GATEWAY_TOKEN` ต้องเป็นคนละค่า สุ่มยาว และห้ามใส่ใน `index.html`, `auth-config.js` หรือ commit ลง Git

> ค่า `CONNECT_TOKEN` รุ่นเดิมเคยอยู่ในหน้าเว็บ จึงต้อง rotate เป็นค่าใหม่ก่อน cutover แม้จะเพิ่ม `AUTH_GATEWAY_TOKEN` แล้วก็ตาม

วาง `google-apps-script/Code.gs` รุ่นใหม่ จากนั้นรัน `setupSheets()` หนึ่งครั้งเพื่อเพิ่มคอลัมน์ `JOBS.requester_google_sub` โดยไม่ลบข้อมูลเดิม แล้ว deploy เป็น Web App เวอร์ชันใหม่

Apps Script ยัง deploy แบบ **Execute as: Me / Who has access: Anyone** ได้ เพราะ browser ไม่ควรเรียก URL นี้โดยตรงอีกต่อไป และทุก user-facing action ต้องผ่านทั้ง `CONNECT_TOKEN` กับ `AUTH_GATEWAY_TOKEN`

## 3. ตั้งค่า Gateway

คัดลอก `.env.example` เป็น `.env` เฉพาะในเครื่อง/ระบบ deploy แล้วกรอกค่าให้ครบ ห้าม commit `.env`

ค่าหลัก:

- `GOOGLE_CLIENT_ID`: Web OAuth Client ID
- `ALLOWED_GOOGLE_DOMAIN=spu.ac.th`
- `ALLOWED_ORIGINS`: origin หน้าเว็บแบบ exact match คั่นหลายค่าด้วย comma
- `SESSION_SECRET`: ค่าสุ่มอย่างน้อย 32 bytes
- `APPS_SCRIPT_URL`: URL deployment ที่ลงท้าย `/exec`
- `APPS_SCRIPT_CONNECT_TOKEN`: ตรงกับ `CONNECT_TOKEN`
- `APPS_SCRIPT_GATEWAY_TOKEN`: ตรงกับ `AUTH_GATEWAY_TOKEN`
- `REQUIRE_PERSONNEL=false`: คงพฤติกรรมเดิม คือบัญชี Workspace ในโดเมนเข้าได้ทั้งหมด
- `REQUIRE_PERSONNEL=true`: อนุญาตเฉพาะอีเมลที่พบใน `Personnel_Information`

Gateway จะไม่ยอมเริ่มทำงานถ้ายังใช้ placeholder จาก `.env.example`, secret/token สั้นกว่า 32 bytes, นำ secret เดียวกันมาใช้ซ้ำ, ใช้ HTTP origin ใน production หรือปิด Secure cookie ใน production

เมื่อ `REQUIRE_PERSONNEL=false` การอ่าน `Personnel_Information` ใช้เพียงเติมชื่อ/หน่วยงาน หากชีตยังไม่พร้อม Google Login จะยังสำเร็จ แต่การเรียกฟังก์ชันงานอื่นยังต้องเชื่อม Apps Script ได้ตามปกติ

สำหรับ production ให้เก็บ `SESSION_SECRET`, token ทั้งสอง และ Apps Script URL ใน Secret Manager หรือ secret store ของแพลตฟอร์ม

## 4. รันและทดสอบในเครื่อง

```powershell
cd auth-gateway
npm ci
npm test
npm run start:local
```

`start:local` ใช้ความสามารถ `--env-file` ของ Node 20 เพื่ออ่าน `.env` ส่วน `npm start` สำหรับ production ที่แพลตฟอร์ม inject environment variables/Secret Manager ให้ process อยู่แล้ว

ถ้า Gateway รันที่ `http://localhost:8080` และหน้าเว็บรันคนละ origin ให้แก้ `auth-config.js` ชั่วคราวเป็น:

```js
window.SPU_AUTH_CONFIG = Object.freeze({ apiBaseUrl: "http://localhost:8080" });
```

และเพิ่ม origin ของหน้าเว็บ local ใน `ALLOWED_ORIGINS` กับ Authorized JavaScript origins ของ OAuth client สำหรับ client ที่ใช้ทดสอบ

## 5. การวาง production

แนวทางแนะนำคือให้หน้าเว็บและ Gateway อยู่ภายใต้ site เดียวกัน แล้ว reverse proxy `/auth/*`, `/api`, `/health` ไป Gateway จากนั้นปล่อย `apiBaseUrl` เป็นค่าว่าง ซึ่งลดปัญหา third-party cookies

หากต้องใช้คนละ origin:

- ตั้ง `auth-config.js` เป็น HTTPS URL ของ Gateway
- ใส่ origin หน้าเว็บใน `ALLOWED_ORIGINS`
- ต้องส่ง cookie ด้วย `credentials: include` (หน้าเว็บทำไว้แล้ว)
- ถ้าอยู่คนละ site จริง ๆ ให้ใช้ `COOKIE_SAME_SITE=none` และ `COOKIE_SECURE=true`; browser บางนโยบายอาจยังบล็อก third-party cookie จึงแนะนำ same-site

Container image สร้างได้จาก Dockerfile ในโฟลเดอร์นี้ และใช้กับ Cloud Run ได้ พอร์ตมาตรฐานคือ `8080` ตัว Gateway endpoint ต้องเปิดให้ browser ติดต่อได้ แต่ข้อมูลจะถูกป้องกันด้วย Google login, signed session, Origin allowlist และ role policy

## 6. ลำดับ cutover ที่ปลอดภัย

1. สำรอง Apps Script deployment เดิมและ Sheet
2. ตั้ง Script Properties ใหม่ทั้งหมดและ rotate token เดิม
3. วาง `Code.gs`, รัน `setupSheets()`, deploy เวอร์ชันใหม่
4. deploy Gateway พร้อม secrets
5. ตั้ง OAuth authorized origin และ `ALLOWED_ORIGINS`
6. ตั้ง `auth-config.js` ให้ชี้ Gateway (หรือใช้ reverse proxy same-origin)
7. ทดสอบด้วยบัญชี SPU อย่างน้อย 2 บัญชีว่าไม่สามารถเห็นงานของกันและกัน
8. ทดสอบ ADMIN และ EXECUTIVE ว่า mutation/read ถูกจำกัดตามบทบาท
9. ตรวจ DevTools ว่า browser ไม่มี Apps Script URL, CONNECT_TOKEN หรือ session token ที่อ่านได้ด้วย JavaScript

## 7. ข้อจำกัดที่ตั้งใจไว้

- การ Login ผู้รับบริการใช้ Google จริง แต่ ADMIN/EXECUTIVE ยังใช้บัญชี/รหัสผ่านในแท็บ `USERS` ผ่าน Gateway ในระยะนี้
- รหัสผ่าน ADMIN/EXECUTIVE รุ่นเดิมยังรองรับ salted SHA-256 และ plaintext fallback เพื่อไม่ให้บัญชีเดิมถูกล็อก ควรย้ายทุกบัญชีและปิด fallback ในงาน hardening ระยะถัดไป; ห้ามถือว่าส่วนนี้แข็งแรงเท่า Google Login
- session เป็นแบบ stateless; เมื่อ `REQUIRE_PERSONNEL=true` การนำ requester ออกจากทะเบียนมีผลเมื่อหน้าเว็บตรวจ session ครั้งถัดไป/รีโหลด ส่วนการบังคับออกทันทีทุกอุปกรณ์ทำได้ด้วยการหมุน `SESSION_SECRET`
- staff ถูกตรวจสถานะ active/role ซ้ำใน Apps Script ทุก action จึงถูกระงับได้ทันที
- in-memory rate limit ช่วยลด brute force ต่อ instance แต่ production ที่มีหลาย instance ควรเสริม Cloud Armor หรือ distributed rate limiting
- `/api` มี rate limit ทั่วไป และจำกัด `submitJob`, `uploadFile`, `sendReportEmail` เข้มขึ้นต่อ instance; quota ส่วนกลางยังควรป้องกันเพิ่มที่ load balancer/Cloud Armor
- Apps Script ปิดการส่ง action ผ่าน `GET` เพื่อไม่ให้ token หลุดไปยัง URL, browser history หรือ proxy log; Gateway ใช้ `POST` เท่านั้น
- งานเก่าที่คอลัมน์ `requester_google_sub` ยังว่างจะเทียบเจ้าของด้วยอีเมลเพื่อให้ดูประวัติเดิมได้ หากองค์กรมีการนำอีเมลเก่ากลับมาใช้กับคนใหม่ ควร backfill Google subject หรือกำหนดนโยบาย archive ก่อน cutover
