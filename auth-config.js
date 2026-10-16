/**
 * Public browser configuration for SPU Printing Service authentication.
 *
 * This file must never contain CONNECT_TOKEN, AUTH_GATEWAY_TOKEN, session
 * secrets, Spreadsheet IDs, or Drive Folder IDs. Those values belong only in
 * the auth-gateway environment.
 *
 * Production recommendation: reverse-proxy /auth and /api to the gateway on
 * the same origin and leave apiBaseUrl empty. For local development, set it
 * to the gateway origin, for example "http://localhost:8080".
 */
window.SPU_AUTH_CONFIG = Object.freeze({
  // สำหรับทดสอบในเครื่อง: ใช้ hostname เดียวกับที่หน้าเว็บเปิด (127.0.0.1 หรือ localhost) พอร์ต 8080
  apiBaseUrl: (typeof window !== "undefined" && window.location && (window.location.hostname === "127.0.0.1" || window.location.hostname === "localhost"))
    ? "http://" + window.location.hostname + ":8080"
    : "http://localhost:8080"
});
