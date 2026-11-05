// Monthly Excel backup: loads the app's own index.html in jsdom, feeds it the latest state
// from Supabase, runs the app's own exportFullExcelReport() (so the file is identical to the
// manual export) and uploads the result to a Google Drive folder.
const fs = require("fs");
const path = require("path");
const { JSDOM, VirtualConsole } = require("jsdom");
const ExcelJS = require("exceljs");

const env = n => { const v = process.env[n]; if (!v) throw new Error("חסר סוד/משתנה: " + n); return v; };
const wait = ms => new Promise(r => setTimeout(r, ms));

// Israel date helpers (the workflow runs in UTC)
function ilParts(d) {
  const f = new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Jerusalem", year: "numeric", month: "2-digit", day: "2-digit" });
  return f.format(d); // YYYY-MM-DD
}

async function fetchState() {
  const url = env("SUPABASE_URL").replace(/\/$/, "");
  const key = env("SUPABASE_SERVICE_ROLE_KEY");
  let q = `${url}/rest/v1/budget_data?select=user_id,data,updated_at&order=updated_at.desc&limit=1`;
  if (process.env.BUDGET_USER_ID) q += `&user_id=eq.${process.env.BUDGET_USER_ID}`;
  // New-style keys (sb_secret_...) are not JWTs: send them only as `apikey`.
  // Legacy service_role keys (JWT, start with "eyJ") also go in Authorization.
  const headers = { apikey: key };
  if (key.startsWith("eyJ")) headers.Authorization = "Bearer " + key;
  const r = await fetch(q, { headers });
  if (!r.ok) throw new Error("Supabase " + r.status + " " + (await r.text()).slice(0, 200));
  const rows = await r.json();
  if (!rows.length) throw new Error("לא נמצאה שורת נתונים ב-budget_data");
  console.log("נטען מצב מהענן, עודכן:", rows[0].updated_at);
  return rows[0].data;
}

async function buildXlsx(state, htmlPath) {
  const html = fs.readFileSync(htmlPath, "utf8");
  const vc = new VirtualConsole();
  vc.on("jsdomError", () => {});
  let captured = null;
  const dom = new JSDOM(html, {
    runScripts: "dangerously", pretendToBeVisual: true,
    url: "https://ilanshahaf-droid.github.io/My-budget-new/", virtualConsole: vc,
    beforeParse(w) {
      w.localStorage.setItem("budgetV36", JSON.stringify(state));
      w.alert = m => { throw new Error("alert: " + m); };
      w.confirm = () => true; w.scrollTo = () => {};
      w.matchMedia = w.matchMedia || (() => ({ matches: false, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} }));
      w.HTMLElement.prototype.scrollIntoView = () => {};
      Object.defineProperty(w.HTMLElement.prototype, "innerText", { get() { return this.textContent; }, set(v) { this.textContent = v; } });
      w.URL.createObjectURL = b => { captured = b; return "blob:capture"; };
      w.URL.revokeObjectURL = () => {};
      w.HTMLAnchorElement.prototype.click = function () {};
      // never let the headless copy talk to the cloud
      w.supabase = undefined;
    }
  });
  const w = dom.window;
  // ExcelJS must live in the SAME realm as the page code (cross-realm arrays/Dates break addRow),
  // so evaluate its browser bundle inside the window instead of passing the Node module in.
  w.eval(fs.readFileSync(require.resolve("exceljs/dist/exceljs.min.js"), "utf8"));
  if (!w.ExcelJS) throw new Error("ExcelJS לא נטען בחלון");
  await wait(3500);
  let alertErr = null;
  w.alert = m => { alertErr = m; };
  await w.exportFullExcelReport();
  if (alertErr) throw new Error("הייצוא נכשל: " + alertErr);
  if (!captured) throw new Error("הייצוא לא החזיר קובץ");
  const buf = await new Promise((res, rej) => {
    const fr = new w.FileReader();
    fr.onload = () => res(Buffer.from(fr.result)); fr.onerror = () => rej(fr.error);
    fr.readAsArrayBuffer(captured);
  });
  w.close();
  return buf;
}

async function sanityCheck(buf) {
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf);
  const names = wb.worksheets.map(s => s.name);
  const tx = wb.getWorksheet("תנועות מלאות");
  console.log("גיליונות:", names.join(", "), "| תנועות:", tx ? tx.rowCount - 1 : 0);
  if (!tx || tx.rowCount < 2) throw new Error("הקובץ ריק מתנועות — לא מעלה");
}

async function driveToken() {
  const r = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: env("GOOGLE_CLIENT_ID"), client_secret: env("GOOGLE_CLIENT_SECRET"),
      refresh_token: env("GOOGLE_REFRESH_TOKEN"), grant_type: "refresh_token"
    })
  });
  const j = await r.json();
  if (!r.ok) throw new Error("Google token: " + JSON.stringify(j));
  return j.access_token;
}

async function uploadToDrive(buf, name) {
  const token = await driveToken();
  const folder = env("DRIVE_FOLDER_ID");
  const boundary = "bkp" + Date.now();
  const meta = JSON.stringify({ name, parents: [folder] });
  const mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const body = Buffer.concat([
    Buffer.from(`--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${meta}\r\n--${boundary}\r\nContent-Type: ${mime}\r\n\r\n`),
    buf, Buffer.from(`\r\n--${boundary}--`)
  ]);
  const r = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id,name,webViewLink", {
    method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "multipart/related; boundary=" + boundary }, body
  });
  const j = await r.json();
  if (!r.ok) throw new Error("Drive upload: " + JSON.stringify(j));
  console.log("הועלה לדרייב:", j.name, j.id);
}

(async () => {
  const now = new Date();
  const today = ilParts(now), tomorrow = ilParts(new Date(now.getTime() + 864e5));
  const manual = process.env.FORCE === "true" || process.env.GITHUB_EVENT_NAME === "workflow_dispatch";
  if (!manual && today.slice(0, 7) === tomorrow.slice(0, 7)) { console.log("לא סוף חודש (" + today + ") — מדלג"); return; }

  const htmlCandidates = [process.env.APP_HTML, "index.html", "pwa/index.html"].filter(Boolean);
  const htmlFound = htmlCandidates.find(f => fs.existsSync(f));
  if (!htmlFound) throw new Error("לא נמצא index.html של האפליקציה");
  const html = path.resolve(htmlFound);
  const state = process.env.LOCAL_STATE_FILE ? JSON.parse(fs.readFileSync(process.env.LOCAL_STATE_FILE, "utf8")) : await fetchState();
  const buf = await buildXlsx(state, html);
  await sanityCheck(buf);
  const name = `דוח-תקציב-מלא-${today}.xlsx`;
  if (process.env.DRY_RUN === "true") { fs.writeFileSync(process.env.DRY_OUT || "out.xlsx", buf); console.log("DRY_RUN: נשמר מקומית", name); return; }
  await uploadToDrive(buf, name);
})().catch(e => { console.error("❌", e.message); process.exit(1); });
