import { createServerFn } from "@tanstack/react-start";
import * as XLSX from "xlsx";
import { z } from "zod";
import {
  archiveBook, buildRows, bumpYear, nextSerial, parseBook, s, ARCHIVE_PREFIX, nextListedYear,
  type Book, type InRec, type Kind, type Lists, type OutRec,
} from "./msg";

const LOVABLE_API = "https://connector-gateway.lovable.dev/google_drive";
const GOOGLE_API = "https://www.googleapis.com";
const FOLDER_NAME = "nedal.MSG - مرفقات";

/** Direct Google OAuth (used outside Lovable, e.g. on Netlify): access token cached until shortly before expiry. */
let token: { value: string; exp: number } | null = null;
async function googleToken() {
  if (token && token.exp > Date.now()) return token.value;
  const res = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: process.env["GOOGLE_CLIENT_ID"] ?? "",
      client_secret: process.env["GOOGLE_CLIENT_SECRET"] ?? "",
      refresh_token: process.env["GOOGLE_REFRESH_TOKEN"] ?? "",
      grant_type: "refresh_token",
    }),
  });
  if (!res.ok) {
    console.error(`Google OAuth [${res.status}]: ${await res.text()}`);
    throw new Error(`تعذر تسجيل الدخول إلى جوجل درايف [${res.status}]`);
  }
  const j = (await res.json()) as { access_token: string; expires_in: number };
  token = { value: j.access_token, exp: Date.now() + (j.expires_in - 60) * 1000 };
  return token.value;
}

async function connection(extra?: Record<string, string>) {
  const a = process.env["LOVABLE_API_KEY"];
  const b = process.env["GOOGLE_DRIVE_API_KEY"];
  if (a && b) return { api: LOVABLE_API, headers: { Authorization: `Bearer ${a}`, "X-Connection-Api-Key": b, ...extra } };
  if (process.env["GOOGLE_CLIENT_ID"] && process.env["GOOGLE_CLIENT_SECRET"] && process.env["GOOGLE_REFRESH_TOKEN"]) {
    return { api: GOOGLE_API, headers: { Authorization: `Bearer ${await googleToken()}`, ...extra } };
  }
  throw new Error("جوجل درايف غير مربوط");
}
async function drive(path: string, init: RequestInit = {}) {
  const c = await connection(init.headers as Record<string, string>);
  const res = await fetch(`${c.api}${path}`, { ...init, headers: c.headers });
  if (!res.ok) {
    const t = await res.text();
    console.error(`Drive [${res.status}] ${path}: ${t}`);
    throw new Error(`تعذر الاتصال بجوجل درايف [${res.status}]`);
  }
  return res;
}

async function findFileId() {
  const q = encodeURIComponent("(name contains 'msg') and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false");
  const j = (await (await drive(`/drive/v3/files?q=${q}&orderBy=modifiedTime desc&fields=files(id,name)`)).json()) as { files: { id: string }[] };
  const f = j.files[0];
  if (!f) throw new Error("لم يتم العثور على ملف nedal.msg في الدرايف");
  return f.id;
}

type Loaded = { id: string; wb: XLSX.WorkBook };
async function load(): Promise<Loaded> {
  const id = await findFileId();
  const buf = await (await drive(`/drive/v3/files/${id}?alt=media`)).arrayBuffer();
  return { id, wb: XLSX.read(buf, { type: "array", cellDates: true }) };
}
async function save({ id, wb }: Loaded) {
  const out = XLSX.write(wb, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
  await drive(`/upload/drive/v3/files/${id}?uploadType=media`, {
    method: "PATCH",
    headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
    body: out,
  });
}

function sheetName(wb: XLSX.WorkBook, name: string) {
  const n = wb.SheetNames.find((x) => x.trim().toLowerCase() === name);
  if (!n) throw new Error(`الورقة ${name} غير موجودة في الملف`);
  return n;
}
/** Rows padded so index 0 is column A / row 1. Dates normalised to yyyy-mm-dd. */
function rows(wb: XLSX.WorkBook, name: string): unknown[][] {
  const ws = wb.Sheets[sheetName(wb, name)];
  if (!ws) throw new Error("الورقة غير موجودة");
  const st = XLSX.utils.decode_range(ws["!ref"] ?? "A1").s;
  const data = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null, blankrows: true, raw: true });
  const fix = (v: unknown) => (v instanceof Date ? v.toISOString().slice(0, 10) : v);
  return [...Array(st.r).fill([]), ...data.map((r) => [...Array(st.c).fill(null), ...r.map(fix)])];
}

function readData(wb: XLSX.WorkBook) {
  const r = rows(wb, "data").slice(1);
  const col = (i: number) => [...new Set(r.map((x) => s(x?.[i])).filter(Boolean))];
  const lists: Lists = { year: col(1)[0] ?? "", units: col(2), authorities: col(3), enclosures: col(4), actions: col(5), institution: col(0)[0] ?? "" };
  const metaSheet = wb.Sheets["_nedal_meta"];
  const activeYear = metaSheet ? s(metaSheet["A1"]?.v) : "";
  if (activeYear && col(1).includes(activeYear)) lists.year = activeYear;
  return { lists, years: col(1), archiveCodes: col(8), ids: col(6), admin: col(7) };
}

function roleOf(wb: XLSX.WorkBook, code: string): "admin" | "user" | null {
  const { ids, admin } = readData(wb);
  const c = code.trim();
  if (!c) return null;
  if (admin.includes(c)) return "admin";
  if (ids.includes(c)) return "user";
  return null;
}
function require(wb: XLSX.WorkBook, code: string, need: "user" | "admin") {
  const r = roleOf(wb, code);
  if (!r || (need === "admin" && r !== "admin")) throw new Error("رمز الدخول غير صحيح");
  return r;
}

function readBook<K extends Kind>(wb: XLSX.WorkBook, kind: K) {
  return parseBook(kind, rows(wb, kind));
}
function writeBook<K extends Kind>(wb: XLSX.WorkBook, kind: K, book: Book<K extends "in" ? InRec : OutRec>) {
  wb.Sheets[sheetName(wb, kind)] = XLSX.utils.aoa_to_sheet(buildRows(kind, book));
}

/** Appends a typed-in value to the matching data column if it is new ("قيمة للاضافة"). */
function addToList(wb: XLSX.WorkBook, colIdx: number, value: string) {
  const v = value.trim();
  if (!v) return;
  const r = rows(wb, "data");
  if (r.slice(1).some((x) => s(x?.[colIdx]) === v)) return;
  let rowIdx = 1;
  while (rowIdx < r.length && s(r[rowIdx]?.[colIdx])) rowIdx++;
  XLSX.utils.sheet_add_aoa(wb.Sheets[sheetName(wb, "data")]!, [[v]], { origin: { r: rowIdx, c: colIdx } });
}

const auth = z.object({ code: z.string().min(1).max(40) });

export const login = createServerFn({ method: "POST" })
  .inputValidator((d) => auth.parse(d))
  .handler(async ({ data }) => {
    const { wb } = await load();
    const role = roleOf(wb, data.code);
    if (!role) throw new Error("رمز الدخول غير صحيح");
    return { role };
  });

export const getMsgData = createServerFn({ method: "POST" })
  .inputValidator((d) => auth.parse(d))
  .handler(async ({ data }) => {
    const { wb } = await load();
    const role = require(wb, data.code, "user");
    const inB = readBook(wb, "in");
    const outB = readBook(wb, "out");
    const { lists } = readData(wb);
    return {
      role,
      lists,
      nextIn: nextSerial(inB.active),
      nextOut: nextSerial(outB.active),
      in: role === "admin" ? inB : { active: inB.active, archives: [] },
      out: role === "admin" ? outB : { active: outB.active, archives: [] },
    };
  });

const inFields = z.object({
  fileNo: z.string().max(100), date: z.string().max(20), subject: z.string().min(1).max(500), status: z.string().max(100).optional(), reminderDate: z.string().max(20).optional(),
  sender: z.string().max(200), unit: z.string().max(200), action: z.string().max(1000), notes: z.string().max(2000), attachment: z.string().max(500),
});
const outFields = z.object({
  fileNo: z.string().max(100), date: z.string().max(20), subject: z.string().min(1).max(500), status: z.string().max(100).optional(), reminderDate: z.string().max(20).optional(),
  authority: z.string().max(200), enclosures: z.string().max(200), attachment: z.string().max(500),
});

export const addIn = createServerFn({ method: "POST" })
  .inputValidator((d) => auth.extend({ rec: inFields }).parse(d))
  .handler(async ({ data }) => {
    const L = await load();
    require(L.wb, data.code, "user");
    const book = readBook(L.wb, "in");
    const serial = nextSerial(book.active);
    book.active.push({ serial, ...data.rec, status: data.rec.status ?? "", reminderDate: data.rec.reminderDate ?? "", year: readData(L.wb).lists.year, institution: readData(L.wb).lists.institution });
    writeBook(L.wb, "in", book);
    addToList(L.wb, 3, data.rec.sender);
    addToList(L.wb, 2, data.rec.unit);
    addToList(L.wb, 5, data.rec.action);
    await save(L);
    return { serial };
  });

export const addOut = createServerFn({ method: "POST" })
  .inputValidator((d) => auth.extend({ rec: outFields }).parse(d))
  .handler(async ({ data }) => {
    const L = await load();
    require(L.wb, data.code, "user");
    const book = readBook(L.wb, "out");
    const serial = nextSerial(book.active);
    book.active.push({ serial, ...data.rec, status: data.rec.status ?? "", reminderDate: data.rec.reminderDate ?? "", year: readData(L.wb).lists.year, institution: readData(L.wb).lists.institution });
    writeBook(L.wb, "out", book);
    addToList(L.wb, 3, data.rec.authority);
    addToList(L.wb, 4, data.rec.enclosures);
    await save(L);
    return { serial };
  });

/** Admin edit of any record (active or archived), identified by archive label + serial. */
export const updateRecord = createServerFn({ method: "POST" })
  .inputValidator((d) =>
    auth.extend({
      kind: z.enum(["in", "out"]),
      archive: z.string().nullable(),
      serial: z.number().int(),
      patch: z.record(z.string(), z.string().max(2000)),
    }).parse(d),
  )
  .handler(async ({ data }) => {
    const L = await load();
    require(L.wb, data.code, data.archive === null ? "user" : "admin");
    const book = readBook(L.wb, data.kind) as Book<InRec | OutRec>;
    const list = data.archive === null ? book.active : book.archives.find((a) => a.label === data.archive)?.rows;
    const rec = list?.find((r) => r.serial === data.serial);
    if (!rec) throw new Error("لم يتم العثور على المراسلة");
    for (const [k, v] of Object.entries(data.patch)) if (!["serial", "year", "institution"].includes(k) && (k in rec || ["status", "reminderDate"].includes(k))) (rec as unknown as Record<string, unknown>)[k] = v;
    writeBook(L.wb, data.kind, book as never);
    await save(L);
    return { ok: true };
  });

export const closeYear = createServerFn({ method: "POST" })
  .inputValidator((d) => auth.extend({ archiveCode: z.string().min(1) }).parse(d))
  .handler(async ({ data }) => {
    const L = await load();
    require(L.wb, data.code, "admin");
    const { lists, years, archiveCodes } = readData(L.wb);
    if (!archiveCodes.includes(data.archiveCode.trim())) throw new Error("رقم الأرشفة غير صحيح");
    const next = nextListedYear(years, lists.year);
    if (!next) throw new Error("أضف العام الدراسي التالي في ورقة data قبل الأرشفة");
    writeBook(L.wb, "in", archiveBook(readBook(L.wb, "in"), lists.year));
    writeBook(L.wb, "out", archiveBook(readBook(L.wb, "out"), lists.year));
    const meta = XLSX.utils.aoa_to_sheet([[next]]);
    if (!L.wb.SheetNames.includes("_nedal_meta")) XLSX.utils.book_append_sheet(L.wb, meta, "_nedal_meta");
    else L.wb.Sheets["_nedal_meta"] = meta;
    await save(L);
    return { archived: ARCHIVE_PREFIX + lists.year, year: next };
  });

async function folderId(kind: string) {
  const FOLDER_NAME = kind === "out" ? "الصادر" : "الوارد";
  const q = encodeURIComponent(`name='${FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false`);
  const j = (await (await drive(`/drive/v3/files?q=${q}&fields=files(id)`)).json()) as { files: { id: string }[] };
  if (j.files[0]) return j.files[0].id;
  const c = (await (await drive(`/drive/v3/files?fields=id`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: FOLDER_NAME, mimeType: "application/vnd.google-apps.folder" }),
  })).json()) as { id: string };
  return c.id;
}

/** Uploads an attachment to Drive named after the subject; returns a shareable view link. */
export const uploadAttachment = createServerFn({ method: "POST" })
  .inputValidator((d) => {
    if (!(d instanceof FormData)) throw new Error("بيانات غير صالحة");
    const file = d.get("file");
    if (!(file instanceof File)) throw new Error("لم يتم اختيار ملف");
    if (file.size > 25 * 1024 * 1024) throw new Error("حجم الملف أكبر من 25 ميغابايت");
    return { code: s(d.get("code")), subject: s(d.get("subject")) || file.name, kind: s(d.get("kind")), file };
  })
  .handler(async ({ data }) => {
    const { wb } = await load();
    require(wb, data.code, "user");
    const parent = await folderId(data.kind);
    const ext = data.file.name.includes(".") ? data.file.name.slice(data.file.name.lastIndexOf(".")) : "";
    const meta = { name: `وارد . ${data.subject} . ${readData(wb).lists.year}${ext}`, parents: [parent] };
    const boundary = "nedalmsg" + Date.now();
    const enc = new TextEncoder();
    const body = new Blob([
      enc.encode(`--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n--${boundary}\r\nContent-Type: ${data.file.type || "application/octet-stream"}\r\n\r\n`),
      await data.file.arrayBuffer(),
      enc.encode(`\r\n--${boundary}--`),
    ]);
    const f = (await (await drive(`/upload/drive/v3/files?uploadType=multipart&fields=id`, {
      method: "POST",
      headers: { "Content-Type": `multipart/related; boundary=${boundary}` },
      body,
    })).json()) as { id: string };
    return { link: `https://drive.google.com/file/d/${f.id}/view` };
  });

export const getPublicInfo = createServerFn({ method: "GET" }).handler(async () => {
  const { wb } = await load();
  const { lists } = readData(wb);
  return { year: lists.year, institution: lists.institution };
});

export const readAttachment = createServerFn({ method: "POST" })
  .inputValidator((d) => auth.extend({ id: z.string().regex(/^[\w-]+$/) }).parse(d))
  .handler(async ({ data }) => {
    const { wb } = await load();
    const role = require(wb, data.code, "user");
    const allowed = [readBook(wb, "in"), readBook(wb, "out")].some((book) =>
      book.active.some((r) => r.attachment.includes(data.id)) || (role === "admin" && book.archives.some((a) => a.rows.some((r) => r.attachment.includes(data.id)))));
    if (!allowed) throw new Error("المرفق غير مرتبط بمراسلة متاحة");
    const res = await drive(`/drive/v3/files/${data.id}?alt=media`);
    const bytes = new Uint8Array(await res.arrayBuffer());
    let binary = "";
    for (let i = 0; i < bytes.length; i += 8192) binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
    return { base64: btoa(binary), type: res.headers.get("content-type") ?? "application/octet-stream" };
  });
