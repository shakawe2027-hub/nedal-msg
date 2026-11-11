import { createServerFn } from "@tanstack/react-start";
import * as XLSX from "xlsx";
import { z } from "zod";
import {
  archiveBook, buildRows, bumpYear, nextSerial, parseBook, parseTell, s, ARCHIVE_PREFIX, nextListedYear,
  type Book, type InRec, type Kind, type Lists, type OutRec,
} from "./msg";

const API = "https://connector-gateway.lovable.dev/google_drive";
const FOLDER_NAME = "nedal.MSG - مرفقات";

function headers(extra?: Record<string, string>) {
  const a = process.env["LOVABLE_API_KEY"];
  const b = process.env["GOOGLE_DRIVE_API_KEY"];
  if (!a || !b) throw new Error("جوجل درايف غير مربوط");
  return { Authorization: `Bearer ${a}`, "X-Connection-Api-Key": b, ...extra };
}
async function drive(path: string, init: RequestInit = {}) {
  const res = await fetch(`${API}${path}`, { ...init, headers: headers(init.headers as Record<string, string>) });
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
  return { lists, years: col(1), archiveCodes: col(8), ids: col(6), admin: col(7), viewers: col(9), senderEmail: col(10)[0] ?? "", whatsapp: col(11).map((x) => x.replace(/\.0$/, "").replace(/\D/g, ""))[0] ?? "" };
}

function roleOf(wb: XLSX.WorkBook, code: string): "admin" | "user" | "viewer" | null {
  const { ids, admin, viewers } = readData(wb);
  const c = code.trim();
  if (!c) return null;
  if (admin.includes(c)) return "admin";
  if (ids.includes(c)) return "user";
  if (viewers.includes(c)) return "viewer";
  return null;
}
function require(wb: XLSX.WorkBook, code: string, need: "read" | "user" | "admin") {
  const r = roleOf(wb, code);
  if (r === "viewer" && need !== "read") throw new Error("حساب المستكشف للاستعراض والطباعة فقط");
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
    const role = require(wb, data.code, "read");
    const inB = readBook(wb, "in");
    const outB = readBook(wb, "out");
    const { lists } = readData(wb);
    return {
      role,
      lists,
      years: readData(wb).years,
      whatsapp: readData(wb).whatsapp,
      nextIn: nextSerial(inB.active),
      nextOut: nextSerial(outB.active),
      in: role !== "user" ? inB : { active: inB.active, archives: [] },
      out: role !== "user" ? outB : { active: outB.active, archives: [] },
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
  .inputValidator((d) => auth.extend({ archiveCode: z.string().min(1), year: z.string().min(1).max(40) }).parse(d))
  .handler(async ({ data }) => {
    const L = await load();
    require(L.wb, data.code, "admin");
    const { lists, years, archiveCodes } = readData(L.wb);
    if (!archiveCodes.includes(data.archiveCode.trim())) throw new Error("رقم الأرشفة غير صحيح");
    if (!years.includes(data.year)) throw new Error("العام الدراسي غير موجود في ورقة data");
    const next = nextListedYear(years, data.year);
    if (!next) throw new Error("أضف العام الدراسي التالي في ورقة data قبل الأرشفة");
    for (const kind of ["in", "out"] as const) {
      const book = readBook(L.wb, kind) as Book<InRec | OutRec>;
      for (const r of book.active) if (!r.year) r.year = lists.year;
      const take = book.active.filter((r) => r.year === data.year);
      if (take.length) { book.active = book.active.filter((r) => r.year !== data.year); book.archives.unshift({ label: data.year, rows: take }); }
      writeBook(L.wb, kind, book as never);
    }
    const meta = XLSX.utils.aoa_to_sheet([[next]]);
    if (!L.wb.SheetNames.includes("_nedal_meta")) XLSX.utils.book_append_sheet(L.wb, meta, "_nedal_meta");
    else L.wb.Sheets["_nedal_meta"] = meta;
    await save(L);
    return { archived: ARCHIVE_PREFIX + data.year, year: next };
  });

async function folderId(kind: string) {
  const FOLDER_NAME = kind === "health" ? "متابعات الصحة المدرسية" : kind === "out" ? "الصادر" : "الوارد";
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
    const meta = { name: `${data.kind === "out" ? "صادر" : "وارد"} . ${data.subject} . ${readData(wb).lists.year}${ext}`, parents: [parent] };
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

/** School directory from the "tell" sheet (school, principal, phones, region). */
export const getSchools = createServerFn({ method: "POST" })
  .inputValidator((d) => auth.parse(d))
  .handler(async ({ data }) => {
    const { wb } = await load();
    require(wb, data.code, "read");
    return { schools: parseTell(rows(wb, "tell")) };
  });

export const readAttachment = createServerFn({ method: "POST" })
  .inputValidator((d) => auth.extend({ id: z.string().regex(/^[\w-]+$/) }).parse(d))
  .handler(async ({ data }) => {
    const { wb } = await load();
    const role = require(wb, data.code, "read");
    const allowed = [readBook(wb, "in"), readBook(wb, "out")].some((book) =>
      book.active.some((r) => r.attachment.includes(data.id)) || (role !== "user" && book.archives.some((a) => a.rows.some((r) => r.attachment.includes(data.id)))));
    if (!allowed) throw new Error("المرفق غير مرتبط بمراسلة متاحة");
    const res = await drive(`/drive/v3/files/${data.id}?alt=media`);
    const bytes = new Uint8Array(await res.arrayBuffer());
    let binary = "";
    for (let i = 0; i < bytes.length; i += 8192) binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
    return { base64: btoa(binary), type: res.headers.get("content-type") ?? "application/octet-stream" };
  });

/** Admin: permanently deletes every incoming and outgoing record of one school year (active and archived). */
export const deleteYear = createServerFn({ method: "POST" })
  .inputValidator((d) => auth.extend({ archiveCode: z.string().min(1), year: z.string().min(1).max(40), kind: z.enum(["in", "out"]) }).parse(d))
  .handler(async ({ data }) => {
    const L = await load();
    require(L.wb, data.code, "admin");
    const { lists, archiveCodes } = readData(L.wb);
    if (!archiveCodes.includes(data.archiveCode.trim())) throw new Error("رقم الأرشفة غير صحيح");
    let removed = 0;
    for (const kind of [data.kind]) {
      const book = readBook(L.wb, kind) as Book<InRec | OutRec>;
      const before = book.active.length + book.archives.reduce((n, a) => n + a.rows.length, 0);
      book.active = book.active.filter((r) => (r.year || lists.year) !== data.year);
      book.archives = book.archives.filter((a) => a.label !== data.year).map((a) => ({ ...a, rows: a.rows.filter((r) => r.year !== data.year) }));
      removed += before - book.active.length - book.archives.reduce((n, a) => n + a.rows.length, 0);
      writeBook(L.wb, kind, book as never);
    }
    await save(L);
    return { removed };
  });

/** Admin: undo a year-end archive — moves the archived year back to active and makes it the current year. */
export const undoArchive = createServerFn({ method: "POST" })
  .inputValidator((d) => auth.extend({ archiveCode: z.string().min(1), year: z.string().min(1).max(40) }).parse(d))
  .handler(async ({ data }) => {
    const L = await load();
    require(L.wb, data.code, "admin");
    if (!readData(L.wb).archiveCodes.includes(data.archiveCode.trim())) throw new Error("رقم الأرشفة غير صحيح");
    let moved = 0;
    for (const kind of ["in", "out"] as const) {
      const book = readBook(L.wb, kind) as Book<InRec | OutRec>;
      const arc = book.archives.find((a) => a.label === data.year);
      if (!arc) continue;
      book.archives = book.archives.filter((a) => a !== arc);
      book.active = [...arc.rows, ...book.active].sort((a, b) => a.serial - b.serial);
      moved += arc.rows.length;
      writeBook(L.wb, kind, book as never);
    }
    if (!moved) throw new Error("لا يوجد أرشيف لهذا العام");
    const meta = XLSX.utils.aoa_to_sheet([[data.year]]);
    if (!L.wb.SheetNames.includes("_nedal_meta")) XLSX.utils.book_append_sheet(L.wb, meta, "_nedal_meta");
    else L.wb.Sheets["_nedal_meta"] = meta;
    await save(L);
    return { moved };
  });

const b64 = (bytes: Uint8Array) => { let bin = ""; for (let i = 0; i < bytes.length; i += 8192) bin += String.fromCharCode(...bytes.subarray(i, i + 8192)); return btoa(bin); };
const hdr = (v: string) => (/^[\x00-\x7F]*$/.test(v) ? v : `=?UTF-8?B?${b64(new TextEncoder().encode(v))}?=`);

/** Sends a report or a letter (with its Drive attachment) by Gmail from the data-sheet sender address. */
export const sendMail = createServerFn({ method: "POST" })
  .inputValidator((d) => auth.extend({ to: z.string().email().max(200), subject: z.string().min(1).max(300), html: z.string().max(400000), attachmentId: z.string().regex(/^[\w-]+$/).optional() }).parse(d))
  .handler(async ({ data }) => {
    const { wb } = await load();
    require(wb, data.code, "user");
    const { senderEmail } = readData(wb);
    const mailKey = process.env["GOOGLE_MAIL_API_KEY"];
    if (!mailKey) throw new Error("بريد الإرسال غير مربوط بعد");
    const B = "nedalmail" + Date.now();
    const parts = [`From: ${hdr("nedal.MSG")} <${senderEmail}>`, `To: ${data.to}`, `Subject: ${hdr(data.subject)}`, "MIME-Version: 1.0", `Content-Type: multipart/mixed; boundary=${B}`, "", `--${B}`, "Content-Type: text/html; charset=UTF-8", "Content-Transfer-Encoding: base64", "", b64(new TextEncoder().encode(data.html))];
    if (data.attachmentId) {
      const meta = (await (await drive(`/drive/v3/files/${data.attachmentId}?fields=name,mimeType`)).json()) as { name: string; mimeType: string };
      const bytes = new Uint8Array(await (await drive(`/drive/v3/files/${data.attachmentId}?alt=media`)).arrayBuffer());
      parts.push(`--${B}`, `Content-Type: ${meta.mimeType}; name="${hdr(meta.name)}"`, `Content-Disposition: attachment; filename="${hdr(meta.name)}"`, "Content-Transfer-Encoding: base64", "", b64(bytes));
    }
    parts.push(`--${B}--`);
    const raw = b64(new TextEncoder().encode(parts.join("\r\n"))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
    const res = await fetch("https://connector-gateway.lovable.dev/google_mail/gmail/v1/users/me/messages/send", {
      method: "POST", headers: headers({ "X-Connection-Api-Key": mailKey, "Content-Type": "application/json" }), body: JSON.stringify({ raw }),
    });
    if (!res.ok) { console.error(`Gmail [${res.status}]: ${await res.text()}`); throw new Error(`تعذر إرسال البريد [${res.status}]`); }
    return { ok: true };
  });
