// Pure correspondence logic shared by server and client (no I/O).

export const ARCHIVE_PREFIX = "مراسلات منتهية-أرشيف ";

export const IN_HEADER = ["عداد متسلسل", "رقم ملف الوارد", "تاريخ الوارد", "المبحث", "الجهة/القسم - المرسل", "المديرية/الوحدة", "اجراء", "ملاحظات المسؤول", "نسخة مرفقة"];
export const OUT_HEADER = ["الرقم الصادر", "رقم الملف", "التاريخ", "المبحث", "جهة الاختصاص", "الملحقات", "نسخة مرفقة"];

export type InRec = { serial: number; fileNo: string; date: string; subject: string; sender: string; unit: string; action: string; notes: string; attachment: string; year?: string; institution?: string; status?: string; reminderDate?: string };
export type OutRec = { serial: number; fileNo: string; date: string; subject: string; authority: string; enclosures: string; attachment: string; year?: string; institution?: string; status?: string; reminderDate?: string };
export type Kind = "in" | "out";
export type RecOf<K extends Kind> = K extends "in" ? InRec : OutRec;

export type Book<T> = { active: T[]; archives: { label: string; rows: T[] }[] };

export type Lists = { year: string; units: string[]; authorities: string[]; enclosures: string[]; actions: string[]; institution: string };

export const s = (v: unknown) => (v === null || v === undefined ? "" : String(v).trim());

function toRec(kind: Kind, r: unknown[]): InRec | OutRec {
  const serial = Number(r[0]) || 0;
  if (kind === "in")
    return { serial, fileNo: s(r[1]), date: s(r[2]), subject: s(r[3]), sender: s(r[4]), unit: s(r[5]), action: s(r[6]), notes: s(r[7]), attachment: s(r[8]) };
  return { serial, fileNo: s(r[1]), date: s(r[2]), subject: s(r[3]), authority: s(r[4]), enclosures: s(r[5]), attachment: s(r[6]) };
}

export function recToRow(kind: Kind, x: InRec | OutRec): (string | number)[] {
  if (kind === "in") {
    const r = x as InRec;
    return [r.serial, r.fileNo, r.date, r.subject, r.sender, r.unit, r.action, r.notes, r.attachment];
  }
  const r = x as OutRec;
  return [r.serial, r.fileNo, r.date, r.subject, r.authority, r.enclosures, r.attachment];
}

/** Rows after a marker row ("مراسلات منتهية-أرشيف YYYY") belong to that archive. Row 0 is the header. */
export function parseBook<K extends Kind>(kind: K, rows: unknown[][]): Book<RecOf<K>> {
  const book: Book<RecOf<K>> = { active: [], archives: [] };
  let cur: RecOf<K>[] = book.active;
  for (const r of rows.slice(1)) {
    const a = s(r?.[0]);
    if (a.startsWith(ARCHIVE_PREFIX.trim())) {
      const sec = { label: a.replace(ARCHIVE_PREFIX.trim(), "").trim(), rows: [] as RecOf<K>[] };
      book.archives.push(sec);
      cur = sec.rows;
      continue;
    }
    if (!r || r.every((c) => s(c) === "")) continue;
    const header = rows[0] ?? [];
    const offset = header.some((c) => s(c) === "الوحدة/المؤسسة") ? 2 : 0;
    const rec = toRec(kind, r.slice(offset));
    rec.institution = offset ? s(r[0]) : "";
    rec.year = offset ? s(r[1]) : "";
    rec.status = s(r[header.findIndex((c) => s(c) === "حالة المتابعة")]);
    rec.reminderDate = s(r[header.findIndex((c) => s(c) === "تاريخ التنبيه")]);
    cur.push(rec as RecOf<K>);
  }
  return book;
}

export function buildRows<K extends Kind>(kind: K, book: Book<RecOf<K>>): (string | number)[][] {
  const out: (string | number)[][] = [["الوحدة/المؤسسة", "العام الدراسي", ...(kind === "in" ? [...IN_HEADER, "حالة المتابعة", "تاريخ التنبيه"] : OUT_HEADER)]];
  for (const r of book.active) out.push([r.institution ?? "", r.year ?? "", ...recToRow(kind, r), ...(kind === "in" ? [r.status ?? "", r.reminderDate ?? ""] : [])]);
  for (const a of book.archives) {
    out.push([]);
    out.push([ARCHIVE_PREFIX + a.label]);
    for (const r of a.rows) out.push([r.institution ?? "", r.year ?? "", ...recToRow(kind, r), ...(kind === "in" ? [r.status ?? "", r.reminderDate ?? ""] : [])]);
  }
  return out;
}

export const nextSerial = (rows: { serial: number }[]) => rows.reduce((m, r) => Math.max(m, r.serial), 0) + 1;

/** Moves active rows into a new archive (newest archive first) so serials restart at 1. */
export function archiveBook<T>(book: Book<T>, label: string): Book<T> {
  if (!book.active.length) return book;
  return { active: [], archives: [{ label, rows: book.active }, ...book.archives] };
}

/** "2027/2026" -> "2028/2027" (keeps the order the sheet uses). */
export function bumpYear(y: string): string {
  return y.replace(/\d{4}/g, (n) => String(Number(n) + 1));
}

export function driveFileId(link: string): string | null {
  return link.match(/\/d\/([\w-]+)/)?.[1] ?? link.match(/[?&]id=([\w-]+)/)?.[1] ?? null;
}

export function nextListedYear(years: string[], current: string): string | null {
  const i = years.indexOf(current);
  return i >= 0 ? years[i + 1] ?? null : null;
}

export type School = { no: number; school: string; principal: string; schoolPhone: string; unified: string; principalPhone: string; region: string };

/** Parses the "tell" sheet: a region header row ("اذنا (24)") precedes its numbered school rows. */
export function parseTell(rows: unknown[][]): School[] {
  const out: School[] = [];
  let region = "";
  for (const r of rows) {
    if (!r || r.every((c) => s(c) === "")) continue;
    if (typeof r[0] === "string" && /\(\d+\)/.test(r[0])) {
      region = s(r[0]).replace(/\s*\(\d+\)\s*$/, "");
      continue;
    }
    const no = Number(r[0]);
    if (!no || !s(r[1])) continue;
    out.push({ no, school: s(r[1]), principal: s(r[2]), schoolPhone: s(r[3]), unified: s(r[4]), principalPhone: s(r[5]), region });
  }
  return out;
}
