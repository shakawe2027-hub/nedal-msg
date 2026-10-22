import { createServerFn } from "@tanstack/react-start";
import * as XLSX from "xlsx";
import { GRADE_NAMES, gradeFromName, parseMonthLabel, type SchoolData, type SchoolClass, type Teacher, type Month } from "./school";

const GATEWAY_URL = "https://connector-gateway.lovable.dev/google_drive/drive/v3";
const FILE_NAME = "n2028";

const str = (v: unknown) => (v === null || v === undefined ? "" : String(v).trim());

async function gw(path: string) {
  const LOVABLE_API_KEY = process.env.LOVABLE_API_KEY;
  const DRIVE_KEY = process.env.GOOGLE_DRIVE_API_KEY;
  if (!LOVABLE_API_KEY || !DRIVE_KEY) throw new Error("Google Drive غير مربوط");
  const res = await fetch(`${GATEWAY_URL}${path}`, {
    headers: { Authorization: `Bearer ${LOVABLE_API_KEY}`, "X-Connection-Api-Key": DRIVE_KEY },
  });
  if (!res.ok) throw new Error(`Drive [${res.status}]: ${await res.text()}`);
  return res;
}

export const getSchoolData = createServerFn({ method: "GET" }).handler(async (): Promise<SchoolData> => {
  const q = encodeURIComponent(`name contains '${FILE_NAME}' and trashed=false`);
  const list = (await (await gw(`/files?q=${q}&orderBy=modifiedTime desc&fields=files(id,name)`)).json()) as {
    files: { id: string; name: string }[];
  };
  const file = list.files[0];
  if (!file) throw new Error(`لم يتم العثور على الملف ${FILE_NAME}`);
  const buf = await (await gw(`/files/${file.id}?alt=media`)).arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const rows = (name: string) =>
    XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[name], { header: 1, defval: null, blankrows: true });

  const classes: SchoolClass[] = [];
  for (const name of wb.SheetNames) {
    const m = name.trim().match(/^(10|11|12)\.(.+)$/);
    if (!m) continue;
    const r = rows(name);
    const grade = Number(m[1]);
    const section = m[2].trim();
    const students = r
      .slice(3)
      .map((row) => ({ name: str(row?.[1]), id: str(row?.[2]) }))
      .filter((s) => s.name && !s.name.startsWith("="));
    classes.push({ key: name.trim(), title: `${GRADE_NAMES[grade]} / ${section}`, grade, section, students });
  }

  const teachers: Teacher[] = [];
  const months: Month[] = [];
  const terms: string[] = [];
  const shmf = wb.SheetNames.find((n) => n.trim() === "ش.م.ف");
  if (shmf) {
    for (const row of rows(shmf).slice(1)) {
      const year = Number(row?.[1]);
      const ml = parseMonthLabel(str(row?.[2]));
      if (ml && year) months.push({ ...ml, year });
      const term = str(row?.[3]);
      if (term) terms.push(term);
      const gName = str(row?.[4]);
      const sec = str(row?.[5]);
      const tName = str(row?.[6]);
      const g = gradeFromName(gName);
      if (g && sec && tName) teachers.push({ name: tName, gradeName: gName, section: sec, classKey: `${g}.${sec}` });
    }
  }
  months.sort((a, b) => a.year - b.year || a.num - b.num);
  return { classes, teachers, months, terms };
});
