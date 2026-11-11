import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { z } from "zod";

export type Student = { row: number; name: string; id: string; birth: string; address: string };
export const referralFields = z.object({
  studentId: z.string().min(1), destination: z.string().trim().min(1).max(200),
  address: z.string().trim().min(1).max(200), studentAddress: z.string().trim().min(1).max(200),
  reasons: z.string().trim().min(1).max(1000), year: z.string().trim().min(1).max(40),
  date: z.string().regex(/^\d{4}-\d{2}-\d{2}$/), headerDate: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
});
export type ReferralInput = z.infer<typeof referralFields>;
export type Referral = ReferralInput & { name: string; birth: string; age: number; school: string; directorate: string; principal: string };
export type HealthArchive = { id: string; name: string; studentId: string; date: string; reasons: string; fileId: string };
export function ageAt(birth: string, today: string) {
  const [y = 0, m = 0, d = 0] = birth.split("-").map(Number);
  const [ty = 0, tm = 0, td = 0] = today.split("-").map(Number);
  return Math.max(0, ty - y - (tm < m || (tm === m && td < d) ? 1 : 0));
}
export const matchArabic = (v: string) => v.replace(/[ً-ْـ]/g, "").replace(/[أإآ]/g, "ا").replace(/ة/g, "ه").replace(/ى/g, "ي").trim();
export const displayDate = (date: string) => date.split("-").reverse().join(" / ");
const escapeXml = (v: string) => v.replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;" })[c] ?? c);
function cellText(cell: string, value: string) {
  let first = true;
  return cell.replace(/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/g, () => {
    const text = first ? escapeXml(value) : ""; first = false;
    return `<w:t xml:space="preserve">${text}</w:t>`;
  });
}
/** Fill the original Word cells without recreating its letterhead, tables or images. */
export function fillReferral(template: Uint8Array, r: Referral): Uint8Array {
  const files = unzipSync(template);
  const document = files["word/document.xml"];
  if (!document) throw new Error("نموذج Word غير صالح");
  const values: Record<number, Record<number, string>> = {
    0: { 0: `الجهة المحول اليها الطالب:- ${r.destination}` },
    1: { 1: r.school, 3: r.directorate },
    2: { 1: r.name, 3: String(r.age), 5: r.address, 7: r.year },
    3: { 1: r.studentAddress, 3: r.studentId, 5: displayDate(r.birth) },
    4: { 1: r.reasons }, 6: { 1: displayDate(r.date), 3: r.principal },
  };
  let tableIndex = 0;
  files["word/document.xml"] = strToU8(strFromU8(document).replace(/<w:tbl[ >][\s\S]*?<\/w:tbl>/g, table => {
    if (tableIndex++ !== 0) return table;
    let rowIndex = 0;
    return table.replace(/<w:tr[ >][\s\S]*?<\/w:tr>/g, row => {
      const v = values[rowIndex++]; let cellIndex = 0;
      return row.replace(/<w:tc[ >][\s\S]*?<\/w:tc>/g, cell => {
        const value = v?.[cellIndex++]; return value === undefined ? cell : cellText(cell, value);
      });
    });
  }));
  for (const name of Object.keys(files)) {
    const file = files[name];
    if (!file || !/^word\/header\d+\.xml$/.test(name)) continue;
    files[name] = strToU8(strFromU8(file).replace(/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/g, text => /\.{6,}/.test(text) ? `<w:t>${displayDate(r.headerDate)}</w:t>` : text));
  }
  return zipSync(files);
}