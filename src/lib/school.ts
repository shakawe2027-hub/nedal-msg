// Pure helpers shared by server parsing and the printable forms.

export type Student = { name: string; id: string };
export type SchoolClass = { key: string; title: string; grade: number; section: string; students: Student[] };
export type Teacher = { name: string; classKey: string; gradeName: string; section: string };
export type Month = { num: number; name: string; year: number };
export type SchoolData = { classes: SchoolClass[]; teachers: Teacher[]; months: Month[]; terms: string[] };

export const GRADE_NAMES: Record<number, string> = { 10: "العاشر", 11: "الحادي عشر", 12: "الثاني عشر" };

export function gradeFromName(name: string): number | null {
  const n = name.trim();
  if (n.includes("العاشر")) return 10;
  if (n.includes("الحادي")) return 11;
  if (n.includes("الثاني عشر")) return 12;
  return null;
}

/** "10. تشرين الأول / أكتوبر (October): 31 يوماً" -> { num: 10, name: "تشرين الأول / أكتوبر" } */
export function parseMonthLabel(label: string): { num: number; name: string } | null {
  const m = label.match(/^\s*(\d{1,2})\s*\.\s*([^(:]+)/);
  if (!m) return null;
  return { num: Number(m[1]), name: m[2].trim() };
}

export function daysInMonth(year: number, month: number): number {
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

const AR_DAYS = ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"];

export type DayCell = { day: number; weekday: string | null; weekend: boolean; outside: boolean };

/** Always 31 columns; days beyond the month's length are marked outside. Friday & Saturday are weekend. */
export function monthDays(year: number, month: number): DayCell[] {
  const len = daysInMonth(year, month);
  return Array.from({ length: 31 }, (_, i) => {
    const day = i + 1;
    if (day > len) return { day, weekday: null, weekend: false, outside: true };
    const wd = new Date(Date.UTC(year, month - 1, day)).getUTCDay();
    return { day, weekday: AR_DAYS[wd] ?? null, weekend: wd === 5 || wd === 6, outside: false };
  });
}

export function academicYear(year: number, month: number): string {
  return month >= 8 ? `${year}/${year + 1}` : `${year - 1}/${year}`;
}
