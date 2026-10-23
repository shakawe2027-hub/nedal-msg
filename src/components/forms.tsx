import { academicYear, monthDays, type SchoolClass, type Month } from "@/lib/school";

export type FormKind = "follow-all" | "follow-lang" | "marks-all" | "marks-lang" | "absence";

export const FORM_LABELS: Record<FormKind, string> = {
  "follow-all": "متابعة جميع المواد",
  "follow-lang": "متابعة لغات 11+12",
  "marks-all": "علامات جميع الصفوف",
  "marks-lang": "علامات لغات 11+12",
  absence: "غياب",
};

import type { CSSProperties, ReactNode } from "react";

function Sheet({ rows, landscape, extra, children }: { rows: number; landscape: boolean; extra?: string; children: ReactNode }) {
  return (
    <div className={`sheet ${landscape ? "sheet-landscape" : "sheet-portrait"} ${extra ?? ""}`} style={{ "--rows": rows } as CSSProperties}>
      {children}
      <div className="sheet-footer">by: nedal qazzaz</div>
    </div>
  );
}

function Field({ label, value }: { label: string; value?: string }) {
  return (
    <span className="sheet-field">
      <b>{label}:</b> <span className="sheet-field-value">{value || "\u00A0"}</span>
    </span>
  );
}

export function FollowForm({ cls, term, lang }: { cls: SchoolClass; term: string; lang: boolean }) {
  const w = lang ? 3 : 2;
  const groups = [
    { t: "واجبات مدرسية", n: 3 },
    { t: "أوراق عمل وأبحاث ومهمات", n: 3 },
    { t: "التعاون ومشاركة الطالب الصفية واللاصفية", n: 2 },
    { t: "سلوك الطالب والحفاظ على الأدوات والأجهزة", n: 2 },
  ];
  return (
    <Sheet rows={cls.students.length + 2} landscape>
      <div className="sheet-head">
        <h2>{lang ? "متابعة التقويم النوعي — اللغات" : "متابعة التقويم النوعي — جميع المواد"}</h2>
        <div className="sheet-fields">
          <Field label="الصف" value={cls.title} />
          <Field label="الفصل" value={term} />
          <Field label="المادة" />
          <Field label="المعلم" />
        </div>
      </div>
      <table className="sheet-table">
        <thead>
          <tr>
            <th rowSpan={2} className="col-num">الرقم</th>
            <th rowSpan={2} className="col-name">الاســم</th>
            {groups.map((g) => [
              <th key={g.t} colSpan={g.n} className="col-group">{g.t}</th>,
              <th key={g.t + "s"} className="col-sum">المجموع</th>,
            ])}
            <th className="col-sum">مجموع التقويم النوعي</th>
          </tr>
          <tr>
            {groups.map((g) => [
              ...Array.from({ length: g.n }, (_, i) => <th key={g.t + i}>{w}</th>),
              <th key={g.t + "sw"} className="col-sum">{w * g.n}</th>,
            ])}
            <th className="col-sum">{w * 10}</th>
          </tr>
        </thead>
        <tbody>
          {cls.students.map((s, i) => (
            <tr key={i}>
              <td>{i + 1}</td>
              <td className="cell-name">{s.name}</td>
              {Array.from({ length: 15 }, (_, j) => <td key={j} />)}
            </tr>
          ))}
        </tbody>
      </table>
    </Sheet>
  );
}

export function MarksForm({ cls, term, lang }: { cls: SchoolClass; term: string; lang: boolean }) {
  const cols = lang
    ? ["اختبارات قصيرة 15%", "اختبار نصف الفصل 30%", "اختبارات قصيرة 15%", "التقويم النوعي 30%", "اختبار نهاية الفصل 60%"]
    : ["اختبارات قصيرة 10%", "اختبار نصف الفصل 20%", "اختبارات قصيرة 10%", "التقويم النوعي 20%", "اختبار نهاية الفصل 40%"];
  const weights = lang ? [15, 30, 15, 30, 60, 150] : [10, 20, 10, 20, 40, 100];
  return (
    <Sheet rows={cls.students.length + 3} landscape={false}>
      <div className="sheet-head">
        <h2>{lang ? "كشف علامات اللغات" : "كشف علامات الطلبة"}</h2>
        <div className="sheet-fields">
          <Field label="الصف والشعبة" value={cls.title} />
          <Field label="الفصل" value={term} />
          <Field label="المادة" />
          <Field label="المعلم" />
        </div>
      </div>
      <table className="sheet-table">
        <thead>
          <tr>
            <th rowSpan={3} className="col-num">الرقم</th>
            <th rowSpan={3} className="col-name">اسم الطالب</th>
            <th colSpan={5} className="col-group">علامات الفصل {term}</th>
            <th rowSpan={2} className="col-sum">معدل الفصل {term}</th>
          </tr>
          <tr>{cols.map((c, i) => <th key={i}>{c}</th>)}</tr>
          <tr>{weights.map((w, i) => <th key={i} className={i === 5 ? "col-sum" : ""}>{w}</th>)}</tr>
        </thead>
        <tbody>
          {cls.students.map((s, i) => (
            <tr key={i}>
              <td>{i + 1}</td>
              <td className="cell-name">{s.name}</td>
              {Array.from({ length: 6 }, (_, j) => <td key={j} />)}
            </tr>
          ))}
        </tbody>
      </table>
    </Sheet>
  );
}

export function AbsenceForm({ cls, term, month, teacher }: { cls: SchoolClass; term: string; month: Month; teacher: string }) {
  const days = monthDays(month.year, month.num);
  const dc = (d: (typeof days)[number]) => (d.outside ? "day-out" : d.weekend ? "day-weekend" : "");
  return (
    <Sheet rows={cls.students.length + 5} landscape extra="sheet-absence">
      <div className="sheet-head">
        <h2>جدول الحضور</h2>
        <div className="sheet-fields">
          <Field label="مربي الصف" value={teacher} />
          <Field label="الصف" value={cls.title} />
          <Field label="الشهر" value={`${month.name} ${month.year}`} />
          <Field label="الفصل" value={term} />
          <Field label="السنة الدراسية" value={academicYear(month.year, month.num)} />
        </div>
      </div>
      <table className="sheet-table">
        <thead>
          <tr>
            <th rowSpan={3} className="col-num">الرقم</th>
            <th className="col-name">التاريخ</th>
            {days.map((d) => <th key={d.day} className={dc(d)}>{d.day}</th>)}
            <th rowSpan={3} className="col-sum">مجموع الغياب</th>
          </tr>
          <tr>
            <th className="col-name">اليوم</th>
            {days.map((d) => <th key={d.day} className={`day-name ${dc(d)}`}><span>{d.weekday ?? ""}</span></th>)}
          </tr>
          <tr>
            <th className="col-name">اسم الطالب</th>
            {days.map((d) => <th key={d.day} className={dc(d)} />)}
          </tr>
        </thead>
        <tbody>
          {cls.students.map((s, i) => (
            <tr key={i}>
              <td>{i + 1}</td>
              <td className="cell-name">{s.name}</td>
              {days.map((d) => <td key={d.day} className={dc(d)} />)}
              <td />
            </tr>
          ))}
          <tr>
            <td />
            <td className="cell-name"><b>مجموع الحضور اليومي</b></td>
            {days.map((d) => <td key={d.day} className={dc(d)} />)}
            <td />
          </tr>
        </tbody>
      </table>
    </Sheet>
  );
}
