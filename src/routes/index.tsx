import { createFileRoute } from "@tanstack/react-router";
import { queryOptions, useSuspenseQuery, useQueryClient } from "@tanstack/react-query";
import { useEffect, useMemo, useState } from "react";
import { useBulkExport, type ExportJob } from "@/components/bulk-export";
import { getSchoolData } from "@/lib/school.functions";
import { AbsenceForm, FollowForm, MarksForm, FORM_LABELS, type FormKind } from "@/components/forms";

const schoolQuery = queryOptions({ queryKey: ["school"], queryFn: () => getSchoolData(), staleTime: 5 * 60_000 });

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "نماذج الصفوف للطباعة — n2028" },
      { name: "description", content: "إنشاء نماذج المتابعة والعلامات والغياب لكل صف وطباعتها مباشرة من ملف المدرسة." },
      { property: "og:title", content: "نماذج الصفوف للطباعة" },
      { property: "og:description", content: "نماذج متابعة وعلامات وغياب جاهزة للطباعة لكل شعبة." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary" },
    ],
  }),
  loader: ({ context }) => context.queryClient.ensureQueryData(schoolQuery),
  component: Index,
  errorComponent: ({ error }) => (
    <div className="app-shell"><div className="panel"><h1 className="app-title">تعذر قراءة الملف</h1><p className="text-muted-foreground mt-2">{(error as Error).message}</p></div></div>
  ),
});

function Index() {
  const { data } = useSuspenseQuery(schoolQuery);
  const qc = useQueryClient();
  const [kind, setKind] = useState<FormKind>("follow-all");
  const [classKey, setClassKey] = useState(data.classes[0]?.key ?? "");
  const [teacherIdx, setTeacherIdx] = useState(0);
  const [term, setTerm] = useState(data.terms[0] ?? "الأول");
  const now = new Date();
  const defMonth = Math.max(0, data.months.findIndex((m) => m.num === now.getMonth() + 1));
  const [monthIdx, setMonthIdx] = useState(defMonth);

  const isLang = kind === "follow-lang" || kind === "marks-lang";
  const classOptions = useMemo(() => data.classes.filter((c) => !isLang || c.grade >= 11), [data, isLang]);
  useEffect(() => {
    if (!classOptions.some((c) => c.key === classKey)) setClassKey(classOptions[0]?.key ?? "");
  }, [classOptions, classKey]);

  const teacher = data.teachers[teacherIdx];
  const cls = kind === "absence"
    ? data.classes.find((c) => c.key === teacher?.classKey)
    : data.classes.find((c) => c.key === classKey);
  const month = data.months[monthIdx];
  const bulk = useBulkExport();
  const termList = data.terms.length ? data.terms : ["الأول", "الثاني"];
  const termFor = (num: number) => (num >= 8 ? termList[0] : termList[1] ?? termList[0]) ?? "";
  const exportClasses = (k: FormKind) => {
    const lang = k.endsWith("lang");
    const jobs: ExportJob[] = data.classes
      .filter((c) => !lang || c.grade >= 11)
      .map((c) => ({
        key: k + c.key,
        landscape: k.startsWith("follow"),
        node: k.startsWith("follow") ? <FollowForm cls={c} term={term} lang={lang} /> : <MarksForm cls={c} term={term} lang={lang} />,
      }));
    bulk.start(jobs, `${FORM_LABELS[k]}.pdf`);
  };
  const exportAbsence = () => {
    const jobs: ExportJob[] = [];
    for (const t of data.teachers) {
      const c = data.classes.find((x) => x.key === t.classKey);
      if (!c) continue;
      for (const m of data.months)
        jobs.push({ key: t.classKey + m.year + m.num, landscape: true, node: <AbsenceForm cls={c} term={termFor(m.num)} month={m} teacher={t.name} /> });
    }
    bulk.start(jobs, "نماذج غياب الطلاب.pdf");
  };
  const landscape = kind === "absence" || kind.startsWith("follow");

  return (
    <div className="app-shell" dir="rtl">
      {!bulk.busy && <style>{`@page { size: A4 ${landscape ? "landscape" : "portrait"}; margin: 8mm; }`}</style>}
      <header className="no-print app-header">
        <div>
          <p className="app-kicker">ملف n2028 · Google Drive</p>
          <h1 className="app-title">نماذج الصفوف للطباعة</h1>
        </div>
        <button className="btn-ghost" onClick={() => qc.invalidateQueries({ queryKey: ["school"] })}>تحديث من الملف</button>
      </header>

      <section className="no-print panel">
        <div className="kind-tabs">
          {(Object.keys(FORM_LABELS) as FormKind[]).map((k) => (
            <button key={k} className="kind-tab" data-active={k === kind} onClick={() => setKind(k)}>{FORM_LABELS[k]}</button>
          ))}
        </div>
        <div className="controls">
          {kind === "absence" ? (
            <>
              <label className="control">مربي الصف
                <select value={teacherIdx} onChange={(e) => setTeacherIdx(Number(e.target.value))}>
                  {data.teachers.map((t, i) => <option key={i} value={i}>{t.name} — {t.classKey}</option>)}
                </select>
              </label>
              <label className="control">الشهر
                <select value={monthIdx} onChange={(e) => {
                  const i = Number(e.target.value);
                  setMonthIdx(i);
                  const m = data.months[i];
                  const ts = data.terms.length ? data.terms : ["الأول", "الثاني"];
                  if (m) setTerm((m.num >= 8 ? ts[0] : ts[1] ?? ts[0]) ?? "الأول");
                }}>
                  {data.months.map((m, i) => <option key={i} value={i}>{m.name} {m.year}</option>)}
                </select>
              </label>
            </>
          ) : (
            <label className="control">الصف
              <select value={classKey} onChange={(e) => setClassKey(e.target.value)}>
                {classOptions.map((c) => <option key={c.key} value={c.key}>{c.key} — {c.title}</option>)}
              </select>
            </label>
          )}
          <label className="control">الفصل
            <select value={term} onChange={(e) => setTerm(e.target.value)}>
              {(data.terms.length ? data.terms : ["الأول", "الثاني"]).map((t) => <option key={t}>{t}</option>)}
            </select>
          </label>
          <button className="btn-print" onClick={() => window.print()} disabled={!cls}>طباعة</button>
        </div>
        <div className="bulk-bar">
          {(["follow-all", "follow-lang", "marks-all", "marks-lang"] as FormKind[]).map((k) => (
            <button key={k} className="btn-bulk" disabled={bulk.busy} onClick={() => exportClasses(k)}>{FORM_LABELS[k]} PDF</button>
          ))}
          <button className="btn-bulk-main" disabled={bulk.busy} onClick={exportAbsence}>نماذج غياب الطلاب PDF</button>
          {bulk.busy && <span className="hint mt-0">جارٍ تجهيز {bulk.progress} صفحة… اختر "حفظ كملف PDF" في نافذة الطباعة</span>}
        </div>
        {cls && <p className="hint">{cls.students.length} طالباً في {cls.title}</p>}
      </section>

      {bulk.stage}
      <main className="paper-wrap">
        {!cls ? (
          <p className="hint no-print">لا توجد ورقة أسماء مطابقة لهذا الاختيار.</p>
        ) : kind === "absence" ? (
          month && <AbsenceForm cls={cls} term={term} month={month} teacher={teacher?.name ?? ""} />
        ) : kind.startsWith("follow") ? (
          <FollowForm cls={cls} term={term} lang={isLang} />
        ) : (
          <MarksForm cls={cls} term={term} lang={isLang} />
        )}
      </main>
    </div>
  );
}
