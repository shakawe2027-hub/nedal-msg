import { useEffect, useRef, useState } from "react";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { useServerFn } from "@tanstack/react-start";
import { Printer, Save, Search, FileHeart } from "lucide-react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Field, today } from "./fields";
import { getHealthData, getHealthDocument, saveHealthReferral } from "@/lib/msg.functions";
import { ageAt, displayDate, matchArabic, type ReferralInput, type Student } from "@/lib/health";
import type { Session } from "@/lib/session";

export function OriginalDocument({ base64 }: { base64: string }) {
  const ref = useRef<HTMLDivElement>(null);
  const [error, setError] = useState("");
  useEffect(() => {
    let cancelled = false;
    const container = ref.current;
    if (!container) return;
    container.replaceChildren();
    setError("");
    import("docx-preview").then(async ({ renderAsync }) => {
      if (cancelled) return;
      const bytes = Uint8Array.from(atob(base64), c => c.charCodeAt(0));
      await renderAsync(bytes, container, undefined, { className: "health-docx", inWrapper: true, ignoreLastRenderedPageBreak: true, breakPages: false, renderHeaders: true, renderFooters: true, useBase64URL: true });
      const page = container.querySelector<HTMLElement>("section.health-docx");
      if (page) {
        const last = page.querySelector("article > p:last-child");
        if (last?.textContent?.includes("المعالج")) last.classList.add("health-signature-line");
        page.style.setProperty("--print-zoom", String(Math.min(1, 1050 / page.scrollHeight)));
      }
    }).catch(() => setError("تعذر عرض النموذج الأصلي"));
    return () => { cancelled = true; };
  }, [base64]);
  return <>{error && <p className="err">{error}</p>}<div ref={ref} className="health-document" /></>;
}

export function HealthWorkspace({ session }: { session: Session }) {
  const q = useQuery({ queryKey: ["health", session.code], queryFn: () => getHealthData({ data: { code: session.code } }) });
  const qc = useQueryClient();
  const preview = useServerFn(getHealthDocument);
  const save = useServerFn(saveHealthReferral);
  const [term, setTerm] = useState("");
  const [searched, setSearched] = useState(false);
  const [student, setStudent] = useState<Student | null>(null);
  const [month, setMonth] = useState("");
  const [busy, setBusy] = useState(false);
  const [doc, setDoc] = useState("");
  const [f, setF] = useState<ReferralInput>({ studentId: "", destination: "", address: "", studentAddress: "", reasons: "مدير المدرسة – منسق الصحة المدرسية – مدرسة ذكور صلاح الدين الثانوية – دورا", year: "", date: today(), headerDate: today() });
  useEffect(() => { if (q.data?.info.year) setF(p => ({ ...p, year: q.data.info.year })); }, [q.data?.info.year]);
  const set = (k: keyof ReferralInput, v: string) => { setF(p => ({ ...p, [k]: v })); setDoc(""); };
  const results = (q.data?.students ?? []).filter(s => matchArabic(s.name).includes(matchArabic(term)) || s.id.includes(term.trim()));
  const archives = (q.data?.archives ?? []).filter(a => !month || a.date.startsWith(month)).sort((a, b) => b.date.localeCompare(a.date));
  const months = [...new Set(q.data?.archives.map(a => a.date.slice(0, 7)) ?? [])].sort().reverse();
  const openArchive = async (fileId: string) => {
    setBusy(true);
    try { setDoc((await preview({ data: { code: session.code, fileId } })).base64); }
    catch (e) { toast.error((e as Error).message); } finally { setBusy(false); }
  };
  return <div className="health-workspace">
    <div className="page-heading no-print"><h1><FileHeart /> متابعات الصحة المدرسية</h1></div>
    {q.isLoading && <p>جارٍ تحميل بيانات الطلاب…</p>}
    {q.error && <p className="err">{(q.error as Error).message}</p>}
    {session.role !== "viewer" && <form className="health-form no-print" onSubmit={async e => {
      e.preventDefault(); setBusy(true);
      try {
        const result = await save({ data: { code: session.code, input: f } });
        setDoc(result.base64); await qc.invalidateQueries({ queryKey: ["health"] });
        toast.success("تمت أرشفة النموذج وحفظ أسباب التحويل في بيانات الطالب");
      } catch (error) { toast.error((error as Error).message); } finally { setBusy(false); }
    }}>
      <div className="health-search">
        <input aria-label="اسم الطالب أو رقم الهوية" role="combobox" aria-expanded={searched && !!term.trim()} aria-controls="student-options" autoComplete="off" className="input" value={term} onChange={e => { setTerm(e.target.value); setSearched(true); }} placeholder="اسم الطالب أو رقم الهوية" />
        <Button type="button" variant="outline" onClick={() => setSearched(true)} disabled={!term.trim()}><Search /> بحث</Button>
      </div>
      {searched && !!term.trim() && <div className="student-results" role="listbox" id="student-options">{results.slice(0, 30).map(s => <Button role="option" key={s.id} variant="ghost" type="button" onClick={() => {
        setStudent(s); setF(p => ({ ...p, studentId: s.id, address: s.address })); setDoc(""); setSearched(false);
      }}>{s.name} · <bdi>{s.id}</bdi> · <bdi>{displayDate(s.birth)}</bdi></Button>)}{!results.length && <p>لا توجد نتائج مطابقة</p>}</div>}
      <div className="grid-form">
        <Field label="اسم الطالب"><input className="input" readOnly value={student?.name ?? ""} /></Field>
        <Field label="رقم هوية الطالب"><input className="input" readOnly value={student?.id ?? ""} /></Field>
        <Field label="تاريخ الميلاد"><input className="input" readOnly value={student?.birth ?? ""} /></Field>
        <Field label="العمر"><input className="input" readOnly value={student ? ageAt(student.birth, today()) : ""} /></Field>
        <Field label="اسم المدرسة"><input className="input" readOnly value={q.data?.info.school ?? ""} /></Field>
        <Field label="المديرية"><input className="input" readOnly value={q.data?.info.directorate ?? ""} /></Field>
        <Field label="اسم مدير المدرسة"><input className="input" readOnly value={q.data?.info.principal ?? ""} /></Field>
        <Field label="السنة الدراسية" required><input className="input" required value={f.year} onChange={e => set("year", e.target.value)} /></Field>
        <Field label="الجهة المحول إليها الطالب" required><input className="input" required value={f.destination} onChange={e => set("destination", e.target.value)} /></Field>
        <Field label="العنوان" required><input className="input" required value={f.address} onChange={e => set("address", e.target.value)} /></Field>
        <Field label="عنوان الطالب" required><input className="input" required value={f.studentAddress} onChange={e => set("studentAddress", e.target.value)} /></Field>
        <Field label="تاريخ الصدور" required><input className="input" type="date" required value={f.date} onChange={e => set("date", e.target.value)} /></Field>
        <Field label="التاريخ في رأس الورقة" required><input className="input" type="date" required value={f.headerDate} onChange={e => set("headerDate", e.target.value)} /></Field>
        <Field label="أسباب التحويل" required><textarea className="input" required rows={3} value={f.reasons} onChange={e => set("reasons", e.target.value)} /></Field>
      </div>
      <div className="row-gap">
        <Button type="button" variant="outline" disabled={busy || !student} onClick={async () => {
          setBusy(true); try { setDoc((await preview({ data: { code: session.code, input: f } })).base64); }
          catch (e) { toast.error((e as Error).message); } finally { setBusy(false); }
        }}>معاينة النموذج</Button>
        <Button type="submit" disabled={busy || !student}><Save />{busy ? "جارٍ التنفيذ…" : "حفظ وأرشفة النموذج"}</Button>
      </div>
    </form>}
    <section className="health-archive no-print">
      <div className="health-archive-head"><h2>التقارير الصادرة ({archives.length})</h2><select aria-label="شهر الصدور" className="input" value={month} onChange={e => setMonth(e.target.value)}><option value="">جميع الأشهر</option>{months.map(m => <option key={m}>{m}</option>)}</select></div>
      {archives.map(a => <div className="health-archive-row" key={a.id}><b>{a.name}</b><span>{displayDate(a.date)}</span><Button variant="outline" disabled={busy} onClick={() => openArchive(a.fileId)}>عرض النموذج</Button></div>)}
      {!archives.length && <p className="muted">لا توجد نماذج مؤرشفة</p>}
    </section>
    {doc && <><div className="row-gap no-print"><Button variant="outline" onClick={() => window.print()}><Printer /> طباعة صفحة واحدة</Button></div><OriginalDocument base64={doc} /></>}
  </div>;
}