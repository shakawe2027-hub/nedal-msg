import { createFileRoute, Link } from "@tanstack/react-router";
import { useQuery } from "@tanstack/react-query";
import { useServerFn } from "@tanstack/react-start";
import { useState } from "react";
import { CalendarDays, Printer } from "lucide-react";
import { toast } from "sonner";
import { Shell } from "@/components/msg/shell";
import { OriginalDocument } from "@/components/msg/health-workspace";
import { Button } from "@/components/ui/button";
import { getHealthData, getHealthDocument } from "@/lib/msg.functions";
import { displayDate, matchArabic } from "@/lib/health";
import type { Session } from "@/lib/session";
export const Route = createFileRoute("/health-monthly")({
  head: () => ({ meta: [{ title: "رصد الإصابات الشهرية — nedal.Health" }, { name: "description", content: "أرشيف تحويلات الطلاب الصحية حسب الشهر والسنة وسبب التحويل." }, { property: "og:title", content: "رصد الإصابات الشهرية — nedal.Health" }, { property: "og:description", content: "متابعة التقارير الصحية المؤرشفة وأسباب التحويل الشهرية." }, { property: "og:type", content: "website" }, { name: "twitter:card", content: "summary" }] }),
  component: () => <Shell>{s => <Monthly session={s} />}</Shell>,
});
function Monthly({ session }: { session: Session }) {
  const q = useQuery({ queryKey: ["health", session.code], queryFn: () => getHealthData({ data: { code: session.code } }) });
  const fetchDoc = useServerFn(getHealthDocument);
  const [month, setMonth] = useState("");
  const [term, setTerm] = useState("");
  const [doc, setDoc] = useState("");
  const [busy, setBusy] = useState(false);
  const all = q.data?.archives ?? [];
  const months = [...new Set(all.map(a => a.date.slice(0, 7)))].sort().reverse();
  const monthLabel = (m: string) => new Date(`${m}-01T12:00:00`).toLocaleDateString("ar-PS", { month: "long", year: "numeric" });
  return <div className="health-workspace">
    <div className="health-archive-head no-print"><h1>رصد الإصابات الشهرية</h1><Button asChild variant="outline"><Link to="/health">تقرير تحويل جديد</Link></Button></div>
    <div className="health-search no-print"><input aria-label="بحث في التقارير الصحية" className="input" placeholder="اسم الطالب أو رقم الهوية" value={term} onChange={e => setTerm(e.target.value)} /><input aria-label="الشهر والسنة" className="input" type="month" value={month} onChange={e => setMonth(e.target.value)} /></div>
    {q.error && <p className="err">{(q.error as Error).message}</p>}
    {q.isLoading && <p>جارٍ تحميل الأرشيف…</p>}
    {months.filter(m => !month || m === month).map(m => {
      const rows = all.filter(a => a.date.startsWith(m) && (matchArabic(a.name).includes(matchArabic(term)) || a.studentId.includes(term)));
      return <details className="monthly-group no-print" key={m} open><summary><CalendarDays />{monthLabel(m)}<span>{rows.length} تقرير</span></summary>
        {rows.map(a => <article key={a.id} className="monthly-report"><div><h2>{a.name}</h2><p><bdi>{a.studentId}</bdi> · {displayDate(a.date)}</p><p className="referral-reason">{a.reasons}</p></div><Button variant="outline" disabled={busy} onClick={async () => {
          setBusy(true); try { setDoc((await fetchDoc({ data: { code: session.code, fileId: a.fileId } })).base64); }
          catch (e) { toast.error((e as Error).message); } finally { setBusy(false); }
        }}>عرض التقرير</Button></article>)}
      </details>;
    })}
    {!q.isLoading && !all.length && <p className="muted no-print">لا توجد تقارير مؤرشفة</p>}
    {doc && <><Button className="no-print" variant="outline" onClick={() => window.print()}><Printer />طباعة صفحة واحدة</Button><OriginalDocument base64={doc} /></>}
  </div>;
}