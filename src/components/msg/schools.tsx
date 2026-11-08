import { useMemo, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { Phone, School as SchoolIcon, Search, X } from "lucide-react";
import { getSchools } from "@/lib/msg.functions";

/** Arabic-insensitive match: drops diacritics/tatweel and unifies alef/taa-marbuta/alef-maqsura. */
const norm = (v: string) =>
  v.replace(/[ً-ْـ]/g, "").replace(/[أإآ]/g, "ا").replace(/ة/g, "ه").replace(/ى/g, "ي").trim();

/** Searchable directory of schools and principals from the "tell" sheet. */
export function SchoolsModal({ code, onClose }: { code: string; onClose: () => void }) {
  const q = useQuery({ queryKey: ["schools", code], queryFn: () => getSchools({ data: { code } }), staleTime: 300_000 });
  const [term, setTerm] = useState("");
  const all = useMemo(() => q.data?.schools ?? [], [q.data]);
  const results = useMemo(() => {
    const t = norm(term);
    if (!t) return all;
    return all.filter((r) => norm(r.school).includes(t) || norm(r.principal).includes(t));
  }, [all, term]);
  const shown = results.slice(0, 60);
  return (
    <div className="modal-back" onClick={onClose}>
      <div className="modal modal-sm schools-modal" dir="rtl" onClick={(e) => e.stopPropagation()}>
        <div className="row-gap" style={{ justifyContent: "space-between" }}>
          <h2 className="schools-title"><SchoolIcon /> دليل المدارس</h2>
          <button className="btn btn-ghost" onClick={onClose}><X /> إغلاق</button>
        </div>
        <div className="schools-search">
          <Search />
          <input className="input" autoFocus placeholder="ابحث باسم المدرسة أو اسم المدير…" value={term} onChange={(e) => setTerm(e.target.value)} />
        </div>
        {q.isLoading && <p className="muted">جارٍ تحميل الدليل…</p>}
        {q.error && <p className="err">{(q.error as Error).message}</p>}
        {q.data && (
          <p className="muted">
            {term ? `نتائج البحث: ${results.length}` : `عدد المدارس: ${all.length}`}
            {results.length > shown.length ? ` — يعرض أول ${shown.length}` : ""}
          </p>
        )}
        <div className="schools-list">
          {shown.map((r) => (
            <div key={`${r.region}-${r.no}`} className="school-card">
              <div className="school-card-head"><b>{r.school}</b><span className="chip">{r.region}</span></div>
              <p className="school-principal">{r.principal}</p>
              <div className="school-phones">
                {r.schoolPhone && <a href={`tel:${r.schoolPhone}`}><Phone /> هاتف المدرسة: <bdi>{r.schoolPhone}</bdi></a>}
                {r.principalPhone && <a href={`tel:${r.principalPhone}`}><Phone /> تلفون المدير: <bdi>{r.principalPhone}</bdi></a>}
                {r.unified && <a href={`tel:${r.unified}`}><Phone /> الوطنية الموحد: <bdi>{r.unified}</bdi></a>}
              </div>
            </div>
          ))}
          {q.data && !results.length && <p className="muted">لا توجد نتائج مطابقة</p>}
        </div>
      </div>
    </div>
  );
}
