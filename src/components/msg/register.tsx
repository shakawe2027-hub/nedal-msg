import { Link } from "@tanstack/react-router";
import { useEffect, useMemo, useState, type ReactNode } from "react";
import { useQueryClient } from "@tanstack/react-query";
import { toast } from "sonner";
import * as XLSX from "xlsx";
import { Archive, Download, FileDown, Paperclip, Pencil, Printer, X, KeyRound, MailOpen } from "lucide-react";
import { closeYear, login, updateRecord, readAttachment } from "@/lib/msg.functions";
import { saveSession, useMsgData, type Session } from "@/lib/session";
import { driveFileId, IN_HEADER, OUT_HEADER, type InRec, type Kind, type OutRec } from "@/lib/msg";
import { Shell, Signature } from "@/components/msg/shell";


export function AdminPrompt({ session }: { session: Session }) {
  const [code, setCode] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  return (
    <form
      className="sheet-card gate-card"
      onSubmit={async (e) => {
        e.preventDefault();
        setBusy(true);
        setErr("");
        try {
          const { role } = await login({ data: { code } });
          if (role !== "admin") throw new Error("هذا الرمز ليس رمز صفحة المعاينة");
          saveSession({ code: code.trim(), role });
          window.location.reload();
        } catch (e2) {
          setErr((e2 as Error).message);
        } finally {
          setBusy(false);
        }
      }}
    >
      <span className="login-key"><KeyRound /></span>
      <h1>صفحة المعاينة والتقارير</h1>
      <p className="muted">أدخل رمز ورقة المعاينة للمتابعة</p>
      <input className="input input-lg" inputMode="numeric" value={code} onChange={(e) => setCode(e.target.value)} />
      {err && <p className="err">{err}</p>}
      <button className="btn btn-primary" disabled={busy || !code.trim()}>{busy ? "جارٍ التحقق…" : "فتح"}</button>
      <p className="muted text-sm">أنت مسجل برقم {session.code}</p>
    </form>
  );
}

type Row = { kind: Kind; archive: string | null; r: InRec | OutRec };

export function Reports({ session, onlyKind, archivesOnly = false }: { session: Session; onlyKind?: Kind; archivesOnly?: boolean }) {
  const q = useMsgData(session.code);
  const qc = useQueryClient();
  const [kind, setKind] = useState<"all" | Kind>(onlyKind ?? "all");
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [auth, setAuth] = useState("");
  const [sender, setSender] = useState("");
  const [orientation, setOrientation] = useState("landscape");
  const [search, setSearch] = useState("");
  const [view, setView] = useState<Row | null>(null);
  const [edit, setEdit] = useState<Row | null>(null);

  const rows = useMemo<Row[]>(() => {
    const out: Row[] = [];
    for (const [k, book] of [["in", q.data?.in], ["out", q.data?.out]] as const) {
      if (!book) continue;
      if (!archivesOnly) for (const r of book.active) out.push({ kind: k, archive: null, r });
      if (archivesOnly) for (const a of book.archives) for (const r of a.rows) out.push({ kind: k, archive: a.label, r });
    }
    return out;
  }, [q.data]);

  const filtered = useMemo(
    () =>
      rows.filter(({ kind: k, r }) => {
        if (kind !== "all" && k !== kind) return false;
        if (from && r.date < from) return false;
        if (to && r.date > to) return false;
        if (auth && (k !== "out" || (r as OutRec).authority !== auth)) return false;
        if (sender && (k !== "in" || (r as InRec).sender !== sender)) return false;
        if (search && ![r.subject, r.fileNo, r.date].some((x) => x.includes(search))) return false;
        return true;
      }),
    [rows, kind, from, to, auth, sender, search],
  );

  const stats = useMemo(() => {
    const by = (k: Kind) => rows.filter((x) => x.kind === k);
    const top = (k: Kind, key: "authority" | "sender") => {
      const m = new Map<string, number>();
      for (const { r } of by(k)) {
        const v = (r as OutRec & InRec)[key];
        if (v) m.set(v, (m.get(v) ?? 0) + 1);
      }
      return [...m.entries()].sort((a, b) => b[1] - a[1]).slice(0, 5);
    };
    return {
      inN: by("in").length,
      outN: by("out").length,
      attN: rows.filter((x) => x.r.attachment).length,
      archN: rows.filter((x) => x.archive).length,
      topAuth: top("out", "authority"),
      topSender: top("in", "sender"),
    };
  }, [rows]);

  const exportXlsx = (k: Kind) => {
    const list = filtered.filter((x) => x.kind === k).map(({ r }) => r);
    const ws = XLSX.utils.json_to_sheet(list, { header: ["serial", "fileNo", "date", "subject", ...(k === "in" ? ["sender", "unit", "action", "notes", "status", "reminderDate"] : ["authority", "enclosures"]), "attachment"] });
    ws["!cols"] = Array(10).fill({ wch: 22 });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, k === "in" ? "الوارد" : "الصادر");
    XLSX.writeFile(wb, `nedal.MSG-${k === "in" ? "الوارد" : "الصادر"}.xlsx`);
  };

  const doCloseYear = async () => {
    if (!confirm("سيتم أرشفة كل المراسلات الحالية والبدء بأرقام متسلسلة جديدة للعام القادم. متابعة؟")) return;
    try {
      const archiveCode = prompt("أدخل رقم الأرشفة");
      if (!archiveCode) return;
      const r = await closeYear({ data: { code: session.code, archiveCode } });
      toast.success(`تمت الأرشفة: ${r.archived} — العام الجديد ${r.year}`);
      qc.invalidateQueries({ queryKey: ["msg"] });
    } catch (e) {
      toast.error((e as Error).message);
    }
  };

  if (session.role !== "admin" && !onlyKind) return <AdminPrompt session={session} />;
  if (q.isLoading) return <div className="shell-loading">جارٍ قراءة ملف المراسلات…</div>;
  if (q.error || !q.data?.in || !q.data?.out) return <p className="err">{(q.error as Error)?.message ?? "لا يمكن فتح التقارير بهذا الرمز"}</p>;

  const L = q.data.lists;
  return (
    <div className={`reports ${onlyKind ? "register" : "reports-light"}`}>
      <style>{`@page { size: A4 ${orientation}; margin: 12mm; }`}</style>
      <div className="sheet-card-head no-print">
        <div className="sheet-card-title"><span className="icon-chip"><MailOpen /></span><div><h1>{onlyKind === "in" ? "سجل المراسلات الواردة" : onlyKind === "out" ? "سجل الكتب الصادرة" : archivesOnly ? "أرشيف الأعوام السابقة" : "المعاينة والتقارير"}</h1><p className="muted">نظام n.MSG · العام {L.year}</p></div></div>
        <div className="row-gap">
          <button className="btn btn-ghost" onClick={() => exportXlsx("in")}><FileDown /> تصدير الوارد Excel</button>
          <button className="btn btn-ghost" onClick={() => exportXlsx("out")}><FileDown /> تصدير الصادر Excel</button>
          {session.role === "admin" && !onlyKind && <Link className="btn btn-ghost" to="/archives">أرشيف الأعوام السابقة</Link>}
          {session.role === "admin" && !onlyKind && <button className="btn btn-gold" onClick={doCloseYear}><Archive /> إنهاء العام {L.year} وأرشفته</button>}
        </div>
      </div>

      {!onlyKind && <div className="stats no-print">
        <Stat n={stats.inN} t="مراسلة واردة" />
        <Stat n={stats.outN} t="كتاب صادر" />
        <Stat n={stats.attN} t="مرفق مخزن" />
        <Stat n={stats.archN} t="مراسلة مؤرشفة" />
        <div className="stat stat-list"><b>أكثر جهات الاختصاص</b>{stats.topAuth.map(([n, c]) => <span key={n} className="stat-item">{n} <i>{c}</i></span>)}</div>
        <div className="stat stat-list"><b>أكثر الجهات إرسالاً</b>{stats.topSender.map(([n, c]) => <span key={n} className="stat-item">{n} <i>{c}</i></span>)}</div>
      </div>}

      <div className="filters no-print">
        <select className="input" value={kind} onChange={(e) => setKind(e.target.value as typeof kind)}>
          <option value="all">الوارد والصادر</option><option value="in">الوارد فقط</option><option value="out">الصادر فقط</option>
        </select>
        <input className="input" type="date" value={from} onChange={(e) => setFrom(e.target.value)} aria-label="من تاريخ" />
        <input className="input" type="date" value={to} onChange={(e) => setTo(e.target.value)} aria-label="إلى تاريخ" />
        <select className="input" value={auth} onChange={(e) => setAuth(e.target.value)}>
          <option value="">كل جهات الاختصاص</option>{L.authorities.map((a) => <option key={a}>{a}</option>)}
        </select>
        <select className="input" value={sender} onChange={(e) => setSender(e.target.value)}>
          <option value="">كل الجهات المرسلة</option>{L.authorities.map((a) => <option key={a}>{a}</option>)}
        </select>
        <input className="input" placeholder="بحث بالمبحث أو رقم الملف" value={search} onChange={(e) => setSearch(e.target.value)} />
        <select className="input" aria-label="اتجاه الطباعة" value={orientation} onChange={(e) => setOrientation(e.target.value)}><option value="portrait">بالطول</option><option value="landscape">بالعرض</option></select>
        <button className="btn btn-ghost" onClick={() => setKind("in")}>البحث عن الوارد</button><button className="btn btn-ghost" onClick={() => setKind("out")}>البحث عن الصادر</button><button className="btn btn-ghost" onClick={() => setKind("all")}>البحث عن الصادر والوارد</button>
        <button className="btn btn-primary" onClick={() => window.print()}><Printer /> طباعة التقرير</button>
      </div>

      <PrintReport rows={filtered} title={kind === "all" ? "تقرير الصادر والوارد" : kind === "in" ? "تقرير الوارد" : "تقرير الصادر"} year={L.year} />

      <div className="table-card no-print">
        <table className="tbl">
          <thead>
            <tr><th>#</th><th>النوع</th><th>رقم الملف</th><th>التاريخ</th><th>المبحث</th><th>الجهة</th><th>مرفق</th><th>أرشيف</th><th /></tr>
          </thead>
          <tbody>
            {filtered.map((row) => (
              <tr key={row.kind + (row.archive ?? "") + row.r.serial}>
                <td>{row.r.serial}</td>
                <td><span className={`chip ${row.kind === "in" ? "chip-in" : "chip-out"}`}>{row.kind === "in" ? "وارد" : "صادر"}</span></td>
                <td>{row.r.fileNo}</td>
                <td>{row.r.date}</td>
                <td className="cell-wide">{row.r.subject}</td>
                <td>{row.kind === "in" ? (row.r as InRec).sender : (row.r as OutRec).authority}</td>
                <td>{row.r.attachment && <button className="btn-icon" title="استعراض المرفق" onClick={() => setView(row)}><Paperclip /></button>}</td>
                <td>{row.archive && <span className="chip chip-arch">{row.archive}</span>}</td>
                <td><button className="btn-icon" title="تعديل / إجراء" onClick={() => setEdit(row)}><Pencil /></button></td>
              </tr>
            ))}
            {!filtered.length && <tr><td colSpan={9} className="muted">لا توجد مراسلات مطابقة</td></tr>}
          </tbody>
        </table>
      </div>

      {view && <Viewer code={session.code} row={view} onClose={() => setView(null)} />}
      {edit && <EditModal session={session} row={edit} onClose={(changed) => { setEdit(null); if (changed) qc.invalidateQueries({ queryKey: ["msg"] }); }} />}
    </div>
  );
}

function Stat({ n, t }: { n: number; t: string }) {
  return <div className="stat"><b className="stat-n">{n}</b><span className="muted">{t}</span></div>;
}

/** Centered attachment preview (Drive) with the record's data to the side. */
function Viewer({ code, row, onClose }: { code: string; row: Row; onClose: () => void }) {
  const id = driveFileId(row.r.attachment);
  const fields: [string, string][] =
    row.kind === "in"
      ? Object.entries({ "رقم ملف الوارد": row.r.fileNo, "تاريخ الوارد": row.r.date, المبحث: row.r.subject, "الجهة/القسم - المرسل": (row.r as InRec).sender, "المديرية/الوحدة": (row.r as InRec).unit, اجراء: (row.r as InRec).action, "ملاحظات المسؤول": (row.r as InRec).notes })
      : Object.entries({ "رقم الملف": row.r.fileNo, التاريخ: row.r.date, المبحث: row.r.subject, "جهة الاختصاص": (row.r as OutRec).authority, الملحقات: (row.r as OutRec).enclosures });
  return (
    <div className="modal-back" onClick={onClose}>
      <div className="modal viewer" dir="rtl" onClick={(e) => e.stopPropagation()}>
        <aside className="viewer-side">
          <span className={`chip ${row.kind === "in" ? "chip-in" : "chip-out"}`}>{row.kind === "in" ? "وارد" : "صادر"} #{row.r.serial}</span>
          {fields.map(([k, v]) => v && <div key={k} className="viewer-field"><b>{k}</b><span>{v}</span></div>)}
          <a className="btn btn-ghost" href={row.r.attachment} target="_blank" rel="noreferrer"><Download /> فتح في جوجل درايف</a>
          <button className="btn btn-ghost" onClick={onClose}><X /> إغلاق</button>
        </aside>
        <div className="viewer-body">
          {id ? <PrivatePreview id={id} code={code} /> : <p className="muted">لا يمكن معاينة هذا الرابط</p>}
        </div>
      </div>
    </div>
  );
}

function EditModal({ session, row, onClose }: { session: Session; row: Row; onClose: (changed: boolean) => void }) {
  const isIn = row.kind === "in";
  const [f, setF] = useState<Record<string, string>>(() => Object.fromEntries(Object.entries(row.r).flatMap(([k,v]) => !["serial","year","institution"].includes(k) && typeof v === "string" ? [[k, v]] : [])) );
  const [busy, setBusy] = useState(false);
  const keys = isIn ? ["fileNo", "date", "subject", "sender", "unit", "action", "notes", "status", "reminderDate"] : ["fileNo", "date", "subject", "authority", "enclosures"];
  const labels: Record<string, string> = { fileNo: "رقم الملف", date: "التاريخ", subject: "المبحث", sender: "الجهة/القسم - المرسل", unit: "المديرية/الوحدة", action: "اجراء", notes: "ملاحظات المسؤول", authority: "جهة الاختصاص", enclosures: "الملحقات", status: "حالة المتابعة", reminderDate: "تاريخ التنبيه" };
  return (
    <div className="modal-back" onClick={() => onClose(false)}>
      <div className="modal modal-sm" dir="rtl" onClick={(e) => e.stopPropagation()}>
        <h2>تعديل {isIn ? "الوارد" : "الصادر"} #{row.r.serial}</h2>
        {keys.map((k) => (
          <label key={k} className="field">
            <span className="field-label">{labels[k]}</span>
            {k === "notes" || k === "action" ? (
              <textarea className="input" rows={2} value={f[k] ?? ""} onChange={(e) => setF((p) => ({ ...p, [k]: e.target.value }))} />
            ) : (
              <input className="input" type={k === "date" ? "date" : "text"} value={f[k] ?? ""} onChange={(e) => setF((p) => ({ ...p, [k]: e.target.value }))} />
            )}
          </label>
        ))}
        <div className="row-gap">
          <button
            className="btn btn-primary"
            disabled={busy}
            onClick={async () => {
              setBusy(true);
              try {
                await updateRecord({ data: { code: session.code, kind: row.kind, archive: row.archive, serial: row.r.serial, patch: f } });
                toast.success("تم حفظ التعديل");
                onClose(true);
              } catch (e) {
                toast.error((e as Error).message);
              } finally {
                setBusy(false);
              }
            }}
          >
            {busy ? "جارٍ الحفظ…" : "حفظ"}
          </button>
          <button className="btn btn-ghost" onClick={() => onClose(false)}>إلغاء</button>
        </div>
      </div>
    </div>
  );
}

/** Print-only report table, A4 landscape, signed by/nedal.qazzaz. */
function PrintReport({ rows, title, year }: { rows: Row[]; title: string; year: string }) {
  const cell = (r: ReactNode, i: number) => <td key={i}>{r}</td>;
  return (
    <div className="print-report" dir="rtl">
      <h2>{title}</h2>
      <p>nedal.MSG · العام {year}</p>
      <table>
        <thead>
          <tr><th>#</th><th>النوع</th><th>رقم الملف</th><th>التاريخ</th><th>المبحث</th><th>الجهة</th><th>الملحقات/الإجراء</th></tr>
        </thead>
        <tbody>
          {rows.map(({ kind: k, r }) => (
            <tr key={k + r.serial + (r.date || "")}>
              {[r.serial, k === "in" ? "وارد" : "صادر", r.fileNo, r.date, r.subject, k === "in" ? (r as InRec).sender : (r as OutRec).authority, k === "in" ? (r as InRec).action : (r as OutRec).enclosures].map(cell)}
            </tr>
          ))}
        </tbody>
      </table>
      <Signature />
    </div>
  );
}

function PrivatePreview({ id, code }: { id: string; code: string }) {
  const [src, setSrc] = useState("");
  const [type, setType] = useState("");
  const [error, setError] = useState("");
  useEffect(() => {
    let url = "";
    readAttachment({ data: { id, code } }).then((r) => {
      const bytes = Uint8Array.from(atob(r.base64), (c) => c.charCodeAt(0));
      url = URL.createObjectURL(new Blob([bytes], { type: r.type }));
      setSrc(url); setType(r.type);
    }).catch((e) => setError(e.message));
    return () => { if (url) URL.revokeObjectURL(url); };
  }, [id, code]);
  if (error) return <p>{error}</p>;
  if (!src) return <p>جارٍ تحميل المرفق…</p>;
  if (type.startsWith("image/")) return <img className="attachment-image" src={src} alt="الكتاب المرفق" />;
  if (type.includes("pdf")) return <iframe src={src} title="المرفق" />;
  return <a className="btn btn-primary" href={src} download>تنزيل المرفق لفتحه</a>;
}
