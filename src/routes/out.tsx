import { createFileRoute } from "@tanstack/react-router";
import { useState } from "react";
import { useQueryClient } from "@tanstack/react-query";
import { toast } from "sonner";
import { Send } from "lucide-react";
import { addOut, uploadAttachment } from "@/lib/msg.functions";
import { useMsgData, type Session } from "@/lib/session";
import { WorkspaceSummary } from "@/components/msg/workspace-summary";
import { Reports } from "@/components/msg/register";
import { Shell } from "@/components/msg/shell";
import { Combo, Field, FilePick, SerialBadge, today } from "@/components/msg/fields";

export const Route = createFileRoute("/out")({
  head: () => ({
    meta: [
      { title: "الصادر — nedal.MSG.OUT" },
      { name: "description", content: "تسجيل الكتب الصادرة وملحقاتها في نظام nedal.MSG." },
      { property: "og:title", content: "الصادر — nedal.MSG.OUT" },
      { property: "og:description", content: "إصدار الكتب بترقيم تلقائي ومرفقات." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary" },
    ],
  }),
  component: () => <Shell>{(s) => <OutForm session={s} />}</Shell>,
});

const empty = { fileNo: "", date: "", subject: "", authority: "", enclosures: "" };

function OutForm({ session }: { session: Session }) {
  const q = useMsgData(session.code);
  const qc = useQueryClient();
  const [f, setF] = useState(() => ({ ...empty, date: today() }));
  const [file, setFile] = useState<File | null>(null);
  const [busy, setBusy] = useState(false);
  const set = (k: keyof typeof empty) => (v: string) => setF((p) => ({ ...p, [k]: v }));

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    setBusy(true);
    try {
      let attachment = "";
      if (file) {
        const fd = new FormData();
        fd.set("code", session.code);
        fd.set("subject", f.subject);
        fd.set("file", file);
        fd.set("kind", "out");
        attachment = (await uploadAttachment({ data: fd })).link;
      }
      const { serial } = await addOut({ data: { code: session.code, rec: { ...f, attachment } } });
      toast.success(`تم حفظ الصادر رقم ${serial}`);
      setF({ ...empty, date: today() });
      setFile(null);
      qc.invalidateQueries({ queryKey: ["msg"] });
    } catch (e) {
      toast.error((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const L = q.data?.lists;
  const last = q.data?.out?.active.at(-1);
  return (
    <div className="entry-stack entry-out">
    {last && <div className="last-out no-print">آخر صادر: #{last.serial} · {last.date} · {last.subject}</div>}
    <WorkspaceSummary rows={q.data?.out?.active ?? []} incoming={false} />
    <form className="sheet-card no-print" onSubmit={submit}>
      <div className="sheet-card-head">
        <div className="sheet-card-title"><span className="icon-chip icon-chip-gold"><Send /></span><div><h1>nedal.MSG.OUT</h1><p className="muted">الكتب الصادرة {L?.year && `· العام ${L.year}`}</p></div></div>
        <SerialBadge label="الرقم الصادر" value={q.data?.nextOut ?? "…"} />
      </div>
      {q.error && <p className="err">{(q.error as Error).message}</p>}
      <div className="grid-form">
        <Field label="رقم الملف" required><input className="input" required value={f.fileNo} onChange={(e) => set("fileNo")(e.target.value)} /></Field>
        <Field label="التاريخ" required hint="تاريخ اليوم — قابل للتعديل"><input className="input" type="date" required value={f.date} onChange={(e) => set("date")(e.target.value)} /></Field>
        <Field label="المبحث" required><input className="input" required value={f.subject} onChange={(e) => set("subject")(e.target.value)} /></Field>
        <Field label="جهة الاختصاص"><Combo value={f.authority} onChange={set("authority")} options={L?.authorities ?? []} /></Field>
        <Field label="الملحقات" hint="اختياري — يمكن تركها فارغة"><Combo value={f.enclosures} onChange={set("enclosures")} options={L?.enclosures ?? []} /></Field>
        <div className="grid-span"><FilePick file={file} onChange={setFile} /></div>
      </div>
      <div className="sheet-card-actions">
        <button className="btn btn-gold" disabled={busy || !q.data}>{busy ? "جارٍ الحفظ…" : "حفظ الصادر"}</button>
      </div>
    </form>
    <Reports session={session} onlyKind="out" />
    </div>
  );
}
