import { createFileRoute } from "@tanstack/react-router";
import { useState } from "react";
import { useQueryClient } from "@tanstack/react-query";
import { toast } from "sonner";
import { Inbox } from "lucide-react";
import { addIn, uploadAttachment } from "@/lib/msg.functions";
import { useMsgData, type Session } from "@/lib/session";
import { Reports } from "@/components/msg/register";
import { Shell } from "@/components/msg/shell";
import { Combo, Field, FilePick, SerialBadge, today } from "@/components/msg/fields";

export const Route = createFileRoute("/in")({
  head: () => ({
    meta: [
      { title: "الوارد — nedal.MSG" },
      { name: "description", content: "تسجيل المراسلات الواردة في نظام nedal.MSG." },
      { property: "og:title", content: "الوارد — nedal.MSG" },
      { property: "og:description", content: "تسجيل المراسلات الواردة مع المرفقات." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary" },
    ],
  }),
  component: () => <Shell>{(s) => <InForm session={s} />}</Shell>,
});

const empty = { fileNo: "", date: "", subject: "", sender: "", unit: "", action: "", notes: "", status: "", reminderDate: "" };

function InForm({ session }: { session: Session }) {
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
        fd.set("kind", "in");
        attachment = (await uploadAttachment({ data: fd })).link;
      }
      const { serial } = await addIn({ data: { code: session.code, rec: { ...f, attachment } } });
      toast.success(`تم حفظ الوارد رقم ${serial}`);
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
  return (
    <div className="entry-stack">
    <form className="sheet-card no-print" onSubmit={submit}>
      <div className="sheet-card-head">
        <div className="sheet-card-title"><span className="icon-chip"><Inbox /></span><div><h1>nedal.MSG.IN</h1><p className="muted">المراسلات الواردة {L?.year && `· العام ${L.year}`}</p></div></div>
        <SerialBadge label="عداد متسلسل" value={q.data?.nextIn ?? "…"} />
      </div>
      {q.error && <p className="err">{(q.error as Error).message}</p>}
      <div className="grid-form">
        <Field label="رقم ملف الوارد" required><input className="input" required value={f.fileNo} onChange={(e) => set("fileNo")(e.target.value)} /></Field>
        <Field label="تاريخ الوارد" required><input className="input" type="date" required value={f.date} onChange={(e) => set("date")(e.target.value)} /></Field>
        <Field label="المبحث" required><input className="input" required value={f.subject} onChange={(e) => set("subject")(e.target.value)} /></Field>
        <Field label="الجهة/القسم - المرسل"><Combo value={f.sender} onChange={set("sender")} options={L?.authorities ?? []} /></Field>
        <Field label="المديرية/الوحدة"><Combo value={f.unit} onChange={set("unit")} options={L?.units ?? []} /></Field>
        <Field label="اجراء" hint="اختياري"><Combo value={f.action} onChange={set("action")} options={L?.actions ?? []} placeholder="اكتب الإجراء" /></Field>
        <Field label="ملاحظات المسؤول" hint="اختياري"><textarea className="input" rows={3} value={f.notes} onChange={(e) => set("notes")(e.target.value)} /></Field>
        <Field label="حالة المتابعة" hint="اختياري"><select className="input" value={f.status} onChange={(e) => set("status")(e.target.value)}><option value="">بدون تنبيه</option><option>تم الرد عليه</option><option>يعاد مراجعة</option></select></Field>
        <Field label="تاريخ التنبيه" hint="اختياري"><input className="input" type="date" value={f.reminderDate} onChange={(e) => set("reminderDate")(e.target.value)} /></Field>
        <div className="grid-span"><FilePick file={file} onChange={setFile} /></div>
      </div>
      <div className="sheet-card-actions">
        <button className="btn btn-primary" disabled={busy || !q.data}>{busy ? "جارٍ الحفظ…" : "حفظ الوارد"}</button>
      </div>
    </form>
    <Reports session={session} onlyKind="in" />
    </div>
  );
}
