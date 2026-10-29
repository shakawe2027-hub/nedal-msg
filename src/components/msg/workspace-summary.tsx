import { Files, Paperclip, Clock3 } from "lucide-react";
import type { InRec, OutRec } from "@/lib/msg";

export function WorkspaceSummary({ rows, incoming }: { rows: (InRec | OutRec)[]; incoming: boolean }) {
  const items = [
    { icon: Files, value: rows.length, label: incoming ? "إجمالي الوارد" : "إجمالي الصادر" },
    { icon: Paperclip, value: rows.filter((r) => r.attachment).length, label: "كتب مرفقة" },
    { icon: Clock3, value: incoming ? rows.filter((r) => r.status === "يعاد مراجعة").length : rows.at(-1)?.serial ?? "—", label: incoming ? "بانتظار المراجعة" : "آخر رقم صادر" },
  ];
  return <div className="page-summary no-print">{items.map(({ icon: Icon, value, label }) => <div className="summary-item" key={label}><Icon /><div><b>{value}</b><span>{label}</span></div></div>)}</div>;
}