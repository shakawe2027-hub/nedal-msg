import { useId, type ReactNode } from "react";
import { Paperclip, X } from "lucide-react";

export function Field({ label, required, children, hint }: { label: string; required?: boolean; children: ReactNode; hint?: string }) {
  return (
    <label className="field">
      <span className="field-label">{label}{required && <i className="field-req">*</i>}</span>
      {children}
      {hint && <span className="field-hint">{hint}</span>}
    </label>
  );
}

/** Pick from the data-sheet list or type a new value ("قيمة للاضافة"). */
export function Combo({ value, onChange, options, placeholder }: { value: string; onChange: (v: string) => void; options: string[]; placeholder?: string }) {
  const id = useId();
  return (
    <>
      <input className="input" list={id} value={value} onChange={(e) => onChange(e.target.value)} placeholder={placeholder ?? "اختر من القائمة أو اكتب قيمة جديدة"} />
      <datalist id={id}>{options.map((o) => <option key={o} value={o} />)}</datalist>
    </>
  );
}

export function FilePick({ file, onChange }: { file: File | null; onChange: (f: File | null) => void }) {
  return (
    <div className="filepick">
      {file ? (
        <div className="filepick-chosen">
          <Paperclip /> <span>{file.name}</span>
          <button type="button" onClick={() => onChange(null)} aria-label="إزالة"><X /></button>
        </div>
      ) : (
        <label className="filepick-drop">
          <Paperclip />
          <span>إرفاق نسخة (اختياري) — PDF، Word، Excel، صورة أو أي صيغة</span>
          <input type="file" hidden onChange={(e) => onChange(e.target.files?.[0] ?? null)} />
        </label>
      )}
    </div>
  );
}

export function SerialBadge({ label, value }: { label: string; value: number | string }) {
  return (
    <div className="serial">
      <span className="serial-label">{label}</span>
      <span className="serial-value">{value}</span>
      <span className="serial-lock">تلقائي</span>
    </div>
  );
}

export const today = () => new Date().toISOString().slice(0, 10);
