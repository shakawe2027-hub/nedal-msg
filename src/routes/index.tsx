import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useState } from "react";
import { KeyRound } from "lucide-react";
import { login } from "@/lib/msg.functions";
import { saveSession } from "@/lib/session";
import correspondence from "@/assets/correspondence.jpg";
import { ClockBar } from "@/components/msg/clock";
import { Brand, Signature } from "@/components/msg/shell";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "nedal.MSG — تسجيل الدخول" },
      { name: "description", content: "نظام nedal.MSG لإدارة المراسلات الواردة والصادرة وأرشفتها." },
      { property: "og:title", content: "nedal.MSG — نظام المراسلات" },
      { property: "og:description", content: "تسجيل الوارد والصادر، المرفقات، التقارير والأرشيف." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary" },
    ],
  }),
  component: LoginPage,
});

function LoginPage() {
  const nav = useNavigate();
  const [code, setCode] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    setBusy(true);
    setErr("");
    try {
      const { role } = await login({ data: { code } });
      saveSession({ code: code.trim(), role });
      nav({ to: role === "admin" ? "/reports" : "/in" });
    } catch (e) {
      setErr((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="login" dir="rtl">
      <aside className="login-art">
        <div className="login-art-inner">
          <Brand big />
          <p className="login-tag">نظام إدارة المراسلات الرسمية</p>
          <p className="version-label">النسخة الحالية ( v.010.2026)</p>
          <p>إعداد وتصميم نضال قزاز</p>
        </div>
        <img className="login-photo" src={correspondence} alt="مراسلات إدارية ومستندات رسمية" width={1536} height={1024} />
      </aside>
      <section className="login-panel">
        <ClockBar />
        <form className="login-card" onSubmit={submit}>
          <span className="login-key"><KeyRound /></span>
          <h1>نظام nedal.MSG</h1>
          <p className="muted">أدخل الرقم الوطني، أو رمز صفحة المعاينة</p>
          <input className="input input-lg" inputMode="numeric" autoFocus value={code} onChange={(e) => setCode(e.target.value)} placeholder="الرقم الوطني" />
          {err && <p className="err">{err}</p>}
          <button className="btn btn-primary btn-lg" disabled={busy || !code.trim()}>{busy ? "جارٍ التحقق…" : "دخول"}</button>
          <Signature />
        </form>
      </section>
    </div>
  );
}
