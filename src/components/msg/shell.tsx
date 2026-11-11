import { Link, useNavigate, useRouterState } from "@tanstack/react-router";
import { useEffect, type ReactNode } from "react";
import { Inbox, Send, BarChart3, LogOut, Mail, FileHeart } from "lucide-react";
import { Button } from "@/components/ui/button";
import { ClockBar } from "./clock";
import { saveSession, useSession, type Session } from "@/lib/session";

export function Brand({ big }: { big?: boolean }) {
  return (
    <div className={`brand ${big ? "brand-big" : ""}`}>
      <span className="brand-seal"><Mail /></span>
      <span className="brand-name">nedal<span>.MSG</span></span>
    </div>
  );
}

export function Signature() {
  return <div className="signature">by/nedal.qazzaz</div>;
}

/** Signed-in layout. Redirects to the login page when there is no session. */
export function Shell({ children }: { children: (s: Session) => ReactNode }) {
  const session = useSession();
  const nav = useNavigate();
  const path = useRouterState({ select: state => state.location.pathname });
  useEffect(() => {
    if (session === null) nav({ to: "/" });
  }, [session, nav]);
  if (!session) return <div className="shell-loading">…</div>;
  return (
    <div className="shell" dir="rtl">
      <header className="topbar no-print">
        <Brand />
        <nav className="topnav">
          <Link to="/in" className="topnav-link" activeProps={{ "data-active": "true" }}><Inbox /> الوارد</Link>
          <Link to="/out" className="topnav-link" activeProps={{ "data-active": "true" }}><Send /> الصادر</Link>
          <Link to="/reports" className="topnav-link" activeProps={{ "data-active": "true" }}><BarChart3 /> المعاينة والتقارير</Link>
        </nav>
        <div className="row-gap">
          {(path === "/out" || path === "/health") && <Button asChild variant="ghost" className="topnav-link"><Link to="/health"><FileHeart /> الإصابات المدرسية</Link></Button>}
          <button className="topnav-link" onClick={() => { saveSession(null); nav({ to: "/" }); }}><LogOut /> خروج</button>
        </div>
      </header>
      <ClockBar />
      <main className="shell-main">{children(session)}</main>
      <footer className="shell-foot no-print"><Signature /></footer>
    </div>
  );
}
