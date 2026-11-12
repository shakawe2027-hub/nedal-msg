import { createFileRoute, Link } from "@tanstack/react-router";
import { useState } from "react";
import { Search, Inbox, Send, BarChart3, FileHeart, CalendarDays, ArrowLeft } from "lucide-react";
import { Shell } from "@/components/msg/shell";
import { SchoolsModal } from "@/components/msg/schools";
import { Button } from "@/components/ui/button";
export const Route = createFileRoute("/welcome")({
  head: () => ({ meta: [{ title: "مرحباً بك — nedal.MSG و nedal.Health" }, { name: "description", content: "المراسلات ودليل المدارس والتقارير ومتابعات الصحة المدرسية." }, { property: "og:title", content: "مرحباً بك — nedal.MSG و nedal.Health" }, { property: "og:description", content: "مساحة المدرسة للمراسلات ومتابعة الإصابات المدرسية." }, { property: "og:type", content: "website" }, { name: "twitter:card", content: "summary" }] }),
  component: Welcome,
});
function Welcome() {
  const [directory, setDirectory] = useState(false);
  return <Shell>{session => <div className="welcome-workspace">
    <header className="welcome-heading"><p className="health-brand" dir="ltr">nedal.MSG · nedal.Health</p><h1>مرحباً بك</h1></header>
    <div className="welcome-actions">
      <Button variant="outline" onClick={() => setDirectory(true)}><Search /><span>بحث دليل المدارس</span><ArrowLeft /></Button>
      <Button asChild variant="outline"><Link to="/in"><Inbox /><span>المراسلات الواردة</span><ArrowLeft /></Link></Button>
      <Button asChild variant="outline"><Link to="/out"><Send /><span>المراسلات الصادرة</span><ArrowLeft /></Link></Button>
      <Button asChild variant="outline"><Link to="/reports"><BarChart3 /><span>معاينة التقارير</span><ArrowLeft /></Link></Button>
      <Button asChild variant="outline"><Link to="/health"><FileHeart /><span>تقرير تحويل الحالات الصحية</span><ArrowLeft /></Link></Button>
      <Button asChild variant="outline"><Link to="/health-monthly"><CalendarDays /><span>رصد الإصابات الشهرية</span><ArrowLeft /></Link></Button>
    </div>
    {directory && <SchoolsModal code={session.code} onClose={() => setDirectory(false)} />}
  </div>}</Shell>;
}