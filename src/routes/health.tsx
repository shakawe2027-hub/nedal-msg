import { createFileRoute } from "@tanstack/react-router";
import { Shell } from "@/components/msg/shell";
import { HealthWorkspace } from "@/components/msg/health-workspace";
export const Route = createFileRoute("/health")({
  head: () => ({ meta: [
    { title: "متابعات الصحة المدرسية — nedal.MSG" },
    { name: "description", content: "نماذج تحويل الإصابات المدرسية وأرشيف التقارير الصحية للطلاب." },
    { property: "og:title", content: "متابعات الصحة المدرسية — nedal.MSG" },
    { property: "og:description", content: "تحويل الطلاب ومتابعة نماذج الإصابات المدرسية المؤرشفة." },
    { property: "og:type", content: "website" }, { name: "twitter:card", content: "summary" },
  ] }),
  component: () => <Shell>{session => <HealthWorkspace session={session} />}</Shell>,
});