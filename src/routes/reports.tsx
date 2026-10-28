import { createFileRoute } from "@tanstack/react-router";
import { Shell } from "@/components/msg/shell";
import { Reports } from "@/components/msg/register";
export const Route = createFileRoute("/reports")({
  head: () => ({
    meta: [
      { title: "المعاينة والتقارير — nedal.MSG" },
      { name: "description", content: "تقارير الصادر والوارد، الإحصائيات، المرفقات والأرشيف." },
      { property: "og:title", content: "المعاينة والتقارير — nedal.MSG" },
      { property: "og:description", content: "تقارير المراسلات حسب التاريخ والجهة مع الأرشيف السنوي." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary" },
    ],
  }),
  component: () => <Shell>{(s) => <Reports session={s} />}</Shell>,
});
