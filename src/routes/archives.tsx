import { createFileRoute } from "@tanstack/react-router";
import { Shell } from "@/components/msg/shell";
import { Reports } from "@/components/msg/register";
export const Route = createFileRoute("/archives")({
 head: () => ({ meta: [{ title: "أرشيف الأعوام السابقة — nedal.MSG" }, { name: "description", content: "المراسلات المؤرشفة للأعوام السابقة" }, { property: "og:title", content: "أرشيف nedal.MSG" }, { property: "og:description", content: "تقارير المراسلات للأعوام السابقة" }, { property: "og:type", content: "website" }, { name: "twitter:card", content: "summary" }] }),
 component: () => <Shell>{s => <Reports session={s} archivesOnly />}</Shell>
});
