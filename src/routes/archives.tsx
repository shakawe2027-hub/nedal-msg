import { createFileRoute } from "@tanstack/react-router";
import { Shell } from "@/components/msg/shell";
import { Reports } from "@/components/msg/register";
export const Route = createFileRoute("/archives")({
 head: () => ({ meta: [{ title: "أرشيف الأعوام السابقة — nedal.MSG" }, { name: "description", content: "المراسلات المؤرشفة للأعوام السابقة" }, { property: "og:title", content: "أرشيف nedal.MSG" }, { property: "og:description", content: "تقارير المراسلات للأعوام السابقة" }, { property: "og:type", content: "website" }, { name: "twitter:card", content: "summary" }] }),
 validateSearch: (s: Record<string, unknown>): { year?: string | undefined } => ({ year: typeof s["year"] === "string" && s["year"] ? s["year"] : undefined }),
 component: Archives,
});
function Archives() {
 const { year } = Route.useSearch();
 return <Shell>{s => <Reports session={s} archivesOnly archiveYear={year} />}</Shell>;
}
