import { useEffect, useState, type ReactElement } from "react";

export type ExportJob = { key: string; landscape: boolean; node: ReactElement };

/**
 * Renders every job as its own printed A4 page, then opens the browser's print dialog
 * (choose "Save as PDF"). Uses the same print layout as single forms, so output matches exactly.
 */
export function useBulkExport() {
  const [run, setRun] = useState<{ jobs: ExportJob[]; file: string } | null>(null);

  useEffect(() => {
    if (!run) return;
    const prevTitle = document.title;
    document.title = run.file.replace(/\.pdf$/, "");
    document.body.classList.add("bulk-printing");
    const done = () => {
      document.body.classList.remove("bulk-printing");
      document.title = prevTitle;
      setRun(null);
    };
    const t = setTimeout(async () => {
      await document.fonts.ready;
      window.addEventListener("afterprint", done, { once: true });
      window.print();
    }, 300);
    return () => clearTimeout(t);
  }, [run]);

  const start = (jobs: ExportJob[], file: string) => {
    if (jobs.length && !run) setRun({ jobs, file });
  };

  const stage = run ? (
    <div className="bulk-stage">
      <style>{`@page { size: A4 ${run.jobs[0]?.landscape ? "landscape" : "portrait"}; margin: 8mm; }`}</style>
      {run.jobs.map((j) => (
        <div key={j.key} className={`bulk-page ${j.landscape ? "bulk-landscape" : "bulk-portrait"}`}>{j.node}</div>
      ))}
    </div>
  ) : null;

  return { start, busy: !!run, progress: run ? String(run.jobs.length) : "", stage };
}
