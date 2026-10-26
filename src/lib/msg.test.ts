import { describe, expect, it } from "vitest";
import { archiveBook, buildRows, bumpYear, nextSerial, parseBook, IN_HEADER } from "./msg";

const row = (n: number) => [n, `F${n}`, "2026-10-01", `موضوع ${n}`, "قسم المالية", "الوزارة", "", "", ""];

describe("correspondence book", () => {
  it("serial continues from the highest active number", () => {
    const b = parseBook("in", [IN_HEADER, row(1), row(2), row(3)]);
    expect(nextSerial(b.active)).toBe(4);
  });
  it("closing the year archives old mail and restarts serials at 1", () => {
    const b = archiveBook(parseBook("in", [IN_HEADER, row(1), row(2)]), "2027/2026");
    expect(nextSerial(b.active)).toBe(1);
    const rows = buildRows("in", b);
    expect(rows).toContainEqual(["مراسلات منتهية-أرشيف 2027/2026"]);
    const again = parseBook("in", rows);
    expect(again.active).toHaveLength(0);
    expect(again.archives[0]!.label).toBe("2027/2026");
    expect(again.archives[0]!.rows).toHaveLength(2);
  });
  it("school year moves forward by one", () => {
    expect(bumpYear("2027/2026")).toBe("2028/2027");
  });
});
