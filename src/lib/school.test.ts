import { describe, expect, it } from "vitest";
import { monthDays, parseMonthLabel, daysInMonth } from "./school";

describe("absence calendar", () => {
  it("February 2027 has 28 days, rest are outside", () => {
    const d = monthDays(2027, 2);
    expect(d.filter((x) => !x.outside)).toHaveLength(28);
    expect(d[28].outside).toBe(true);
  });
  it("leap February 2028 has 29 days", () => {
    expect(daysInMonth(2028, 2)).toBe(29);
  });
  it("Friday 23 Oct 2026 and Saturday 24 are shaded, Sunday 25 is not", () => {
    const d = monthDays(2026, 10);
    expect(d[22]).toMatchObject({ weekday: "الجمعة", weekend: true });
    expect(d[23]).toMatchObject({ weekday: "السبت", weekend: true });
    expect(d[24].weekend).toBe(false);
  });
  it("parses month labels from the sheet", () => {
    expect(parseMonthLabel("10. تشرين الأول / أكتوبر (October): 31 يوماً")).toEqual({ num: 10, name: "تشرين الأول / أكتوبر" });
  });
});
