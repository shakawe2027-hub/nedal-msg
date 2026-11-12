import { describe, expect, it } from "vitest";
import { ageAt, referralFields } from "./health";
describe("health referral rules", () => {
  it("calculates completed years from birth date, not the issue date", () => {
    expect(ageAt("2010-11-13", "2026-11-12")).toBe(15);
    expect(ageAt("2010-11-12", "2026-11-12")).toBe(16);
  });
  it("requires the manually supplied student address", () => {
    expect(referralFields.safeParse({ studentId: "123456789", destination: "مركز صحي", address: "دورا", studentAddress: "", reasons: "إصابة", year: "2026/2027", date: "2026-11-12", headerDate: "2026-11-12" }).success).toBe(false);
  });
});