import { checkHealthInsurance } from "../healthInsurance";

describe("checkHealthInsurance", () => {
  it("đủ 10 Tháng hoạt động trở lên thì được hưởng", () => {
    expect(checkHealthInsurance({ activeMonthsInLast12: 10 }).eligible).toBe(
      true
    );
    expect(checkHealthInsurance({ activeMonthsInLast12: 12 }).eligible).toBe(
      true
    );
  });

  it("chưa đủ thì báo còn thiếu bao nhiêu tháng", () => {
    const result = checkHealthInsurance({ activeMonthsInLast12: 7 });
    expect(result.eligible).toBe(false);
    expect(result.monthsShort).toBe(3);
    expect(result.reason).toContain("3");
  });
});
