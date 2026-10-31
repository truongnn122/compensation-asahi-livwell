import { calcMonthlyBonus, resolveBonusColumn } from "../monthlyBonus";

const base = {
  fyc: 26,
  k2Percent: null as number | null,
  oneUpPercent: 0,
  actualWorkingDays: 22,
};

describe("resolveBonusColumn", () => {
  it("có K2 thì xét theo K2, 1Up có thể kéo lên bậc cao hơn", () => {
    expect(resolveBonusColumn(85, 0)).toBe(1);
    expect(resolveBonusColumn(75, 0)).toBe(2);
    expect(resolveBonusColumn(65, 0)).toBe(3);
    expect(resolveBonusColumn(55, 0)).toBe(4);
    expect(resolveBonusColumn(55, 25)).toBe(1);
    expect(resolveBonusColumn(45, 0)).toBe(0);
  });

  it("chưa có K2 thì chỉ 1Up mới nâng bậc", () => {
    expect(resolveBonusColumn(null, 30)).toBe(1);
    expect(resolveBonusColumn(null, 15)).toBe(2);
    expect(resolveBonusColumn(null, 10)).toBe("NA");
  });
});

describe("calcMonthlyBonus", () => {
  it("tra bảng theo bậc FYC và cột chất lượng", () => {
    const result = calcMonthlyBonus({ ...base, k2Percent: 85 });
    expect(result.column).toBe(1);
    expect(result.ratePercent).toBe(50);
    expect(result.amount).toBeCloseTo(13, 6);
  });

  it("chưa có K2 thì dùng cột riêng", () => {
    const result = calcMonthlyBonus({ ...base, fyc: 12 });
    expect(result.column).toBe("NA");
    expect(result.ratePercent).toBe(15);
    expect(result.amount).toBeCloseTo(1.8, 6);
  });

  it("K2 dưới 50% thì không được thưởng", () => {
    const result = calcMonthlyBonus({ ...base, k2Percent: 45 });
    expect(result.eligible).toBe(false);
    expect(result.amount).toBe(0);
  });

  it("chấm công dưới 4 ngày thì không được thưởng", () => {
    const result = calcMonthlyBonus({
      ...base,
      k2Percent: 85,
      actualWorkingDays: 3,
    });
    expect(result.eligible).toBe(false);
    expect(result.reason).toContain("Chấm công");
  });

  it("FYC dưới 5 triệu thì chưa vào bảng", () => {
    const result = calcMonthlyBonus({ ...base, fyc: 4.9, k2Percent: 85 });
    expect(result.eligible).toBe(false);
    expect(result.amount).toBe(0);
  });
});
