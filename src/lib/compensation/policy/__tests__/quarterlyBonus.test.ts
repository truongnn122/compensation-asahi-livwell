import { calcQuarterlyBonus } from "../quarterlyBonus";

const base = {
  quarterFYC: 60,
  k2Percent: null as number | null,
  oneUpPercent: 0,
  actualWorkingDays: 22,
};

describe("calcQuarterlyBonus", () => {
  it("tra bảng quý theo bậc FYC và cột chất lượng", () => {
    const result = calcQuarterlyBonus({ ...base, k2Percent: 85 });
    expect(result.column).toBe(1);
    expect(result.ratePercent).toBe(35);
    expect(result.amount).toBeCloseTo(21, 6);
  });

  it("ngưỡng cao nhất từ 90 triệu", () => {
    const result = calcQuarterlyBonus({
      ...base,
      quarterFYC: 90,
      k2Percent: 85,
    });
    expect(result.ratePercent).toBe(50);
    expect(result.amount).toBeCloseTo(45, 6);
  });

  it("chưa có K2 thì dùng cột riêng", () => {
    const result = calcQuarterlyBonus({ ...base, quarterFYC: 40 });
    expect(result.column).toBe("NA");
    expect(result.ratePercent).toBe(10);
    expect(result.amount).toBeCloseTo(4, 6);
  });

  it("FYC quý dưới 30 triệu thì chưa vào bảng", () => {
    const result = calcQuarterlyBonus({
      ...base,
      quarterFYC: 29,
      k2Percent: 85,
    });
    expect(result.eligible).toBe(false);
    expect(result.amount).toBe(0);
  });

  it("K2 dưới 50% thì không được thưởng", () => {
    const result = calcQuarterlyBonus({ ...base, k2Percent: 45 });
    expect(result.eligible).toBe(false);
  });

  it("chấm công dưới 4 ngày thì không được thưởng", () => {
    const result = calcQuarterlyBonus({
      ...base,
      k2Percent: 85,
      actualWorkingDays: 3,
    });
    expect(result.eligible).toBe(false);
    expect(result.reason).toContain("Chấm công");
  });
});
