import {
  calcIncomeSupport,
  calcQualityBonus,
  calcSignOnBonus,
} from "../mdrtBuyout";

describe("calcSignOnBonus", () => {
  it("tra đúng mức theo hồ sơ thành tích", () => {
    expect(
      calcSignOnBonus({ record: "fyp700", previousPersistencyPercent: 80 })
        .amount
    ).toBe(100);
    expect(
      calcSignOnBonus({ record: "mdrt", previousPersistencyPercent: 80 }).amount
    ).toBe(150);
    expect(
      calcSignOnBonus({ record: "mdrt2", previousPersistencyPercent: 80 })
        .amount
    ).toBe(200);
    expect(
      calcSignOnBonus({ record: "cot", previousPersistencyPercent: 80 }).amount
    ).toBe(300);
  });

  it("tỷ lệ duy trì ở công ty cũ dưới 60% thì không được", () => {
    const result = calcSignOnBonus({
      record: "cot",
      previousPersistencyPercent: 59,
    });
    expect(result.amount).toBe(0);
    expect(result.reason).toContain("60%");
  });
});

describe("calcIncomeSupport", () => {
  const base = {
    monthSinceJoining: 3,
    fyp: 80,
    policies: 2,
    attendancePercent: 90,
  };

  it("FYP ≥ 60 triệu và 2 hợp đồng được 15%", () => {
    expect(calcIncomeSupport(base).amount).toBeCloseTo(12, 6);
  });

  it("FYP ≥ 100 triệu nhưng chỉ 2 hợp đồng thì vẫn ở mức 15%", () => {
    expect(calcIncomeSupport({ ...base, fyp: 120 }).ratePercent).toBe(15);
  });

  it("FYP ≥ 100 triệu và 3 hợp đồng được 20%, cắt ở trần 40 triệu", () => {
    const result = calcIncomeSupport({ ...base, fyp: 250, policies: 3 });
    expect(result.ratePercent).toBe(20);
    expect(result.uncappedAmount).toBeCloseTo(50, 6);
    expect(result.amount).toBe(40);
    expect(result.capped).toBe(true);
  });

  it("ngoài 12 tháng đầu thì dừng", () => {
    expect(calcIncomeSupport({ ...base, monthSinceJoining: 13 }).amount).toBe(
      0
    );
  });

  it("chấm công dưới 70% thì không được", () => {
    expect(
      calcIncomeSupport({ ...base, attendancePercent: 60 }).reason
    ).toContain("70%");
  });
});

describe("calcQualityBonus", () => {
  it("nhân thưởng MDRT năm 1 với hệ số K2", () => {
    expect(
      calcQualityBonus({ yearOneFYP: 2500, k2Percent: 85 }).amount
    ).toBeCloseTo(550, 6);
    expect(
      calcQualityBonus({ yearOneFYP: 1300, k2Percent: 65 }).amount
    ).toBeCloseTo(125, 6);
  });

  it("K2 dưới 60% thì mất hẳn", () => {
    expect(calcQualityBonus({ yearOneFYP: 4000, k2Percent: 59 }).amount).toBe(
      0
    );
  });
});
