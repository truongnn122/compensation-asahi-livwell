import {
  FYC_PER_POLICY,
  calcCommission,
  calcRYCRate,
  estimateFYC,
} from "../commission";

describe("calcRYCRate — bảng Mục A.I.2", () => {
  it("≥2 tháng hoạt động và K2 ≥80% → 120%", () => {
    expect(
      calcRYCRate({ activeMonthsInQuarter: 2, k2Percent: 80 }).ratePercent
    ).toBe(120);
    expect(
      calcRYCRate({ activeMonthsInQuarter: 3, k2Percent: 95 }).ratePercent
    ).toBe(120);
  });

  it("K2 ≥80% nhưng chỉ 1 tháng hoạt động thì rơi về 100%", () => {
    expect(
      calcRYCRate({ activeMonthsInQuarter: 1, k2Percent: 85 }).ratePercent
    ).toBe(100);
  });

  it("≥1 tháng hoạt động và K2 ≥60% → 100%", () => {
    expect(
      calcRYCRate({ activeMonthsInQuarter: 1, k2Percent: 60 }).ratePercent
    ).toBe(100);
    expect(
      calcRYCRate({ activeMonthsInQuarter: 3, k2Percent: 75 }).ratePercent
    ).toBe(100);
  });

  it("≥1 tháng hoạt động và K2 <60% → 80%", () => {
    expect(
      calcRYCRate({ activeMonthsInQuarter: 1, k2Percent: 59.9 }).ratePercent
    ).toBe(80);
    expect(
      calcRYCRate({ activeMonthsInQuarter: 3, k2Percent: 0 }).ratePercent
    ).toBe(80);
  });

  it("không có tháng hoạt động nào thì không được chi trả", () => {
    const result = calcRYCRate({ activeMonthsInQuarter: 0, k2Percent: 90 });
    expect(result.eligible).toBe(false);
    expect(result.ratePercent).toBe(0);
  });

  it("chưa có K2 (N/A) xếp vào mức 100%", () => {
    expect(
      calcRYCRate({ activeMonthsInQuarter: 1, k2Percent: null }).ratePercent
    ).toBe(100);
  });
});

describe("calcCommission", () => {
  it("RYC thực nhận = RYC gốc × tỷ lệ chi trả", () => {
    const result = calcCommission({
      fyc: 28,
      rycBase: 5,
      activeMonthsInQuarter: 2,
      k2Percent: 85,
    });
    expect(result.rycRate.ratePercent).toBe(120);
    expect(result.ryc).toBe(6);
    expect(result.total).toBe(34);
  });

  it("không đủ tháng hoạt động thì chỉ còn FYC", () => {
    const result = calcCommission({
      fyc: 28,
      rycBase: 5,
      activeMonthsInQuarter: 0,
      k2Percent: 85,
    });
    expect(result.ryc).toBe(0);
    expect(result.total).toBe(28);
  });
});

describe("estimateFYC", () => {
  it(`ước tính ${FYC_PER_POLICY} triệu cho mỗi hợp đồng`, () => {
    expect(estimateFYC(7)).toBe(28);
    expect(estimateFYC(0)).toBe(0);
  });
});
