import {
  DEPARTMENT_MONTHLY_BONUS_CAP,
  calcDepartmentMonthlyBonus,
  calcDepartmentQuarterlyBonus,
  departmentGrowthPercent,
} from "../departmentBonus";

describe("calcDepartmentMonthlyBonus", () => {
  const base = {
    departmentFYC: 100,
    activityPercent: 60,
    k2Percent: 85,
    oneUpPercent: 0,
  };

  it("tra đúng ô theo tỷ lệ hoạt động và cột chất lượng", () => {
    expect(calcDepartmentMonthlyBonus(base).ratePercent).toBe(12);
    expect(
      calcDepartmentMonthlyBonus({ ...base, activityPercent: 45 }).ratePercent
    ).toBe(11);
    expect(
      calcDepartmentMonthlyBonus({ ...base, activityPercent: 30 }).ratePercent
    ).toBe(10);
  });

  it("tỷ lệ hoạt động dưới 30% thì chưa có thưởng", () => {
    const result = calcDepartmentMonthlyBonus({ ...base, activityPercent: 29 });
    expect(result.eligible).toBe(false);
    expect(result.amount).toBe(0);
  });

  it("cắt ở trần 25 triệu và đánh dấu đã cắt", () => {
    const result = calcDepartmentMonthlyBonus({ ...base, departmentFYC: 500 });
    expect(result.uncappedAmount).toBe(60);
    expect(result.amount).toBe(DEPARTMENT_MONTHLY_BONUS_CAP);
    expect(result.capped).toBe(true);
  });

  it("dưới trần thì không đánh dấu cắt", () => {
    const result = calcDepartmentMonthlyBonus({ ...base, departmentFYC: 100 });
    expect(result.amount).toBe(12);
    expect(result.capped).toBe(false);
  });

  it("K2 ban dưới 40% thì không được nhận", () => {
    const result = calcDepartmentMonthlyBonus({ ...base, k2Percent: 39 });
    expect(result.eligible).toBe(false);
    expect(result.reason).toContain("K2");
  });

  it("chưa có K2 và 1Up thấp thì không có cột nào áp dụng", () => {
    const result = calcDepartmentMonthlyBonus({
      ...base,
      k2Percent: null,
      oneUpPercent: 0,
    });
    expect(result.column).toBe("NA");
    expect(result.eligible).toBe(false);
  });

  it("chưa có K2 nhưng 1Up đủ cao thì vẫn vào bảng", () => {
    const result = calcDepartmentMonthlyBonus({
      ...base,
      k2Percent: null,
      oneUpPercent: 25,
    });
    expect(result.column).toBe(1);
    expect(result.ratePercent).toBe(12);
  });
});

describe("departmentGrowthPercent", () => {
  it("mẫu số dưới 175 triệu được ấn định ở 175", () => {
    expect(departmentGrowthPercent(350, 100)).toBe(100);
    expect(departmentGrowthPercent(350, 0)).toBe(100);
  });

  it("mẫu số từ 175 trở lên thì dùng số thật", () => {
    expect(departmentGrowthPercent(400, 200)).toBe(100);
  });
});

describe("calcDepartmentQuarterlyBonus", () => {
  it("nền dưới 300 triệu thì mốc tăng trưởng 5% chưa được thưởng", () => {
    const result = calcDepartmentQuarterlyBonus({
      departmentQuarterFYC: 210,
      lastYearSameQuarterFYC: 200,
      k2Percent: 70,
    });
    expect(result.growthPercent).toBeCloseTo(5, 6);
    expect(result.eligible).toBe(false);
  });

  it("nền từ 300 triệu trở lên thì mốc 5% được 1%", () => {
    const result = calcDepartmentQuarterlyBonus({
      departmentQuarterFYC: 315,
      lastYearSameQuarterFYC: 300,
      k2Percent: 70,
    });
    expect(result.ratePercent).toBe(1);
    expect(result.amount).toBeCloseTo(3.15, 6);
  });

  it("tăng trưởng từ 20% trở lên được 4% ở cả hai nền", () => {
    expect(
      calcDepartmentQuarterlyBonus({
        departmentQuarterFYC: 240,
        lastYearSameQuarterFYC: 200,
        k2Percent: 70,
      }).ratePercent
    ).toBe(4);
  });

  it("K2 ban dưới 60% thì không được nhận", () => {
    const result = calcDepartmentQuarterlyBonus({
      departmentQuarterFYC: 400,
      lastYearSameQuarterFYC: 200,
      k2Percent: 59,
    });
    expect(result.eligible).toBe(false);
    expect(result.reason).toContain("K2");
  });

  it("ban chưa có K2 thì miễn xét ngưỡng", () => {
    const result = calcDepartmentQuarterlyBonus({
      departmentQuarterFYC: 400,
      lastYearSameQuarterFYC: 200,
      k2Percent: null,
    });
    expect(result.eligible).toBe(true);
  });
});
