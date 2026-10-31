import {
  activeLPTurnFactor,
  calcManagerFA,
  managerAllowancePerDay,
} from "../managerAllowance";

describe("managerAllowancePerDay", () => {
  it("ba bậc theo FYC nhóm trực tiếp", () => {
    expect(managerAllowancePerDay(10, 2)).toBe(0.3);
    expect(managerAllowancePerDay(30, 2)).toBe(0.5);
    expect(managerAllowancePerDay(60, 2)).toBe(0.8);
    expect(managerAllowancePerDay(200, 2)).toBe(0.8);
  });

  it("dưới 10 triệu thì chưa có bậc nào", () => {
    expect(managerAllowancePerDay(9, 2)).toBe(0);
  });

  it("chưa đủ 2 LP hoạt động thì không có phụ cấp dù FYC cao", () => {
    expect(managerAllowancePerDay(200, 1)).toBe(0);
  });
});

describe("activeLPTurnFactor", () => {
  it("cộng hệ số theo lượt LP hoạt động tích luỹ", () => {
    expect(activeLPTurnFactor(0)).toBe(100);
    expect(activeLPTurnFactor(49)).toBe(100);
    expect(activeLPTurnFactor(50)).toBe(105);
    expect(activeLPTurnFactor(100)).toBe(115);
  });
});

describe("calcManagerFA", () => {
  const base = {
    teamFYC: 30,
    activeLPInMonth: 3,
    cumulativeActiveTurns: 0,
    k2Percent: null,
    attendancePercent: 100,
  };

  it("nhân đủ bốn thừa số", () => {
    const result = calcManagerFA(base);
    expect(result.allowancePerWorkingDay).toBe(0.5);
    expect(result.workingDayCoefficient).toBe(22);
    expect(result.amount).toBeCloseTo(0.5 * 22, 6);
  });

  it("K2 cao và nhiều lượt LP thì cộng hệ số", () => {
    const result = calcManagerFA({
      ...base,
      k2Percent: 85,
      cumulativeActiveTurns: 100,
    });
    expect(result.amount).toBeCloseTo(0.5 * 22 * 1.2 * 1.15, 6);
  });

  it("chấm công thấp làm giảm hệ số ngày làm việc", () => {
    const result = calcManagerFA({ ...base, attendancePercent: 50 });
    expect(result.workingDayCoefficient).toBe(11);
    expect(result.amount).toBeCloseTo(0.5 * 11, 6);
  });

  it("thiếu LP hoạt động thì không đạt điều kiện", () => {
    const result = calcManagerFA({ ...base, activeLPInMonth: 1 });
    expect(result.eligible).toBe(false);
    expect(result.reason).toContain("LP nhóm trực tiếp hoạt động");
  });

  it("FYC nhóm chưa đạt 10 triệu thì không có phụ cấp", () => {
    const result = calcManagerFA({ ...base, teamFYC: 5 });
    expect(result.eligible).toBe(false);
    expect(result.amount).toBe(0);
  });
});
