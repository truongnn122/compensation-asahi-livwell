import {
  calcAttendancePercent,
  calcFA,
  calcPartTimeAttendancePercent,
  k2Factor,
  mdrtFactor,
  policyFactor,
} from "../fixedAllowance";

describe("calcAttendancePercent", () => {
  it("(số ngày + 1) ÷ 22, chặn trên ở 100%", () => {
    expect(calcAttendancePercent(20)).toBeCloseTo(95.45, 2);
    expect(calcAttendancePercent(21)).toBe(100);
    expect(calcAttendancePercent(22)).toBe(100);
    expect(calcAttendancePercent(0)).toBeCloseTo(4.55, 2);
  });
});

const base = {
  fyc: 26,
  policies: 2,
  referredAgents: 0,
  k2Percent: null as number | null,
  attendancePercent: 100,
  mdrtLastYear: false,
};

describe("các hệ số thành phần", () => {
  it("%Hệ số K2", () => {
    expect(k2Factor(null)).toBe(100);
    expect(k2Factor(85)).toBe(120);
    expect(k2Factor(70)).toBe(110);
    expect(k2Factor(60)).toBe(100);
    expect(k2Factor(59.9)).toBe(50);
  });

  it("%Hệ số MDRT", () => {
    expect(mdrtFactor(true)).toBe(120);
    expect(mdrtFactor(false)).toBe(100);
  });

  it("%Hệ số hợp đồng", () => {
    expect(policyFactor(2)).toBe(100);
    expect(policyFactor(1)).toBe(70);
    expect(policyFactor(0.5)).toBe(50);
    expect(policyFactor(0)).toBe(0);
  });
});

describe("calcFA", () => {
  it("ví dụ minh họa trong chính sách: 0,5 × 22 × 100% × 100% × 100% = 11 triệu", () => {
    const result = calcFA(base);
    expect(result.allowancePerWorkingDay).toBe(0.5);
    expect(result.workingDayCoefficient).toBe(22);
    expect(result.amount).toBeCloseTo(11, 6);
  });

  it("có LP giới thiệu thì mở mức phụ cấp/ngày cao hơn", () => {
    expect(calcFA({ ...base, referredAgents: 1 }).allowancePerWorkingDay).toBe(
      0.7
    );
  });

  it("1 hợp đồng chỉ còn 70% hệ số hợp đồng", () => {
    const result = calcFA({ ...base, policies: 1 });
    expect(result.policyFactor).toBe(70);
    expect(result.amount).toBeCloseTo(7.7, 6);
  });

  it("dưới 0,5 hợp đồng thì không có FA", () => {
    const result = calcFA({ ...base, policies: 0 });
    expect(result.eligible).toBe(false);
    expect(result.amount).toBe(0);
  });

  it("chấm công dưới 70% thì không đạt điều kiện", () => {
    const result = calcFA({ ...base, attendancePercent: 69 });
    expect(result.eligible).toBe(false);
    expect(result.reason).toContain("Chấm công");
  });

  it("K2 cao và đạt MDRT thì cộng hệ số", () => {
    const result = calcFA({ ...base, k2Percent: 85, mdrtLastYear: true });
    expect(result.amount).toBeCloseTo(0.5 * 22 * 1.2 * 1.2, 6);
  });
});

describe("calcPartTimeAttendancePercent", () => {
  it("số ngày ÷ 4, không cộng +1", () => {
    expect(calcPartTimeAttendancePercent(1)).toBe(25);
    expect(calcPartTimeAttendancePercent(2)).toBe(50);
    expect(calcPartTimeAttendancePercent(3)).toBe(75);
    expect(calcPartTimeAttendancePercent(4)).toBe(100);
  });

  it("chặn trên ở 100% và dưới ở 0%", () => {
    expect(calcPartTimeAttendancePercent(6)).toBe(100);
    expect(calcPartTimeAttendancePercent(-2)).toBe(0);
  });
});
