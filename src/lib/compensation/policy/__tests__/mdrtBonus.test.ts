import { calcMDRTBonus, nextMDRTMilestone } from "../mdrtBonus";

const base = {
  consecutiveTimes: 2,
  k2Percent: null as number | null,
  partTimeOrManager: false,
};

describe("calcMDRTBonus", () => {
  it("lấy mốc cao nhất đã đạt", () => {
    expect(calcMDRTBonus({ ...base, consecutiveTimes: 2 }).amount).toBe(50);
    expect(calcMDRTBonus({ ...base, consecutiveTimes: 4 }).amount).toBe(50);
    expect(calcMDRTBonus({ ...base, consecutiveTimes: 5 }).amount).toBe(500);
    expect(calcMDRTBonus({ ...base, consecutiveTimes: 9 }).amount).toBe(1000);
  });

  it("mốc 1 lần chỉ dành cho LP bán thời gian hoặc UM+", () => {
    const fullTime = calcMDRTBonus({ ...base, consecutiveTimes: 1 });
    expect(fullTime.eligible).toBe(false);
    expect(fullTime.reason).toContain("1 lần");

    const partTime = calcMDRTBonus({
      ...base,
      consecutiveTimes: 1,
      partTimeOrManager: true,
    });
    expect(partTime.amount).toBe(20);
  });

  it("chưa đạt lần nào thì không có thưởng", () => {
    const result = calcMDRTBonus({ ...base, consecutiveTimes: 0 });
    expect(result.eligible).toBe(false);
    expect(result.milestone).toBeNull();
  });

  it("K2 trung bình năm dưới 70% thì không được thưởng", () => {
    const result = calcMDRTBonus({
      ...base,
      consecutiveTimes: 5,
      k2Percent: 65,
    });
    expect(result.eligible).toBe(false);
    expect(result.reason).toContain("K2");
  });

  it("chưa có K2 thì miễn xét điều kiện K2", () => {
    expect(
      calcMDRTBonus({ ...base, consecutiveTimes: 5, k2Percent: null }).amount
    ).toBe(500);
  });
});

describe("nextMDRTMilestone", () => {
  it("chỉ ra mốc kế tiếp và còn thiếu mấy lần", () => {
    expect(nextMDRTMilestone(2, false)).toEqual({
      times: 5,
      amount: 500,
      remaining: 3,
    });
    expect(nextMDRTMilestone(6, false)).toEqual({
      times: 8,
      amount: 1000,
      remaining: 2,
    });
  });

  it("mốc 1 lần chỉ hiện với người được hưởng nó", () => {
    expect(nextMDRTMilestone(0, true)).toEqual({
      times: 1,
      amount: 20,
      remaining: 1,
    });
    expect(nextMDRTMilestone(0, false)).toEqual({
      times: 2,
      amount: 50,
      remaining: 2,
    });
  });

  it("đã ở mốc cao nhất thì không còn mốc nào", () => {
    expect(nextMDRTMilestone(8, false)).toBeNull();
  });
});
