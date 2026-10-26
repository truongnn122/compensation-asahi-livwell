import { calcIndirectSupport, promotionSupportRate } from "../indirectSupport";

describe("promotionSupportRate", () => {
  it("75% trong sáu tháng đầu", () => {
    expect(promotionSupportRate(1)).toBe(75);
    expect(promotionSupportRate(6)).toBe(75);
  });

  it("25% từ tháng 7 đến tháng 12", () => {
    expect(promotionSupportRate(7)).toBe(25);
    expect(promotionSupportRate(12)).toBe(25);
  });

  it("dừng từ tháng 13", () => {
    expect(promotionSupportRate(13)).toBe(0);
  });
});

describe("calcIndirectSupport", () => {
  it("thưởng quản lý gián tiếp luôn bằng 50%", () => {
    const result = calcIndirectSupport({
      indirectTeamBonus: 20,
      monthsSinceAppointment: 24,
    });
    expect(result.managementAmount).toBe(10);
    expect(result.promotionAmount).toBe(0);
    expect(result.total).toBe(10);
  });

  it("cộng thêm thưởng thăng tiến trong sáu tháng đầu", () => {
    const result = calcIndirectSupport({
      indirectTeamBonus: 20,
      monthsSinceAppointment: 3,
    });
    expect(result.promotionRatePercent).toBe(75);
    expect(result.total).toBe(10 + 15);
  });
});
