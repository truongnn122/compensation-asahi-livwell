/**
 * Các ví dụ do chính memo in ra. Đây là mốc đối chiếu đáng tin nhất: nếu sửa
 * công thức mà lệch khỏi những con số này thì test báo ngay.
 */
import { calcAttendancePercent, calcFA } from "../fixedAllowance";

describe("Memo ALL/AD/2608 – 001, Mục A.II — ví dụ phụ cấp cố định", () => {
  // "Tháng 10/2026 có 22 ngày làm việc. LP A đi làm 20 ngày làm việc.
  //  %attendance = (20 + 1) ÷ 22 = 95,5%; Hệ số ngày làm việc = 95,5% × 22 = 21,0.
  //  LP A đạt FYC 26 triệu đồng, K2 chưa có (N/A), chưa từng đạt MDRT, có 02 hợp đồng.
  //  FA của LP A = 0,5 triệu × 21,0 × 100% × 100% × 100% = 10,5 triệu đồng."
  const attendancePercent = calcAttendancePercent(20);

  const result = calcFA({
    fyc: 26,
    policies: 2,
    referredAgents: 0,
    k2Percent: null,
    attendancePercent,
    mdrtLastYear: false,
  });

  it("%attendance ra 95,5%", () => {
    expect(attendancePercent).toBeCloseTo(95.45, 2);
    expect(Math.round(attendancePercent * 10) / 10).toBe(95.5);
  });

  it("hệ số ngày làm việc ra 21,0", () => {
    expect(result.workingDayCoefficient).toBeCloseTo(21, 6);
  });

  it("phụ cấp mỗi ngày ra 0,5 triệu", () => {
    expect(result.allowancePerWorkingDay).toBe(0.5);
  });

  it("cả ba hệ số đều 100%", () => {
    expect(result.k2Factor).toBe(100);
    expect(result.mdrtFactor).toBe(100);
    expect(result.policyFactor).toBe(100);
  });

  it("FA ra đúng 10,5 triệu như memo", () => {
    expect(result.amount).toBeCloseTo(10.5, 6);
  });
});
