import {
  bandingFYC,
  calcManagerMonthlyBonus,
  calcManagerQuarterlyBonus,
  convertQuarterFYC,
  resolveManagerColumn,
} from "../managerBonus";

describe("resolveManagerColumn", () => {
  it("chọn cột theo K2", () => {
    expect(resolveManagerColumn(85, 0)).toBe(1);
    expect(resolveManagerColumn(75, 0)).toBe(2);
    expect(resolveManagerColumn(65, 0)).toBe(3);
    expect(resolveManagerColumn(45, 0)).toBe(4);
  });

  it("1Up mở được cột cao hơn khi K2 thấp", () => {
    expect(resolveManagerColumn(45, 25)).toBe(1);
    expect(resolveManagerColumn(45, 15)).toBe(2);
    expect(resolveManagerColumn(45, 10)).toBe(3);
  });

  it("chưa có K2 và 1Up dưới 10 thì dùng bảng NA", () => {
    expect(resolveManagerColumn(null, 0)).toBe("NA");
    expect(resolveManagerColumn(null, 9.9)).toBe("NA");
  });

  it("chưa có K2 nhưng 1Up từ 10 trở lên vẫn dùng bảng chính", () => {
    expect(resolveManagerColumn(null, 10)).toBe(3);
    expect(resolveManagerColumn(null, 25)).toBe(1);
  });
});

describe("bandingFYC", () => {
  it("cộng FYC nhóm gián tiếp cấp 1 vào để tra bậc", () => {
    expect(bandingFYC(40, 10)).toBe(50);
  });

  it("phần cộng thêm không vượt quá 50% FYC nhóm trực tiếp", () => {
    expect(bandingFYC(40, 100)).toBe(60);
  });

  it("không có nhóm gián tiếp thì giữ nguyên", () => {
    expect(bandingFYC(40, 0)).toBe(40);
  });
});

describe("calcManagerMonthlyBonus", () => {
  it("tra đúng ô trong bảng", () => {
    const result = calcManagerMonthlyBonus({
      directFYC: 45,
      k2Percent: 85,
      oneUpPercent: 0,
    });
    expect(result.column).toBe(1);
    expect(result.ratePercent).toBe(46);
    expect(result.amount).toBeCloseTo(45 * 0.46, 6);
  });

  it("lấy bậc cao nhất mà FYC vượt qua", () => {
    expect(
      calcManagerMonthlyBonus({
        directFYC: 200,
        k2Percent: 85,
        oneUpPercent: 0,
      }).ratePercent
    ).toBe(52);
  });

  it("dưới 15 triệu thì chưa có thưởng", () => {
    const result = calcManagerMonthlyBonus({
      directFYC: 14,
      k2Percent: 85,
      oneUpPercent: 0,
    });
    expect(result.eligible).toBe(false);
    expect(result.amount).toBe(0);
  });

  it("K2 dưới 40% thì không được nhận", () => {
    const result = calcManagerMonthlyBonus({
      directFYC: 100,
      k2Percent: 39,
      oneUpPercent: 30,
    });
    expect(result.eligible).toBe(false);
    expect(result.reason).toContain("K2");
  });

  it("tiền tính trên FYC trực tiếp, bậc tra theo FYC đã cộng gián tiếp", () => {
    const result = calcManagerMonthlyBonus({
      directFYC: 20,
      indirectLevel1FYC: 10,
      k2Percent: 85,
      oneUpPercent: 0,
    });
    expect(result.bandingFYC).toBe(30);
    expect(result.ratePercent).toBe(43);
    expect(result.amount).toBeCloseTo(20 * 0.43, 6);
  });
});

describe("calcManagerQuarterlyBonus", () => {
  it("tra đúng ô trong bảng quý", () => {
    const result = calcManagerQuarterlyBonus({
      directFYC: 130,
      k2Percent: 75,
      oneUpPercent: 0,
    });
    expect(result.column).toBe(2);
    expect(result.ratePercent).toBe(22);
    expect(result.amount).toBeCloseTo(130 * 0.22, 6);
  });

  it("bảng NA dùng cho người chưa có K2 và 1Up thấp", () => {
    const result = calcManagerQuarterlyBonus({
      directFYC: 80,
      k2Percent: null,
      oneUpPercent: 0,
    });
    expect(result.column).toBe("NA");
    expect(result.ratePercent).toBe(17);
  });

  it("cũng cộng FYC nhóm gián tiếp cấp 1 để tra bậc, như bảng tháng", () => {
    const withoutIndirect = calcManagerQuarterlyBonus({
      directFYC: 60,
      k2Percent: 85,
      oneUpPercent: 0,
    });
    const withIndirect = calcManagerQuarterlyBonus({
      directFYC: 60,
      indirectLevel1FYC: 30,
      k2Percent: 85,
      oneUpPercent: 0,
    });

    expect(withoutIndirect.bandingFYC).toBe(60);
    expect(withoutIndirect.ratePercent).toBe(20);
    // 60 + 30 = 90 vượt mốc 70 nên lên bậc cao hơn
    expect(withIndirect.bandingFYC).toBe(90);
    expect(withIndirect.ratePercent).toBe(22);
    // Tiền vẫn tính trên FYC nhóm trực tiếp thực tế
    expect(withIndirect.amount).toBeCloseTo(60 * 0.22, 6);
  });

  it("phần gián tiếp cộng vào quý cũng bị chặn ở 50% FYC nhóm trực tiếp", () => {
    const result = calcManagerQuarterlyBonus({
      directFYC: 60,
      indirectLevel1FYC: 500,
      k2Percent: 85,
      oneUpPercent: 0,
    });
    expect(result.bandingFYC).toBe(90);
  });
});

describe("convertQuarterFYC", () => {
  it("quy đổi khi mới hoạt động một phần quý", () => {
    expect(convertQuarterFYC(60, 1)).toBe(180);
    expect(convertQuarterFYC(60, 2)).toBe(90);
  });

  it("đủ ba tháng thì giữ nguyên", () => {
    expect(convertQuarterFYC(60, 3)).toBe(60);
  });
});
