/**
 * Thưởng quản lý nhóm trực tiếp của UM/DM — Mục A.II (tháng) và A.III (quý),
 * memo ALL/AD/2608 – 002.
 *
 * Thưởng = FYC nhóm trực tiếp × Tỷ lệ thưởng, tỷ lệ tra theo bậc FYC và cột
 * chất lượng. Cột chất lượng xét theo K2 hoặc %1Up, lấy mức nào cao hơn.
 */

/** Điều kiện tối thiểu để nhận thưởng quản lý. */
export const MANAGER_BONUS_MIN_K2_PERCENT = 40;

/** Phần FYC nhóm gián tiếp cấp 1 được cộng vào, tối đa 50% FYC nhóm trực tiếp. */
export const INDIRECT_FYC_CAP_RATIO = 0.5;

/** Bốn cột chất lượng, cột 1 là cao nhất. NA là bảng dành cho người chưa có K2. */
export type ManagerBonusColumn = 1 | 2 | 3 | 4 | "NA";

interface BonusRow {
  minFYC: number;
  /** Tỷ lệ theo cột 1..4. */
  rates: [number, number, number, number];
  /** Tỷ lệ của bảng chưa có K2. */
  noK2Rate: number;
}

const MONTHLY_ROWS: BonusRow[] = [
  { minFYC: 15, rates: [40, 30, 25, 12], noK2Rate: 25 },
  { minFYC: 25, rates: [43, 33, 28, 15], noK2Rate: 28 },
  { minFYC: 40, rates: [46, 36, 31, 18], noK2Rate: 31 },
  { minFYC: 60, rates: [49, 39, 34, 21], noK2Rate: 34 },
  { minFYC: 90, rates: [52, 42, 37, 24], noK2Rate: 37 },
];

const QUARTERLY_ROWS: BonusRow[] = [
  { minFYC: 45, rates: [20, 15, 13, 6], noK2Rate: 13 },
  { minFYC: 70, rates: [22, 20, 17, 8], noK2Rate: 17 },
  { minFYC: 125, rates: [24, 22, 20, 9], noK2Rate: 20 },
  { minFYC: 150, rates: [26, 24, 22, 11], noK2Rate: 22 },
  { minFYC: 180, rates: [28, 26, 24, 12], noK2Rate: 24 },
];

/**
 * Chọn cột chất lượng. Bảng chính dùng khi đã có K2 hoặc %1Up ≥ 10; ngược lại
 * (chưa có K2 và 1Up < 10) rơi về bảng NA dành cho năm đầu.
 */
export function resolveManagerColumn(
  k2Percent: number | null,
  oneUpPercent: number
): ManagerBonusColumn {
  if (k2Percent === null && oneUpPercent < 10) return "NA";
  if ((k2Percent !== null && k2Percent >= 80) || oneUpPercent >= 25) return 1;
  if ((k2Percent !== null && k2Percent >= 70) || oneUpPercent >= 15) return 2;
  if ((k2Percent !== null && k2Percent >= 60) || oneUpPercent >= 10) return 3;
  return 4;
}

/**
 * FYC dùng để tra bậc thưởng: FYC nhóm trực tiếp cộng phần FYC nhóm gián tiếp
 * cấp 1 năm đầu, phần cộng thêm không vượt quá 50% FYC nhóm trực tiếp.
 */
export function bandingFYC(
  directFYC: number,
  indirectLevel1FYC: number
): number {
  const allowed = Math.min(
    Math.max(indirectLevel1FYC, 0),
    directFYC * INDIRECT_FYC_CAP_RATIO
  );
  return directFYC + allowed;
}

function lookupRate(
  rows: BonusRow[],
  fyc: number,
  column: ManagerBonusColumn
): number {
  let rate = 0;
  for (const row of rows) {
    if (fyc < row.minFYC) continue;
    rate = column === "NA" ? row.noK2Rate : row.rates[column - 1];
  }
  return rate;
}

export interface ManagerBonusInput {
  /** FYC nhóm trực tiếp của kỳ tính toán, triệu đồng. */
  directFYC: number;
  /** FYC nhóm gián tiếp cấp 1 năm đầu, được cộng vào khi tra bậc. */
  indirectLevel1FYC?: number;
  k2Percent: number | null;
  oneUpPercent: number;
}

export interface ManagerBonusResult {
  eligible: boolean;
  reason: string | null;
  column: ManagerBonusColumn;
  /** FYC đã cộng phần gián tiếp, dùng để tra bậc. */
  bandingFYC: number;
  ratePercent: number;
  amount: number;
}

function calc(
  rows: BonusRow[],
  input: ManagerBonusInput,
  period: string
): ManagerBonusResult {
  const banding = bandingFYC(input.directFYC, input.indirectLevel1FYC ?? 0);
  const column = resolveManagerColumn(input.k2Percent, input.oneUpPercent);
  const ratePercent = lookupRate(rows, banding, column);
  const base = { column, bandingFYC: banding, ratePercent };

  // Ngưỡng K2 chỉ xét khi đã có K2; chưa có K2 thì dùng bảng NA như chính sách quy định.
  if (
    input.k2Percent !== null &&
    input.k2Percent < MANAGER_BONUS_MIN_K2_PERCENT
  ) {
    return {
      ...base,
      eligible: false,
      reason: `Nhóm trực tiếp cần K2 ≥ ${MANAGER_BONUS_MIN_K2_PERCENT}%`,
      amount: 0,
    };
  }

  if (ratePercent === 0) {
    const floor = rows[0].minFYC;
    return {
      ...base,
      eligible: false,
      reason: `FYC nhóm ${period} chưa đạt ${floor} triệu`,
      amount: 0,
    };
  }

  // Tiền thưởng tính trên FYC nhóm trực tiếp thực tế, không tính trên FYC tra bậc.
  return {
    ...base,
    eligible: true,
    reason: null,
    amount: (input.directFYC * ratePercent) / 100,
  };
}

export function calcManagerMonthlyBonus(
  input: ManagerBonusInput
): ManagerBonusResult {
  return calc(MONTHLY_ROWS, input, "tháng");
}

export function calcManagerQuarterlyBonus(
  input: ManagerBonusInput
): ManagerBonusResult {
  return calc(QUARTERLY_ROWS, input, "quý");
}

/**
 * Quy đổi FYC quý cho UM/DM gia nhập giữa quý: tỷ lệ thưởng tra theo FYC quy
 * đổi, còn tiền thưởng vẫn tính trên FYC thực tế.
 */
export function convertQuarterFYC(
  actualFYC: number,
  activeMonthsInQuarter: number
): number {
  if (activeMonthsInQuarter <= 0 || activeMonthsInQuarter >= 3)
    return actualFYC;
  return actualFYC / (activeMonthsInQuarter / 3);
}
