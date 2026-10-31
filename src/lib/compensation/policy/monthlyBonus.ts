/**
 * Thưởng doanh số hàng tháng — Mục A.III của chính sách.
 * Thưởng = FYC tháng × Tỷ lệ thưởng, tỷ lệ tra theo bậc FYC và bậc chất lượng.
 */

/** Điều kiện tối thiểu: chấm công ≥4 ngày làm việc trong tháng. */
export const BONUS_MIN_WORKING_DAYS = 4;
/** Điều kiện tối thiểu: K2 ≥50% (miễn xét khi chưa có K2). */
export const BONUS_MIN_K2_PERCENT = 50;

/**
 * Cột bậc chất lượng:
 *   1 = ≥80% K2 hoặc ≥25% 1Up
 *   2 = ≥70% K2 hoặc ≥15% 1Up
 *   3 = ≥60% K2
 *   4 = ≥50% K2
 *   'NA' = chưa có K2 và 1Up chưa đủ để lên bậc
 *   0 = không đủ điều kiện
 */
export type BonusColumn = 1 | 2 | 3 | 4 | "NA" | 0;

interface BonusRow {
  minFYC: number;
  /** Tỷ lệ theo cột 1..4 (đơn vị %). */
  rates: [number, number, number, number];
  /** Tỷ lệ cho đại lý chưa có K2 (đơn vị %). */
  noK2: number;
}

const MONTHLY_BONUS_TABLE: BonusRow[] = [
  { minFYC: 5, rates: [15, 10, 5, 3], noK2: 5 },
  { minFYC: 10, rates: [25, 20, 15, 8], noK2: 15 },
  { minFYC: 15, rates: [35, 30, 25, 15], noK2: 25 },
  { minFYC: 25, rates: [50, 40, 30, 20], noK2: 30 },
];

export function resolveBonusColumn(
  k2Percent: number | null,
  oneUpPercent: number
): BonusColumn {
  if (k2Percent !== null) {
    if (k2Percent >= 80 || oneUpPercent >= 25) return 1;
    if (k2Percent >= 70 || oneUpPercent >= 15) return 2;
    if (k2Percent >= 60) return 3;
    if (k2Percent >= BONUS_MIN_K2_PERCENT) return 4;
    return 0;
  }
  if (oneUpPercent >= 25) return 1;
  if (oneUpPercent >= 15) return 2;
  return "NA";
}

function lookupRow(fyc: number): BonusRow | null {
  let found: BonusRow | null = null;
  for (const row of MONTHLY_BONUS_TABLE) {
    if (fyc >= row.minFYC) found = row;
    else break;
  }
  return found;
}

export interface MonthlyBonusInput {
  fyc: number;
  k2Percent: number | null;
  oneUpPercent: number;
  actualWorkingDays: number;
}

export interface MonthlyBonusResult {
  eligible: boolean;
  reason: string | null;
  column: BonusColumn;
  ratePercent: number;
  amount: number;
}

export function calcMonthlyBonus(input: MonthlyBonusInput): MonthlyBonusResult {
  const column = resolveBonusColumn(input.k2Percent, input.oneUpPercent);

  if (input.actualWorkingDays < BONUS_MIN_WORKING_DAYS) {
    return {
      eligible: false,
      reason: `Chấm công dưới ${BONUS_MIN_WORKING_DAYS} ngày làm việc`,
      column,
      ratePercent: 0,
      amount: 0,
    };
  }

  if (column === 0) {
    return {
      eligible: false,
      reason: `K2 dưới ${BONUS_MIN_K2_PERCENT}%`,
      column,
      ratePercent: 0,
      amount: 0,
    };
  }

  const row = lookupRow(input.fyc);
  if (!row) {
    return {
      eligible: false,
      reason: `FYC chưa đạt ${MONTHLY_BONUS_TABLE[0].minFYC} triệu`,
      column,
      ratePercent: 0,
      amount: 0,
    };
  }

  const ratePercent = column === "NA" ? row.noK2 : row.rates[column - 1];
  return {
    eligible: true,
    reason: null,
    column,
    ratePercent,
    amount: (input.fyc * ratePercent) / 100,
  };
}
