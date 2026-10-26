/**
 * Thưởng doanh số hàng quý — Mục A.IV của chính sách.
 * Thưởng = FYC quý × Tỷ lệ thưởng. Cách chọn cột bậc chất lượng giống thưởng
 * tháng (Mục A.III), chỉ khác bảng ngưỡng FYC và tỷ lệ.
 */
import {
  BONUS_MIN_K2_PERCENT,
  BONUS_MIN_WORKING_DAYS,
  resolveBonusColumn,
  type BonusColumn,
} from "./monthlyBonus";

interface BonusRow {
  minFYC: number;
  rates: [number, number, number, number];
  noK2: number;
}

const QUARTERLY_BONUS_TABLE: BonusRow[] = [
  { minFYC: 30, rates: [20, 15, 5, 3], noK2: 5 },
  { minFYC: 40, rates: [25, 20, 10, 8], noK2: 10 },
  { minFYC: 60, rates: [35, 30, 20, 15], noK2: 20 },
  { minFYC: 90, rates: [50, 40, 30, 20], noK2: 30 },
];

function lookupRow(fyc: number): BonusRow | null {
  let found: BonusRow | null = null;
  for (const row of QUARTERLY_BONUS_TABLE) {
    if (fyc >= row.minFYC) found = row;
    else break;
  }
  return found;
}

export interface QuarterlyBonusInput {
  /** FYC cả quý (triệu đồng). */
  quarterFYC: number;
  /** %K2 trung bình quý; null = chưa có K2. */
  k2Percent: number | null;
  oneUpPercent: number;
  /** Số ngày chấm công trong kỳ. */
  actualWorkingDays: number;
}

export interface QuarterlyBonusResult {
  eligible: boolean;
  reason: string | null;
  column: BonusColumn;
  ratePercent: number;
  amount: number;
}

export function calcQuarterlyBonus(
  input: QuarterlyBonusInput
): QuarterlyBonusResult {
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

  const row = lookupRow(input.quarterFYC);
  if (!row) {
    return {
      eligible: false,
      reason: `FYC quý chưa đạt ${QUARTERLY_BONUS_TABLE[0].minFYC} triệu`,
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
    amount: (input.quarterFYC * ratePercent) / 100,
  };
}
