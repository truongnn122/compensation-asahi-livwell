/**
 * Thưởng quản lý ban của DM — Mục B, memo ALL/AD/2608 – 002.
 * DM hưởng đủ các khoản của UM trên nhóm trực tiếp, cộng thêm hai khoản này
 * tính trên doanh số ban.
 */
import { resolveManagerColumn, type ManagerBonusColumn } from "./managerBonus";

/** Trần thưởng quản lý ban hàng tháng. */
export const DEPARTMENT_MONTHLY_BONUS_CAP = 25;

export const DEPARTMENT_MONTHLY_MIN_K2_PERCENT = 40;
export const DEPARTMENT_QUARTERLY_MIN_K2_PERCENT = 60;

/** Mẫu số tăng trưởng được ấn định tối thiểu ở mức này. */
export const GROWTH_DENOMINATOR_FLOOR = 175;

/** Mốc chia hai dòng của bảng thưởng tăng trưởng. */
export const GROWTH_ROW_SPLIT = 300;

interface ActivityRow {
  minActivityPercent: number;
  /** Tỷ lệ theo cột chất lượng 1..4. */
  rates: [number, number, number, number];
}

const MONTHLY_ROWS: ActivityRow[] = [
  { minActivityPercent: 30, rates: [10, 8, 6, 3] },
  { minActivityPercent: 45, rates: [11, 8, 7, 4] },
  { minActivityPercent: 60, rates: [12, 9, 8, 5] },
];

export interface DepartmentMonthlyInput {
  /** FYC toàn ban trong tháng, triệu đồng. */
  departmentFYC: number;
  /** %Tỷ lệ hoạt động của ban. */
  activityPercent: number;
  k2Percent: number | null;
  oneUpPercent: number;
}

export interface DepartmentMonthlyResult {
  eligible: boolean;
  reason: string | null;
  column: ManagerBonusColumn;
  ratePercent: number;
  /** Số tiền trước khi áp trần, để thấy rõ phần bị cắt. */
  uncappedAmount: number;
  capped: boolean;
  amount: number;
}

export function calcDepartmentMonthlyBonus(
  input: DepartmentMonthlyInput
): DepartmentMonthlyResult {
  const column = resolveManagerColumn(input.k2Percent, input.oneUpPercent);

  let ratePercent = 0;
  if (column !== "NA") {
    for (const row of MONTHLY_ROWS) {
      if (input.activityPercent >= row.minActivityPercent)
        ratePercent = row.rates[column - 1];
    }
  }

  const uncappedAmount = (input.departmentFYC * ratePercent) / 100;
  const base = { column, ratePercent, uncappedAmount, capped: false };

  if (
    input.k2Percent !== null &&
    input.k2Percent < DEPARTMENT_MONTHLY_MIN_K2_PERCENT
  ) {
    return {
      ...base,
      eligible: false,
      reason: `Ban cần K2 ≥ ${DEPARTMENT_MONTHLY_MIN_K2_PERCENT}%`,
      amount: 0,
    };
  }

  // Bảng thưởng ban không có cột dành cho người chưa có K2, khác bảng nhóm trực tiếp.
  if (column === "NA") {
    return {
      ...base,
      eligible: false,
      reason: "Ban chưa có K2 và %1Up dưới 10%",
      amount: 0,
    };
  }

  if (ratePercent === 0) {
    return {
      ...base,
      eligible: false,
      reason: `Tỷ lệ hoạt động của ban chưa đạt ${MONTHLY_ROWS[0].minActivityPercent}%`,
      amount: 0,
    };
  }

  const amount = Math.min(uncappedAmount, DEPARTMENT_MONTHLY_BONUS_CAP);
  return {
    ...base,
    eligible: true,
    reason: null,
    capped: amount < uncappedAmount,
    amount,
  };
}

/**
 * %Tăng trưởng FYC ban so với cùng kỳ năm trước. Mẫu số thấp hơn 175 triệu
 * (hoặc không có số liệu) được ấn định ở mức 175 triệu.
 */
export function departmentGrowthPercent(
  currentQuarterFYC: number,
  lastYearSameQuarterFYC: number
): number {
  const denominator = Math.max(
    lastYearSameQuarterFYC,
    GROWTH_DENOMINATOR_FLOOR
  );
  const raw = (currentQuarterFYC / denominator - 1) * 100;
  // Chính sách làm tròn mọi % tới 1 chữ số thập phân trước khi so ngưỡng. Không
  // làm tròn thì 240/200 ra 19,999...% và trượt mốc 20% dù thực tế đúng 20%.
  return Math.round(raw * 10) / 10;
}

/** Tỷ lệ thưởng tăng trưởng; null nghĩa là ô "Không áp dụng" trong bảng. */
function growthRate(
  lastYearSameQuarterFYC: number,
  growth: number
): number | null {
  const bigBase = lastYearSameQuarterFYC >= GROWTH_ROW_SPLIT;
  if (growth >= 20) return 4;
  if (growth >= 15) return 3;
  if (growth >= 10) return 2;
  if (growth >= 5) return bigBase ? 1 : null;
  return null;
}

export interface DepartmentQuarterlyInput {
  /** FYC toàn ban trong quý hiện tại, triệu đồng. */
  departmentQuarterFYC: number;
  /** FYC toàn ban của quý cùng kỳ năm trước, triệu đồng. */
  lastYearSameQuarterFYC: number;
  /** K2 trung bình quý của ban. */
  k2Percent: number | null;
}

export interface DepartmentQuarterlyResult {
  eligible: boolean;
  reason: string | null;
  growthPercent: number;
  /** Mẫu số thực dùng sau khi áp sàn 175 triệu. */
  denominator: number;
  ratePercent: number;
  amount: number;
}

export function calcDepartmentQuarterlyBonus(
  input: DepartmentQuarterlyInput
): DepartmentQuarterlyResult {
  const denominator = Math.max(
    input.lastYearSameQuarterFYC,
    GROWTH_DENOMINATOR_FLOOR
  );
  const growth = departmentGrowthPercent(
    input.departmentQuarterFYC,
    input.lastYearSameQuarterFYC
  );
  const rate = growthRate(input.lastYearSameQuarterFYC, growth);
  const base = { growthPercent: growth, denominator, ratePercent: rate ?? 0 };

  // Ban chưa có K2 được miễn xét ngưỡng này theo ghi chú của chính sách.
  if (
    input.k2Percent !== null &&
    input.k2Percent < DEPARTMENT_QUARTERLY_MIN_K2_PERCENT
  ) {
    return {
      ...base,
      eligible: false,
      reason: `Ban cần K2 ≥ ${DEPARTMENT_QUARTERLY_MIN_K2_PERCENT}%`,
      amount: 0,
    };
  }

  if (rate === null) {
    return {
      ...base,
      eligible: false,
      reason: "Tăng trưởng chưa đạt mức được thưởng",
      amount: 0,
    };
  }

  return {
    ...base,
    eligible: true,
    reason: null,
    amount: (input.departmentQuarterFYC * rate) / 100,
  };
}
