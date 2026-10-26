/**
 * Chương trình thu hút MDRT — memo ALL/AD/2608 – 003, hiệu lực 01/10/2026.
 *
 * Dành cho LP có thành tích cao từ công ty khác gia nhập MVI. Ngoài ba khoản
 * dưới đây, LP vẫn hưởng đủ thu nhập LP theo memo 001. Khác các memo trước,
 * memo này đo bằng FYP (phí bảo hiểm năm đầu) chứ không phải FYC.
 */

/* ---------- I. Thưởng gia nhập ---------- */

export type SignOnRecord = "fyp700" | "mdrt" | "mdrt2" | "cot";

/** Mỗi mốc thành tích được đặt tên hạng; `criteria` là điều kiện gốc trong memo. */
export const SIGN_ON_RECORDS: {
  value: SignOnRecord;
  label: string;
  criteria: string;
  amount: number;
}[] = [
  { value: "fyp700", label: "ƯU TÚ", criteria: "FYP ≥ 700 triệu", amount: 100 },
  { value: "mdrt", label: "TINH ANH", criteria: "Đạt MDRT", amount: 150 },
  {
    value: "mdrt2",
    label: "DANH GIÁ",
    criteria: "MDRT 2 năm liên tiếp",
    amount: 200,
  },
  {
    value: "cot",
    label: "ĐỈNH CAO",
    criteria: "COT hoặc MDRT 3 năm liên tiếp",
    amount: 300,
  },
];

/** Ứng viên cần tỷ lệ duy trì hợp đồng tối thiểu này trong 6 tháng gần nhất ở công ty cũ. */
export const SIGN_ON_MIN_PERSISTENCY_PERCENT = 60;

/** Khoản tạm ứng bị thu hồi nếu nghỉ hoặc không duy trì được trong ngần này năm. */
export const SIGN_ON_CLAWBACK_YEARS = 3;

export interface SignOnInput {
  record: SignOnRecord;
  /** Tỷ lệ duy trì hợp đồng 6 tháng gần nhất ở công ty cũ, %. */
  previousPersistencyPercent: number;
}

export interface SignOnResult {
  eligible: boolean;
  reason: string | null;
  amount: number;
}

export function calcSignOnBonus(input: SignOnInput): SignOnResult {
  const amount =
    SIGN_ON_RECORDS.find(item => item.value === input.record)?.amount ?? 0;

  if (amount === 0) {
    return { eligible: false, reason: "Chưa chọn hạng thành tích", amount: 0 };
  }
  if (input.previousPersistencyPercent < SIGN_ON_MIN_PERSISTENCY_PERCENT) {
    return {
      eligible: false,
      reason: `Tỷ lệ duy trì ở công ty cũ dưới ${SIGN_ON_MIN_PERSISTENCY_PERCENT}%`,
      amount: 0,
    };
  }
  return { eligible: true, reason: null, amount };
}

/* ---------- II. Thưởng hỗ trợ thu nhập hàng tháng, 12 tháng đầu ---------- */

export const INCOME_SUPPORT_MONTHS = 12;
export const INCOME_SUPPORT_CAP = 40;
export const INCOME_SUPPORT_MIN_ATTENDANCE_PERCENT = 70;

interface SupportRow {
  minFYP: number;
  minPolicies: number;
  ratePercent: number;
}

/** Xếp từ cao xuống thấp: lấy dòng đầu tiên mà mọi điều kiện đều đạt. */
const SUPPORT_ROWS: SupportRow[] = [
  { minFYP: 100, minPolicies: 3, ratePercent: 20 },
  { minFYP: 60, minPolicies: 2, ratePercent: 15 },
];

export const INCOME_SUPPORT_MIN_FYP = 60;
export const INCOME_SUPPORT_MIN_POLICIES = 2;

export interface IncomeSupportInput {
  /** Tháng thứ mấy kể từ ngày bổ nhiệm, bắt đầu từ 1. */
  monthSinceJoining: number;
  /** FYP trong tháng tính toán, triệu đồng. */
  fyp: number;
  policies: number;
  attendancePercent: number;
}

export interface IncomeSupportResult {
  eligible: boolean;
  reason: string | null;
  ratePercent: number;
  /** Số tiền trước khi áp trần 40 triệu, để thấy rõ phần bị cắt. */
  uncappedAmount: number;
  capped: boolean;
  amount: number;
}

export function calcIncomeSupport(
  input: IncomeSupportInput
): IncomeSupportResult {
  const row = SUPPORT_ROWS.find(
    item => input.fyp >= item.minFYP && input.policies >= item.minPolicies
  );
  const ratePercent = row?.ratePercent ?? 0;
  const uncappedAmount = (input.fyp * ratePercent) / 100;
  const base = { ratePercent, uncappedAmount, capped: false };
  const blocked = (reason: string) => ({
    ...base,
    eligible: false,
    reason,
    amount: 0,
  });

  if (input.monthSinceJoining < 1)
    return blocked("Chưa nhập tháng thứ mấy kể từ bổ nhiệm");
  if (input.monthSinceJoining > INCOME_SUPPORT_MONTHS) {
    return blocked(
      `Chỉ áp dụng trong ${INCOME_SUPPORT_MONTHS} tháng đầu kể từ ngày bổ nhiệm`
    );
  }
  if (input.attendancePercent < INCOME_SUPPORT_MIN_ATTENDANCE_PERCENT) {
    return blocked(`Chấm công dưới ${INCOME_SUPPORT_MIN_ATTENDANCE_PERCENT}%`);
  }
  if (input.fyp < INCOME_SUPPORT_MIN_FYP)
    return blocked(`FYP tháng chưa đạt ${INCOME_SUPPORT_MIN_FYP} triệu`);
  if (!row)
    return blocked(
      `Chưa đạt ${INCOME_SUPPORT_MIN_POLICIES} hợp đồng trong tháng`
    );

  const amount = Math.min(uncappedAmount, INCOME_SUPPORT_CAP);
  return {
    ...base,
    eligible: true,
    reason: null,
    capped: amount < uncappedAmount,
    amount,
  };
}

/* ---------- III. Thưởng chất lượng và duy trì hợp đồng, chốt K2 tháng 26, chi trả tháng 27 ---------- */

export const QUALITY_BONUS_K2_MONTH = 26;
export const QUALITY_BONUS_PAY_MONTH = 27;

const YEAR_ONE_TIERS: { minFYP: number; amount: number }[] = [
  { minFYP: 1200, amount: 250 },
  { minFYP: 2400, amount: 500 },
  { minFYP: 3600, amount: 800 },
];

/** Thưởng MDRT năm 1 theo FYP 12 tháng đầu (triệu đồng). */
export function yearOneMDRTBonus(yearOneFYP: number): number {
  let amount = 0;
  for (const tier of YEAR_ONE_TIERS) {
    if (yearOneFYP >= tier.minFYP) amount = tier.amount;
  }
  return amount;
}

/** %Hệ số K2 riêng của khoản này — khác bảng K2 của FA: dưới 60% là mất hẳn. */
export function qualityK2Factor(k2Percent: number): number {
  if (k2Percent >= 80) return 110;
  if (k2Percent >= 70) return 100;
  if (k2Percent >= 60) return 50;
  return 0;
}

export interface QualityBonusInput {
  /** FYP trong 12 tháng đầu kể từ tháng gia nhập, triệu đồng. */
  yearOneFYP: number;
  /** %K2 của các hợp đồng phát hành trong 12 tháng đầu, chốt tại tháng QUALITY_BONUS_K2_MONTH. */
  k2Percent: number;
}

export interface QualityBonusResult {
  eligible: boolean;
  reason: string | null;
  baseAmount: number;
  k2Factor: number;
  amount: number;
}

export function calcQualityBonus(input: QualityBonusInput): QualityBonusResult {
  const baseAmount = yearOneMDRTBonus(input.yearOneFYP);
  const k2Factor = qualityK2Factor(input.k2Percent);
  const base = { baseAmount, k2Factor };

  if (baseAmount === 0) {
    return {
      ...base,
      eligible: false,
      reason: `FYP năm 1 chưa đạt ${YEAR_ONE_TIERS[0].minFYP.toLocaleString("vi-VN")} triệu`,
      amount: 0,
    };
  }
  if (k2Factor === 0)
    return { ...base, eligible: false, reason: "K2 dưới 60%", amount: 0 };

  return {
    ...base,
    eligible: true,
    reason: null,
    amount: (baseAmount * k2Factor) / 100,
  };
}
