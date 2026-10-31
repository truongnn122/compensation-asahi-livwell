/**
 * Phụ cấp cố định hàng tháng (FA) — Mục A.II của chính sách.
 * FA = Phụ cấp/ngày làm việc × Hệ số ngày làm việc × %Hệ số K2 × %Hệ số MDRT × %Hệ số hợp đồng
 */

/** Số ngày làm việc chuẩn của tháng, dùng quy đổi hệ số ngày làm việc. */
export const STANDARD_WORKING_DAYS = 22;

/** Số ngày làm việc chuẩn của LP bán thời gian. */
export const PART_TIME_STANDARD_WORKING_DAYS = 4;

/** Điều kiện tối thiểu để nhận FA: chấm công ≥70% (hoặc ≥15 ngày làm việc). */
export const FA_MIN_ATTENDANCE_PERCENT = 70;

/**
 * %Chấm công = (số ngày đi làm thực tế + 1) ÷ số ngày làm việc của tháng, tối đa 100%.
 * Mẫu số cố định 22 ngày theo tháng làm việc chuẩn.
 */
export function calcAttendancePercent(actualWorkingDays: number): number {
  const raw = ((actualWorkingDays + 1) / STANDARD_WORKING_DAYS) * 100;
  return Math.min(100, Math.max(0, raw));
}

/**
 * %Chấm công của LP bán thời gian = số ngày đi làm ÷ 4, tối đa 100%.
 * Khác toàn thời gian ở hai điểm: mẫu số là 4 ngày chuẩn và không cộng +1.
 */
export function calcPartTimeAttendancePercent(
  actualWorkingDays: number
): number {
  const raw = (actualWorkingDays / PART_TIME_STANDARD_WORKING_DAYS) * 100;
  return Math.min(100, Math.max(0, raw));
}

interface Tier {
  minFYC: number;
  allowance: number;
}

/** Phụ cấp/ngày — Lựa chọn 1: xét theo FYC cá nhân. */
const ALLOWANCE_OPTION_1: Tier[] = [
  { minFYC: 0, allowance: 0.2 },
  { minFYC: 15, allowance: 0.3 },
  { minFYC: 25, allowance: 0.5 },
];

/** Phụ cấp/ngày — Lựa chọn 2: cần ≥1 LP được giới thiệu. */
const ALLOWANCE_OPTION_2: Tier[] = [
  { minFYC: 0, allowance: 0.3 },
  { minFYC: 15, allowance: 0.5 },
  { minFYC: 25, allowance: 0.7 },
];

function lookupAllowance(tiers: Tier[], fyc: number): number {
  let found = 0;
  for (const tier of tiers) {
    if (fyc >= tier.minFYC) found = tier.allowance;
    else break;
  }
  return found;
}

/** %Hệ số K2 — null (chưa có K2) xếp cùng mức ≥60%. */
export function k2Factor(k2Percent: number | null): number {
  if (k2Percent === null) return 100;
  if (k2Percent >= 80) return 120;
  if (k2Percent >= 70) return 110;
  if (k2Percent >= 60) return 100;
  return 50;
}

/** %Hệ số MDRT */
export function mdrtFactor(achievedLastYear: boolean): number {
  return achievedLastYear ? 120 : 100;
}

/** %Hệ số hợp đồng */
export function policyFactor(policies: number): number {
  if (policies >= 2) return 100;
  if (policies >= 1) return 70;
  if (policies >= 0.5) return 50;
  return 0;
}

export interface FAInput {
  fyc: number;
  policies: number;
  referredAgents: number;
  k2Percent: number | null;
  attendancePercent: number;
  mdrtLastYear: boolean;
}

export interface FAResult {
  eligible: boolean;
  reason: string | null;
  allowancePerWorkingDay: number;
  /** Hệ số ngày làm việc = %chấm công × 22. */
  workingDayCoefficient: number;
  k2Factor: number;
  mdrtFactor: number;
  policyFactor: number;
  amount: number;
}

export function calcFA(input: FAInput): FAResult {
  const allowancePerWorkingDay = Math.max(
    lookupAllowance(ALLOWANCE_OPTION_1, input.fyc),
    input.referredAgents >= 1
      ? lookupAllowance(ALLOWANCE_OPTION_2, input.fyc)
      : 0
  );
  const workingDayCoefficient =
    (input.attendancePercent / 100) * STANDARD_WORKING_DAYS;
  const k2 = k2Factor(input.k2Percent);
  const mdrt = mdrtFactor(input.mdrtLastYear);
  const policy = policyFactor(input.policies);

  const breakdown = {
    allowancePerWorkingDay,
    workingDayCoefficient,
    k2Factor: k2,
    mdrtFactor: mdrt,
    policyFactor: policy,
  };

  if (input.attendancePercent < FA_MIN_ATTENDANCE_PERCENT) {
    return {
      ...breakdown,
      eligible: false,
      reason: `Chấm công dưới ${FA_MIN_ATTENDANCE_PERCENT}%`,
      amount: 0,
    };
  }

  if (policy === 0) {
    return {
      ...breakdown,
      eligible: false,
      reason: "Chưa đạt 0,5 hợp đồng trong tháng",
      amount: 0,
    };
  }

  const amount =
    allowancePerWorkingDay *
    workingDayCoefficient *
    (k2 / 100) *
    (mdrt / 100) *
    (policy / 100);
  return { ...breakdown, eligible: true, reason: null, amount };
}
