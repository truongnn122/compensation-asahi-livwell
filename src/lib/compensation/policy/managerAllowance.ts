/**
 * Phụ cấp cố định hàng tháng của UM/DM — Mục A.I, memo ALL/AD/2608 – 002.
 * FA = Phụ cấp/ngày làm việc × Hệ số ngày làm việc × %Hệ số K2 × %Hệ số Lượt LP hoạt động
 *
 * Khác LP ở hai chỗ: bậc phụ cấp xét theo kết quả của nhóm trực tiếp chứ không
 * phải cá nhân, và thừa số cuối là lượt LP hoạt động thay cho %MDRT × %hợp đồng.
 */
import { STANDARD_WORKING_DAYS } from "./fixedAllowance";

/** Mỗi bậc phụ cấp đều đòi tối thiểu 2 LP nhóm trực tiếp hoạt động trong tháng. */
export const MANAGER_FA_MIN_ACTIVE_LP = 2;

interface AllowanceTier {
  minTeamFYC: number;
  allowance: number;
}

const ALLOWANCE_TIERS: AllowanceTier[] = [
  { minTeamFYC: 10, allowance: 0.3 },
  { minTeamFYC: 30, allowance: 0.5 },
  { minTeamFYC: 60, allowance: 0.8 },
];

/** Phụ cấp/ngày làm việc theo FYC nhóm trực tiếp; 0 nếu chưa đủ LP hoạt động. */
export function managerAllowancePerDay(
  teamFYC: number,
  activeLPInMonth: number
): number {
  if (activeLPInMonth < MANAGER_FA_MIN_ACTIVE_LP) return 0;

  let found = 0;
  for (const tier of ALLOWANCE_TIERS) {
    if (teamFYC >= tier.minTeamFYC) found = tier.allowance;
  }
  return found;
}

/** %Hệ số K2 — giống bảng của LP, N/A xếp cùng mức ≥60%. */
export function managerK2Factor(k2Percent: number | null): number {
  if (k2Percent === null) return 100;
  if (k2Percent >= 80) return 120;
  if (k2Percent >= 70) return 110;
  if (k2Percent >= 60) return 100;
  return 50;
}

/**
 * %Hệ số Lượt LP nhóm trực tiếp hoạt động. Lượt này tính tích luỹ kể từ lần
 * thăng tiến/bổ nhiệm lên UM gần nhất, không phải trong tháng.
 */
export function activeLPTurnFactor(cumulativeActiveTurns: number): number {
  if (cumulativeActiveTurns >= 100) return 115;
  if (cumulativeActiveTurns >= 50) return 105;
  return 100;
}

export interface ManagerFAInput {
  /** FYC nhóm trực tiếp trong tháng tính toán, triệu đồng. */
  teamFYC: number;
  /** Số LP nhóm trực tiếp hoạt động trong tháng. */
  activeLPInMonth: number;
  /** Lượt LP hoạt động tích luỹ từ khi lên UM. */
  cumulativeActiveTurns: number;
  k2Percent: number | null;
  attendancePercent: number;
}

export interface ManagerFAResult {
  eligible: boolean;
  reason: string | null;
  allowancePerWorkingDay: number;
  workingDayCoefficient: number;
  k2Factor: number;
  activeLPFactor: number;
  amount: number;
}

export function calcManagerFA(input: ManagerFAInput): ManagerFAResult {
  const allowancePerWorkingDay = managerAllowancePerDay(
    input.teamFYC,
    input.activeLPInMonth
  );
  const workingDayCoefficient =
    (input.attendancePercent / 100) * STANDARD_WORKING_DAYS;
  const k2 = managerK2Factor(input.k2Percent);
  const activeLPFactor = activeLPTurnFactor(input.cumulativeActiveTurns);

  const breakdown = {
    allowancePerWorkingDay,
    workingDayCoefficient,
    k2Factor: k2,
    activeLPFactor,
  };

  if (input.activeLPInMonth < MANAGER_FA_MIN_ACTIVE_LP) {
    return {
      ...breakdown,
      eligible: false,
      reason: `Cần tối thiểu ${MANAGER_FA_MIN_ACTIVE_LP} LP nhóm trực tiếp hoạt động`,
      amount: 0,
    };
  }

  if (allowancePerWorkingDay === 0) {
    return {
      ...breakdown,
      eligible: false,
      reason: "FYC nhóm trực tiếp chưa đạt 10 triệu",
      amount: 0,
    };
  }

  return {
    ...breakdown,
    eligible: true,
    reason: null,
    amount:
      allowancePerWorkingDay *
      workingDayCoefficient *
      (k2 / 100) *
      (activeLPFactor / 100),
  };
}
