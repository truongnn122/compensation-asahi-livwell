/**
 * Gói bảo hiểm sức khỏe — Mục A.VI của chính sách.
 * Đây là quyền lợi phi tiền mặt: chỉ xét đủ hay chưa đủ điều kiện, không quy
 * ra số tiền. Không được chuyển nhượng hay quy đổi thành tiền mặt.
 */

/** Cần tối thiểu 10 Tháng hoạt động trong 12 tháng gần nhất (theo doanh số thuần). */
export const HEALTH_INSURANCE_MIN_ACTIVE_MONTHS = 10;
export const HEALTH_INSURANCE_WINDOW_MONTHS = 12;

export interface HealthInsuranceInput {
  /** Số Tháng hoạt động trong 12 tháng gần nhất. */
  activeMonthsInLast12: number;
}

export interface HealthInsuranceResult {
  eligible: boolean;
  reason: string | null;
  /** Còn thiếu bao nhiêu tháng hoạt động nữa mới đạt. */
  monthsShort: number;
}

export function checkHealthInsurance(
  input: HealthInsuranceInput
): HealthInsuranceResult {
  const monthsShort = Math.max(
    0,
    HEALTH_INSURANCE_MIN_ACTIVE_MONTHS - input.activeMonthsInLast12
  );

  if (monthsShort > 0) {
    return {
      eligible: false,
      reason: `Còn thiếu ${monthsShort} Tháng hoạt động`,
      monthsShort,
    };
  }

  return { eligible: true, reason: null, monthsShort: 0 };
}
