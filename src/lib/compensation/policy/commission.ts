/**
 * Hoa hồng bảo hiểm — Mục A.I của chính sách.
 * Gồm hoa hồng năm đầu (FYC, chi trả không kèm điều kiện) và hoa hồng tái tục
 * (RYC, chi trả theo quý với tỷ lệ phụ thuộc số Tháng hoạt động và %K2).
 */

/** Hoa hồng năm đầu bình quân trên mỗi hợp đồng (triệu đồng). */
export const FYC_PER_POLICY = 4;

/** Ngưỡng xác định "Tháng hoạt động" (Mục D.i): ≥1 hợp đồng thuần và ≥4 triệu FYC thuần. */
export const ACTIVE_MONTH_MIN_POLICIES = 1;
export const ACTIVE_MONTH_MIN_FYC = 4;

/** FYC ước tính theo số hợp đồng, dùng để tự điền ô FYC khi đổi số hợp đồng. */
export function estimateFYC(policies: number): number {
  return policies * FYC_PER_POLICY;
}

export interface RYCRateInput {
  /** Số Tháng hoạt động trong quý (tháng có ≥1 hợp đồng thuần và ≥4 triệu FYC thuần). */
  activeMonthsInQuarter: number;
  /** %K2 theo đơn vị phần trăm; null = chưa có K2 (N/A). */
  k2Percent: number | null;
}

export interface RYCRateResult {
  /** Tỷ lệ chi trả %RYC: 0, 80, 100 hoặc 120. */
  ratePercent: number;
  eligible: boolean;
  /** Lý do khi không đủ điều kiện chi trả. */
  reason: string | null;
}

/**
 * Bảng tỷ lệ chi trả RYC (Mục A.I.2):
 *   ≥2 Tháng hoạt động và K2 ≥80%  → 120%
 *   ≥1 Tháng hoạt động và K2 ≥60%  → 100%
 *   ≥1 Tháng hoạt động và K2 <60%  → 80%
 * Chưa có K2 (N/A) được xếp vào mức 100%, thống nhất với cách %Hệ số K2 của
 * phụ cấp cố định xử lý trường hợp N/A.
 */
export function calcRYCRate(input: RYCRateInput): RYCRateResult {
  if (input.activeMonthsInQuarter < 1) {
    return {
      ratePercent: 0,
      eligible: false,
      reason: "Cần tối thiểu 1 Tháng hoạt động trong quý",
    };
  }
  if (input.k2Percent === null) {
    return { ratePercent: 100, eligible: true, reason: null };
  }
  if (input.activeMonthsInQuarter >= 2 && input.k2Percent >= 80) {
    return { ratePercent: 120, eligible: true, reason: null };
  }
  if (input.k2Percent >= 60) {
    return { ratePercent: 100, eligible: true, reason: null };
  }
  return { ratePercent: 80, eligible: true, reason: null };
}

export interface CommissionInput {
  /** Hoa hồng năm đầu phát sinh trong tháng (triệu đồng). */
  fyc: number;
  /** Hoa hồng tái tục gốc, trước khi áp tỷ lệ chi trả (triệu đồng). */
  rycBase: number;
  activeMonthsInQuarter: number;
  k2Percent: number | null;
}

export interface CommissionResult {
  fyc: number;
  rycBase: number;
  rycRate: RYCRateResult;
  /** RYC thực nhận = RYC gốc × tỷ lệ chi trả. */
  ryc: number;
  /** Tổng hoa hồng bảo hiểm = FYC + RYC thực nhận. */
  total: number;
}

export function calcCommission(input: CommissionInput): CommissionResult {
  const rycRate = calcRYCRate({
    activeMonthsInQuarter: input.activeMonthsInQuarter,
    k2Percent: input.k2Percent,
  });
  const ryc = (input.rycBase * rycRate.ratePercent) / 100;

  return {
    fyc: input.fyc,
    rycBase: input.rycBase,
    rycRate,
    ryc,
    total: input.fyc + ryc,
  };
}
