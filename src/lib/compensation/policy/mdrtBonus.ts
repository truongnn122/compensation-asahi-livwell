/**
 * Thưởng danh hiệu MDRT — Mục A.V của chính sách.
 * Thưởng theo số lần đạt MDRT liên tiếp, mỗi mốc một mức tiền cố định.
 */

/** Điều kiện tối thiểu: %K2 trung bình năm ≥70%, miễn xét khi chưa có K2. */
export const MDRT_MIN_K2_PERCENT = 70;

interface Milestone {
  times: number;
  amount: number;
  /** Mốc 1 lần chỉ áp dụng cho LP bán thời gian hoặc UM+. */
  partTimeOrManagerOnly?: boolean;
}

/** Sắp xếp giảm dần để lấy mốc cao nhất đã đạt. */
const MILESTONES: Milestone[] = [
  { times: 8, amount: 1000 },
  { times: 5, amount: 500 },
  { times: 2, amount: 50 },
  { times: 1, amount: 20, partTimeOrManagerOnly: true },
];

export interface MDRTBonusInput {
  /** Số lần đạt MDRT liên tiếp, không gián đoạn hợp đồng đại lý. */
  consecutiveTimes: number;
  /** %K2 trung bình năm; null = chưa có K2 (miễn xét điều kiện). */
  k2Percent: number | null;
  /** LP bán thời gian hoặc UM+ — cần cho mốc 1 lần. */
  partTimeOrManager: boolean;
}

export interface MDRTBonusResult {
  eligible: boolean;
  reason: string | null;
  /** Mốc được hưởng, null nếu chưa đạt mốc nào. */
  milestone: number | null;
  amount: number;
}

export function calcMDRTBonus(input: MDRTBonusInput): MDRTBonusResult {
  if (input.k2Percent !== null && input.k2Percent < MDRT_MIN_K2_PERCENT) {
    return {
      eligible: false,
      reason: `%K2 trung bình năm dưới ${MDRT_MIN_K2_PERCENT}%`,
      milestone: null,
      amount: 0,
    };
  }

  const reached = MILESTONES.find(
    m =>
      input.consecutiveTimes >= m.times &&
      (!m.partTimeOrManagerOnly || input.partTimeOrManager)
  );

  if (!reached) {
    const blockedByRole =
      input.consecutiveTimes >= 1 &&
      input.consecutiveTimes < 2 &&
      !input.partTimeOrManager;
    return {
      eligible: false,
      reason: blockedByRole
        ? "Mốc 1 lần chỉ dành cho LP bán thời gian hoặc UM+"
        : "Chưa đạt mốc MDRT liên tiếp nào",
      milestone: null,
      amount: 0,
    };
  }

  return {
    eligible: true,
    reason: null,
    milestone: reached.times,
    amount: reached.amount,
  };
}

export interface NextMDRTMilestone {
  times: number;
  amount: number;
  /** Còn bao nhiêu lần nữa mới tới mốc này. */
  remaining: number;
}

/**
 * Mốc thưởng kế tiếp chưa với tới, để người dùng biết còn cách bao xa. Trả về
 * null khi đã ở mốc cao nhất.
 */
export function nextMDRTMilestone(
  consecutiveTimes: number,
  partTimeOrManager: boolean
): NextMDRTMilestone | null {
  const upcoming = [...MILESTONES]
    .reverse()
    .filter(m => !m.partTimeOrManagerOnly || partTimeOrManager)
    .find(m => consecutiveTimes < m.times);

  if (!upcoming) return null;
  return {
    times: upcoming.times,
    amount: upcoming.amount,
    remaining: upcoming.times - consecutiveTimes,
  };
}
