/**
 * Hỗ trợ phát triển nhóm gián tiếp — Mục A.IV, memo ALL/AD/2608 – 002.
 *
 * Khi một thành viên tách nhóm lên UM, người quản lý cũ được hai khoản tính
 * trên thưởng quản lý nhóm trực tiếp của nhóm gián tiếp cấp 1 đó:
 *  - Thưởng quản lý gián tiếp: 50%, trả lâu dài
 *  - Thưởng thăng tiến: 75% trong 6 tháng đầu, 25% tháng 7–12, sau đó dừng
 */

export const INDIRECT_MANAGEMENT_RATE = 50;

/** %Thưởng thăng tiến theo số tháng kể từ khi nhóm gián tiếp cấp 1 được bổ nhiệm. */
export function promotionSupportRate(monthsSinceAppointment: number): number {
  if (monthsSinceAppointment < 1) return 0;
  if (monthsSinceAppointment <= 6) return 75;
  if (monthsSinceAppointment <= 12) return 25;
  return 0;
}

export interface IndirectSupportInput {
  /** Thưởng quản lý nhóm trực tiếp của nhóm gián tiếp cấp 1 trong kỳ, triệu đồng. */
  indirectTeamBonus: number;
  /** Số tháng kể từ khi nhóm gián tiếp cấp 1 được bổ nhiệm. */
  monthsSinceAppointment: number;
}

export interface IndirectSupportResult {
  managementAmount: number;
  promotionAmount: number;
  promotionRatePercent: number;
  total: number;
}

export function calcIndirectSupport(
  input: IndirectSupportInput
): IndirectSupportResult {
  const base = Math.max(input.indirectTeamBonus, 0);
  const promotionRatePercent = promotionSupportRate(
    input.monthsSinceAppointment
  );
  const managementAmount = (base * INDIRECT_MANAGEMENT_RATE) / 100;
  const promotionAmount = (base * promotionRatePercent) / 100;

  return {
    managementAmount,
    promotionAmount,
    promotionRatePercent,
    total: managementAmount + promotionAmount,
  };
}
