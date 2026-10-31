import type { CompPlan } from "./types";

/**
 * Placeholder policy used until the real scheme is configured. Every number
 * below is an assumption kept under the legal caps of Circular 67/2023/TT-BTC
 * and must be replaced from the Policy tab.
 */
export const DEFAULT_PLAN: CompPlan = {
  planId: "CS-MAU",
  effectiveFrom: "2026-01-01",
  ranks: [
    { code: "FC", name: "Tư vấn viên", level: 1 },
    { code: "UM", name: "Trưởng nhóm", level: 2 },
    { code: "BM", name: "Trưởng phòng", level: 3 },
    { code: "AD", name: "Giám đốc kinh doanh", level: 4 },
  ],
  products: [
    {
      code: "UL",
      name: "Liên kết chung",
      shortTerm: false,
      rates: { year1: 28, year2: 15, year3Plus: 5 },
      legalCapYear1: 30,
    },
    {
      code: "TERM",
      name: "Tử kỳ",
      shortTerm: false,
      rates: { year1: 25, year2: 5, year3Plus: 5 },
      legalCapYear1: 40,
    },
    {
      code: "ENDOW",
      name: "Hỗn hợp",
      shortTerm: false,
      rates: { year1: 25, year2: 10, year3Plus: 5 },
      legalCapYear1: 40,
    },
    {
      code: "HEALTH",
      name: "Bổ trợ sức khỏe",
      shortTerm: true,
      rates: { year1: 18, year2: 18, year3Plus: 18 },
      legalCapYear1: 20,
    },
  ],
  overrideRates: {
    UM: { direct: 20, indirect: 0 },
    BM: { direct: 10, indirect: 10 },
    AD: { direct: 5, indirect: 5 },
  },
  monthlyBonusTiers: [
    { minFyc: 10_000_000, minCases: 2, rate: 10 },
    { minFyc: 20_000_000, minCases: 2, rate: 15 },
    { minFyc: 40_000_000, minCases: 2, rate: 20 },
  ],
  chargebackSchedule: [
    { fromMonth: 1, toMonth: 12, recoverRate: 100 },
    { fromMonth: 13, toMonth: 24, recoverRate: 50 },
  ],
  bonusCapLongTermPct: 30,
  bonusCapShortTermPct: 20,
};
