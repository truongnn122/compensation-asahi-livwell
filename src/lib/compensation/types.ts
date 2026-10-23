/**
 * Domain types for the agent compensation engine.
 *
 * Everything here is plain data so the engine (`engine.ts`) can stay a pure
 * function that is easy to unit test and to run on the client or the server.
 */

export type RankCode = string;

export type Rank = {
  code: RankCode;
  name: string;
  /** 1 = lowest (selling agent). Higher numbers sit above in the hierarchy. */
  level: number;
};

export type Agent = {
  code: string;
  name: string;
  rank: RankCode;
  /** Code of the direct manager, empty for the top of the tree. */
  managerCode: string;
};

export type PaymentMode =
  "ANNUAL" | "SEMI" | "QUARTERLY" | "MONTHLY" | "SINGLE";

export const PAYMENT_MODES: PaymentMode[] = [
  "ANNUAL",
  "SEMI",
  "QUARTERLY",
  "MONTHLY",
  "SINGLE",
];

/** Multiplier from premium per instalment to annualised premium (APE). */
export const APE_FACTOR: Record<PaymentMode, number> = {
  ANNUAL: 1,
  SEMI: 2,
  QUARTERLY: 4,
  MONTHLY: 12,
  SINGLE: 0.1,
};

export type PremiumEventType =
  | "PREMIUM_COLLECTED"
  | "PREMIUM_REFUNDED"
  | "FREE_LOOK_CANCELLED"
  | "POLICY_LAPSED"
  | "POLICY_SURRENDERED";

export const PREMIUM_EVENT_TYPES: PremiumEventType[] = [
  "PREMIUM_COLLECTED",
  "PREMIUM_REFUNDED",
  "FREE_LOOK_CANCELLED",
  "POLICY_LAPSED",
  "POLICY_SURRENDERED",
];

/**
 * One row of the immutable premium event log. A pay period is computed from
 * these events only; nothing is ever edited in place.
 */
export type PremiumEvent = {
  eventId: string;
  type: PremiumEventType;
  policyNo: string;
  /** Selling agent code. */
  agentCode: string;
  productCode: string;
  /** Policy year the premium belongs to (1 = first year). */
  policyYear: number;
  /** Premium amount in VND (always positive). */
  amount: number;
  /** ISO date (YYYY-MM-DD) the event happened. */
  occurredAt: string;
  /** ISO date (YYYY-MM-DD) the policy was issued. */
  issueDate: string;
  paymentMode: PaymentMode;
};

export type Product = {
  code: string;
  name: string;
  /** Contracts of one year or less (riders, health) use a different legal cap. */
  shortTerm: boolean;
  rates: { year1: number; year2: number; year3Plus: number };
  /** Legal maximum first-year commission rate (percent) per Circular 67/2023. */
  legalCapYear1: number;
};

export type OverrideRate = {
  /** Percent of FYC when the manager directly manages the seller. */
  direct: number;
  /** Percent of FYC when the manager is higher up the chain. */
  indirect: number;
};

export type MonthlyBonusTier = {
  minFyc: number;
  minCases: number;
  rate: number;
};

export type ChargebackBand = {
  fromMonth: number;
  toMonth: number;
  /** Percent of previously paid amounts to recover. */
  recoverRate: number;
};

export type CompPlan = {
  planId: string;
  effectiveFrom: string;
  ranks: Rank[];
  products: Product[];
  overrideRates: Record<RankCode, OverrideRate>;
  monthlyBonusTiers: MonthlyBonusTier[];
  chargebackSchedule: ChargebackBand[];
  /** Article 52 caps on bonuses and support, percent of collected premium. */
  bonusCapLongTermPct: number;
  bonusCapShortTermPct: number;
};

export type EarningType = "FYC" | "RYC" | "OVERRIDE" | "BONUS" | "CHARGEBACK";

export type EarningLine = {
  periodId: string;
  agentCode: string;
  type: EarningType;
  /** VND, negative for chargebacks. */
  amount: number;
  ruleId: string;
  sourceEventId: string;
  sourcePolicyNo: string;
  /** Human readable explanation of how the amount was derived. */
  formula: string;
};

/** Amount already paid in an earlier period, used to compute chargebacks. */
export type PaidLine = {
  periodId: string;
  agentCode: string;
  type: EarningType;
  amount: number;
  ruleId: string;
  sourcePolicyNo: string;
};

export type Warning = {
  code: string;
  level: "info" | "warn" | "error";
  message: string;
  agentCode?: string;
  policyNo?: string;
};

export type RunInput = {
  periodId: string;
  agents: Agent[];
  events: PremiumEvent[];
  paidHistory: PaidLine[];
  plan: CompPlan;
};

export type AgentSummary = {
  agent: Agent;
  fyc: number;
  ryc: number;
  override: number;
  bonus: number;
  chargeback: number;
  gross: number;
  cases: number;
  ape: number;
  fyp: number;
};

export type RunOutput = {
  periodId: string;
  lines: EarningLine[];
  warnings: Warning[];
  summaries: AgentSummary[];
  totals: Omit<AgentSummary, "agent">;
  bonusCap: {
    cap: number;
    used: number;
    longTermFyp: number;
    shortTermPremium: number;
  };
};
