import {
  APE_FACTOR,
  type Agent,
  type AgentSummary,
  type CompPlan,
  type EarningLine,
  type PremiumEvent,
  type RunInput,
  type RunOutput,
  type Warning,
} from "./types";

export const RULE_IDS = {
  fyc: "R-FYC",
  ryc: "R-RYC",
  override: "R-OVR",
  monthlyBonus: "R-B-M1",
  chargeback: "R-CB",
  capArticle51: "R-CAP-51",
  capArticle52: "R-CAP-52",
} as const;

const CHARGEBACK_EVENT_TYPES = new Set([
  "FREE_LOOK_CANCELLED",
  "POLICY_LAPSED",
  "POLICY_SURRENDERED",
  "PREMIUM_REFUNDED",
]);

export function formatVnd(amount: number): string {
  return Math.round(amount).toLocaleString("vi-VN");
}

/** Inclusive month count between two YYYY-MM(-DD) dates: same month = 1. */
export function policyMonth(issueDate: string, atDate: string): number {
  const [y1, m1] = issueDate.split("-").map(Number);
  const [y2, m2] = atDate.split("-").map(Number);
  return (y2 - y1) * 12 + (m2 - m1) + 1;
}

/** Walks up the hierarchy from `code`, nearest manager first. Cycles stop the walk. */
export function managersOf(agents: Agent[], code: string): Agent[] {
  const byCode = new Map(agents.map(a => [a.code, a]));
  const chain: Agent[] = [];
  const seen = new Set<string>([code]);
  let current = byCode.get(code);
  while (current && current.managerCode) {
    const manager = byCode.get(current.managerCode);
    if (!manager || seen.has(manager.code)) break;
    seen.add(manager.code);
    chain.push(manager);
    current = manager;
  }
  return chain;
}

function rateForYear(
  plan: CompPlan,
  productCode: string,
  policyYear: number
): { rate: number; cap: number; name: string } | null {
  const product = plan.products.find(p => p.code === productCode);
  if (!product) return null;
  const rate =
    policyYear <= 1
      ? product.rates.year1
      : policyYear === 2
        ? product.rates.year2
        : product.rates.year3Plus;
  return { rate, cap: product.legalCapYear1, name: product.name };
}

/**
 * Pure period calculation. No I/O: everything it needs arrives in `input`
 * and every VND it produces carries the rule and event that created it.
 */
export function runPeriod(input: RunInput): RunOutput {
  const { periodId, agents, events, paidHistory, plan } = input;
  const lines: EarningLine[] = [];
  const warnings: Warning[] = [];
  const byCode = new Map(agents.map(a => [a.code, a]));
  const rankByCode = new Map(plan.ranks.map(r => [r.code, r]));

  const push = (line: Omit<EarningLine, "periodId">) =>
    lines.push({ periodId, ...line });

  // 1) First-year and renewal commission on collected premium.
  for (const ev of events) {
    if (ev.type !== "PREMIUM_COLLECTED") continue;
    const seller = byCode.get(ev.agentCode);
    if (!seller) {
      warnings.push({
        code: "AGENT_NOT_FOUND",
        level: "error",
        message: `Hợp đồng ${ev.policyNo}: không tìm thấy đại lý ${ev.agentCode}.`,
        policyNo: ev.policyNo,
      });
      continue;
    }
    const r = rateForYear(plan, ev.productCode, ev.policyYear);
    if (!r) {
      warnings.push({
        code: "PRODUCT_NOT_FOUND",
        level: "error",
        message: `Hợp đồng ${ev.policyNo}: sản phẩm ${ev.productCode} chưa có trong chính sách.`,
        policyNo: ev.policyNo,
      });
      continue;
    }

    const isFirstYear = ev.policyYear <= 1;
    if (isFirstYear && r.rate > r.cap) {
      warnings.push({
        code: RULE_IDS.capArticle51,
        level: "error",
        message: `${r.name}: tỷ lệ hoa hồng năm 1 ${r.rate}% vượt trần ${r.cap}% (Điều 51 TT 67/2023). Hoa hồng của HĐ ${ev.policyNo} bị chặn.`,
        policyNo: ev.policyNo,
        agentCode: ev.agentCode,
      });
      continue;
    }

    const amount = (ev.amount * r.rate) / 100;
    push({
      agentCode: ev.agentCode,
      type: isFirstYear ? "FYC" : "RYC",
      amount,
      ruleId: isFirstYear ? RULE_IDS.fyc : RULE_IDS.ryc,
      sourceEventId: ev.eventId,
      sourcePolicyNo: ev.policyNo,
      formula: `${formatVnd(ev.amount)} × ${r.rate}% (${r.name}, năm HĐ ${ev.policyYear})`,
    });

    // 2) Override for every manager in the chain, using the hierarchy as of the event.
    if (!isFirstYear) continue;
    const chain = managersOf(agents, ev.agentCode);
    chain.forEach((manager, depth) => {
      const rates = plan.overrideRates[manager.rank];
      if (!rates) return;
      const pct = depth === 0 ? rates.direct : rates.indirect;
      if (pct <= 0) return;
      const rankName = rankByCode.get(manager.rank)?.name ?? manager.rank;
      push({
        agentCode: manager.code,
        type: "OVERRIDE",
        amount: (amount * pct) / 100,
        ruleId: RULE_IDS.override,
        sourceEventId: ev.eventId,
        sourcePolicyNo: ev.policyNo,
        formula: `FYC ${formatVnd(amount)} của ${seller.name} × ${pct}% (${rankName}, ${depth === 0 ? "trực tiếp" : "gián tiếp"})`,
      });
    });
  }

  // 3) Chargebacks for cancellations, lapses and refunds.
  for (const ev of events) {
    if (!CHARGEBACK_EVENT_TYPES.has(ev.type)) continue;
    const month = policyMonth(ev.issueDate, ev.occurredAt);
    const recoverRate =
      ev.type === "FREE_LOOK_CANCELLED"
        ? 100
        : (plan.chargebackSchedule.find(
            b => month >= b.fromMonth && month <= b.toMonth
          )?.recoverRate ?? 0);
    const paid = paidHistory.filter(p => p.sourcePolicyNo === ev.policyNo);
    if (paid.length === 0) {
      warnings.push({
        code: "NO_PAID_HISTORY",
        level: "warn",
        message: `HĐ ${ev.policyNo} (${eventLabel(ev.type)}): không có khoản đã trả trong lịch sử để truy thu.`,
        policyNo: ev.policyNo,
      });
      continue;
    }
    if (recoverRate <= 0) continue;
    for (const p of paid) {
      push({
        agentCode: p.agentCode,
        type: "CHARGEBACK",
        amount: (-p.amount * recoverRate) / 100,
        ruleId: RULE_IDS.chargeback,
        sourceEventId: ev.eventId,
        sourcePolicyNo: ev.policyNo,
        formula: `Đã trả ${formatVnd(p.amount)} (${p.type}, kỳ ${p.periodId}) × ${recoverRate}% thu hồi, ${eventLabel(ev.type)} ở tháng HĐ thứ ${month}`,
      });
    }
  }

  // 4) Monthly personal bonus on FYC tiers.
  const tiers = [...plan.monthlyBonusTiers].sort((a, b) => b.minFyc - a.minFyc);
  for (const agent of agents) {
    const mine = lines.filter(
      l => l.agentCode === agent.code && l.type === "FYC" && l.amount > 0
    );
    if (mine.length === 0) continue;
    const fyc = mine.reduce((s, l) => s + l.amount, 0);
    const cases = new Set(mine.map(l => l.sourcePolicyNo)).size;
    const tier = tiers.find(t => fyc >= t.minFyc && cases >= t.minCases);
    if (!tier) continue;
    push({
      agentCode: agent.code,
      type: "BONUS",
      amount: (fyc * tier.rate) / 100,
      ruleId: RULE_IDS.monthlyBonus,
      sourceEventId: "AGG",
      sourcePolicyNo: "",
      formula: `FYC tháng ${formatVnd(fyc)} (${cases} HĐ) đạt bậc ≥ ${formatVnd(tier.minFyc)} → × ${tier.rate}%`,
    });
  }

  // 5) Article 52 cap on bonuses (checked on this period's collected premium).
  const shortTermCodes = new Set(
    plan.products.filter(p => p.shortTerm).map(p => p.code)
  );
  let longTermFyp = 0;
  let shortTermPremium = 0;
  for (const ev of events) {
    if (ev.type !== "PREMIUM_COLLECTED") continue;
    if (shortTermCodes.has(ev.productCode)) shortTermPremium += ev.amount;
    else if (ev.policyYear <= 1) longTermFyp += ev.amount;
  }
  const cap =
    (longTermFyp * plan.bonusCapLongTermPct) / 100 +
    (shortTermPremium * plan.bonusCapShortTermPct) / 100;
  const used = lines
    .filter(l => l.type === "BONUS")
    .reduce((s, l) => s + l.amount, 0);
  if (cap > 0 && used > cap) {
    warnings.push({
      code: RULE_IDS.capArticle52,
      level: "error",
      message: `Tổng thưởng ${formatVnd(used)} vượt trần Điều 52 (${formatVnd(cap)}).`,
    });
  } else if (cap > 0 && used >= cap * 0.9) {
    warnings.push({
      code: RULE_IDS.capArticle52,
      level: "warn",
      message: `Tổng thưởng đã dùng ${((used / cap) * 100).toFixed(0)}% trần Điều 52.`,
    });
  }

  // 6) Summaries per agent and totals.
  const summaries = summarize(agents, events, lines);
  const totals = summaries.reduce<Omit<AgentSummary, "agent">>(
    (t, s) => ({
      fyc: t.fyc + s.fyc,
      ryc: t.ryc + s.ryc,
      override: t.override + s.override,
      bonus: t.bonus + s.bonus,
      chargeback: t.chargeback + s.chargeback,
      gross: t.gross + s.gross,
      cases: t.cases + s.cases,
      ape: t.ape + s.ape,
      fyp: t.fyp + s.fyp,
    }),
    {
      fyc: 0,
      ryc: 0,
      override: 0,
      bonus: 0,
      chargeback: 0,
      gross: 0,
      cases: 0,
      ape: 0,
      fyp: 0,
    }
  );

  return {
    periodId,
    lines,
    warnings,
    summaries,
    totals,
    bonusCap: { cap, used, longTermFyp, shortTermPremium },
  };
}

export function summarize(
  agents: Agent[],
  events: PremiumEvent[],
  lines: EarningLine[]
): AgentSummary[] {
  const map = new Map<string, AgentSummary>(
    agents.map(agent => [
      agent.code,
      {
        agent,
        fyc: 0,
        ryc: 0,
        override: 0,
        bonus: 0,
        chargeback: 0,
        gross: 0,
        cases: 0,
        ape: 0,
        fyp: 0,
      },
    ])
  );
  for (const line of lines) {
    const s = map.get(line.agentCode);
    if (!s) continue;
    if (line.type === "FYC") s.fyc += line.amount;
    else if (line.type === "RYC") s.ryc += line.amount;
    else if (line.type === "OVERRIDE") s.override += line.amount;
    else if (line.type === "BONUS") s.bonus += line.amount;
    else if (line.type === "CHARGEBACK") s.chargeback += line.amount;
    s.gross += line.amount;
  }
  for (const ev of events) {
    if (ev.type !== "PREMIUM_COLLECTED" || ev.policyYear > 1) continue;
    const s = map.get(ev.agentCode);
    if (!s) continue;
    s.cases += 1;
    s.fyp += ev.amount;
    s.ape += ev.amount * APE_FACTOR[ev.paymentMode];
  }
  return [...map.values()];
}

export function eventLabel(type: PremiumEvent["type"]): string {
  switch (type) {
    case "PREMIUM_COLLECTED":
      return "Thu phí";
    case "PREMIUM_REFUNDED":
      return "Hoàn phí";
    case "FREE_LOOK_CANCELLED":
      return "Hủy trong thời gian cân nhắc";
    case "POLICY_LAPSED":
      return "Mất hiệu lực";
    case "POLICY_SURRENDERED":
      return "Hủy hợp đồng";
  }
}
