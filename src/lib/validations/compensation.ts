import { z } from "zod";

const pct = z.number().min(0).max(100);

export const compPlanSchema = z.object({
  planId: z.string().min(1),
  effectiveFrom: z.string().min(1),
  ranks: z
    .array(
      z.object({
        code: z.string().min(1),
        name: z.string().min(1),
        level: z.number().int().min(1),
      })
    )
    .min(1),
  products: z
    .array(
      z.object({
        code: z.string().min(1),
        name: z.string().min(1),
        shortTerm: z.boolean(),
        rates: z.object({ year1: pct, year2: pct, year3Plus: pct }),
        legalCapYear1: pct,
      })
    )
    .min(1),
  overrideRates: z.record(z.string(), z.object({ direct: pct, indirect: pct })),
  monthlyBonusTiers: z.array(
    z.object({
      minFyc: z.number().min(0),
      minCases: z.number().int().min(0),
      rate: pct,
    })
  ),
  chargebackSchedule: z.array(
    z.object({
      fromMonth: z.number().int().min(1),
      toMonth: z.number().int().min(1),
      recoverRate: pct,
    })
  ),
  bonusCapLongTermPct: pct,
  bonusCapShortTermPct: pct,
});

export type CompPlanInput = z.infer<typeof compPlanSchema>;
