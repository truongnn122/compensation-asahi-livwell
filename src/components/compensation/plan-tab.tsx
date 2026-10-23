"use client";

import { useState } from "react";

import { Plus, RotateCcw, Save, X } from "lucide-react";
import { toast } from "sonner";

import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { Checkbox } from "@/components/ui/checkbox";
import { Input } from "@/components/ui/input";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { useDictionary } from "@/hooks/use-dictionary";
import { DEFAULT_PLAN } from "@/lib/compensation/default-plan";
import type { CompPlan } from "@/lib/compensation/types";
import { saveCompensationPlan } from "@/server/compensation-actions";

function NumberCell({
  id,
  value,
  onChange,
  step = 1,
  width = "w-20",
}: {
  id: string;
  value: number;
  onChange: (n: number) => void;
  step?: number;
  width?: string;
}) {
  return (
    <Input
      id={id}
      type="number"
      step={step}
      min={0}
      value={Number.isFinite(value) ? value : 0}
      onChange={e => onChange(Number(e.target.value) || 0)}
      className={`h-8 ${width} text-right tabular-nums`}
    />
  );
}

export function PlanTab({
  plan,
  source,
  onChange,
  onSaved,
}: {
  plan: CompPlan;
  source: "saved" | "default";
  onChange: (next: CompPlan) => void;
  onSaved: (saved: CompPlan) => void;
}) {
  const t = useDictionary();
  const p = t.compensation.plan;
  const [saving, setSaving] = useState(false);

  const update = (patch: Partial<CompPlan>) => onChange({ ...plan, ...patch });

  const save = async () => {
    setSaving(true);
    try {
      const result = await saveCompensationPlan(plan);
      if (!result.ok) {
        toast.error(result.error);
        return;
      }
      onSaved(result.data);
      toast.success(p.saved);
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="flex flex-col gap-4">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="flex flex-wrap items-center gap-2">
          <Badge variant={source === "saved" ? "default" : "outline"}>
            {source === "saved" ? p.sourceSaved : p.sourceDefault}
          </Badge>
          <label className="flex items-center gap-2 text-sm">
            <span className="text-muted-foreground">{p.planId}</span>
            <Input
              id="plan-id"
              value={plan.planId}
              onChange={e => update({ planId: e.target.value })}
              className="h-8 w-32"
            />
          </label>
          <label className="flex items-center gap-2 text-sm">
            <span className="text-muted-foreground">{p.effectiveFrom}</span>
            <Input
              id="plan-effective-from"
              type="date"
              value={plan.effectiveFrom}
              onChange={e => update({ effectiveFrom: e.target.value })}
              className="h-8 w-40"
            />
          </label>
        </div>
        <div className="flex gap-2">
          <Button
            variant="outline"
            onClick={() => onChange(structuredClone(DEFAULT_PLAN))}
          >
            <RotateCcw />
            {p.reset}
          </Button>
          <Button onClick={save} disabled={saving}>
            <Save />
            {saving ? t.common.saving : p.save}
          </Button>
        </div>
      </div>
      <p className="text-muted-foreground text-sm">{p.description}</p>

      <div className="grid gap-4 xl:grid-cols-2">
        <Card className="xl:col-span-2">
          <CardHeader>
            <CardTitle className="text-base">{p.products}</CardTitle>
          </CardHeader>
          <CardContent className="overflow-x-auto">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>{p.productCols.code}</TableHead>
                  <TableHead>{p.productCols.name}</TableHead>
                  <TableHead>{p.productCols.shortTerm}</TableHead>
                  <TableHead className="text-right">
                    {p.productCols.year1}
                  </TableHead>
                  <TableHead className="text-right">
                    {p.productCols.year2}
                  </TableHead>
                  <TableHead className="text-right">
                    {p.productCols.year3}
                  </TableHead>
                  <TableHead className="text-right">
                    {p.productCols.cap}
                  </TableHead>
                  <TableHead />
                </TableRow>
              </TableHeader>
              <TableBody>
                {plan.products.map((prod, i) => {
                  const set = (patch: Partial<typeof prod>) =>
                    update({
                      products: plan.products.map((x, j) =>
                        j === i ? { ...x, ...patch } : x
                      ),
                    });
                  const over = prod.rates.year1 > prod.legalCapYear1;
                  return (
                    <TableRow
                      key={i}
                      className={over ? "bg-destructive/5" : ""}
                    >
                      <TableCell>
                        <Input
                          id={`prod-code-${i}`}
                          value={prod.code}
                          onChange={e =>
                            set({ code: e.target.value.toUpperCase() })
                          }
                          className="h-8 w-24 font-mono"
                        />
                      </TableCell>
                      <TableCell>
                        <Input
                          id={`prod-name-${i}`}
                          value={prod.name}
                          onChange={e => set({ name: e.target.value })}
                          className="h-8 min-w-40"
                        />
                      </TableCell>
                      <TableCell>
                        <Checkbox
                          id={`prod-short-${i}`}
                          checked={prod.shortTerm}
                          onCheckedChange={v => set({ shortTerm: v === true })}
                        />
                      </TableCell>
                      <TableCell className="text-right">
                        <NumberCell
                          id={`prod-y1-${i}`}
                          value={prod.rates.year1}
                          onChange={n =>
                            set({ rates: { ...prod.rates, year1: n } })
                          }
                        />
                      </TableCell>
                      <TableCell className="text-right">
                        <NumberCell
                          id={`prod-y2-${i}`}
                          value={prod.rates.year2}
                          onChange={n =>
                            set({ rates: { ...prod.rates, year2: n } })
                          }
                        />
                      </TableCell>
                      <TableCell className="text-right">
                        <NumberCell
                          id={`prod-y3-${i}`}
                          value={prod.rates.year3Plus}
                          onChange={n =>
                            set({ rates: { ...prod.rates, year3Plus: n } })
                          }
                        />
                      </TableCell>
                      <TableCell className="text-right">
                        <NumberCell
                          id={`prod-cap-${i}`}
                          value={prod.legalCapYear1}
                          onChange={n => set({ legalCapYear1: n })}
                        />
                      </TableCell>
                      <TableCell>
                        <Button
                          variant="ghost"
                          size="icon-sm"
                          aria-label={p.remove}
                          onClick={() =>
                            update({
                              products: plan.products.filter((_, j) => j !== i),
                            })
                          }
                        >
                          <X />
                        </Button>
                      </TableCell>
                    </TableRow>
                  );
                })}
              </TableBody>
            </Table>
            <Button
              variant="outline"
              size="sm"
              className="mt-3"
              onClick={() =>
                update({
                  products: [
                    ...plan.products,
                    {
                      code: "",
                      name: "",
                      shortTerm: false,
                      rates: { year1: 0, year2: 0, year3Plus: 0 },
                      legalCapYear1: 40,
                    },
                  ],
                })
              }
            >
              <Plus />
              {p.addProduct}
            </Button>
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardTitle className="text-base">{p.ranks}</CardTitle>
          </CardHeader>
          <CardContent className="overflow-x-auto">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>{p.rankCols.code}</TableHead>
                  <TableHead>{p.rankCols.name}</TableHead>
                  <TableHead className="text-right">
                    {p.rankCols.level}
                  </TableHead>
                  <TableHead />
                </TableRow>
              </TableHeader>
              <TableBody>
                {plan.ranks.map((rank, i) => {
                  const set = (patch: Partial<typeof rank>) =>
                    update({
                      ranks: plan.ranks.map((x, j) =>
                        j === i ? { ...x, ...patch } : x
                      ),
                    });
                  return (
                    <TableRow key={i}>
                      <TableCell>
                        <Input
                          id={`rank-code-${i}`}
                          value={rank.code}
                          onChange={e =>
                            set({ code: e.target.value.toUpperCase() })
                          }
                          className="h-8 w-20 font-mono"
                        />
                      </TableCell>
                      <TableCell>
                        <Input
                          id={`rank-name-${i}`}
                          value={rank.name}
                          onChange={e => set({ name: e.target.value })}
                          className="h-8 min-w-36"
                        />
                      </TableCell>
                      <TableCell className="text-right">
                        <NumberCell
                          id={`rank-level-${i}`}
                          value={rank.level}
                          onChange={n => set({ level: n })}
                          width="w-16"
                        />
                      </TableCell>
                      <TableCell>
                        <Button
                          variant="ghost"
                          size="icon-sm"
                          aria-label={p.remove}
                          onClick={() =>
                            update({
                              ranks: plan.ranks.filter((_, j) => j !== i),
                            })
                          }
                        >
                          <X />
                        </Button>
                      </TableCell>
                    </TableRow>
                  );
                })}
              </TableBody>
            </Table>
            <Button
              variant="outline"
              size="sm"
              className="mt-3"
              onClick={() =>
                update({
                  ranks: [
                    ...plan.ranks,
                    { code: "", name: "", level: plan.ranks.length + 1 },
                  ],
                })
              }
            >
              <Plus />
              {p.addRank}
            </Button>
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardTitle className="text-base">{p.overrides}</CardTitle>
          </CardHeader>
          <CardContent className="overflow-x-auto">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>{p.overrideCols.rank}</TableHead>
                  <TableHead className="text-right">
                    {p.overrideCols.direct}
                  </TableHead>
                  <TableHead className="text-right">
                    {p.overrideCols.indirect}
                  </TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {plan.ranks
                  .filter(r => r.level > 1)
                  .map(rank => {
                    const rates = plan.overrideRates[rank.code] ?? {
                      direct: 0,
                      indirect: 0,
                    };
                    const set = (patch: Partial<typeof rates>) =>
                      update({
                        overrideRates: {
                          ...plan.overrideRates,
                          [rank.code]: { ...rates, ...patch },
                        },
                      });
                    return (
                      <TableRow key={rank.code}>
                        <TableCell>
                          {rank.name}{" "}
                          <span className="text-muted-foreground font-mono text-xs">
                            {rank.code}
                          </span>
                        </TableCell>
                        <TableCell className="text-right">
                          <NumberCell
                            id={`ovr-direct-${rank.code}`}
                            value={rates.direct}
                            onChange={n => set({ direct: n })}
                          />
                        </TableCell>
                        <TableCell className="text-right">
                          <NumberCell
                            id={`ovr-indirect-${rank.code}`}
                            value={rates.indirect}
                            onChange={n => set({ indirect: n })}
                          />
                        </TableCell>
                      </TableRow>
                    );
                  })}
              </TableBody>
            </Table>
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardTitle className="text-base">{p.bonus}</CardTitle>
            <CardDescription>{p.bonusHint}</CardDescription>
          </CardHeader>
          <CardContent className="overflow-x-auto">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead className="text-right">
                    {p.bonusCols.minFyc}
                  </TableHead>
                  <TableHead className="text-right">
                    {p.bonusCols.minCases}
                  </TableHead>
                  <TableHead className="text-right">
                    {p.bonusCols.rate}
                  </TableHead>
                  <TableHead />
                </TableRow>
              </TableHeader>
              <TableBody>
                {plan.monthlyBonusTiers.map((tier, i) => {
                  const set = (patch: Partial<typeof tier>) =>
                    update({
                      monthlyBonusTiers: plan.monthlyBonusTiers.map((x, j) =>
                        j === i ? { ...x, ...patch } : x
                      ),
                    });
                  return (
                    <TableRow key={i}>
                      <TableCell className="text-right">
                        <NumberCell
                          id={`tier-fyc-${i}`}
                          value={tier.minFyc}
                          step={1_000_000}
                          width="w-36"
                          onChange={n => set({ minFyc: n })}
                        />
                      </TableCell>
                      <TableCell className="text-right">
                        <NumberCell
                          id={`tier-cases-${i}`}
                          value={tier.minCases}
                          width="w-16"
                          onChange={n => set({ minCases: n })}
                        />
                      </TableCell>
                      <TableCell className="text-right">
                        <NumberCell
                          id={`tier-rate-${i}`}
                          value={tier.rate}
                          onChange={n => set({ rate: n })}
                        />
                      </TableCell>
                      <TableCell>
                        <Button
                          variant="ghost"
                          size="icon-sm"
                          aria-label={p.remove}
                          onClick={() =>
                            update({
                              monthlyBonusTiers: plan.monthlyBonusTiers.filter(
                                (_, j) => j !== i
                              ),
                            })
                          }
                        >
                          <X />
                        </Button>
                      </TableCell>
                    </TableRow>
                  );
                })}
              </TableBody>
            </Table>
            <Button
              variant="outline"
              size="sm"
              className="mt-3"
              onClick={() =>
                update({
                  monthlyBonusTiers: [
                    ...plan.monthlyBonusTiers,
                    { minFyc: 0, minCases: 1, rate: 0 },
                  ],
                })
              }
            >
              <Plus />
              {p.addTier}
            </Button>
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardTitle className="text-base">{p.chargeback}</CardTitle>
          </CardHeader>
          <CardContent className="overflow-x-auto">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead className="text-right">
                    {p.chargebackCols.from}
                  </TableHead>
                  <TableHead className="text-right">
                    {p.chargebackCols.to}
                  </TableHead>
                  <TableHead className="text-right">
                    {p.chargebackCols.rate}
                  </TableHead>
                  <TableHead />
                </TableRow>
              </TableHeader>
              <TableBody>
                {plan.chargebackSchedule.map((band, i) => {
                  const set = (patch: Partial<typeof band>) =>
                    update({
                      chargebackSchedule: plan.chargebackSchedule.map((x, j) =>
                        j === i ? { ...x, ...patch } : x
                      ),
                    });
                  return (
                    <TableRow key={i}>
                      <TableCell className="text-right">
                        <NumberCell
                          id={`cb-from-${i}`}
                          value={band.fromMonth}
                          onChange={n => set({ fromMonth: n })}
                        />
                      </TableCell>
                      <TableCell className="text-right">
                        <NumberCell
                          id={`cb-to-${i}`}
                          value={band.toMonth}
                          onChange={n => set({ toMonth: n })}
                        />
                      </TableCell>
                      <TableCell className="text-right">
                        <NumberCell
                          id={`cb-rate-${i}`}
                          value={band.recoverRate}
                          onChange={n => set({ recoverRate: n })}
                        />
                      </TableCell>
                      <TableCell>
                        <Button
                          variant="ghost"
                          size="icon-sm"
                          aria-label={p.remove}
                          onClick={() =>
                            update({
                              chargebackSchedule:
                                plan.chargebackSchedule.filter(
                                  (_, j) => j !== i
                                ),
                            })
                          }
                        >
                          <X />
                        </Button>
                      </TableCell>
                    </TableRow>
                  );
                })}
              </TableBody>
            </Table>
            <Button
              variant="outline"
              size="sm"
              className="mt-3"
              onClick={() =>
                update({
                  chargebackSchedule: [
                    ...plan.chargebackSchedule,
                    { fromMonth: 25, toMonth: 36, recoverRate: 0 },
                  ],
                })
              }
            >
              <Plus />
              {p.addBand}
            </Button>
          </CardContent>
        </Card>

        <Card className="xl:col-span-2">
          <CardHeader>
            <CardTitle className="text-base">{p.caps}</CardTitle>
          </CardHeader>
          <CardContent className="flex flex-wrap gap-6">
            <label className="flex items-center gap-3 text-sm">
              <NumberCell
                id="cap-long"
                value={plan.bonusCapLongTermPct}
                onChange={n => update({ bonusCapLongTermPct: n })}
              />
              <span>{p.capLong}</span>
            </label>
            <label className="flex items-center gap-3 text-sm">
              <NumberCell
                id="cap-short"
                value={plan.bonusCapShortTermPct}
                onChange={n => update({ bonusCapShortTermPct: n })}
              />
              <span>{p.capShort}</span>
            </label>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
