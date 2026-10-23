"use client";

import { useState } from "react";

import { AlertTriangle, Calculator, Eye } from "lucide-react";

import { EmptyState } from "@/components/empty-state";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import {
  Sheet,
  SheetContent,
  SheetDescription,
  SheetHeader,
  SheetTitle,
} from "@/components/ui/sheet";
import {
  Table,
  TableBody,
  TableCell,
  TableFooter,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { useDictionary } from "@/hooks/use-dictionary";
import { EARNING_TYPE_LABEL } from "@/lib/compensation/excel";
import type {
  AgentSummary,
  CompPlan,
  RunOutput,
} from "@/lib/compensation/types";
import { cn } from "@/lib/utils";

import { fill, vnd } from "./format";

function Kpi({
  label,
  value,
  tone,
}: {
  label: string;
  value: string;
  tone?: "neg";
}) {
  return (
    <Card className="gap-1 py-4">
      <CardHeader className="px-4">
        <CardDescription className="text-xs tracking-wide uppercase">
          {label}
        </CardDescription>
        <CardTitle
          className={cn(
            "text-xl tabular-nums",
            tone === "neg" && "text-destructive"
          )}
        >
          {value}
        </CardTitle>
      </CardHeader>
    </Card>
  );
}

export function ResultsTab({
  result,
  plan,
}: {
  result: RunOutput | null;
  plan: CompPlan;
}) {
  const t = useDictionary();
  const r = t.compensation.results;
  const [selected, setSelected] = useState<AgentSummary | null>(null);

  if (!result) {
    return (
      <EmptyState icon={Calculator} title={r.empty} description={r.emptyHint} />
    );
  }

  const rankName = new Map(plan.ranks.map(x => [x.code, x.name]));
  const { totals, bonusCap } = result;
  const capPct = bonusCap.cap > 0 ? (bonusCap.used / bonusCap.cap) * 100 : 0;
  const lines = selected
    ? result.lines.filter(l => l.agentCode === selected.agent.code)
    : [];
  const money = (n: number) => (
    <span className={cn("tabular-nums", n < 0 && "text-destructive")}>
      {vnd(n)}
    </span>
  );

  return (
    <div className="flex flex-col gap-4">
      <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-6">
        <Kpi label={r.kpi.gross} value={vnd(totals.gross)} />
        <Kpi label={r.kpi.fyc} value={vnd(totals.fyc)} />
        <Kpi label={r.kpi.override} value={vnd(totals.override)} />
        <Kpi label={r.kpi.bonus} value={vnd(totals.bonus)} />
        <Kpi
          label={r.kpi.chargeback}
          value={vnd(totals.chargeback)}
          tone={totals.chargeback < 0 ? "neg" : undefined}
        />
        <Kpi
          label={r.kpi.agents}
          value={String(result.summaries.filter(s => s.gross !== 0).length)}
        />
      </div>

      <p className="text-muted-foreground text-sm">
        {fill(r.capUsed, {
          used: vnd(bonusCap.used),
          cap: vnd(bonusCap.cap),
          pct: capPct.toFixed(1),
        })}
      </p>

      {result.warnings.length > 0 && (
        <Card className="border-amber-500/40">
          <CardHeader>
            <CardTitle className="flex items-center gap-2 text-base">
              <AlertTriangle className="size-4 text-amber-600" />
              {r.warnings}
            </CardTitle>
          </CardHeader>
          <CardContent>
            <ul className="space-y-1 text-sm">
              {result.warnings.map((w, i) => (
                <li key={i} className="flex gap-2">
                  <Badge
                    variant={w.level === "error" ? "destructive" : "secondary"}
                    className="shrink-0"
                  >
                    {w.code}
                  </Badge>
                  <span>{w.message}</span>
                </li>
              ))}
            </ul>
          </CardContent>
        </Card>
      )}

      <Card>
        <CardHeader>
          <CardTitle className="text-base">{r.summaryTitle}</CardTitle>
          <CardDescription>
            {t.compensation.period}: {result.periodId}
          </CardDescription>
        </CardHeader>
        <CardContent className="overflow-x-auto">
          <Table>
            <TableHeader>
              <TableRow>
                <TableHead>{r.cols.agent}</TableHead>
                <TableHead>{r.cols.rank}</TableHead>
                <TableHead className="text-right">{r.cols.cases}</TableHead>
                <TableHead className="text-right">{r.cols.fyp}</TableHead>
                <TableHead className="text-right">{r.cols.fyc}</TableHead>
                <TableHead className="text-right">{r.cols.ryc}</TableHead>
                <TableHead className="text-right">{r.cols.override}</TableHead>
                <TableHead className="text-right">{r.cols.bonus}</TableHead>
                <TableHead className="text-right">
                  {r.cols.chargeback}
                </TableHead>
                <TableHead className="text-right">{r.cols.gross}</TableHead>
                <TableHead />
              </TableRow>
            </TableHeader>
            <TableBody>
              {result.summaries.map(s => (
                <TableRow key={s.agent.code}>
                  <TableCell>
                    <div className="font-medium">{s.agent.name}</div>
                    <div className="text-muted-foreground font-mono text-xs">
                      {s.agent.code}
                    </div>
                  </TableCell>
                  <TableCell>
                    <Badge variant="secondary">
                      {rankName.get(s.agent.rank) ?? s.agent.rank}
                    </Badge>
                  </TableCell>
                  <TableCell className="text-right tabular-nums">
                    {s.cases}
                  </TableCell>
                  <TableCell className="text-right">{money(s.fyp)}</TableCell>
                  <TableCell className="text-right">{money(s.fyc)}</TableCell>
                  <TableCell className="text-right">{money(s.ryc)}</TableCell>
                  <TableCell className="text-right">
                    {money(s.override)}
                  </TableCell>
                  <TableCell className="text-right">{money(s.bonus)}</TableCell>
                  <TableCell className="text-right">
                    {money(s.chargeback)}
                  </TableCell>
                  <TableCell className="text-right font-semibold">
                    {money(s.gross)}
                  </TableCell>
                  <TableCell>
                    <Button
                      variant="ghost"
                      size="sm"
                      onClick={() => setSelected(s)}
                    >
                      <Eye />
                      {r.cols.detail}
                    </Button>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
            <TableFooter>
              <TableRow>
                <TableCell colSpan={2}>{r.total}</TableCell>
                <TableCell className="text-right tabular-nums">
                  {totals.cases}
                </TableCell>
                <TableCell className="text-right">
                  {money(totals.fyp)}
                </TableCell>
                <TableCell className="text-right">
                  {money(totals.fyc)}
                </TableCell>
                <TableCell className="text-right">
                  {money(totals.ryc)}
                </TableCell>
                <TableCell className="text-right">
                  {money(totals.override)}
                </TableCell>
                <TableCell className="text-right">
                  {money(totals.bonus)}
                </TableCell>
                <TableCell className="text-right">
                  {money(totals.chargeback)}
                </TableCell>
                <TableCell className="text-right font-semibold">
                  {money(totals.gross)}
                </TableCell>
                <TableCell />
              </TableRow>
            </TableFooter>
          </Table>
        </CardContent>
      </Card>

      <Sheet
        open={selected !== null}
        onOpenChange={o => !o && setSelected(null)}
      >
        <SheetContent className="w-full overflow-y-auto sm:max-w-2xl">
          {selected && (
            <>
              <SheetHeader>
                <SheetTitle>
                  {r.statementTitle} · {selected.agent.name}
                </SheetTitle>
                <SheetDescription>
                  {rankName.get(selected.agent.rank) ?? selected.agent.rank} ·{" "}
                  {t.compensation.period} {result.periodId} · {r.statementHint}
                </SheetDescription>
              </SheetHeader>
              <div className="px-4 pb-6">
                {lines.length === 0 ? (
                  <p className="text-muted-foreground text-sm">{r.noLines}</p>
                ) : (
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead>{r.lineCols.type}</TableHead>
                        <TableHead>{r.lineCols.policy}</TableHead>
                        <TableHead>{r.lineCols.formula}</TableHead>
                        <TableHead className="text-right">
                          {r.lineCols.amount}
                        </TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {lines.map((l, i) => (
                        <TableRow key={i}>
                          <TableCell>
                            <div>{EARNING_TYPE_LABEL[l.type]}</div>
                            <div className="text-muted-foreground font-mono text-xs">
                              {l.ruleId}
                            </div>
                          </TableCell>
                          <TableCell className="font-mono text-xs">
                            {l.sourcePolicyNo || "—"}
                          </TableCell>
                          <TableCell className="text-xs">{l.formula}</TableCell>
                          <TableCell className="text-right">
                            {money(l.amount)}
                          </TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                    <TableFooter>
                      <TableRow>
                        <TableCell colSpan={3}>{r.total}</TableCell>
                        <TableCell className="text-right font-semibold">
                          {money(selected.gross)}
                        </TableCell>
                      </TableRow>
                    </TableFooter>
                  </Table>
                )}
              </div>
            </>
          )}
        </SheetContent>
      </Sheet>
    </div>
  );
}
