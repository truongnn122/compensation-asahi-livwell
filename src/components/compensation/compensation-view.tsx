"use client";

import { useMemo, useState } from "react";

import { Calculator, Download } from "lucide-react";
import { toast } from "sonner";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { useDictionary } from "@/hooks/use-dictionary";
import { runPeriod } from "@/lib/compensation/engine";
import { buildResultWorkbook, downloadBlob } from "@/lib/compensation/excel";
import type {
  Agent,
  CompPlan,
  PaidLine,
  PremiumEvent,
  RunOutput,
} from "@/lib/compensation/types";

import { DataTab } from "./data-tab";
import { fill, currentPeriod } from "./format";
import { PlanTab } from "./plan-tab";
import { ResultsTab } from "./results-tab";

export type Dataset = {
  agents: Agent[];
  events: PremiumEvent[];
  paidHistory: PaidLine[];
};

const EMPTY: Dataset = { agents: [], events: [], paidHistory: [] };

export function CompensationView({
  initialPlan,
  planSource,
}: {
  initialPlan: CompPlan;
  planSource: "saved" | "default";
}) {
  const t = useDictionary();
  const c = t.compensation;
  const [plan, setPlan] = useState<CompPlan>(initialPlan);
  const [source, setSource] = useState(planSource);
  const [dataset, setDataset] = useState<Dataset>(EMPTY);
  const [periodId, setPeriodId] = useState(currentPeriod());
  const [result, setResult] = useState<RunOutput | null>(null);
  const [tab, setTab] = useState("data");
  const [exporting, setExporting] = useState(false);

  const hasData = dataset.agents.length > 0 && dataset.events.length > 0;

  const periodEvents = useMemo(
    () => dataset.events.filter(e => e.occurredAt.startsWith(periodId)),
    [dataset.events, periodId]
  );

  const calculate = () => {
    if (!hasData) {
      toast.error(c.results.needData);
      setTab("data");
      return;
    }
    const output = runPeriod({
      periodId,
      agents: dataset.agents,
      events: periodEvents.length > 0 ? periodEvents : dataset.events,
      paidHistory: dataset.paidHistory,
      plan,
    });
    setResult(output);
    setTab("results");
    toast.success(
      fill(c.results.calculated, {
        lines: output.lines.length,
        agents: output.summaries.filter(s => s.gross !== 0).length,
      })
    );
  };

  const exportResults = async () => {
    if (!result) return;
    setExporting(true);
    try {
      const blob = await buildResultWorkbook(result, plan);
      downloadBlob(blob, `thu-nhap-dai-ly-${result.periodId}.xlsx`);
    } finally {
      setExporting(false);
    }
  };

  return (
    <div className="flex flex-col gap-4">
      <div className="flex flex-wrap items-end justify-between gap-3">
        <div>
          <h1 className="text-2xl font-semibold">{c.heading}</h1>
          <p className="text-muted-foreground text-sm">{c.subtitle}</p>
        </div>
        <div className="flex flex-wrap items-end gap-2">
          <label className="flex flex-col gap-1 text-xs">
            <span className="text-muted-foreground">{c.period}</span>
            <Input
              id="compensation-period"
              type="month"
              value={periodId}
              onChange={e => setPeriodId(e.target.value)}
              className="h-8 w-40"
            />
          </label>
          <Button onClick={calculate} disabled={!hasData}>
            <Calculator />
            {c.calculate}
          </Button>
          <Button
            variant="outline"
            onClick={exportResults}
            disabled={!result || exporting}
          >
            <Download />
            {c.exportExcel}
          </Button>
        </div>
      </div>

      <Tabs value={tab} onValueChange={setTab}>
        <TabsList>
          <TabsTrigger value="data">
            {c.tabs.data}
            {hasData && (
              <span className="text-muted-foreground ml-1 text-xs">
                ({dataset.events.length})
              </span>
            )}
          </TabsTrigger>
          <TabsTrigger value="plan">{c.tabs.plan}</TabsTrigger>
          <TabsTrigger value="results">{c.tabs.results}</TabsTrigger>
        </TabsList>
        <TabsContent value="data">
          <DataTab
            plan={plan}
            dataset={dataset}
            onChange={next => {
              setDataset(next);
              setResult(null);
            }}
          />
        </TabsContent>
        <TabsContent value="plan">
          <PlanTab
            plan={plan}
            source={source}
            onChange={next => {
              setPlan(next);
              setResult(null);
            }}
            onSaved={saved => {
              setPlan(saved);
              setSource("saved");
            }}
          />
        </TabsContent>
        <TabsContent value="results">
          <ResultsTab result={result} plan={plan} />
        </TabsContent>
      </Tabs>
    </div>
  );
}
