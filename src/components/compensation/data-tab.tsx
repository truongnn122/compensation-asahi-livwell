"use client";

import { useRef, useState } from "react";

import { Database, FileSpreadsheet, Trash2, Upload } from "lucide-react";
import { toast } from "sonner";

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
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { useDictionary } from "@/hooks/use-dictionary";
import { eventLabel } from "@/lib/compensation/engine";
import {
  buildTemplateWorkbook,
  downloadBlob,
  PAYMENT_MODE_LABEL,
  parseWorkbook,
} from "@/lib/compensation/excel";
import {
  SAMPLE_AGENTS,
  SAMPLE_EVENTS,
  SAMPLE_PAID_HISTORY,
} from "@/lib/compensation/sample-data";
import type { CompPlan } from "@/lib/compensation/types";

import type { Dataset } from "./compensation-view";
import { fill, vnd } from "./format";

export function DataTab({
  plan,
  dataset,
  onChange,
}: {
  plan: CompPlan;
  dataset: Dataset;
  onChange: (next: Dataset) => void;
}) {
  const t = useDictionary();
  const d = t.compensation.data;
  const fileRef = useRef<HTMLInputElement>(null);
  const [errors, setErrors] = useState<string[]>([]);
  const [busy, setBusy] = useState(false);

  const productName = new Map(plan.products.map(p => [p.code, p.name]));
  const rankName = new Map(plan.ranks.map(r => [r.code, r.name]));
  const agentName = new Map(dataset.agents.map(a => [a.code, a.name]));

  const downloadTemplate = async () => {
    setBusy(true);
    try {
      const blob = await buildTemplateWorkbook(plan, {
        agents: SAMPLE_AGENTS,
        events: SAMPLE_EVENTS,
        paidHistory: SAMPLE_PAID_HISTORY,
      });
      downloadBlob(blob, "mau-du-lieu-thu-nhap.xlsx");
    } finally {
      setBusy(false);
    }
  };

  const handleFile = async (file: File) => {
    setBusy(true);
    try {
      const parsed = await parseWorkbook(await file.arrayBuffer(), plan);
      setErrors(parsed.errors);
      if (parsed.errors.length > 0) {
        toast.error(d.importErrors);
        return;
      }
      onChange({
        agents: parsed.agents,
        events: parsed.events,
        paidHistory: parsed.paidHistory,
      });
      toast.success(
        fill(d.imported, {
          agents: parsed.agents.length,
          events: parsed.events.length,
        })
      );
    } catch {
      toast.error(d.importFailed);
    } finally {
      setBusy(false);
      if (fileRef.current) fileRef.current.value = "";
    }
  };

  const useSample = () => {
    setErrors([]);
    onChange({
      agents: SAMPLE_AGENTS,
      events: SAMPLE_EVENTS,
      paidHistory: SAMPLE_PAID_HISTORY,
    });
    toast.success(d.sampleLoaded);
  };

  const clear = () => {
    setErrors([]);
    onChange({ agents: [], events: [], paidHistory: [] });
  };

  const hasData = dataset.agents.length > 0 || dataset.events.length > 0;

  return (
    <div className="flex flex-col gap-4">
      <div className="flex flex-wrap gap-2">
        <Button variant="outline" onClick={downloadTemplate} disabled={busy}>
          <FileSpreadsheet />
          {d.downloadTemplate}
        </Button>
        <Button
          variant="outline"
          onClick={() => fileRef.current?.click()}
          disabled={busy}
        >
          <Upload />
          {d.importExcel}
        </Button>
        <input
          ref={fileRef}
          id="compensation-import-file"
          type="file"
          accept=".xlsx"
          className="hidden"
          onChange={e => {
            const file = e.target.files?.[0];
            if (file) void handleFile(file);
          }}
        />
        <Button variant="secondary" onClick={useSample} disabled={busy}>
          <Database />
          {d.useSample}
        </Button>
        {hasData && (
          <Button variant="ghost" onClick={clear} disabled={busy}>
            <Trash2 />
            {d.clear}
          </Button>
        )}
      </div>

      {errors.length > 0 && (
        <Card className="border-destructive/40">
          <CardHeader>
            <CardTitle className="text-destructive text-base">
              {d.importErrors}
            </CardTitle>
          </CardHeader>
          <CardContent>
            <ul className="list-disc space-y-1 pl-5 text-sm">
              {errors.slice(0, 30).map((err, i) => (
                <li key={i}>{err}</li>
              ))}
              {errors.length > 30 && (
                <li className="text-muted-foreground">+{errors.length - 30}</li>
              )}
            </ul>
          </CardContent>
        </Card>
      )}

      {!hasData ? (
        <EmptyState
          icon={FileSpreadsheet}
          title={d.empty}
          description={d.emptyHint}
        />
      ) : (
        <div className="grid gap-4 lg:grid-cols-5">
          <Card className="lg:col-span-2">
            <CardHeader>
              <CardTitle className="text-base">{d.agentsTitle}</CardTitle>
              <CardDescription>
                {fill(d.agentsCount, { count: dataset.agents.length })}
              </CardDescription>
            </CardHeader>
            <CardContent className="overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>{d.cols.code}</TableHead>
                    <TableHead>{d.cols.name}</TableHead>
                    <TableHead>{d.cols.rank}</TableHead>
                    <TableHead>{d.cols.manager}</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {dataset.agents.map(a => (
                    <TableRow key={a.code}>
                      <TableCell className="font-mono text-xs">
                        {a.code}
                      </TableCell>
                      <TableCell>{a.name}</TableCell>
                      <TableCell>
                        <Badge variant="secondary">
                          {rankName.get(a.rank) ?? a.rank}
                        </Badge>
                      </TableCell>
                      <TableCell className="text-muted-foreground">
                        {a.managerCode
                          ? (agentName.get(a.managerCode) ?? a.managerCode)
                          : "—"}
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </CardContent>
          </Card>

          <Card className="lg:col-span-3">
            <CardHeader>
              <CardTitle className="text-base">{d.eventsTitle}</CardTitle>
              <CardDescription>
                {fill(d.eventsCount, { count: dataset.events.length })}
              </CardDescription>
            </CardHeader>
            <CardContent className="overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>{d.cols.policy}</TableHead>
                    <TableHead>{d.cols.type}</TableHead>
                    <TableHead>{d.cols.code}</TableHead>
                    <TableHead>{d.cols.product}</TableHead>
                    <TableHead className="text-right">{d.cols.year}</TableHead>
                    <TableHead>{d.cols.mode}</TableHead>
                    <TableHead className="text-right">
                      {d.cols.amount}
                    </TableHead>
                    <TableHead>{d.cols.occurredAt}</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {dataset.events.map(e => (
                    <TableRow key={e.eventId}>
                      <TableCell className="font-mono text-xs">
                        {e.policyNo}
                      </TableCell>
                      <TableCell>
                        <Badge
                          variant={
                            e.type === "PREMIUM_COLLECTED"
                              ? "secondary"
                              : "destructive"
                          }
                        >
                          {eventLabel(e.type)}
                        </Badge>
                      </TableCell>
                      <TableCell>
                        {agentName.get(e.agentCode) ?? e.agentCode}
                      </TableCell>
                      <TableCell>
                        {productName.get(e.productCode) ?? e.productCode}
                      </TableCell>
                      <TableCell className="text-right tabular-nums">
                        {e.policyYear}
                      </TableCell>
                      <TableCell>{PAYMENT_MODE_LABEL[e.paymentMode]}</TableCell>
                      <TableCell className="text-right tabular-nums">
                        {vnd(e.amount)}
                      </TableCell>
                      <TableCell className="text-muted-foreground">
                        {e.occurredAt}
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </CardContent>
          </Card>

          {dataset.paidHistory.length > 0 && (
            <Card className="lg:col-span-5">
              <CardHeader>
                <CardTitle className="text-base">{d.paidTitle}</CardTitle>
                <CardDescription>
                  {fill(d.paidCount, { count: dataset.paidHistory.length })}
                </CardDescription>
              </CardHeader>
              <CardContent className="overflow-x-auto">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>{d.cols.period}</TableHead>
                      <TableHead>{d.cols.code}</TableHead>
                      <TableHead>{d.cols.type}</TableHead>
                      <TableHead>{d.cols.policy}</TableHead>
                      <TableHead>{d.cols.rule}</TableHead>
                      <TableHead className="text-right">
                        {d.cols.amount}
                      </TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {dataset.paidHistory.map((p, i) => (
                      <TableRow key={`${p.sourcePolicyNo}-${p.agentCode}-${i}`}>
                        <TableCell>{p.periodId}</TableCell>
                        <TableCell>
                          {agentName.get(p.agentCode) ?? p.agentCode}
                        </TableCell>
                        <TableCell>{p.type}</TableCell>
                        <TableCell className="font-mono text-xs">
                          {p.sourcePolicyNo}
                        </TableCell>
                        <TableCell className="text-muted-foreground">
                          {p.ruleId}
                        </TableCell>
                        <TableCell className="text-right tabular-nums">
                          {vnd(p.amount)}
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </CardContent>
            </Card>
          )}
        </div>
      )}
    </div>
  );
}
