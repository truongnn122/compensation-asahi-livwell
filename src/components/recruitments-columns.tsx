"use client";

import Link from "next/link";

import { IconEye, IconTrash } from "@tabler/icons-react";
import type { ColumnDef } from "@tanstack/react-table";

import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import type { Dictionary } from "@/lib/i18n/dictionaries";
import { formatDate } from "@/lib/utils";
import type { TRecruitmentSubmission } from "@/server/recruitment-actions";

const STATUS_VARIANTS: Record<string, "default" | "secondary" | "destructive"> =
  {
    new: "secondary",
    contacted: "default",
    hired: "default",
    rejected: "destructive",
  };

export function createRecruitmentsColumns({
  t,
  onDelete,
}: {
  t: Dictionary;
  onDelete: (submission: TRecruitmentSubmission) => void;
}): ColumnDef<TRecruitmentSubmission & { id: string }>[] {
  return [
    {
      accessorKey: "fullName",
      header: t.recruitmentsList.columns.name,
    },
    {
      accessorKey: "mobile1",
      header: t.recruitmentsList.columns.phone,
    },
    {
      accessorKey: "email",
      header: t.recruitmentsList.columns.email,
    },
    {
      accessorKey: "positionApplied",
      header: t.recruitmentsList.columns.position,
      cell: ({ row }) =>
        t.recruitmentsList.positionLabels[row.original.positionApplied] ??
        row.original.positionApplied,
    },
    {
      accessorKey: "managerName",
      header: t.recruitmentsList.columns.manager,
    },
    {
      accessorKey: "status",
      header: t.recruitmentsList.columns.status,
      cell: ({ row }) => (
        <Badge variant={STATUS_VARIANTS[row.original.status]}>
          {t.recruitmentsList.statusLabels[row.original.status]}
        </Badge>
      ),
    },
    {
      accessorKey: "submittedAt",
      header: t.recruitmentsList.columns.submittedAt,
      cell: ({ row }) => formatDate(row.original.submittedAt),
    },
    {
      id: "actions",
      header: "",
      cell: ({ row }) => (
        <div className="flex justify-end gap-1">
          <Button variant="ghost" size="icon-sm" asChild>
            <Link href={`/recruitments/${row.original.id}`}>
              <IconEye className="size-4" />
              <span className="sr-only">{t.recruitmentsList.viewSr}</span>
            </Link>
          </Button>
          <Button
            variant="ghost"
            size="icon-sm"
            onClick={() => onDelete(row.original)}
          >
            <IconTrash className="size-4" />
            <span className="sr-only">{t.recruitmentsList.deleteSr}</span>
          </Button>
        </div>
      ),
      enableSorting: false,
      enableHiding: false,
    },
  ];
}
