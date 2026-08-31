"use client";

import Link from "next/link";

import { IconEye, IconTrash } from "@tabler/icons-react";
import type { ColumnDef } from "@tanstack/react-table";

import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { formatDate } from "@/lib/utils";
import type { TRecruitmentSubmission } from "@/server/recruitment-actions";

const POSITION_LABELS: Record<string, string> = {
  life_planner: "Life Planner (LP)",
  sales_director: "Giám Đốc Ban Kinh Doanh (DM)",
  sales_manager: "Trưởng Phòng Kinh Doanh (UM)",
  other: "Khác",
};

const STATUS_LABELS: Record<string, string> = {
  new: "Mới",
  contacted: "Đã liên hệ",
  hired: "Đã tuyển",
  rejected: "Từ chối",
};

const STATUS_VARIANTS: Record<string, "default" | "secondary" | "destructive"> =
  {
    new: "secondary",
    contacted: "default",
    hired: "default",
    rejected: "destructive",
  };

export function createRecruitmentsColumns({
  onDelete,
}: {
  onDelete: (submission: TRecruitmentSubmission) => void;
}): ColumnDef<TRecruitmentSubmission & { id: string }>[] {
  return [
    {
      accessorKey: "fullName",
      header: "Họ và tên",
    },
    {
      accessorKey: "mobile1",
      header: "Điện thoại",
    },
    {
      accessorKey: "email",
      header: "Email",
    },
    {
      accessorKey: "positionApplied",
      header: "Vị trí ứng tuyển",
      cell: ({ row }) =>
        POSITION_LABELS[row.original.positionApplied] ??
        row.original.positionApplied,
    },
    {
      accessorKey: "managerName",
      header: "Quản lý",
    },
    {
      accessorKey: "status",
      header: "Trạng thái",
      cell: ({ row }) => (
        <Badge variant={STATUS_VARIANTS[row.original.status]}>
          {STATUS_LABELS[row.original.status]}
        </Badge>
      ),
    },
    {
      accessorKey: "submittedAt",
      header: "Ngày nộp",
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
              <span className="sr-only">Xem chi tiết</span>
            </Link>
          </Button>
          <Button
            variant="ghost"
            size="icon-sm"
            onClick={() => onDelete(row.original)}
          >
            <IconTrash className="size-4" />
            <span className="sr-only">Xóa</span>
          </Button>
        </div>
      ),
      enableSorting: false,
      enableHiding: false,
    },
  ];
}

export { POSITION_LABELS, STATUS_LABELS };
