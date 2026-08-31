"use client";

import { IconEdit, IconTrash } from "@tabler/icons-react";
import type { ColumnDef } from "@tanstack/react-table";

import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { ROLE_LABELS } from "@/lib/permissions";
import { formatDate } from "@/lib/utils";
import type { TAppUser } from "@/server/user-actions";

export function createUsersColumns({
  onEdit,
  onDelete,
}: {
  onEdit: (user: TAppUser) => void;
  onDelete: (user: TAppUser) => void;
}): ColumnDef<TAppUser & { id: string }>[] {
  return [
    {
      accessorKey: "name",
      header: "Họ và tên",
    },
    {
      accessorKey: "email",
      header: "Email",
    },
    {
      accessorKey: "role",
      header: "Vai trò",
      cell: ({ row }) => (
        <Badge variant="secondary">{ROLE_LABELS[row.original.role]}</Badge>
      ),
    },
    {
      accessorKey: "createdAt",
      header: "Ngày tạo",
      cell: ({ row }) => formatDate(row.original.createdAt),
    },
    {
      id: "actions",
      header: "",
      cell: ({ row }) => (
        <div className="flex justify-end gap-1">
          <Button
            variant="ghost"
            size="icon-sm"
            onClick={() => onEdit(row.original)}
          >
            <IconEdit className="size-4" />
            <span className="sr-only">Chỉnh sửa</span>
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
