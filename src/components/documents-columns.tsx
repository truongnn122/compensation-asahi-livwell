"use client";

import { Download, Trash2 } from "lucide-react";
import type { ColumnDef } from "@tanstack/react-table";

import { Button } from "@/components/ui/button";
import { formatBytes, formatDate } from "@/lib/utils";
import type { TDocument } from "@/server/documents-actions";

export function createDocumentsColumns({
  onDownload,
  onDelete,
}: {
  onDownload: (id: string) => void;
  onDelete: (id: string) => void;
}): ColumnDef<TDocument>[] {
  return [
    {
      accessorKey: "fileName",
      header: "Tên tệp",
    },
    {
      accessorKey: "size",
      header: "Kích thước",
      cell: ({ row }) => formatBytes(row.original.size),
    },
    {
      accessorKey: "uploadedByEmail",
      header: "Người tải lên",
    },
    {
      accessorKey: "uploadedAt",
      header: "Ngày tải lên",
      cell: ({ row }) => formatDate(row.original.uploadedAt),
    },
    {
      id: "actions",
      header: "",
      cell: ({ row }) => (
        <div className="flex justify-end gap-1">
          <Button
            variant="ghost"
            size="icon-sm"
            onClick={() => onDownload(row.original.id)}
          >
            <Download />
            <span className="sr-only">Tải xuống</span>
          </Button>
          <Button
            variant="ghost"
            size="icon-sm"
            onClick={() => onDelete(row.original.id)}
          >
            <Trash2 />
            <span className="sr-only">Xóa</span>
          </Button>
        </div>
      ),
      enableSorting: false,
      enableHiding: false,
    },
  ];
}
