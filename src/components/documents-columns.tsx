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
      header: "File name",
    },
    {
      accessorKey: "size",
      header: "Size",
      cell: ({ row }) => formatBytes(row.original.size),
    },
    {
      accessorKey: "uploadedByEmail",
      header: "Uploaded by",
    },
    {
      accessorKey: "uploadedAt",
      header: "Uploaded",
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
            <span className="sr-only">Download</span>
          </Button>
          <Button
            variant="ghost"
            size="icon-sm"
            onClick={() => onDelete(row.original.id)}
          >
            <Trash2 />
            <span className="sr-only">Delete</span>
          </Button>
        </div>
      ),
      enableSorting: false,
      enableHiding: false,
    },
  ];
}
