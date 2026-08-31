"use client";

import { useState } from "react";

import { toast } from "sonner";

import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from "@/components/ui/alert-dialog";
import { DataTable } from "@/components/ui/data-table/data-table";
import { createRecruitmentsColumns } from "@/components/recruitments-columns";
import {
  deleteRecruitmentSubmission,
  type TRecruitmentSubmission,
} from "@/server/recruitment-actions";

export function RecruitmentsView({
  initialSubmissions,
}: {
  initialSubmissions: TRecruitmentSubmission[];
}) {
  const [submissions, setSubmissions] = useState(initialSubmissions);
  const [deleting, setDeleting] = useState<TRecruitmentSubmission | null>(null);

  const handleDelete = async () => {
    if (!deleting) return;
    const target = deleting;
    setDeleting(null);
    setSubmissions(prev => prev.filter(s => s.id !== target.id));

    const result = await deleteRecruitmentSubmission(target.id);
    if (!result.ok) {
      setSubmissions(prev => [target, ...prev]);
      toast.error(result.error);
    } else {
      toast.success("Đã xóa hồ sơ ứng viên");
    }
  };

  const columns = createRecruitmentsColumns({
    onDelete: setDeleting,
  });

  return (
    <div className="flex flex-col gap-4">
      <DataTable
        data={submissions.map(s => ({ ...s, id: s.id }))}
        columns={columns}
        emptyMessage="Chưa có hồ sơ ứng viên nào."
        enableColumnVisibility={false}
      />

      <AlertDialog
        open={!!deleting}
        onOpenChange={open => !open && setDeleting(null)}
      >
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Xóa hồ sơ ứng viên?</AlertDialogTitle>
            <AlertDialogDescription>
              Hồ sơ của {deleting?.fullName} sẽ bị xóa vĩnh viễn, bao gồm mọi
              tệp đính kèm. Hành động này không thể hoàn tác.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Hủy</AlertDialogCancel>
            <AlertDialogAction onClick={handleDelete}>Xóa</AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}
