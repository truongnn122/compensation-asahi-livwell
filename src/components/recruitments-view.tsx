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
import { RecruitmentDetailDialog } from "@/components/recruitment-detail-dialog";
import { createRecruitmentsColumns } from "@/components/recruitments-columns";
import {
  deleteRecruitmentSubmission,
  updateRecruitmentSubmissionStatus,
  type TRecruitmentSubmission,
} from "@/server/recruitment-actions";

export function RecruitmentsView({
  initialSubmissions,
}: {
  initialSubmissions: TRecruitmentSubmission[];
}) {
  const [submissions, setSubmissions] = useState(initialSubmissions);
  const [viewing, setViewing] = useState<TRecruitmentSubmission | null>(null);
  const [deleting, setDeleting] = useState<TRecruitmentSubmission | null>(null);
  const [isSaving, setIsSaving] = useState(false);

  const handleStatusChange = async (id: string, status: string) => {
    setIsSaving(true);
    try {
      const result = await updateRecruitmentSubmissionStatus(id, status);
      if (!result.ok) {
        toast.error(result.error);
        return;
      }
      setSubmissions(prev =>
        prev.map(s =>
          s.id === id
            ? { ...s, status: status as TRecruitmentSubmission["status"] }
            : s
        )
      );
      setViewing(null);
      toast.success("Đã cập nhật trạng thái");
    } finally {
      setIsSaving(false);
    }
  };

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
    onView: setViewing,
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

      <RecruitmentDetailDialog
        submission={viewing}
        open={!!viewing}
        onOpenChange={open => !open && setViewing(null)}
        onStatusChange={handleStatusChange}
        isSaving={isSaving}
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
