"use client";

import { useState } from "react";

import { zodResolver } from "@hookform/resolvers/zod";
import { useForm } from "react-hook-form";
import { useRouter } from "next/navigation";
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
import { Button } from "@/components/ui/button";
import { Field, FieldLabel } from "@/components/ui/field";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { RecruitmentFormFields } from "@/components/recruitment-form-fields";
import { STATUS_LABELS } from "@/components/recruitments-columns";
import { formatDate } from "@/lib/utils";
import {
  recruitmentSchema,
  type RecruitmentValues,
} from "@/lib/validations/recruitment";
import {
  deleteRecruitmentSubmission,
  getRecruitmentAttachmentUrl,
  updateRecruitmentSubmission,
  updateRecruitmentSubmissionStatus,
  type TAttachment,
  type TRecruitmentSubmission,
} from "@/server/recruitment-actions";
import type { TManagerOption } from "@/server/user-actions";

export function RecruitmentDetailView({
  submission,
  managers,
}: {
  submission: TRecruitmentSubmission;
  managers: TManagerOption[];
}) {
  const router = useRouter();
  const [status, setStatus] = useState(submission.status);
  const [isSaving, setIsSaving] = useState(false);
  const [isDeleting, setIsDeleting] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);

  const {
    register,
    handleSubmit,
    control,
    watch,
    setValue,
    formState: { errors },
  } = useForm<RecruitmentValues>({
    resolver: zodResolver(recruitmentSchema),
    defaultValues: submission,
  });

  const handleDownload = async (attachment: TAttachment) => {
    const result = await getRecruitmentAttachmentUrl(attachment.storagePath);
    if (!result.ok) {
      toast.error(result.error);
      return;
    }
    window.open(result.data.url, "_blank");
  };

  const onSubmit = async (values: RecruitmentValues) => {
    setIsSaving(true);
    try {
      const result = await updateRecruitmentSubmission(submission.id, values);
      if (!result.ok) {
        toast.error(result.error);
        return;
      }
      if (status !== submission.status) {
        const statusResult = await updateRecruitmentSubmissionStatus(
          submission.id,
          status
        );
        if (!statusResult.ok) {
          toast.error(statusResult.error);
          return;
        }
      }
      toast.success("Đã lưu thay đổi");
      router.refresh();
    } finally {
      setIsSaving(false);
    }
  };

  const handleDelete = async () => {
    setConfirmDelete(false);
    setIsDeleting(true);
    const result = await deleteRecruitmentSubmission(submission.id);
    if (!result.ok) {
      toast.error(result.error);
      setIsDeleting(false);
      return;
    }
    toast.success("Đã xóa hồ sơ ứng viên");
    router.push("/recruitments");
  };

  return (
    <form onSubmit={handleSubmit(onSubmit)} className="flex flex-col gap-6">
      <div className="rounded-xl border bg-card p-6 shadow-sm md:p-8">
        <div className="grid gap-4 sm:grid-cols-2">
          <Field>
            <FieldLabel>Trạng thái</FieldLabel>
            <Select
              value={status}
              onValueChange={v =>
                setStatus(v as TRecruitmentSubmission["status"])
              }
            >
              <SelectTrigger className="w-full">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                {Object.entries(STATUS_LABELS).map(([value, label]) => (
                  <SelectItem key={value} value={value}>
                    {label}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </Field>
          <Field>
            <FieldLabel>Ngày nộp</FieldLabel>
            <p className="text-muted-foreground text-sm">
              {formatDate(submission.submittedAt)}
            </p>
          </Field>
        </div>
      </div>

      <RecruitmentFormFields
        control={control}
        register={register}
        errors={errors}
        watch={watch}
        setValue={setValue}
        managers={managers}
        onDownloadAttachment={handleDownload}
      />

      <div className="flex items-center justify-between gap-2">
        <Button
          type="button"
          variant="destructive"
          disabled={isDeleting}
          onClick={() => setConfirmDelete(true)}
        >
          Xóa hồ sơ
        </Button>
        <Button type="submit" disabled={isSaving}>
          {isSaving ? "Đang lưu..." : "Lưu thay đổi"}
        </Button>
      </div>

      <AlertDialog open={confirmDelete} onOpenChange={setConfirmDelete}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Xóa hồ sơ ứng viên?</AlertDialogTitle>
            <AlertDialogDescription>
              Hồ sơ của {submission.fullName} sẽ bị xóa vĩnh viễn, bao gồm mọi
              tệp đính kèm. Hành động này không thể hoàn tác.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Hủy</AlertDialogCancel>
            <AlertDialogAction onClick={handleDelete}>Xóa</AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </form>
  );
}
