"use client";

import { useState } from "react";

import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { formatDate } from "@/lib/utils";
import type { TRecruitmentSubmission } from "@/server/recruitment-actions";
import {
  POSITION_LABELS,
  STATUS_LABELS,
} from "@/components/recruitments-columns";

function Row({ label, value }: { label: string; value?: string | null }) {
  if (!value) return null;
  return (
    <div className="grid grid-cols-3 gap-2 text-sm">
      <span className="text-muted-foreground">{label}</span>
      <span className="col-span-2">{value}</span>
    </div>
  );
}

export function RecruitmentDetailDialog({
  submission,
  open,
  onOpenChange,
  onStatusChange,
  isSaving,
}: {
  submission: TRecruitmentSubmission | null;
  open: boolean;
  onOpenChange: (open: boolean) => void;
  onStatusChange: (id: string, status: string) => Promise<void>;
  isSaving: boolean;
}) {
  const [status, setStatus] = useState<string>(submission?.status ?? "new");

  if (!submission) return null;

  return (
    <Dialog
      open={open}
      onOpenChange={next => {
        onOpenChange(next);
        if (next) setStatus(submission.status);
      }}
    >
      <DialogContent className="max-h-[85vh] overflow-y-auto sm:max-w-xl">
        <DialogHeader>
          <DialogTitle>{submission.fullName}</DialogTitle>
          <DialogDescription>
            Nộp ngày {formatDate(submission.submittedAt)}
          </DialogDescription>
        </DialogHeader>

        <div className="flex flex-col gap-2">
          <Row label="Điện thoại" value={submission.mobile1} />
          <Row label="Email" value={submission.email} />
          <Row label="Quản lý" value={submission.managerName} />
          <Row label="Ngày sinh" value={submission.dateOfBirth} />
          <Row label="Số CCCD" value={submission.idNumber} />
          <Row label="Địa chỉ thường trú" value={submission.permanentAddress} />
          <Row
            label="Vị trí ứng tuyển"
            value={POSITION_LABELS[submission.positionApplied]}
          />
          <Row label="Công ty hiện tại" value={submission.currentCompanyName} />
          <Row
            label="Đính kèm"
            value={`${submission.attachments.length} tệp`}
          />
        </div>

        <div className="space-y-1.5">
          <Label>Trạng thái</Label>
          <Select value={status} onValueChange={setStatus}>
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
        </div>

        <DialogFooter>
          <Button
            disabled={isSaving || status === submission.status}
            onClick={() => onStatusChange(submission.id, status)}
          >
            {isSaving ? "Đang lưu..." : "Cập nhật trạng thái"}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
