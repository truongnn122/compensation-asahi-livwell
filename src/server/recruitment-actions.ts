"use server";

import { randomUUID } from "crypto";

import { adminDb, adminStorage } from "@/lib/firebase/admin";
import { ActionResult } from "@/lib/types";
import {
  recruitmentSchema,
  type RecruitmentValues,
} from "@/lib/validations/recruitment";

const COLLECTION = "recruitment_submissions";
const MAX_UPLOAD_BYTES = 5 * 1024 * 1024;
const ALLOWED_CONTENT_TYPES = [
  "image/jpeg",
  "image/png",
  "image/webp",
  "image/heic",
  "application/pdf",
  "application/msword",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
];

export type TAttachment = {
  storagePath: string;
  fileName: string;
  size: number;
  contentType: string;
};

export async function uploadRecruitmentAttachment(
  formData: FormData
): Promise<ActionResult<TAttachment>> {
  const file = formData.get("file");
  if (!(file instanceof File)) {
    return { ok: false, error: "Không có tệp nào được chọn." };
  }
  if (file.size > MAX_UPLOAD_BYTES) {
    return { ok: false, error: "Mỗi tệp tối đa 5MB." };
  }
  const contentType = file.type || "application/octet-stream";
  if (!ALLOWED_CONTENT_TYPES.includes(contentType)) {
    return { ok: false, error: "Định dạng tệp không được hỗ trợ." };
  }

  const storagePath = `recruitment/${randomUUID()}-${file.name}`;

  try {
    const buffer = Buffer.from(await file.arrayBuffer());
    await adminStorage.bucket().file(storagePath).save(buffer, {
      contentType,
    });

    return {
      ok: true,
      data: { storagePath, fileName: file.name, size: file.size, contentType },
    };
  } catch {
    return { ok: false, error: "Không thể tải tệp lên." };
  }
}

export async function submitRecruitmentForm(
  values: RecruitmentValues
): Promise<ActionResult<{ id: string }>> {
  const parsed = recruitmentSchema.safeParse(values);
  if (!parsed.success) {
    return { ok: false, error: "Dữ liệu không hợp lệ." };
  }

  try {
    const id = randomUUID();
    await adminDb
      .collection(COLLECTION)
      .doc(id)
      .set({
        ...parsed.data,
        submittedAt: new Date().toISOString(),
      });

    return { ok: true, data: { id } };
  } catch {
    return {
      ok: false,
      error: "Không thể gửi thông tin. Vui lòng thử lại sau.",
    };
  }
}
