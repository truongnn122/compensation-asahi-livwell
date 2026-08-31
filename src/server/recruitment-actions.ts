"use server";

import { randomUUID } from "crypto";

import { adminDb, adminStorage } from "@/lib/firebase/admin";
import { getSessionUser } from "@/lib/firebase/session";
import { canAccessRecruitments } from "@/lib/permissions";
import { ActionResult } from "@/lib/types";
import {
  recruitmentSchema,
  type RecruitmentValues,
} from "@/lib/validations/recruitment";

const COLLECTION = "recruitment_submissions";
const STATUS_VALUES = ["new", "contacted", "hired", "rejected"] as const;
export type TRecruitmentStatus = (typeof STATUS_VALUES)[number];
const SIGNED_URL_TTL_MS = 15 * 60 * 1000;

export type TRecruitmentSubmission = RecruitmentValues & {
  id: string;
  submittedAt: string;
  status: TRecruitmentStatus;
};

async function requireRecruitmentAccess() {
  const user = await getSessionUser();
  if (!user) return { ok: false as const, error: "Not authenticated." };
  if (!canAccessRecruitments(user.role)) {
    return {
      ok: false as const,
      error: "Bạn không có quyền thực hiện thao tác này.",
    };
  }
  return { ok: true as const, user };
}
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
        status: "new" satisfies TRecruitmentStatus,
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

export async function listRecruitmentSubmissions(): Promise<
  ActionResult<TRecruitmentSubmission[]>
> {
  const check = await requireRecruitmentAccess();
  if (!check.ok) return check;

  try {
    const snapshot = await adminDb
      .collection(COLLECTION)
      .orderBy("submittedAt", "desc")
      .limit(200)
      .get();

    const submissions = snapshot.docs.map(
      doc => ({ id: doc.id, ...doc.data() }) as TRecruitmentSubmission
    );

    return { ok: true, data: submissions };
  } catch {
    return { ok: false, error: "Không thể tải danh sách ứng viên." };
  }
}

export async function getRecruitmentSubmission(
  id: string
): Promise<ActionResult<TRecruitmentSubmission>> {
  const check = await requireRecruitmentAccess();
  if (!check.ok) return check;

  try {
    const doc = await adminDb.collection(COLLECTION).doc(id).get();
    if (!doc.exists) {
      return { ok: false, error: "Không tìm thấy hồ sơ ứng viên." };
    }

    return {
      ok: true,
      data: { id: doc.id, ...doc.data() } as TRecruitmentSubmission,
    };
  } catch {
    return { ok: false, error: "Không thể tải hồ sơ ứng viên." };
  }
}

export async function updateRecruitmentSubmission(
  id: string,
  values: RecruitmentValues
): Promise<ActionResult> {
  const check = await requireRecruitmentAccess();
  if (!check.ok) return check;

  const parsed = recruitmentSchema.safeParse(values);
  if (!parsed.success) {
    return { ok: false, error: "Dữ liệu không hợp lệ." };
  }

  try {
    await adminDb.collection(COLLECTION).doc(id).update(parsed.data);
    return { ok: true, data: null };
  } catch {
    return { ok: false, error: "Không thể cập nhật hồ sơ ứng viên." };
  }
}

export async function getRecruitmentAttachmentUrl(
  storagePath: string
): Promise<ActionResult<{ url: string }>> {
  const check = await requireRecruitmentAccess();
  if (!check.ok) return check;

  try {
    const [url] = await adminStorage
      .bucket()
      .file(storagePath)
      .getSignedUrl({
        action: "read",
        expires: Date.now() + SIGNED_URL_TTL_MS,
      });

    return { ok: true, data: { url } };
  } catch {
    return { ok: false, error: "Không thể tạo liên kết tải xuống." };
  }
}

export async function updateRecruitmentSubmissionStatus(
  id: string,
  status: string
): Promise<ActionResult> {
  const check = await requireRecruitmentAccess();
  if (!check.ok) return check;

  if (!STATUS_VALUES.includes(status as TRecruitmentStatus)) {
    return { ok: false, error: "Trạng thái không hợp lệ." };
  }

  try {
    await adminDb.collection(COLLECTION).doc(id).update({ status });
    return { ok: true, data: null };
  } catch {
    return { ok: false, error: "Không thể cập nhật trạng thái." };
  }
}

export async function deleteRecruitmentSubmission(
  id: string
): Promise<ActionResult> {
  const check = await requireRecruitmentAccess();
  if (!check.ok) return check;

  try {
    const ref = adminDb.collection(COLLECTION).doc(id);
    const doc = await ref.get();
    const attachments =
      (doc.data()?.attachments as TRecruitmentSubmission["attachments"]) ?? [];

    await Promise.all(
      attachments.map(a =>
        adminStorage
          .bucket()
          .file(a.storagePath)
          .delete({ ignoreNotFound: true })
      )
    );
    await ref.delete();

    return { ok: true, data: null };
  } catch {
    return { ok: false, error: "Không thể xóa hồ sơ ứng viên." };
  }
}
