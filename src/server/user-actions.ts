"use server";

import { adminAuth, adminDb } from "@/lib/firebase/admin";
import { getSessionUser } from "@/lib/firebase/session";
import { isRole, type Role } from "@/lib/permissions";
import { ActionResult } from "@/lib/types";

const COLLECTION = "users";

export type TAppUser = {
  uid: string;
  email: string;
  name: string;
  role: Role;
  createdAt: string;
};

async function requireAdmin() {
  const user = await getSessionUser();
  if (!user) return { ok: false as const, error: "Not authenticated." };
  if (user.role !== "admin") {
    return {
      ok: false as const,
      error: "Bạn không có quyền thực hiện thao tác này.",
    };
  }
  return { ok: true as const, user };
}

export type TManagerOption = { uid: string; name: string };

/**
 * Public (unauthenticated) — the recruitment application form needs to let
 * a candidate pick which AD (Quản lý Đại lý) they're applying under. This
 * intentionally exposes only { uid, name }, never email or role, and has
 * no requireAdmin() gate since anonymous candidates must be able to call it.
 */
export async function listRecruitmentManagers(): Promise<
  ActionResult<TManagerOption[]>
> {
  try {
    // Sorted client-side rather than via .orderBy("name") to avoid needing
    // a composite Firestore index for this where+orderBy combination — the
    // AD roster is small enough that this is negligible.
    const snapshot = await adminDb
      .collection(COLLECTION)
      .where("role", "==", "ad")
      .get();

    const managers = snapshot.docs
      .map(doc => ({ uid: doc.id, name: doc.data().name as string }))
      .sort((a, b) => a.name.localeCompare(b.name));

    return { ok: true, data: managers };
  } catch {
    return { ok: false, error: "Không thể tải danh sách quản lý." };
  }
}

export async function listUsers(): Promise<ActionResult<TAppUser[]>> {
  const check = await requireAdmin();
  if (!check.ok) return check;

  try {
    const snapshot = await adminDb
      .collection(COLLECTION)
      .orderBy("createdAt", "desc")
      .limit(200)
      .get();

    const users = snapshot.docs.map(
      doc => ({ uid: doc.id, ...doc.data() }) as TAppUser
    );

    return { ok: true, data: users };
  } catch {
    return { ok: false, error: "Không thể tải danh sách người dùng." };
  }
}

export async function createUser(input: {
  email: string;
  password: string;
  name: string;
  role: string;
}): Promise<ActionResult<TAppUser>> {
  const check = await requireAdmin();
  if (!check.ok) return check;

  if (!isRole(input.role)) {
    return { ok: false, error: "Vai trò không hợp lệ." };
  }
  if (input.password.length < 8) {
    return { ok: false, error: "Mật khẩu phải có ít nhất 8 ký tự." };
  }

  try {
    const record = await adminAuth.createUser({
      email: input.email,
      password: input.password,
      displayName: input.name,
    });

    const createdAt = new Date().toISOString();
    await adminDb.collection(COLLECTION).doc(record.uid).set({
      email: input.email,
      name: input.name,
      role: input.role,
      createdAt,
    });

    return {
      ok: true,
      data: {
        uid: record.uid,
        email: input.email,
        name: input.name,
        role: input.role,
        createdAt,
      },
    };
  } catch (error) {
    const code = (error as { code?: string }).code;
    if (code === "auth/email-already-exists") {
      return { ok: false, error: "Email này đã được sử dụng." };
    }
    return { ok: false, error: "Không thể tạo người dùng." };
  }
}

export async function updateUser(
  uid: string,
  input: { name: string; role: string }
): Promise<ActionResult> {
  const check = await requireAdmin();
  if (!check.ok) return check;

  if (!isRole(input.role)) {
    return { ok: false, error: "Vai trò không hợp lệ." };
  }

  try {
    await adminDb.collection(COLLECTION).doc(uid).update({
      name: input.name,
      role: input.role,
    });
    await adminAuth.updateUser(uid, { displayName: input.name });

    return { ok: true, data: null };
  } catch {
    return { ok: false, error: "Không thể cập nhật người dùng." };
  }
}

export async function deleteUser(uid: string): Promise<ActionResult> {
  const check = await requireAdmin();
  if (!check.ok) return check;

  if (uid === check.user.uid) {
    return { ok: false, error: "Bạn không thể xóa chính tài khoản của mình." };
  }

  try {
    await adminDb.collection(COLLECTION).doc(uid).delete();
    await adminAuth.deleteUser(uid);
    return { ok: true, data: null };
  } catch {
    return { ok: false, error: "Không thể xóa người dùng." };
  }
}
