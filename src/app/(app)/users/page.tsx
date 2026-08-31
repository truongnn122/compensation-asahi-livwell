import type { Metadata } from "next";
import { redirect } from "next/navigation";

import { getSessionUser } from "@/lib/firebase/session";
import { canAccessUsers } from "@/lib/permissions";
import { listUsers } from "@/server/user-actions";
import { UsersView } from "@/components/users-view";

export const metadata: Metadata = {
  title: "Người dùng",
};

export default async function UsersPage() {
  const sessionUser = await getSessionUser();
  if (!sessionUser || !canAccessUsers(sessionUser.role)) {
    redirect("/dashboard");
  }

  const result = await listUsers();
  const users = result.ok ? result.data : [];

  return (
    <div className="flex flex-col gap-4">
      <h1 className="text-2xl font-semibold">Quản lý người dùng</h1>
      <UsersView initialUsers={users} />
    </div>
  );
}
