import type { Metadata } from "next";
import { redirect } from "next/navigation";

import { getSessionUser } from "@/lib/firebase/session";
import { canAccessRecruitments } from "@/lib/permissions";
import { listRecruitmentSubmissions } from "@/server/recruitment-actions";
import { RecruitmentsView } from "@/components/recruitments-view";

export const metadata: Metadata = {
  title: "Ứng viên tuyển dụng",
};

export default async function RecruitmentsPage() {
  const sessionUser = await getSessionUser();
  if (!sessionUser || !canAccessRecruitments(sessionUser.role)) {
    redirect("/dashboard");
  }

  const result = await listRecruitmentSubmissions();
  const submissions = result.ok ? result.data : [];

  return (
    <div className="flex flex-col gap-4">
      <h1 className="text-2xl font-semibold">Ứng viên tuyển dụng</h1>
      <RecruitmentsView initialSubmissions={submissions} />
    </div>
  );
}
