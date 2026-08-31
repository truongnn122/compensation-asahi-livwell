import type { Metadata } from "next";
import { redirect } from "next/navigation";

import { getSessionUser } from "@/lib/firebase/session";
import { canAccessDocuments } from "@/lib/permissions";
import { DocumentsView } from "@/components/documents-view";
import { listDocuments } from "@/server/documents-actions";

export const metadata: Metadata = {
  title: "Tài liệu",
};

export default async function DocumentsPage() {
  const sessionUser = await getSessionUser();
  if (!sessionUser || !canAccessDocuments(sessionUser.role)) {
    redirect("/dashboard");
  }

  const result = await listDocuments();
  const documents = result.ok ? result.data : [];

  return (
    <div className="flex flex-col gap-4">
      <h1 className="text-2xl font-semibold">Tài liệu</h1>
      <DocumentsView initialDocuments={documents} />
    </div>
  );
}
