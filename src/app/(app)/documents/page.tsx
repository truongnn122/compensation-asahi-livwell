import type { Metadata } from "next";

import { DocumentsView } from "@/components/documents-view";
import { listDocuments } from "@/server/documents-actions";

export const metadata: Metadata = {
  title: "Tài liệu",
};

export default async function DocumentsPage() {
  const result = await listDocuments();
  const documents = result.ok ? result.data : [];

  return (
    <div className="flex flex-col gap-4">
      <h1 className="text-2xl font-semibold">Tài liệu</h1>
      <DocumentsView initialDocuments={documents} />
    </div>
  );
}
