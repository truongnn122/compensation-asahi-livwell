"use server";

import { randomUUID } from "crypto";

import { adminDb, adminStorage } from "@/lib/firebase/admin";
import { getSessionUser } from "@/lib/firebase/session";
import { ActionResult } from "@/lib/types";

const COLLECTION = "documents";
const MAX_UPLOAD_BYTES = 5 * 1024 * 1024;
const SIGNED_URL_TTL_MS = 15 * 60 * 1000;

export type TDocument = {
  id: string;
  fileName: string;
  contentType: string;
  size: number;
  storagePath: string;
  uploadedByUid: string;
  uploadedByEmail: string;
  uploadedAt: string;
};

export async function listDocuments(): Promise<ActionResult<TDocument[]>> {
  const user = await getSessionUser();
  if (!user) return { ok: false, error: "Not authenticated." };

  try {
    const snapshot = await adminDb
      .collection(COLLECTION)
      .orderBy("uploadedAt", "desc")
      .limit(100)
      .get();

    const documents = snapshot.docs.map(
      doc => ({ id: doc.id, ...doc.data() }) as TDocument
    );

    return { ok: true, data: documents };
  } catch {
    return { ok: false, error: "Unable to load documents." };
  }
}

export async function uploadDocument(
  formData: FormData
): Promise<ActionResult<TDocument>> {
  const user = await getSessionUser();
  if (!user) return { ok: false, error: "Not authenticated." };

  const file = formData.get("file");
  if (!(file instanceof File)) {
    return { ok: false, error: "No file provided." };
  }
  if (file.size > MAX_UPLOAD_BYTES) {
    return { ok: false, error: "File must be 5MB or smaller." };
  }

  const storagePath = `documents/${user.uid}/${randomUUID()}-${file.name}`;
  const contentType = file.type || "application/octet-stream";

  try {
    const buffer = Buffer.from(await file.arrayBuffer());

    await adminStorage.bucket().file(storagePath).save(buffer, {
      contentType,
    });

    const uploadedAt = new Date().toISOString();

    console.log({ fileName: file.name, storagePath, uploadedAt, COLLECTION });

    const docRef = await adminDb.collection(COLLECTION).add({
      fileName: file.name,
      contentType,
      size: file.size,
      storagePath,
      uploadedByUid: user.uid,
      uploadedByEmail: user.email ?? "",
      uploadedAt,
    });

    return {
      ok: true,
      data: {
        id: docRef.id,
        fileName: file.name,
        contentType,
        size: file.size,
        storagePath,
        uploadedByUid: user.uid,
        uploadedByEmail: user.email ?? "",
        uploadedAt,
      },
    };
  } catch (error) {
    console.error("Error generating download link:", error);

    return { ok: false, error: "Unable to upload the document." };
  }
}

export async function deleteDocument(id: string): Promise<ActionResult> {
  const user = await getSessionUser();
  if (!user) return { ok: false, error: "Not authenticated." };

  try {
    const docRef = adminDb.collection(COLLECTION).doc(id);
    const doc = await docRef.get();
    if (!doc.exists) return { ok: false, error: "Document not found." };

    const { storagePath } = doc.data() as TDocument;
    await adminStorage
      .bucket()
      .file(storagePath)
      .delete({ ignoreNotFound: true });
    await docRef.delete();

    return { ok: true, data: null };
  } catch {
    return { ok: false, error: "Unable to delete the document." };
  }
}

export async function getDocumentDownloadUrl(
  id: string
): Promise<ActionResult<{ url: string }>> {
  const user = await getSessionUser();
  if (!user) return { ok: false, error: "Not authenticated." };

  try {
    const doc = await adminDb.collection(COLLECTION).doc(id).get();
    if (!doc.exists) return { ok: false, error: "Document not found." };

    const { storagePath } = doc.data() as TDocument;
    const [url] = await adminStorage
      .bucket()
      .file(storagePath)
      .getSignedUrl({
        action: "read",
        expires: Date.now() + SIGNED_URL_TTL_MS,
      });

    return { ok: true, data: { url } };
  } catch (error) {
    console.error("Error generating download link:", error);
    return { ok: false, error: "Unable to generate a download link." };
  }
}
