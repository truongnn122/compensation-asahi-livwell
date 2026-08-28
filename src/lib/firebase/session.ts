import "server-only";

import { cookies } from "next/headers";

import { adminAuth } from "@/lib/firebase/admin";
import { SESSION_COOKIE_NAME } from "@/lib/firebase/session-cookie";

export type TSessionUser = {
  uid: string;
  email: string | null;
  name?: string;
  picture?: string;
};

export async function getSessionUser(): Promise<TSessionUser | null> {
  const cookieStore = await cookies();
  const cookie = cookieStore.get(SESSION_COOKIE_NAME)?.value;
  if (!cookie) return null;

  try {
    const decoded = await adminAuth.verifySessionCookie(cookie, true);
    return {
      uid: decoded.uid,
      email: decoded.email ?? null,
      name: decoded.name,
      picture: decoded.picture,
    };
  } catch {
    return null;
  }
}
