import type { Metadata } from "next";
import { redirect } from "next/navigation";

import { getSessionUser } from "@/lib/firebase/session";
import { canAccessSettings } from "@/lib/permissions";

export const metadata: Metadata = {
  title: "Cài đặt",
};

export default async function SettingsPage() {
  const sessionUser = await getSessionUser();
  if (!sessionUser || !canAccessSettings(sessionUser.role)) {
    redirect("/dashboard");
  }

  return <h1 className="text-2xl font-semibold">Cài đặt</h1>;
}
