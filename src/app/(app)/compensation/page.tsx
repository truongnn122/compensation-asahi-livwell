import type { Metadata } from "next";
import { redirect } from "next/navigation";

import { CompensationView } from "@/components/compensation/compensation-view";
import { getSessionUser } from "@/lib/firebase/session";
import { getDictionary } from "@/lib/i18n/get-dictionary";
import { canAccessCompensation } from "@/lib/permissions";
import { getCompensationPlan } from "@/server/compensation-actions";

export async function generateMetadata(): Promise<Metadata> {
  const dict = await getDictionary();
  return { title: dict.pages.compensation.title };
}

export default async function CompensationPage() {
  const sessionUser = await getSessionUser();
  if (!sessionUser || !canAccessCompensation(sessionUser.role)) {
    redirect("/dashboard");
  }

  const { plan, source } = await getCompensationPlan();

  return <CompensationView initialPlan={plan} planSource={source} />;
}
