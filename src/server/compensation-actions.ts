"use server";

import { adminDb } from "@/lib/firebase/admin";
import { getSessionUser } from "@/lib/firebase/session";
import { getDictionary } from "@/lib/i18n/get-dictionary";
import { canAccessCompensation } from "@/lib/permissions";
import { DEFAULT_PLAN } from "@/lib/compensation/default-plan";
import type { CompPlan } from "@/lib/compensation/types";
import { compPlanSchema } from "@/lib/validations/compensation";
import type { ActionResult } from "@/lib/types";

const COLLECTION = "compensationPlans";
const CURRENT_DOC = "current";

async function requireAccess() {
  const dict = await getDictionary();
  const user = await getSessionUser();
  if (!user) return { ok: false as const, error: dict.errors.notAuthenticated };
  if (!canAccessCompensation(user.role))
    return { ok: false as const, error: dict.errors.forbidden };
  return { ok: true as const, user, dict };
}

/**
 * Returns the saved plan, or the placeholder default when nothing has been
 * saved yet (and when Firestore is unreachable, so the page still renders).
 */
export async function getCompensationPlan(): Promise<{
  plan: CompPlan;
  source: "saved" | "default";
}> {
  try {
    const doc = await adminDb.collection(COLLECTION).doc(CURRENT_DOC).get();
    if (!doc.exists) return { plan: DEFAULT_PLAN, source: "default" };
    const parsed = compPlanSchema.safeParse(doc.data()?.plan);
    if (!parsed.success) return { plan: DEFAULT_PLAN, source: "default" };
    return { plan: parsed.data, source: "saved" };
  } catch {
    return { plan: DEFAULT_PLAN, source: "default" };
  }
}

export async function saveCompensationPlan(
  input: CompPlan
): Promise<ActionResult<CompPlan>> {
  const check = await requireAccess();
  if (!check.ok) return check;
  const { user, dict } = check;

  const parsed = compPlanSchema.safeParse(input);
  if (!parsed.success) {
    return { ok: false, error: dict.errors.compensation.invalidPlan };
  }

  try {
    await adminDb
      .collection(COLLECTION)
      .doc(CURRENT_DOC)
      .set({
        plan: parsed.data,
        updatedAt: new Date().toISOString(),
        updatedByUid: user.uid,
        updatedByEmail: user.email ?? "",
      });
    return { ok: true, data: parsed.data };
  } catch (error) {
    console.error("Error saving compensation plan:", error);
    return { ok: false, error: dict.errors.compensation.saveFailed };
  }
}
