import { z } from "zod";

import type { Dictionary } from "@/lib/i18n/dictionaries";

export function buildRecruitmentSchema(
  t: Dictionary["recruitmentForm"]["validation"]
) {
  const recruitmentObjectSchema = z.object({
    fullName: z.string().min(1, t.fullNameRequired),
    dateOfBirth: z.string().min(1, t.dateOfBirthRequired),
    idNumber: z.string().min(1, t.idNumberRequired),
    idIssueDate: z.string().optional(),
    idIssuePlace: z.string().optional(),
    permanentAddress: z.string().min(1, t.permanentAddressRequired),
    contactAddress: z.string().optional(),
    mobile1: z.string().min(1, t.mobile1Required),
    mobile2: z.string().optional(),
    email: z.email(t.emailInvalid),
    managerUid: z.string().min(1, t.managerRequired),
    managerName: z.string().min(1, t.managerRequired),

    positionApplied: z.enum([
      "life_planner",
      "sales_director",
      "sales_manager",
      "other",
    ]),
    positionOther: z.string().optional(),

    referralChannel: z
      .array(z.enum(["ads", "fanpage", "website", "other"]))
      .min(1, t.referralChannelRequired),
    referralOther: z.string().optional(),

    familyStatus: z.enum([
      "single",
      "married",
      "single_dependent",
      "married_children",
      "other",
    ]),
    childrenCount: z.string().optional(),
    familyStatusOther: z.string().optional(),

    pepStatus: z.enum(["no", "yes"], {
      error: t.pepStatusRequired,
    }),
    pepRelationship: z.string().optional(),
    pepFullName: z.string().optional(),
    pepPosition: z.string().optional(),
    pepOrganization: z.string().optional(),

    currentCompanyName: z.string().optional(),
    currentManagerName: z.string().optional(),
    currentManagerContact: z.string().optional(),
    previousCompanyName: z.string().optional(),
    previousManagerName: z.string().optional(),
    previousManagerContact: z.string().optional(),

    q1: z.string().min(1, t.questionRequired),
    q2: z.string().min(1, t.questionRequired),
    q3: z.string().min(1, t.questionRequired),
    q4: z.string().min(1, t.questionRequired),
    q5Training: z
      .array(z.enum(["lpfc", "sales_skills", "sales_management"]))
      .min(1, t.trainingRequired),
    q6: z.string().min(1, t.questionRequired),

    attachments: z
      .array(
        z.object({
          storagePath: z.string(),
          fileName: z.string(),
          size: z.number(),
          contentType: z.string(),
        })
      )
      .max(5, t.attachmentsMax),

    commitmentVoluntary: z.literal(true, {
      error: t.commitmentRequired,
    }),
    commitmentDataConsent: z.literal(true, {
      error: t.commitmentRequired,
    }),
    signatureName: z.string().min(1, t.signatureNameRequired),
  });

  return recruitmentObjectSchema.superRefine((data, ctx) => {
    if (data.pepStatus !== "yes") return;

    const requiredWhenPep: {
      field:
        "pepRelationship" | "pepFullName" | "pepPosition" | "pepOrganization";
      message: string;
    }[] = [
      { field: "pepRelationship", message: t.pepRelationshipRequired },
      { field: "pepFullName", message: t.pepFullNameRequired },
      { field: "pepPosition", message: t.pepPositionRequired },
      { field: "pepOrganization", message: t.pepOrganizationRequired },
    ];

    for (const { field, message } of requiredWhenPep) {
      if (!data[field]) {
        ctx.addIssue({ code: "custom", path: [field], message });
      }
    }
  });
}

export type RecruitmentValues = z.infer<
  ReturnType<typeof buildRecruitmentSchema>
>;
