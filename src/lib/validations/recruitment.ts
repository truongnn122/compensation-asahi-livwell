import { z } from "zod";

const recruitmentObjectSchema = z.object({
  fullName: z.string().min(1, "Vui lòng nhập họ và tên"),
  dateOfBirth: z.string().min(1, "Vui lòng nhập ngày sinh"),
  idNumber: z.string().min(1, "Vui lòng nhập số CCCD"),
  idIssueDate: z.string().optional(),
  idIssuePlace: z.string().optional(),
  permanentAddress: z.string().min(1, "Vui lòng nhập địa chỉ thường trú"),
  contactAddress: z.string().optional(),
  mobile1: z.string().min(1, "Vui lòng nhập số điện thoại"),
  mobile2: z.string().optional(),
  email: z.email("Vui lòng nhập địa chỉ email hợp lệ"),

  positionApplied: z.enum([
    "life_planner",
    "sales_director",
    "sales_manager",
    "other",
  ]),
  positionOther: z.string().optional(),

  referralChannel: z
    .array(z.enum(["ads", "fanpage", "website", "other"]))
    .min(1, "Vui lòng chọn ít nhất một lựa chọn"),
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
    error: "Vui lòng chọn một lựa chọn",
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

  q1: z.string().min(1, "Vui lòng trả lời câu hỏi này"),
  q2: z.string().min(1, "Vui lòng trả lời câu hỏi này"),
  q3: z.string().min(1, "Vui lòng trả lời câu hỏi này"),
  q4: z.string().min(1, "Vui lòng trả lời câu hỏi này"),
  q5Training: z
    .array(z.enum(["lpfc", "sales_skills", "sales_management"]))
    .min(1, "Vui lòng chọn ít nhất một khóa đào tạo"),
  q6: z.string().min(1, "Vui lòng trả lời câu hỏi này"),

  attachments: z
    .array(
      z.object({
        storagePath: z.string(),
        fileName: z.string(),
        size: z.number(),
        contentType: z.string(),
      })
    )
    .max(5, "Tối đa 5 tệp"),

  commitmentVoluntary: z.literal(true, {
    error: "Vui lòng xác nhận cam kết này",
  }),
  commitmentDataConsent: z.literal(true, {
    error: "Vui lòng xác nhận cam kết này",
  }),
  signatureName: z.string().min(1, "Vui lòng nhập họ tên xác nhận"),
});

export const recruitmentSchema = recruitmentObjectSchema.superRefine(
  (data, ctx) => {
    if (data.pepStatus !== "yes") return;

    const requiredWhenPep: {
      field: "pepRelationship" | "pepFullName" | "pepPosition" | "pepOrganization";
      message: string;
    }[] = [
      { field: "pepRelationship", message: "Vui lòng nhập mối quan hệ với PEP" },
      { field: "pepFullName", message: "Vui lòng nhập họ và tên" },
      { field: "pepPosition", message: "Vui lòng nhập chức vụ" },
      { field: "pepOrganization", message: "Vui lòng nhập đơn vị công tác" },
    ];

    for (const { field, message } of requiredWhenPep) {
      if (!data[field]) {
        ctx.addIssue({ code: "custom", path: [field], message });
      }
    }
  }
);

export type RecruitmentValues = z.infer<typeof recruitmentObjectSchema>;
