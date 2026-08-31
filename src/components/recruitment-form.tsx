"use client";

import { useState } from "react";
import type { ReactNode } from "react";

import { zodResolver } from "@hookform/resolvers/zod";
import { IconCloudUpload, IconLoader2, IconX } from "@tabler/icons-react";
import { Controller, useForm } from "react-hook-form";

import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import {
  Field,
  FieldDescription,
  FieldError,
  FieldGroup,
  FieldLabel,
  FieldLegend,
  FieldSet,
} from "@/components/ui/field";
import { Input } from "@/components/ui/input";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { Textarea } from "@/components/ui/textarea";
import { ToggleGroup, ToggleGroupItem } from "@/components/ui/toggle-group";
import { formatBytes } from "@/lib/utils";
import {
  submitRecruitmentForm,
  uploadRecruitmentAttachment,
} from "@/server/recruitment-actions";
import {
  recruitmentSchema,
  type RecruitmentValues,
} from "@/lib/validations/recruitment";

function SectionCard({
  number,
  title,
  description,
  children,
}: {
  number: number;
  title: string;
  description?: string;
  children: ReactNode;
}) {
  return (
    <div className="rounded-xl border bg-card p-6 shadow-sm md:p-8">
      <FieldSet>
        <div className="mb-1 flex items-center gap-3">
          <span className="bg-foreground text-background flex size-7 shrink-0 items-center justify-center rounded-full text-sm font-semibold">
            {number}
          </span>
          <FieldLegend className="mb-0">{title}</FieldLegend>
        </div>
        {description && <FieldDescription>{description}</FieldDescription>}
        {children}
      </FieldSet>
    </div>
  );
}

const POSITION_OPTIONS = [
  { value: "life_planner", label: "Life Planner (LP)" },
  { value: "sales_director", label: "Giám Đốc Ban Kinh Doanh (DM)" },
  { value: "sales_manager", label: "Trưởng Phòng Kinh Doanh (UM)" },
  { value: "other", label: "Khác" },
] as const;

const REFERRAL_OPTIONS = [
  { value: "ads", label: "Quảng cáo" },
  { value: "fanpage", label: "Fanpage" },
  { value: "website", label: "Website tuyển dụng" },
  { value: "other", label: "Khác" },
] as const;

const FAMILY_STATUS_OPTIONS = [
  { value: "single", label: "Độc thân" },
  { value: "married", label: "Đã kết hôn" },
  {
    value: "single_dependent",
    label: "Độc thân, có người phụ thuộc (VD: cha mẹ già)",
  },
  { value: "married_children", label: "Đã kết hôn, có con" },
  { value: "other", label: "Tình trạng khác" },
] as const;

const TRAINING_OPTIONS = [
  { value: "lpfc", label: "Khóa Khởi đầu sự nghiệp (LPFC – 5 ngày)" },
  {
    value: "sales_skills",
    label: "Khóa xây dựng kỹ năng bán hàng (3 buổi trong 3 tháng)",
  },
  {
    value: "sales_management",
    label: "Hoạt động quản lý bán hàng định kỳ (hàng tuần)",
  },
] as const;

const OPEN_QUESTIONS: { name: "q1" | "q2" | "q3" | "q4" | "q6"; label: string }[] =
  [
    {
      name: "q1",
      label:
        "Trong cuộc sống, anh/chị đã từng thấy người thân hoặc người xung quanh nhận được quyền lợi từ bảo hiểm nhân thọ chưa?",
    },
    { name: "q2", label: "Anh/chị nhìn nhận như thế nào về bảo hiểm nhân thọ?" },
    {
      name: "q3",
      label:
        "Người thân, bạn bè sẽ nói gì nếu anh/chị trở thành tư vấn viên bảo hiểm nhân thọ?",
    },
    {
      name: "q4",
      label: "Ai là 10 người đầu tiên anh/chị sẽ trò chuyện về bảo hiểm nhân thọ?",
    },
    {
      name: "q6",
      label:
        "Anh/chị mong MVI hỗ trợ gì để thành công, tăng thu nhập, có cơ hội thăng tiến, nâng cao kỹ năng và kiến thức?",
    },
  ];

const MAX_FILES = 5;
const MAX_FILE_BYTES = 5 * 1024 * 1024;
const ACCEPTED_TYPES =
  "image/jpeg,image/png,image/webp,image/heic,application/pdf,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document";

export function RecruitmentForm() {
  const [submitted, setSubmitted] = useState(false);
  const [formError, setFormError] = useState<string | null>(null);
  const [uploading, setUploading] = useState(false);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const {
    register,
    handleSubmit,
    control,
    watch,
    setValue,
    formState: { errors, isSubmitting },
  } = useForm<RecruitmentValues>({
    resolver: zodResolver(recruitmentSchema),
    defaultValues: {
      pepStatus: "no",
      referralChannel: [],
      q5Training: [],
      attachments: [],
      commitmentVoluntary: undefined,
      commitmentDataConsent: undefined,
    },
  });

  const positionApplied = watch("positionApplied");
  const referralChannel = watch("referralChannel");
  const familyStatus = watch("familyStatus");
  const pepStatus = watch("pepStatus");
  const attachments = watch("attachments");

  const handleFilesSelected = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const files = Array.from(event.target.files ?? []);
    event.target.value = "";
    if (!files.length) return;

    setUploadError(null);
    if (attachments.length + files.length > MAX_FILES) {
      setUploadError(`Tối đa ${MAX_FILES} tệp.`);
      return;
    }
    const oversized = files.find(f => f.size > MAX_FILE_BYTES);
    if (oversized) {
      setUploadError("Mỗi tệp tối đa 5MB.");
      return;
    }

    setUploading(true);
    try {
      for (const file of files) {
        const formData = new FormData();
        formData.set("file", file);
        const result = await uploadRecruitmentAttachment(formData);
        if (!result.ok) {
          setUploadError(result.error);
          continue;
        }
        setValue("attachments", [...attachments, result.data]);
      }
    } finally {
      setUploading(false);
    }
  };

  const removeAttachment = (storagePath: string) => {
    setValue(
      "attachments",
      attachments.filter(a => a.storagePath !== storagePath)
    );
  };

  const onSubmit = async (values: RecruitmentValues) => {
    setFormError(null);
    const result = await submitRecruitmentForm(values);
    if (!result.ok) {
      setFormError(result.error);
      return;
    }
    setSubmitted(true);
  };

  if (submitted) {
    return (
      <div className="flex flex-col items-center gap-2 rounded-xl border bg-card p-10 text-center shadow-sm">
        <h2 className="text-xl font-semibold">Cảm ơn bạn!</h2>
        <p className="text-muted-foreground max-w-md text-sm">
          Thông tin ứng tuyển của bạn đã được gửi thành công. Đội ngũ tuyển
          dụng Asahi Livwell sẽ liên hệ với bạn trong thời gian sớm nhất.
        </p>
      </div>
    );
  }

  return (
    <form onSubmit={handleSubmit(onSubmit)} className="flex flex-col gap-6">
      <SectionCard number={1} title="Thông tin cá nhân">
        <FieldGroup>
          <Field data-invalid={!!errors.fullName}>
            <FieldLabel htmlFor="fullName">Họ và tên *</FieldLabel>
            <Input
              id="fullName"
              placeholder="Nguyễn Văn A"
              {...register("fullName")}
            />
            <FieldError errors={errors.fullName ? [errors.fullName] : undefined} />
          </Field>

          <div className="grid gap-4 sm:grid-cols-2">
            <Field data-invalid={!!errors.dateOfBirth}>
              <FieldLabel htmlFor="dateOfBirth">Ngày sinh *</FieldLabel>
              <Input id="dateOfBirth" type="date" {...register("dateOfBirth")} />
              <FieldError
                errors={errors.dateOfBirth ? [errors.dateOfBirth] : undefined}
              />
            </Field>
            <Field data-invalid={!!errors.idNumber}>
              <FieldLabel htmlFor="idNumber">Số CCCD *</FieldLabel>
              <Input
                id="idNumber"
                placeholder="9 hoặc 12 số"
                {...register("idNumber")}
              />
              <FieldError
                errors={errors.idNumber ? [errors.idNumber] : undefined}
              />
            </Field>
          </div>

          <div className="grid gap-4 sm:grid-cols-2">
            <Field>
              <FieldLabel htmlFor="idIssueDate">Ngày cấp *</FieldLabel>
              <Input id="idIssueDate" type="date" {...register("idIssueDate")} />
            </Field>
            <Field>
              <FieldLabel htmlFor="idIssuePlace">Nơi cấp *</FieldLabel>
              <Input
                id="idIssuePlace"
                placeholder="Cục Cảnh sát QLHC về TTXH"
                {...register("idIssuePlace")}
              />
            </Field>
          </div>

          <Field data-invalid={!!errors.permanentAddress}>
            <FieldLabel htmlFor="permanentAddress">
              Địa chỉ thường trú *
            </FieldLabel>
            <Input id="permanentAddress" {...register("permanentAddress")} />
            <FieldError
              errors={
                errors.permanentAddress ? [errors.permanentAddress] : undefined
              }
            />
          </Field>

          <Field>
            <FieldLabel htmlFor="contactAddress">
              Địa chỉ liên hệ (nếu khác thường trú)
            </FieldLabel>
            <Input id="contactAddress" {...register("contactAddress")} />
          </Field>

          <div className="grid gap-4 sm:grid-cols-2">
            <Field data-invalid={!!errors.mobile1}>
              <FieldLabel htmlFor="mobile1">Di động 1 *</FieldLabel>
              <Input
                id="mobile1"
                placeholder="09xxxxxxxx"
                {...register("mobile1")}
              />
              <FieldError
                errors={errors.mobile1 ? [errors.mobile1] : undefined}
              />
            </Field>
            <Field>
              <FieldLabel htmlFor="mobile2">Di động 2</FieldLabel>
              <Input id="mobile2" {...register("mobile2")} />
            </Field>
          </div>

          <Field data-invalid={!!errors.email}>
            <FieldLabel htmlFor="email">Email *</FieldLabel>
            <Input
              id="email"
              type="email"
              placeholder="ban@email.com"
              {...register("email")}
            />
            <FieldError errors={errors.email ? [errors.email] : undefined} />
          </Field>
        </FieldGroup>
      </SectionCard>

      <SectionCard number={2} title="Vị trí ứng tuyển & kênh biết đến">
        <FieldGroup>
          <Controller
            control={control}
            name="positionApplied"
            render={({ field }) => (
              <Field data-invalid={!!errors.positionApplied}>
                <FieldLabel>Vị trí ứng tuyển tại MVI *</FieldLabel>
                <RadioGroup
                  value={field.value}
                  onValueChange={field.onChange}
                  className="grid gap-3 sm:grid-cols-2"
                >
                  {POSITION_OPTIONS.map(opt => (
                    <div key={opt.value} className="flex items-center gap-2">
                      <RadioGroupItem
                        value={opt.value}
                        id={`position-${opt.value}`}
                      />
                      <FieldLabel
                        htmlFor={`position-${opt.value}`}
                        className="font-normal"
                      >
                        {opt.label}
                      </FieldLabel>
                    </div>
                  ))}
                </RadioGroup>
                <FieldError
                  errors={
                    errors.positionApplied ? [errors.positionApplied] : undefined
                  }
                />
              </Field>
            )}
          />
          {positionApplied === "other" && (
            <Field>
              <FieldLabel htmlFor="positionOther">Vui lòng ghi rõ</FieldLabel>
              <Input id="positionOther" {...register("positionOther")} />
            </Field>
          )}

          <Controller
            control={control}
            name="referralChannel"
            render={({ field }) => (
              <Field data-invalid={!!errors.referralChannel}>
                <FieldLabel>
                  Ứng viên biết đến chương trình tuyển dụng qua *
                </FieldLabel>
                <div className="grid gap-3 sm:grid-cols-2">
                  {REFERRAL_OPTIONS.map(opt => (
                    <div key={opt.value} className="flex items-center gap-2">
                      <Checkbox
                        id={`referral-${opt.value}`}
                        checked={field.value.includes(opt.value)}
                        onCheckedChange={checked => {
                          field.onChange(
                            checked
                              ? [...field.value, opt.value]
                              : field.value.filter(v => v !== opt.value)
                          );
                        }}
                      />
                      <FieldLabel
                        htmlFor={`referral-${opt.value}`}
                        className="font-normal"
                      >
                        {opt.label}
                      </FieldLabel>
                    </div>
                  ))}
                </div>
                <FieldError
                  errors={
                    errors.referralChannel ? [errors.referralChannel] : undefined
                  }
                />
              </Field>
            )}
          />
          {referralChannel.includes("other") && (
            <Field>
              <FieldLabel htmlFor="referralOther">Vui lòng ghi rõ</FieldLabel>
              <Input id="referralOther" {...register("referralOther")} />
            </Field>
          )}
        </FieldGroup>
      </SectionCard>

      <SectionCard number={3} title="Tình trạng gia đình">
        <FieldGroup>
          <Controller
            control={control}
            name="familyStatus"
            render={({ field }) => (
              <Field data-invalid={!!errors.familyStatus}>
                <RadioGroup
                  value={field.value}
                  onValueChange={field.onChange}
                  className="grid gap-3 sm:grid-cols-2"
                >
                  {FAMILY_STATUS_OPTIONS.map(opt => (
                    <div key={opt.value} className="flex items-center gap-2">
                      <RadioGroupItem
                        value={opt.value}
                        id={`family-${opt.value}`}
                      />
                      <FieldLabel
                        htmlFor={`family-${opt.value}`}
                        className="font-normal"
                      >
                        {opt.label}
                      </FieldLabel>
                    </div>
                  ))}
                </RadioGroup>
                <FieldError
                  errors={errors.familyStatus ? [errors.familyStatus] : undefined}
                />
              </Field>
            )}
          />
          {familyStatus === "married_children" && (
            <Field>
              <FieldLabel htmlFor="childrenCount">Số lượng con</FieldLabel>
              <Input id="childrenCount" {...register("childrenCount")} />
            </Field>
          )}
          {familyStatus === "other" && (
            <Field>
              <FieldLabel htmlFor="familyStatusOther">Vui lòng ghi rõ</FieldLabel>
              <Input id="familyStatusOther" {...register("familyStatusOther")} />
            </Field>
          )}
        </FieldGroup>
      </SectionCard>

      <SectionCard number={4} title="Thông tin mối quan hệ với PEP">
        <FieldGroup>
          <div className="bg-muted text-muted-foreground rounded-lg p-4 text-sm">
            <span className="text-foreground font-semibold">
              Định nghĩa PEP:
            </span>{" "}
            là cá nhân có ảnh hưởng chính trị, người nắm giữ chức vụ cấp cao
            trong cơ quan Nhà nước ở tất cả các quốc gia (VD: Nguyên thủ
            quốc gia, bộ trưởng, thứ trưởng, Đại biểu quốc hội, Đại sứ, Quản
            lý cấp cao của cơ quan/doanh nghiệp Nhà nước...), hoặc người
            thân/cá nhân có quan hệ mật thiết với PEP (vợ/chồng, con, bạn
            đời, cha/mẹ, cha mẹ của bạn đời, anh/chị/em ruột hoặc cùng cha
            khác mẹ hoặc cùng mẹ khác cha). Nếu không tự xác định được, vui
            lòng liên hệ bộ phận Nhân sự để được tham vấn.
          </div>
          <Controller
            control={control}
            name="pepStatus"
            render={({ field }) => (
              <Field>
                <FieldLabel>
                  Anh/chị có phải PEP hoặc có mối quan hệ với PEP như định
                  nghĩa trên không? *
                </FieldLabel>
                <ToggleGroup
                  type="single"
                  variant="outline"
                  value={field.value}
                  onValueChange={v => v && field.onChange(v)}
                >
                  <ToggleGroupItem value="no">Không</ToggleGroupItem>
                  <ToggleGroupItem value="yes">Có</ToggleGroupItem>
                </ToggleGroup>
              </Field>
            )}
          />
          {pepStatus === "yes" && (
            <FieldGroup>
              <Field data-invalid={!!errors.pepRelationship}>
                <FieldLabel htmlFor="pepRelationship">
                  Mối quan hệ với PEP *
                </FieldLabel>
                <Input
                  id="pepRelationship"
                  placeholder="VD: Cha/mẹ, vợ/chồng, bản thân là PEP..."
                  {...register("pepRelationship")}
                />
                <FieldError
                  errors={
                    errors.pepRelationship ? [errors.pepRelationship] : undefined
                  }
                />
              </Field>
              <div className="grid gap-4 sm:grid-cols-2">
                <Field data-invalid={!!errors.pepFullName}>
                  <FieldLabel htmlFor="pepFullName">
                    Họ và tên (của PEP) *
                  </FieldLabel>
                  <Input id="pepFullName" {...register("pepFullName")} />
                  <FieldError
                    errors={
                      errors.pepFullName ? [errors.pepFullName] : undefined
                    }
                  />
                </Field>
                <Field data-invalid={!!errors.pepPosition}>
                  <FieldLabel htmlFor="pepPosition">Chức vụ *</FieldLabel>
                  <Input id="pepPosition" {...register("pepPosition")} />
                  <FieldError
                    errors={
                      errors.pepPosition ? [errors.pepPosition] : undefined
                    }
                  />
                </Field>
              </div>
              <Field data-invalid={!!errors.pepOrganization}>
                <FieldLabel htmlFor="pepOrganization">
                  Đơn vị công tác *
                </FieldLabel>
                <Input id="pepOrganization" {...register("pepOrganization")} />
                <FieldError
                  errors={
                    errors.pepOrganization ? [errors.pepOrganization] : undefined
                  }
                />
              </Field>
            </FieldGroup>
          )}
        </FieldGroup>
      </SectionCard>

      <SectionCard number={5} title="Về công việc của bạn">
        <FieldGroup>
          <p className="text-sm font-medium">Công ty hiện tại</p>
          <div className="grid gap-4 sm:grid-cols-3">
            <Field>
              <FieldLabel htmlFor="currentCompanyName">
                Tên công ty
              </FieldLabel>
              <Input
                id="currentCompanyName"
                {...register("currentCompanyName")}
              />
            </Field>
            <Field>
              <FieldLabel htmlFor="currentManagerName">
                Tên quản lý
              </FieldLabel>
              <Input
                id="currentManagerName"
                {...register("currentManagerName")}
              />
            </Field>
            <Field>
              <FieldLabel htmlFor="currentManagerContact">
                Liên hệ quản lý
              </FieldLabel>
              <Input
                id="currentManagerContact"
                {...register("currentManagerContact")}
              />
            </Field>
          </div>

          <p className="text-sm font-medium">Công ty trước đây</p>
          <div className="grid gap-4 sm:grid-cols-3">
            <Field>
              <FieldLabel htmlFor="previousCompanyName">
                Tên công ty
              </FieldLabel>
              <Input
                id="previousCompanyName"
                {...register("previousCompanyName")}
              />
            </Field>
            <Field>
              <FieldLabel htmlFor="previousManagerName">
                Tên quản lý
              </FieldLabel>
              <Input
                id="previousManagerName"
                {...register("previousManagerName")}
              />
            </Field>
            <Field>
              <FieldLabel htmlFor="previousManagerContact">
                Liên hệ quản lý
              </FieldLabel>
              <Input
                id="previousManagerContact"
                {...register("previousManagerContact")}
              />
            </Field>
          </div>
        </FieldGroup>
      </SectionCard>

      <SectionCard number={6} title="Vui lòng trả lời các câu hỏi sau">
        <FieldGroup>
          {OPEN_QUESTIONS.map((q, i) => (
            <Field key={q.name} data-invalid={!!errors[q.name]}>
              <FieldLabel htmlFor={q.name}>
                {i + 1}) {q.label} *
              </FieldLabel>
              <Textarea id={q.name} rows={3} {...register(q.name)} />
              <FieldError
                errors={errors[q.name] ? [errors[q.name]!] : undefined}
              />
            </Field>
          ))}

          <Controller
            control={control}
            name="q5Training"
            render={({ field }) => (
              <Field data-slot="checkbox-group" data-invalid={!!errors.q5Training}>
                <FieldLabel>
                  5) Anh/chị có sẵn sàng tham gia các khóa đào tạo sau để
                  phát triển bản thân? *
                </FieldLabel>
                {TRAINING_OPTIONS.map(opt => (
                  <Field
                    key={opt.value}
                    orientation="horizontal"
                    className="items-start"
                  >
                    <Checkbox
                      checked={field.value.includes(opt.value)}
                      onCheckedChange={checked => {
                        field.onChange(
                          checked
                            ? [...field.value, opt.value]
                            : field.value.filter(v => v !== opt.value)
                        );
                      }}
                      id={`training-${opt.value}`}
                    />
                    <FieldLabel
                      htmlFor={`training-${opt.value}`}
                      className="font-normal"
                    >
                      {opt.label}
                    </FieldLabel>
                  </Field>
                ))}
                <FieldError
                  errors={errors.q5Training ? [errors.q5Training] : undefined}
                />
              </Field>
            )}
          />

          <Field data-invalid={!!errors.q6}>
            <FieldLabel htmlFor="q6">
              6) {OPEN_QUESTIONS.find(q => q.name === "q6")?.label}
            </FieldLabel>
            <Textarea id="q6" rows={3} {...register("q6")} />
            <FieldError errors={errors.q6 ? [errors.q6] : undefined} />
          </Field>
        </FieldGroup>
      </SectionCard>

      <SectionCard
        number={7}
        title="Upload hồ sơ đầu vào"
        description={`Đính kèm hình ảnh hoặc tài liệu liên quan (CCCD, bằng cấp, CV...). Tối đa ${MAX_FILES} file, mỗi file không quá 5MB. Hỗ trợ: Hình ảnh JPG, Hình ảnh PNG, Hình ảnh WEBP, Hình ảnh HEIC, Tài liệu PDF, Tài liệu DOC, Tài liệu DOCX.`}
      >
        <Field>
          <label
            className={`flex flex-col items-center justify-center gap-1 rounded-lg border border-dashed py-10 text-center ${
              uploading || attachments.length >= MAX_FILES
                ? "cursor-not-allowed opacity-50"
                : "hover:bg-muted/50 cursor-pointer"
            }`}
          >
            {uploading ? (
              <IconLoader2 className="text-muted-foreground size-6 animate-spin" />
            ) : (
              <IconCloudUpload className="text-muted-foreground size-6" />
            )}
            <span className="text-sm font-medium">Nhấp để chọn file</span>
            <span className="text-muted-foreground text-xs">
              Hình ảnh hoặc tài liệu (PDF, DOC, DOCX)
            </span>
            <input
              type="file"
              multiple
              accept={ACCEPTED_TYPES}
              className="sr-only"
              onChange={handleFilesSelected}
              disabled={uploading || attachments.length >= MAX_FILES}
            />
          </label>
          {uploadError && (
            <p className="text-destructive text-sm">{uploadError}</p>
          )}
          {attachments.length > 0 && (
            <ul className="mt-2 flex flex-col gap-2">
              {attachments.map(a => (
                <li
                  key={a.storagePath}
                  className="bg-muted flex items-center justify-between gap-2 rounded-md px-3 py-2 text-sm"
                >
                  <span className="truncate">
                    {a.fileName}{" "}
                    <span className="text-muted-foreground">
                      ({formatBytes(a.size)})
                    </span>
                  </span>
                  <button
                    type="button"
                    onClick={() => removeAttachment(a.storagePath)}
                    className="text-muted-foreground hover:text-foreground shrink-0"
                  >
                    <IconX className="size-4" />
                    <span className="sr-only">Xóa tệp</span>
                  </button>
                </li>
              ))}
            </ul>
          )}
        </Field>
      </SectionCard>

      <SectionCard number={8} title="Cam kết của ứng viên">
        <FieldGroup data-slot="checkbox-group">
          <Field orientation="horizontal" className="items-start">
            <Controller
              control={control}
              name="commitmentVoluntary"
              render={({ field }) => (
                <Checkbox
                  id="commitmentVoluntary"
                  checked={field.value ?? false}
                  onCheckedChange={v => field.onChange(v === true)}
                />
              )}
            />
            <FieldLabel htmlFor="commitmentVoluntary" className="font-normal">
              Tôi xác nhận việc tìm hiểu cơ hội nghề nghiệp này và ứng tuyển
              làm đại lý tại MVI hoàn toàn là quyết định tự nguyện của cá
              nhân tôi, không do bất kỳ cá nhân hay tổ chức nào chi phối.
            </FieldLabel>
          </Field>
          <FieldError
            errors={
              errors.commitmentVoluntary
                ? [errors.commitmentVoluntary]
                : undefined
            }
          />

          <Field orientation="horizontal" className="items-start">
            <Controller
              control={control}
              name="commitmentDataConsent"
              render={({ field }) => (
                <Checkbox
                  id="commitmentDataConsent"
                  checked={field.value ?? false}
                  onCheckedChange={v => field.onChange(v === true)}
                />
              )}
            />
            <FieldLabel
              htmlFor="commitmentDataConsent"
              className="font-normal"
            >
              Tôi đồng ý cho phép xử lý dữ liệu cá nhân (PDPD) để hệ thống
              thu thập, lưu trữ và xử lý hồ sơ của tôi.
            </FieldLabel>
          </Field>
          <FieldError
            errors={
              errors.commitmentDataConsent
                ? [errors.commitmentDataConsent]
                : undefined
            }
          />

          <Field data-invalid={!!errors.signatureName}>
            <FieldLabel htmlFor="signatureName">
              Họ và tên xác nhận (chữ ký điện tử) *
            </FieldLabel>
            <Input id="signatureName" {...register("signatureName")} />
            <FieldError
              errors={errors.signatureName ? [errors.signatureName] : undefined}
            />
          </Field>
        </FieldGroup>
      </SectionCard>

      <FieldError>{formError}</FieldError>

      <Button type="submit" disabled={isSubmitting} className="w-full">
        {isSubmitting ? "Đang gửi..." : "Gửi phiếu thông tin"}
      </Button>
    </form>
  );
}
