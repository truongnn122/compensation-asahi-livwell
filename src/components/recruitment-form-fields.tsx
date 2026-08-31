"use client";

import { useState } from "react";
import type { ReactNode } from "react";

import {
  IconCloudUpload,
  IconDownload,
  IconLoader2,
  IconX,
} from "@tabler/icons-react";
import {
  Controller,
  type Control,
  type FieldErrors,
  type UseFormRegister,
  type UseFormSetValue,
  type UseFormWatch,
} from "react-hook-form";

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
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Textarea } from "@/components/ui/textarea";
import { useDictionary } from "@/hooks/use-dictionary";
import type { Dictionary } from "@/lib/i18n/dictionaries";
import { formatBytes } from "@/lib/utils";
import {
  uploadRecruitmentAttachment,
  type TAttachment,
} from "@/server/recruitment-actions";
import type { TManagerOption } from "@/server/user-actions";
import type { RecruitmentValues } from "@/lib/validations/recruitment";

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

const MAX_FILES = 5;
const MAX_FILE_BYTES = 5 * 1024 * 1024;
const ACCEPTED_TYPES =
  "image/jpeg,image/png,image/webp,image/heic,application/pdf,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document";

function getOpenQuestions(
  t: Dictionary
): { name: "q1" | "q2" | "q3" | "q4" | "q6"; label: string }[] {
  return [
    { name: "q1", label: t.recruitmentForm.openQuestions.q1 },
    { name: "q2", label: t.recruitmentForm.openQuestions.q2 },
    { name: "q3", label: t.recruitmentForm.openQuestions.q3 },
    { name: "q4", label: t.recruitmentForm.openQuestions.q4 },
    { name: "q6", label: t.recruitmentForm.openQuestions.q6 },
  ];
}

export function RecruitmentFormFields({
  control,
  register,
  errors,
  watch,
  setValue,
  managers,
  onDownloadAttachment,
}: {
  control: Control<RecruitmentValues>;
  register: UseFormRegister<RecruitmentValues>;
  errors: FieldErrors<RecruitmentValues>;
  watch: UseFormWatch<RecruitmentValues>;
  setValue: UseFormSetValue<RecruitmentValues>;
  managers: TManagerOption[];
  onDownloadAttachment?: (attachment: TAttachment) => void;
}) {
  const t = useDictionary();
  const [uploading, setUploading] = useState(false);
  const [uploadError, setUploadError] = useState<string | null>(null);

  const POSITION_OPTIONS = [
    {
      value: "life_planner",
      label: t.recruitmentForm.options.position.life_planner,
    },
    {
      value: "sales_director",
      label: t.recruitmentForm.options.position.sales_director,
    },
    {
      value: "sales_manager",
      label: t.recruitmentForm.options.position.sales_manager,
    },
    { value: "other", label: t.recruitmentForm.options.position.other },
  ] as const;

  const REFERRAL_OPTIONS = [
    { value: "ads", label: t.recruitmentForm.options.referral.ads },
    { value: "fanpage", label: t.recruitmentForm.options.referral.fanpage },
    { value: "website", label: t.recruitmentForm.options.referral.website },
    { value: "other", label: t.recruitmentForm.options.referral.other },
  ] as const;

  const FAMILY_STATUS_OPTIONS = [
    { value: "single", label: t.recruitmentForm.options.familyStatus.single },
    { value: "married", label: t.recruitmentForm.options.familyStatus.married },
    {
      value: "single_dependent",
      label: t.recruitmentForm.options.familyStatus.single_dependent,
    },
    {
      value: "married_children",
      label: t.recruitmentForm.options.familyStatus.married_children,
    },
    { value: "other", label: t.recruitmentForm.options.familyStatus.other },
  ] as const;

  const TRAINING_OPTIONS = [
    { value: "lpfc", label: t.recruitmentForm.options.training.lpfc },
    {
      value: "sales_skills",
      label: t.recruitmentForm.options.training.sales_skills,
    },
    {
      value: "sales_management",
      label: t.recruitmentForm.options.training.sales_management,
    },
  ] as const;

  const OPEN_QUESTIONS = getOpenQuestions(t);

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
      setUploadError(t.recruitmentForm.maxFilesError(MAX_FILES));
      return;
    }
    const oversized = files.find(f => f.size > MAX_FILE_BYTES);
    if (oversized) {
      setUploadError(t.recruitmentForm.maxFileSizeError);
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
        setValue("attachments", [...watch("attachments"), result.data]);
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

  return (
    <>
      <SectionCard number={1} title={t.recruitmentForm.section1.title}>
        <FieldGroup>
          <Field data-invalid={!!errors.fullName}>
            <FieldLabel htmlFor="fullName">
              {t.recruitmentForm.section1.fullName}
            </FieldLabel>
            <Input
              id="fullName"
              placeholder={t.recruitmentForm.section1.fullNamePlaceholder}
              {...register("fullName")}
            />
            <FieldError
              errors={errors.fullName ? [errors.fullName] : undefined}
            />
          </Field>

          <div className="grid gap-4 sm:grid-cols-2">
            <Field data-invalid={!!errors.dateOfBirth}>
              <FieldLabel htmlFor="dateOfBirth">
                {t.recruitmentForm.section1.dateOfBirth}
              </FieldLabel>
              <Input
                id="dateOfBirth"
                type="date"
                {...register("dateOfBirth")}
              />
              <FieldError
                errors={errors.dateOfBirth ? [errors.dateOfBirth] : undefined}
              />
            </Field>
            <Field data-invalid={!!errors.idNumber}>
              <FieldLabel htmlFor="idNumber">
                {t.recruitmentForm.section1.idNumber}
              </FieldLabel>
              <Input
                id="idNumber"
                placeholder={t.recruitmentForm.section1.idNumberPlaceholder}
                {...register("idNumber")}
              />
              <FieldError
                errors={errors.idNumber ? [errors.idNumber] : undefined}
              />
            </Field>
          </div>

          <div className="grid gap-4 sm:grid-cols-2">
            <Field>
              <FieldLabel htmlFor="idIssueDate">
                {t.recruitmentForm.section1.idIssueDate}
              </FieldLabel>
              <Input
                id="idIssueDate"
                type="date"
                {...register("idIssueDate")}
              />
            </Field>
            <Field>
              <FieldLabel htmlFor="idIssuePlace">
                {t.recruitmentForm.section1.idIssuePlace}
              </FieldLabel>
              <Input
                id="idIssuePlace"
                placeholder={t.recruitmentForm.section1.idIssuePlacePlaceholder}
                {...register("idIssuePlace")}
              />
            </Field>
          </div>

          <Field data-invalid={!!errors.permanentAddress}>
            <FieldLabel htmlFor="permanentAddress">
              {t.recruitmentForm.section1.permanentAddress}
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
              {t.recruitmentForm.section1.contactAddress}
            </FieldLabel>
            <Input id="contactAddress" {...register("contactAddress")} />
          </Field>

          <div className="grid gap-4 sm:grid-cols-2">
            <Field data-invalid={!!errors.mobile1}>
              <FieldLabel htmlFor="mobile1">
                {t.recruitmentForm.section1.mobile1}
              </FieldLabel>
              <Input
                id="mobile1"
                placeholder={t.recruitmentForm.section1.mobile1Placeholder}
                {...register("mobile1")}
              />
              <FieldError
                errors={errors.mobile1 ? [errors.mobile1] : undefined}
              />
            </Field>
            <Field>
              <FieldLabel htmlFor="mobile2">
                {t.recruitmentForm.section1.mobile2}
              </FieldLabel>
              <Input id="mobile2" {...register("mobile2")} />
            </Field>
          </div>

          <Field data-invalid={!!errors.email}>
            <FieldLabel htmlFor="email">
              {t.recruitmentForm.section1.email}
            </FieldLabel>
            <Input
              id="email"
              type="email"
              placeholder={t.recruitmentForm.section1.emailPlaceholder}
              {...register("email")}
            />
            <FieldError errors={errors.email ? [errors.email] : undefined} />
          </Field>

          <Controller
            control={control}
            name="managerUid"
            render={({ field }) => (
              <Field data-invalid={!!errors.managerUid}>
                <FieldLabel>
                  {t.recruitmentForm.section1.managerLabel}
                </FieldLabel>
                <Select
                  value={field.value}
                  onValueChange={uid => {
                    field.onChange(uid);
                    const manager = managers.find(m => m.uid === uid);
                    setValue("managerName", manager?.name ?? "");
                  }}
                >
                  <SelectTrigger className="w-full">
                    <SelectValue
                      placeholder={
                        t.recruitmentForm.section1.managerPlaceholder
                      }
                    />
                  </SelectTrigger>
                  <SelectContent>
                    {managers.map(manager => (
                      <SelectItem key={manager.uid} value={manager.uid}>
                        {manager.name}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
                <FieldError
                  errors={errors.managerUid ? [errors.managerUid] : undefined}
                />
              </Field>
            )}
          />
        </FieldGroup>
      </SectionCard>

      <SectionCard number={2} title={t.recruitmentForm.section2.title}>
        <FieldGroup>
          <Controller
            control={control}
            name="positionApplied"
            render={({ field }) => (
              <Field data-invalid={!!errors.positionApplied}>
                <FieldLabel>
                  {t.recruitmentForm.section2.positionLabel}
                </FieldLabel>
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
                    errors.positionApplied
                      ? [errors.positionApplied]
                      : undefined
                  }
                />
              </Field>
            )}
          />
          {positionApplied === "other" && (
            <Field>
              <FieldLabel htmlFor="positionOther">
                {t.recruitmentForm.section2.specifyOther}
              </FieldLabel>
              <Input id="positionOther" {...register("positionOther")} />
            </Field>
          )}

          <Controller
            control={control}
            name="referralChannel"
            render={({ field }) => (
              <Field data-invalid={!!errors.referralChannel}>
                <FieldLabel>
                  {t.recruitmentForm.section2.referralLabel}
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
                    errors.referralChannel
                      ? [errors.referralChannel]
                      : undefined
                  }
                />
              </Field>
            )}
          />
          {referralChannel.includes("other") && (
            <Field>
              <FieldLabel htmlFor="referralOther">
                {t.recruitmentForm.section2.specifyOther}
              </FieldLabel>
              <Input id="referralOther" {...register("referralOther")} />
            </Field>
          )}
        </FieldGroup>
      </SectionCard>

      <SectionCard number={3} title={t.recruitmentForm.section3.title}>
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
                  errors={
                    errors.familyStatus ? [errors.familyStatus] : undefined
                  }
                />
              </Field>
            )}
          />
          {familyStatus === "married_children" && (
            <Field>
              <FieldLabel htmlFor="childrenCount">
                {t.recruitmentForm.section3.childrenCount}
              </FieldLabel>
              <Input id="childrenCount" {...register("childrenCount")} />
            </Field>
          )}
          {familyStatus === "other" && (
            <Field>
              <FieldLabel htmlFor="familyStatusOther">
                {t.recruitmentForm.section3.specifyOther}
              </FieldLabel>
              <Input
                id="familyStatusOther"
                {...register("familyStatusOther")}
              />
            </Field>
          )}
        </FieldGroup>
      </SectionCard>

      <SectionCard number={4} title={t.recruitmentForm.section4.title}>
        <FieldGroup>
          <div className="bg-muted text-muted-foreground rounded-lg p-4 text-sm">
            <span className="text-foreground font-semibold">
              {t.recruitmentForm.section4.definitionLabel}
            </span>{" "}
            {t.recruitmentForm.section4.definitionText}
          </div>
          <Controller
            control={control}
            name="pepStatus"
            render={({ field }) => (
              <Field>
                <FieldLabel>
                  {t.recruitmentForm.section4.questionLabel}
                </FieldLabel>
                <RadioGroup
                  value={field.value}
                  onValueChange={field.onChange}
                  className="flex gap-6"
                >
                  <div className="flex items-center gap-2">
                    <RadioGroupItem value="no" id="pepStatus-no" />
                    <FieldLabel htmlFor="pepStatus-no" className="font-normal">
                      {t.recruitmentForm.section4.no}
                    </FieldLabel>
                  </div>
                  <div className="flex items-center gap-2">
                    <RadioGroupItem value="yes" id="pepStatus-yes" />
                    <FieldLabel htmlFor="pepStatus-yes" className="font-normal">
                      {t.recruitmentForm.section4.yes}
                    </FieldLabel>
                  </div>
                </RadioGroup>
              </Field>
            )}
          />
          {pepStatus === "yes" && (
            <FieldGroup>
              <Field data-invalid={!!errors.pepRelationship}>
                <FieldLabel htmlFor="pepRelationship">
                  {t.recruitmentForm.section4.relationship}
                </FieldLabel>
                <Input
                  id="pepRelationship"
                  placeholder={
                    t.recruitmentForm.section4.relationshipPlaceholder
                  }
                  {...register("pepRelationship")}
                />
                <FieldError
                  errors={
                    errors.pepRelationship
                      ? [errors.pepRelationship]
                      : undefined
                  }
                />
              </Field>
              <div className="grid gap-4 sm:grid-cols-2">
                <Field data-invalid={!!errors.pepFullName}>
                  <FieldLabel htmlFor="pepFullName">
                    {t.recruitmentForm.section4.fullName}
                  </FieldLabel>
                  <Input id="pepFullName" {...register("pepFullName")} />
                  <FieldError
                    errors={
                      errors.pepFullName ? [errors.pepFullName] : undefined
                    }
                  />
                </Field>
                <Field data-invalid={!!errors.pepPosition}>
                  <FieldLabel htmlFor="pepPosition">
                    {t.recruitmentForm.section4.position}
                  </FieldLabel>
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
                  {t.recruitmentForm.section4.organization}
                </FieldLabel>
                <Input id="pepOrganization" {...register("pepOrganization")} />
                <FieldError
                  errors={
                    errors.pepOrganization
                      ? [errors.pepOrganization]
                      : undefined
                  }
                />
              </Field>
            </FieldGroup>
          )}
        </FieldGroup>
      </SectionCard>

      <SectionCard number={5} title={t.recruitmentForm.section5.title}>
        <FieldGroup>
          <p className="text-sm font-medium">
            {t.recruitmentForm.section5.currentCompany}
          </p>
          <div className="grid gap-4 sm:grid-cols-3">
            <Field>
              <FieldLabel htmlFor="currentCompanyName">
                {t.recruitmentForm.section5.companyName}
              </FieldLabel>
              <Input
                id="currentCompanyName"
                {...register("currentCompanyName")}
              />
            </Field>
            <Field>
              <FieldLabel htmlFor="currentManagerName">
                {t.recruitmentForm.section5.managerName}
              </FieldLabel>
              <Input
                id="currentManagerName"
                {...register("currentManagerName")}
              />
            </Field>
            <Field>
              <FieldLabel htmlFor="currentManagerContact">
                {t.recruitmentForm.section5.managerContact}
              </FieldLabel>
              <Input
                id="currentManagerContact"
                {...register("currentManagerContact")}
              />
            </Field>
          </div>

          <p className="text-sm font-medium">
            {t.recruitmentForm.section5.previousCompany}
          </p>
          <div className="grid gap-4 sm:grid-cols-3">
            <Field>
              <FieldLabel htmlFor="previousCompanyName">
                {t.recruitmentForm.section5.companyName}
              </FieldLabel>
              <Input
                id="previousCompanyName"
                {...register("previousCompanyName")}
              />
            </Field>
            <Field>
              <FieldLabel htmlFor="previousManagerName">
                {t.recruitmentForm.section5.managerName}
              </FieldLabel>
              <Input
                id="previousManagerName"
                {...register("previousManagerName")}
              />
            </Field>
            <Field>
              <FieldLabel htmlFor="previousManagerContact">
                {t.recruitmentForm.section5.managerContact}
              </FieldLabel>
              <Input
                id="previousManagerContact"
                {...register("previousManagerContact")}
              />
            </Field>
          </div>
        </FieldGroup>
      </SectionCard>

      <SectionCard number={6} title={t.recruitmentForm.section6.title}>
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
              <Field
                data-slot="checkbox-group"
                data-invalid={!!errors.q5Training}
              >
                <FieldLabel>
                  {t.recruitmentForm.section6.trainingQuestion}
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
        title={t.recruitmentForm.section7.title}
        description={t.recruitmentForm.section7.description(MAX_FILES)}
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
            <span className="text-sm font-medium">
              {t.recruitmentForm.section7.dropzoneTitle}
            </span>
            <span className="text-muted-foreground text-xs">
              {t.recruitmentForm.section7.dropzoneSubtitle}
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
                  <span className="flex shrink-0 items-center gap-2">
                    {onDownloadAttachment && (
                      <button
                        type="button"
                        onClick={() => onDownloadAttachment(a)}
                        className="text-muted-foreground hover:text-foreground"
                      >
                        <IconDownload className="size-4" />
                        <span className="sr-only">
                          {t.recruitmentForm.section7.downloadSr}
                        </span>
                      </button>
                    )}
                    <button
                      type="button"
                      onClick={() => removeAttachment(a.storagePath)}
                      className="text-muted-foreground hover:text-foreground"
                    >
                      <IconX className="size-4" />
                      <span className="sr-only">
                        {t.recruitmentForm.section7.removeSr}
                      </span>
                    </button>
                  </span>
                </li>
              ))}
            </ul>
          )}
        </Field>
      </SectionCard>

      <SectionCard number={8} title={t.recruitmentForm.section8.title}>
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
              {t.recruitmentForm.section8.voluntary}
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
            <FieldLabel htmlFor="commitmentDataConsent" className="font-normal">
              {t.recruitmentForm.section8.dataConsent}
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
              {t.recruitmentForm.section8.signature}
            </FieldLabel>
            <Input id="signatureName" {...register("signatureName")} />
            <FieldError
              errors={errors.signatureName ? [errors.signatureName] : undefined}
            />
          </Field>
        </FieldGroup>
      </SectionCard>
    </>
  );
}
