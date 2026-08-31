"use client";

import { useState } from "react";

import { zodResolver } from "@hookform/resolvers/zod";
import { useForm } from "react-hook-form";

import { Button } from "@/components/ui/button";
import { FieldError } from "@/components/ui/field";
import { RecruitmentFormFields } from "@/components/recruitment-form-fields";
import { submitRecruitmentForm } from "@/server/recruitment-actions";
import type { TManagerOption } from "@/server/user-actions";
import {
  recruitmentSchema,
  type RecruitmentValues,
} from "@/lib/validations/recruitment";

export function RecruitmentForm({ managers }: { managers: TManagerOption[] }) {
  const [submitted, setSubmitted] = useState(false);
  const [formError, setFormError] = useState<string | null>(null);
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
          Thông tin ứng tuyển của bạn đã được gửi thành công. Đội ngũ tuyển dụng
          Asahi Livwell sẽ liên hệ với bạn trong thời gian sớm nhất.
        </p>
      </div>
    );
  }

  return (
    <form onSubmit={handleSubmit(onSubmit)} className="flex flex-col gap-6">
      <RecruitmentFormFields
        control={control}
        register={register}
        errors={errors}
        watch={watch}
        setValue={setValue}
        managers={managers}
      />

      <FieldError>{formError}</FieldError>

      <Button type="submit" disabled={isSubmitting} className="w-full">
        {isSubmitting ? "Đang gửi..." : "Gửi phiếu thông tin"}
      </Button>
    </form>
  );
}
