"use client";

import { useEffect } from "react";

import { zodResolver } from "@hookform/resolvers/zod";
import { Controller, useForm } from "react-hook-form";
import { z } from "zod";

import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import {
  Field,
  FieldError,
  FieldGroup,
  FieldLabel,
} from "@/components/ui/field";
import { Input } from "@/components/ui/input";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { ROLE_OPTIONS } from "@/lib/permissions";
import type { TAppUser } from "@/server/user-actions";

const formSchema = z.object({
  email: z.string(),
  password: z.string(),
  name: z.string().min(1, "Vui lòng nhập tên"),
  role: z.enum(["admin", "ad"]),
});

export type UserFormValues = z.infer<typeof formSchema>;
export type UserEditValues = UserFormValues;

export function UserFormDialog({
  open,
  onOpenChange,
  user,
  onSubmit,
  isSubmitting,
}: {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  user: TAppUser | null;
  onSubmit: (values: UserFormValues | UserEditValues) => Promise<void>;
  isSubmitting: boolean;
}) {
  const isEdit = !!user;
  const {
    register,
    handleSubmit,
    control,
    reset,
    formState: { errors },
  } = useForm<UserFormValues>({
    resolver: zodResolver(formSchema),
    defaultValues: { role: "ad", email: "", password: "", name: "" },
  });

  useEffect(() => {
    if (open) {
      reset(
        user
          ? {
              name: user.name,
              role: user.role,
              email: user.email,
              password: "",
            }
          : { name: "", role: "ad", email: "", password: "" }
      );
    }
  }, [open, user, reset]);

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent>
        <form onSubmit={handleSubmit(onSubmit)}>
          <DialogHeader>
            <DialogTitle>
              {isEdit ? "Chỉnh sửa người dùng" : "Thêm người dùng"}
            </DialogTitle>
            <DialogDescription>
              {isEdit
                ? "Cập nhật tên và vai trò của người dùng."
                : "Tạo tài khoản đăng nhập mới cho hệ thống."}
            </DialogDescription>
          </DialogHeader>

          <FieldGroup className="py-4">
            {!isEdit && (
              <Field data-invalid={!!errors.email}>
                <FieldLabel htmlFor="email">Email</FieldLabel>
                <Input
                  id="email"
                  type="email"
                  required
                  {...register("email")}
                />
                <FieldError
                  errors={errors.email ? [errors.email] : undefined}
                />
              </Field>
            )}
            {!isEdit && (
              <Field data-invalid={!!errors.password}>
                <FieldLabel htmlFor="password">Mật khẩu</FieldLabel>
                <Input
                  id="password"
                  type="password"
                  required
                  minLength={8}
                  {...register("password")}
                />
                <FieldError
                  errors={errors.password ? [errors.password] : undefined}
                />
              </Field>
            )}
            <Field data-invalid={!!errors.name}>
              <FieldLabel htmlFor="name">Họ và tên</FieldLabel>
              <Input id="name" {...register("name")} />
              <FieldError errors={errors.name ? [errors.name] : undefined} />
            </Field>
            <Field data-invalid={!!errors.role}>
              <FieldLabel>Vai trò</FieldLabel>
              <Controller
                control={control}
                name="role"
                render={({ field }) => (
                  <Select value={field.value} onValueChange={field.onChange}>
                    <SelectTrigger className="w-full">
                      <SelectValue placeholder="Chọn vai trò" />
                    </SelectTrigger>
                    <SelectContent>
                      {ROLE_OPTIONS.map(opt => (
                        <SelectItem key={opt.value} value={opt.value}>
                          {opt.label}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                )}
              />
              <FieldError errors={errors.role ? [errors.role] : undefined} />
            </Field>
          </FieldGroup>

          <DialogFooter>
            <Button type="submit" disabled={isSubmitting}>
              {isSubmitting ? "Đang lưu..." : "Lưu"}
            </Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}
