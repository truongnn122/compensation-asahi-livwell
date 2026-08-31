import type { Metadata } from "next";
import Image from "next/image";

import { RecruitmentForm } from "@/components/recruitment-form";
import { listRecruitmentManagers } from "@/server/user-actions";

export const metadata: Metadata = {
  title: "Phiếu thông tin tuyển dụng",
};

export default async function RecruitmentPage() {
  const result = await listRecruitmentManagers();
  const managers = result.ok ? result.data : [];

  return (
    <div className="relative min-h-svh overflow-hidden">
      <video
        className="fixed inset-0 h-full w-full object-cover"
        src="/brand/videos/looping-av1.mp4"
        autoPlay
        loop
        muted
        playsInline
      />
      <div className="fixed inset-0 bg-[oklch(0.283_0.121_260.9)]/55" />
      <div className="fixed inset-x-0 top-0 h-40 bg-gradient-to-b from-[oklch(0.283_0.121_260.9)]/80 to-transparent" />

      <main className="relative mx-auto max-w-3xl px-6 pb-10">
        <div className="flex flex-col items-center gap-6 py-16 text-center text-white">
          <div className="flex w-full items-center justify-between gap-3">
            <Image
              src="/brand/logo/ASAHI_HORIZONTAL_VECTOR.svg"
              alt="Asahi Livwell"
              width={124}
              height={51}
              className="h-9 w-auto"
              priority
            />
            <p className="text-xs font-medium tracking-widest text-white/70 uppercase">
              Asahi – Hồ sơ đại lý · CT-01
            </p>
          </div>
          <div className="flex flex-col gap-2">
            <h1 className="text-2xl font-semibold text-balance sm:text-3xl">
              Phiếu thông tin tuyển dụng &amp; sàng lọc ứng viên
            </h1>
            <p className="mx-auto max-w-xl text-sm text-white/70">
              Vui lòng điền đầy đủ và chính xác thông tin bên dưới. Các trường
              có dấu * là bắt buộc.
            </p>
          </div>
        </div>

        <RecruitmentForm managers={managers} />
      </main>
    </div>
  );
}
