import { ReactNode } from "react";

import Image from "next/image";

const QUOTE =
  '"The journey of a thousand miles begins with a single step." — Lao Tzu';

export default function AuthLayout({
  children,
}: Readonly<{ children: ReactNode }>) {
  return (
    <div className="flex min-h-svh w-full">
      <div className="relative hidden w-[60%] overflow-hidden lg:flex lg:flex-col">
        <video
          className="absolute inset-0 h-full w-full object-cover"
          src="/brand/videos/looping-av1.mp4"
          autoPlay
          loop
          muted
          playsInline
        />
        <div className="bg-primary/55 absolute inset-0" />
        <div className="from-primary/80 absolute inset-x-0 bottom-0 h-28 bg-gradient-to-t to-transparent" />

        <div className="text-primary-foreground relative flex h-full flex-col justify-between p-10">
          <Image
            src="/brand/logo/ASAHI_HORIZONTAL_VECTOR.svg"
            alt="Asahi Livwell"
            width={124}
            height={51}
            className="mr-auto h-10 w-auto self-start"
            priority
          />

          <div className="max-w-md space-y-3">
            <Image
              src="/brand/images/pink_circle.svg"
              alt=""
              width={44}
              height={28}
              className="h-7 w-11"
            />
            <h2 className="text-3xl leading-tight font-semibold text-balance">
              A new dawn unfolds
            </h2>
            <p className="text-primary-foreground/70 text-sm">
              Sign in to manage compensation and benefits for the Asahi
              Livwell team.
            </p>
          </div>

          <div className="overflow-hidden whitespace-nowrap">
            <div className="animate-marquee flex w-max items-center text-sm">
              {[0, 1].map(i => (
                <span key={i} className="flex shrink-0 items-center">
                  <span className="px-6">{QUOTE}</span>
                  <span className="text-primary-foreground/40 px-6">—</span>
                </span>
              ))}
            </div>
          </div>
        </div>
      </div>
      <div className="relative flex w-full flex-col items-center justify-center p-6 md:p-10 lg:w-[40%]">
        <video
          className="absolute inset-0 h-full w-full object-cover lg:hidden"
          src="/brand/videos/looping-av1.mp4"
          autoPlay
          loop
          muted
          playsInline
        />
        <div className="bg-primary/55 absolute inset-0 lg:hidden" />
        <Image
          src="/brand/logo/ASAHI_HORIZONTAL_VECTOR.svg"
          alt="Asahi Livwell"
          width={124}
          height={51}
          className="absolute top-6 left-6 h-9 w-auto lg:hidden"
        />
        <div className="bg-background/95 relative w-full max-w-sm rounded-xl p-6 shadow-lg backdrop-blur-sm">
          {children}
        </div>
      </div>
    </div>
  );
}
