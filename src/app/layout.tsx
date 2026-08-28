import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import { ThemeProvider } from "@/components/theme-provider";
import { Toaster } from "@/components/ui/sonner";

import "./globals.css";
import {
  THEME_MODE_VALUES,
  THEME_PRESET_VALUES,
  ThemeMode,
  ThemePreset,
} from "@/types/preferences/theme";
import { getPreference } from "@/server/server-actions";
import { PreferencesStoreProvider } from "@/stores/preferences/preferences-provider";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: "Compasation",
  description: "Compasation Asahi Livwel",
};

export default async function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  const themeMode = await getPreference<ThemeMode>(
    "theme_mode",
    THEME_MODE_VALUES,
    "light"
  );
  const themePreset = await getPreference<ThemePreset>(
    "theme_preset",
    THEME_PRESET_VALUES,
    "default"
  );

  return (
    <html
      lang="en"
      className={themeMode === "dark" ? "dark" : ""}
      data-theme-preset={themePreset}
      suppressHydrationWarning
    >
      <body
        className={`${geistSans.variable} ${geistMono.variable} min-h-screen  antialiased`}
      >
        <PreferencesStoreProvider
          themeMode={themeMode}
          themePreset={themePreset}
        >
          {children}
          <Toaster />
        </PreferencesStoreProvider>
      </body>
    </html>
  );
}
