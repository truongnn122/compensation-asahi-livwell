import { ThemeMode, ThemePreset } from "@/types/preferences/theme";
import { createStore } from "zustand/vanilla";

export type PreferencesState = {
  themeMode: ThemeMode;
  themePreset: ThemePreset;
  setThemeMode: (mode: ThemeMode) => void;
  setThemePreset: (preset: ThemePreset) => void;
};

export const createPreferencesStore = (init?: Partial<PreferencesState>) =>
  createStore<PreferencesState>()(set => ({
    themeMode: init?.themeMode ?? "light",
    themePreset: init?.themePreset ?? "asahi-livwell",
    setThemeMode: mode => set({ themeMode: mode }),
    setThemePreset: preset => set({ themePreset: preset }),
  }));
