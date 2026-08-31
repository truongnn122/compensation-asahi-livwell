import { getPreference } from "@/server/server-actions";
import { LANGUAGE_VALUES, type Language } from "@/types/preferences/language";

import { dictionaries, type Dictionary } from "./dictionaries";

export async function getDictionary(): Promise<Dictionary> {
  const language = await getPreference<Language>(
    "language",
    LANGUAGE_VALUES,
    "vi"
  );
  return dictionaries[language];
}
