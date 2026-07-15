import type { Language } from './LanguageContext';

export interface LocalizedText {
  en: string;
  ja: string;
}

export function localize(language: Language, text: LocalizedText): string {
  return text[language];
}

export function getPromptLanguageInstruction(language: Language): string {
  return language === 'ja'
    ? 'OUTPUT LANGUAGE: Write all user-facing titles, descriptions, labels, connection text, workflow text, findings, recommendations, and narrative content in Japanese. Keep official Microsoft/Azure/Fabric product names, API names, code, commands, identifiers, and resource names in English.'
    : 'OUTPUT LANGUAGE: Write all user-facing titles, descriptions, labels, connection text, workflow text, findings, recommendations, and narrative content in English. Keep official Microsoft/Azure/Fabric product names, API names, code, commands, identifiers, and resource names unchanged.';
}
