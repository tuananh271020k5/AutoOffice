import React, {
  createContext, useCallback, useEffect, useMemo, useState, type ReactNode,
} from 'react';
import { translationService } from './service.ts';
import { detectLocale } from './detect.ts';
import { loadStoredLocale, saveStoredLocale } from './storage.ts';
import { LOCALES, isLocaleId, type LocaleId } from './registry.ts';
import type { TranslationKey } from './keys.generated.ts';
import type { TranslationParams } from './types.ts';
import { makeFormatters, type Formatters } from './format.ts';

export interface LanguageContextValue {
  locale: LocaleId;
  direction: 'ltr' | 'rtl';
  t: (key: TranslationKey, params?: TranslationParams) => string;
  setLocale: (id: LocaleId) => Promise<void>;
  formatters: Formatters;
}

export const LanguageContext = createContext<LanguageContextValue | null>(null);

export interface LanguageProviderProps {
  children: ReactNode;
  /** When provided, skips detection and storage on mount. */
  initialLocale?: LocaleId;
}

function resolveDefaultStartupLocale(): LocaleId {
  const saved = loadStoredLocale();
  if (saved && isLocaleId(saved)) return saved;

  const envLocale = (import.meta.env?.VITE_DEFAULT_LOCALE as string | undefined)?.trim();
  if (envLocale && isLocaleId(envLocale)) return envLocale;

  // Prefer Vietnamese ('vi') by default when no saved preference exists,
  // while still supporting detectLocale for explicit detection calls.
  if (isLocaleId('vi')) return 'vi';
  return detectLocale({ saved });
}

export function LanguageProvider({ children, initialLocale }: LanguageProviderProps) {
  const seed = initialLocale ?? resolveDefaultStartupLocale();
  const [locale, setLocaleState] = useState<LocaleId>(seed);
  const [rev, force] = useState(0);

  // Initial load. setLocale is idempotent + cached so re-runs are cheap.
  useEffect(() => {
    let cancelled = false;
    void (async () => {
      await translationService.setLocale(seed);
      if (!cancelled) force(n => n + 1);
    })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Apply <html lang>/<dir> whenever locale changes.
  useEffect(() => {
    const meta = LOCALES[locale];
    document.documentElement.setAttribute('lang', locale);
    document.documentElement.setAttribute('dir', meta.direction);
  }, [locale]);

  const setLocale = useCallback(async (id: LocaleId) => {
    await translationService.setLocale(id);
    saveStoredLocale(id);
    setLocaleState(id);
    announce(id);
  }, []);

  const value = useMemo<LanguageContextValue>(() => ({
    locale,
    direction: LOCALES[locale].direction,
    t: (k, p) => translationService.t(k, p),
    setLocale,
    formatters: makeFormatters(locale),
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }), [locale, setLocale, rev]);

  return <LanguageContext.Provider value={value}>{children}</LanguageContext.Provider>;
}

function announce(locale: LocaleId): void {
  const node = document.createElement('div');
  node.setAttribute('role', 'status');
  node.setAttribute('aria-live', 'polite');
  node.style.cssText =
    'position:absolute;left:-10000px;width:1px;height:1px;overflow:hidden;';
  node.textContent = translationService.t('accessibility.languageChanged', {
    language: LOCALES[locale].nativeName,
  });
  document.body.appendChild(node);
  setTimeout(() => node.remove(), 1000);
}
